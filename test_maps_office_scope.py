"""Adversarial office attribution checks; no browser or network requests."""
import copy
import unittest

from maps_office_scope import UNCLASSIFIED, classify_maps_location, maps_target_location


class MapsOfficeScopeTests(unittest.TestCase):
    def record(self, **extra):
        return {"Company Name": "Example Industries", "City": "Boston",
                "State": "Massachusetts", "Country": "United States", **extra}

    def details(self, **extra):
        return {"company_name": "Example Industries",
                "address": "123 Main St, Boston, MA 02110, USA",
                "phone": "+1 617 555 1234",
                "address_link": "https://www.google.com/maps/place/example",
                "company_matched": True, "location_matched": True, **extra}

    def classify(self, record=None, **details):
        return classify_maps_location(self.record() if record is None else record,
                                      self.details(**details))

    def test_contract_has_evidence_and_attribution_fields(self):
        result = self.classify()
        self.assertEqual(set(result), {
            "scope", "basis", "location_basis", "source_url", "address", "phone",
            "prospect_location", "is_mergeable", "reason", "prospect_office_verified"})
        self.assertEqual(result["prospect_location"], "Boston, Massachusetts, United States")
        self.assertTrue(result["is_mergeable"])
        self.assertFalse(result["prospect_office_verified"])

    def test_prospect_locality_is_only_query_guidance(self):
        target = maps_target_location(self.record())
        self.assertEqual(target["location"], "Boston, Massachusetts, United States")
        self.assertEqual(target["basis"], "Prospect locality (search guidance only)")
        self.assertEqual(target["company_address"], "")
        result = self.classify()
        self.assertEqual(result["scope"], UNCLASSIFIED)
        self.assertIn("unverified", result["location_basis"])

    def test_saved_company_address_precedes_prospect_locality(self):
        target = maps_target_location(self.record(Address="10 Broadgate, London, UK"))
        self.assertEqual(target["location"], "10 Broadgate, London, UK")
        self.assertEqual(target["basis"], "Saved company address")
        self.assertEqual(target["prospect_location"], "Boston, Massachusetts, United States")

    def test_unknown_markers_do_not_become_geography(self):
        for missing in (None, "", "-", "N/A", "unknown", "nan", "not available"):
            with self.subTest(missing=missing):
                target = maps_target_location({"Address": missing, "City": missing,
                                               "State": missing, "Country": missing})
                self.assertEqual(target["location"], "")
                self.assertEqual(target["basis"], "No geographic evidence")

    def test_explicit_headquarters_name_labels(self):
        for label in ("Headquarters", "Global Headquarters", "HQ", "Head Office",
                      "World Headquarters", "Corporate Head Quarters"):
            with self.subTest(label=label):
                result = self.classify(company_name="Example Industries - " + label)
                self.assertEqual(result["scope"], "Headquarters")
                self.assertTrue(result["is_mergeable"])
                self.assertFalse(result["prospect_office_verified"])

    def test_explicit_headquarters_prefix(self):
        self.assertEqual(self.classify(company_name="Headquarters: Example Industries")
                         ["scope"], "Headquarters")

    def test_explicit_local_office_name_labels(self):
        for label in ("Branch", "Branch Office", "Local Office", "Regional Office",
                      "Sales Office", "Regional Headquarters", "District HQ"):
            with self.subTest(label=label):
                result = self.classify(company_name="Example Industries - " + label)
                self.assertEqual(result["scope"], "Local office")
                self.assertTrue(result["is_mergeable"])
                self.assertFalse(result["prospect_office_verified"])

    def test_corporate_office_does_not_imply_headquarters(self):
        self.assertEqual(self.classify(company_name="Example Industries Corporate Office")
                         ["scope"], UNCLASSIFIED)

    def test_office_word_in_company_name_is_not_a_label(self):
        for name in ("Headquarters", "HQ Industries", "Branch Office Supplies"):
            with self.subTest(name=name):
                result = self.classify(self.record(**{"Company Name": name}), company_name=name)
                self.assertEqual(result["scope"], UNCLASSIFIED)

    def test_company_name_must_be_a_whole_token_sequence(self):
        result = self.classify(self.record(**{"Company Name": "Acme"}),
                               company_name="Acmeology Headquarters")
        self.assertEqual(result["scope"], UNCLASSIFIED)

    def test_unvalidated_alias_does_not_provide_scope(self):
        self.assertEqual(self.classify(company_name="EI Headquarters")["scope"], UNCLASSIFIED)

    def test_case_accents_and_separator_variants(self):
        result = self.classify(self.record(**{"Company Name": "Exemple Société"}),
                               company_name="EXEMPLE SOCIETE — HEADQUARTERS")
        self.assertEqual(result["scope"], "Headquarters")

    def test_free_text_and_street_names_do_not_infer_scope(self):
        for suffix in ("Headquarters Road", "near Headquarters", "former Headquarters",
                       "not Headquarters", "Headquarters closed", "Office Park"):
            with self.subTest(suffix=suffix):
                self.assertEqual(self.classify(company_name="Example Industries " + suffix)
                                 ["scope"], UNCLASSIFIED)

    def test_reviews_queries_and_prospect_job_title_are_not_scope_evidence(self):
        result = self.classify(self.record(**{"Job Title": "Headquarters Director"}),
                               query="Example Industries Headquarters", reviews="Global HQ",
                               text="Our headquarters is in Boston", category="Corporate HQ")
        self.assertEqual(result["scope"], UNCLASSIFIED)

    def test_short_place_scope_label_is_used(self):
        result = self.classify(scope_label="Regional office")
        self.assertEqual(result["scope"], "Local office")
        self.assertIn("Maps place", result["basis"])

    def test_place_scope_prose_is_not_a_short_label(self):
        for label in ("formerly headquarters", "headquarters in London", "not a local office"):
            with self.subTest(label=label):
                self.assertEqual(self.classify(scope_label=label)["scope"], UNCLASSIFIED)

    def test_matching_saved_address_explicit_label_is_evidence(self):
        record = self.record(Address="Headquarters: 123 Main St, Boston, MA 02110, USA")
        result = self.classify(record, company_address_matched=True)
        self.assertEqual(result["scope"], "Headquarters")
        self.assertTrue(result["is_mergeable"])
        self.assertIn("saved company address", result["location_basis"])

    def test_unmatched_saved_address_does_not_supply_headquarters_label(self):
        record = self.record(Address="Headquarters: 10 Broadgate, London, UK")
        result = self.classify(record, company_address_matched=False)
        self.assertEqual(result["scope"], UNCLASSIFIED)
        self.assertFalse(result["is_mergeable"])
        self.assertIn("saved company address", result["reason"])

    def test_street_named_headquarters_does_not_supply_scope(self):
        record = self.record(Address="123 Headquarters Road, Boston, MA 02110, USA")
        result = self.classify(record, company_address_matched=True)
        self.assertEqual(result["scope"], UNCLASSIFIED)

    def test_conflicting_place_and_saved_labels_block_merge(self):
        record = self.record(Address="Headquarters, 123 Main St, Boston, MA 02110, USA")
        result = self.classify(record, company_name="Example Industries Local Office",
                               company_address_matched=True)
        self.assertEqual(result["scope"], "Conflicting office scope")
        self.assertFalse(result["is_mergeable"])

    def test_conflicting_labels_inside_saved_address_block_merge(self):
        record = self.record(Address="Headquarters, Local office, 123 Main St, Boston, USA")
        result = self.classify(record, company_address_matched=True)
        self.assertEqual(result["scope"], "Conflicting office scope")
        self.assertFalse(result["is_mergeable"])

    def test_conflicting_place_name_and_place_label_block_merge(self):
        result = self.classify(company_name="Example Industries Headquarters",
                               scope_label="Local office")
        self.assertEqual(result["scope"], "Conflicting office scope")
        self.assertFalse(result["is_mergeable"])

    def test_duplicate_consistent_labels_do_not_create_conflict(self):
        record = self.record(Address="Headquarters, 123 Main St, Boston, USA")
        result = self.classify(record, company_name="Example Industries Headquarters",
                               scope_label="HQ", company_address_matched=True)
        self.assertEqual(result["scope"], "Headquarters")
        self.assertTrue(result["is_mergeable"])

    def test_registered_office_does_not_establish_operating_contact(self):
        for label in ("Registered Office", "Registered address", "Registered agent"):
            with self.subTest(label=label):
                result = self.classify(scope_label=label)
                self.assertEqual(result["scope"], "Registered office")
                self.assertFalse(result["is_mergeable"])

    def test_identity_and_geography_require_literal_true(self):
        for field in ("company_matched", "location_matched"):
            for value in (False, None, 1, "True", "true", "yes", [], {}):
                with self.subTest(field=field, value=value):
                    self.assertFalse(self.classify(**{field: value})["is_mergeable"])

    def test_saved_address_match_requires_literal_true(self):
        record = self.record(Address="123 Main St, Boston, MA, USA")
        for value in (False, None, 1, "True"):
            with self.subTest(value=value):
                self.assertFalse(self.classify(record, company_address_matched=value)["is_mergeable"])

    def test_rejection_overrides_stale_positive_flags(self):
        for rejection in ({"rejected": True}, {"is_rejected": True},
                          {"rejection_reason": "Wrong company domain"}):
            with self.subTest(rejection=rejection):
                result = self.classify(**rejection)
                self.assertFalse(result["is_mergeable"])
                self.assertIn("rejected", result["reason"])

    def test_matching_name_and_score_do_not_replace_validation(self):
        result = self.classify(company_matched=False, name_score=100)
        self.assertFalse(result["is_mergeable"])

    def test_missing_company_never_merges(self):
        for value in (None, "-", "unknown"):
            with self.subTest(value=value):
                self.assertFalse(self.classify(self.record(**{"Company Name": value}))["is_mergeable"])

    def test_phone_without_address_never_merges(self):
        for value in (None, "-", "", "not available", {}):
            with self.subTest(value=value):
                result = self.classify(address=value)
                self.assertFalse(result["is_mergeable"])
                self.assertIn("without an address", result["reason"])

    def test_address_without_phone_can_merge(self):
        result = self.classify(phone=None)
        self.assertEqual(result["phone"], "")
        self.assertTrue(result["is_mergeable"])

    def test_missing_or_unsafe_source_blocks_merge(self):
        for value in (None, "-", "javascript:alert(1)", "file:///tmp/place.html",
                      "https://localhost/maps", "http://127.0.0.1/maps", "https://[bad/"):
            with self.subTest(value=value):
                result = self.classify(address_link=value)
                self.assertFalse(result["is_mergeable"])
                self.assertEqual(result["source_url"], "")

    def test_valid_fallback_source_after_invalid_address_link(self):
        result = self.classify(address_link="javascript:alert(1)",
                               source_url="https://www.google.com/maps/place/verified")
        self.assertTrue(result["is_mergeable"])
        self.assertEqual(result["source_url"], "https://www.google.com/maps/place/verified")

    def test_no_company_or_prospect_locality_does_not_invent_relationship(self):
        result = self.classify({"Company Name": "Example Industries"})
        self.assertEqual(result["scope"], UNCLASSIFIED)
        self.assertEqual(result["prospect_location"], "")
        self.assertIn("no prospect office relationship", result["location_basis"])
        self.assertFalse(result["prospect_office_verified"])

    def test_flags_cannot_claim_prospect_office_verified(self):
        result = self.classify(prospect_office_verified=True,
                               company_name="Example Industries Headquarters")
        self.assertFalse(result["prospect_office_verified"])

    def test_malformed_containers_do_not_raise(self):
        for record, details in ((None, None), ([], []), ("bad", 123), ({}, {})):
            with self.subTest(record=record, details=details):
                self.assertFalse(classify_maps_location(record, details)["is_mergeable"])

    def test_inputs_are_not_mutated(self):
        record, details = self.record(), self.details()
        originals = copy.deepcopy((record, details))
        classify_maps_location(record, details)
        self.assertEqual((record, details), originals)


if __name__ == "__main__":
    unittest.main()
