"""Label company contacts without treating a prospect's locality as an office.

This module makes no network requests. Identity and geography must already have
been checked by the Maps collector, which supplies literal boolean validation
flags. A nearby employer listing does not prove where an individual works.
"""
from __future__ import annotations

import re
import unicodedata
from collections.abc import Mapping

from company_public_data import public_url


UNKNOWN = {"", "-", "none", "null", "nan", "n/a", "unknown", "not available"}
UNCLASSIFIED = "Company location (unclassified)"


def _text(value) -> str:
    if not isinstance(value, (str, int, float)) or isinstance(value, bool):
        return ""
    value = str(value).strip()
    return "" if value.casefold() in UNKNOWN else value


def _normal(value) -> str:
    value = unicodedata.normalize("NFKD", _text(value)).casefold()
    value = "".join(char for char in value if not unicodedata.combining(char))
    return re.sub(r"[^\w]+", " ", value).strip()


def maps_target_location(record) -> dict:
    """Keep a saved company address and the prospect's locality separate.

    ``location`` is suitable for the existing Maps query builder. Its ``basis``
    explains why it was selected; this is never a claim about office type.
    """
    record = record if isinstance(record, Mapping) else {}
    company_address = _text(record.get("Address"))
    prospect_location = ", ".join(
        part for key in ("City", "State", "Country")
        if (part := _text(record.get(key)))
    )
    if company_address:
        basis = "Saved company address"
        location = company_address
    elif prospect_location:
        basis = "Prospect locality (search guidance only)"
        location = prospect_location
    else:
        basis = "No geographic evidence"
        location = ""
    return {"location": location, "basis": basis,
            "company_address": company_address,
            "prospect_location": prospect_location}


def _scope_label(value) -> str:
    """Recognize a short, explicit office label, never prose or a street name."""
    value = _normal(value)
    if re.fullmatch(r"(?:(?:global|world|worldwide|corporate|international) )?"
                    r"(?:headquarters|head quarters|head office|hq)", value):
        return "Headquarters"
    if re.fullmatch(r"(?:(?:regional|local|branch|sales|district|satellite) )"
                    r"(?:office|headquarters|head office|hq)|"
                    r"(?:branch|branch office|local office|regional office)", value):
        return "Local office"
    # A registered address can be an accountant's office or a legal mailbox.
    if value in {"registered office", "registered address", "registered agent"}:
        return "Registered office"
    return ""


def _listing_scope(expected_company: str, listing_name: str) -> str:
    """Only office descriptors around the matched company name are evidence."""
    expected, actual = _normal(expected_company), _normal(listing_name)
    if not actual:
        return ""
    if expected and actual == expected:
        return ""
    if expected:
        descriptor = re.sub(r"(?<!\w)" + re.escape(expected) + r"(?!\w)",
                            " ", actual, count=1).strip()
        descriptor = re.sub(r"\s+", " ", descriptor)
        if descriptor != actual:
            return _scope_label(descriptor)
    return ""


def _address_scopes(address: str) -> set[str]:
    """A standalone label or labelled address is evidence; 'HQ Road' is not."""
    scopes = set()
    for part in re.split(r"[,\n]+", address):
        label = part.split(":", 1)[0].strip()
        if scope := _scope_label(label):
            scopes.add(scope)
    return scopes


def classify_maps_location(record, details) -> dict:
    """Classify one Maps place before its address or phone is attached.

    Required collector facts: ``company_matched`` and ``location_matched`` must
    be the literal ``True``. ``company_address_matched`` allows a saved company
    address's explicit office label to contribute. ``scope_label`` is optional
    and must come from the place itself, not a query, review or search snippet.

    Successful validation identifies a *company location*. Even a matching
    prospect city never verifies that person's employment at that office.
    """
    record = record if isinstance(record, Mapping) else {}
    details = details if isinstance(details, Mapping) else {}
    target = maps_target_location(record)
    address, phone = _text(details.get("address")), _text(details.get("phone"))
    source = next((url for key in ("address_link", "source_url")
                   if (url := _text(details.get(key))) and public_url(url)), "")
    company = _text(record.get("Company Name"))
    scope_evidence = []
    listing_scope = _listing_scope(company, _text(details.get("company_name")))
    if listing_scope:
        scope_evidence.append((listing_scope, "Explicit office descriptor in Maps place name"))
    explicit_scope = _scope_label(details.get("scope_label"))
    if explicit_scope:
        scope_evidence.append((explicit_scope, "Explicit office label on Maps place"))
    if details.get("company_address_matched") is True:
        for existing_scope in sorted(_address_scopes(target["company_address"])):
            scope_evidence.append((existing_scope, "Explicit label on matching saved company address"))

    scopes = {scope for scope, _ in scope_evidence}
    conflict = len(scopes) > 1
    scope = "Conflicting office scope" if conflict else next(iter(scopes), UNCLASSIFIED)
    basis = "; ".join(dict.fromkeys(reason for _, reason in scope_evidence))
    if not basis:
        basis = "No explicit headquarters or branch label in available place evidence"
    if details.get("company_address_matched") is True and target["company_address"]:
        location_basis = "Matches saved company address; prospect locality was not used as office identity"
    elif target["company_address"]:
        location_basis = "Saved company address exists; same-office match has not been established"
    elif target["prospect_location"]:
        location_basis = "Prospect locality narrowed search; prospect employment at this office is unverified"
    else:
        location_basis = "Company place evidence only; no prospect office relationship established"

    reason = "Validated company place; office type is labelled separately"
    mergeable = True
    if (details.get("rejected") is True or details.get("is_rejected") is True
            or _text(details.get("rejection_reason"))):
        mergeable, reason = False, "The Maps collector rejected this candidate"
    elif not company or details.get("company_matched") is not True:
        mergeable, reason = False, "Company identity has not been validated"
    elif details.get("location_matched") is not True:
        mergeable, reason = False, "Candidate geography has not been validated"
    elif target["company_address"] and details.get("company_address_matched") is not True:
        mergeable, reason = False, "Candidate has not been matched to the saved company address"
    elif not address:
        mergeable, reason = False, "A phone without an address cannot be assigned to an office"
    elif not source:
        mergeable, reason = False, "A public place source is required for office contact evidence"
    elif conflict:
        mergeable, reason = False, "Conflicting headquarters and office labels need review"
    elif scope == "Registered office":
        mergeable, reason = False, "Registered address alone does not establish an operating office contact"

    return {
        "scope": scope, "basis": basis, "location_basis": location_basis,
        "source_url": source, "address": address, "phone": phone,
        "prospect_location": target["prospect_location"],
        "is_mergeable": mergeable, "reason": reason,
        "prospect_office_verified": False,
    }
