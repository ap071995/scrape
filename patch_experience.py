from pathlib import Path
p = Path('Scraper.py')
s = p.read_text(encoding='utf-8')
a = s.index('    def _capture_standard_profile(')
b = s.index('    def _parse_standard_current_experience(', a)
s = s[:a] + '''    def _read_standard_experience_dom(self) -> dict[str, Any]:
        return self.page.evaluate(r"""() => {
            const clean = v => (v || '').replace(/\\u00a0/g, ' ').replace(/[ \\t]+/g, ' ').trim();
            const main = document.querySelector('main') || document.body;
            const date = /\\b\\d{4}\\s*[-–—]\\s*Present\\b/i;
            const heading = Array.from(main.querySelectorAll('h2,h3,[role="heading"],span'))
                .find(el => /^Experience$/i.test(clean(el.innerText)));
            const marker = main.querySelector('#experience');
            let section = heading?.closest('section') || marker?.closest('section');
            if (!section) {
                let node = heading?.parentElement || marker?.parentElement;
                while (node && node !== main) {
                    if (date.test(node.innerText || '') && node.querySelector('a[href*="/company/"]')) {
                        section = node; break;
                    }
                    node = node.parentElement;
                }
            }
            if (!section && /\\/details\\/experience/.test(location.pathname)) section = main;
            if (!section) return {experienceText:'',currentText:'',companyLinks:[],sectionFound:false};
            const candidates = Array.from(section.querySelectorAll('li, [role="listitem"], article, div'))
                .filter(el => date.test(el.innerText || '') &&
                    clean(el.innerText).split('\\n').filter(Boolean).length >= 3);
            // Prefer a focused complete card with an employer link. For grouped
            // employers, prefer the leaf role and recover its enclosing company.
            const leafRoles = candidates.filter(el => !candidates.some(other =>
                other !== el && el.contains(other) && date.test(other.innerText || '') &&
                clean(other.innerText).split('\\n').filter(Boolean).length >= 3));
            leafRoles.sort((a,b) => clean(a.innerText).length - clean(b.innerText).length);
            const role = leafRoles[0];
            let node = role, companyLinks = [];
            while (node && node !== section) {
                companyLinks = Array.from(node.querySelectorAll('a[href*="/company/"]')).map(a => ({
                    text: clean(a.innerText || a.getAttribute('aria-label') || a.querySelector('img')?.alt),
                    href:a.href
                }));
                if (companyLinks.length) break;
                node = node.parentElement;
            }
            return {experienceText:clean(section.innerText), currentText:clean(role?.innerText),
                companyLinks, sectionFound:true};
        }""") or {}

    def _capture_standard_profile(self, profile_url: str) -> dict[str, Any]:
        profile_url = canonical_profile_url(profile_url)
        self.page.goto(profile_url, wait_until='domcontentloaded', timeout=PAGE_TIMEOUT_MS)
        self._wait_for_manual_login_if_needed()
        self.page.wait_for_timeout(LINK_PROFILE_SETTLE_MS)
        profile_name = self._standard_profile_name()
        data = {}
        for attempt in range(2):
            # The dedicated Experience page is a normal LinkedIn UI route and
            # avoids dependence on the top-profile page's lazy section wrapper.
            if attempt:
                self.page.goto(profile_url + '/details/experience/',
                               wait_until='domcontentloaded', timeout=PAGE_TIMEOUT_MS)
                self._wait_for_manual_login_if_needed()
                self.page.wait_for_timeout(LINK_PROFILE_SETTLE_MS)
            for _ in range(LINK_PROFILE_MAX_SCROLLS):
                if self.stop_event.is_set() or not self._wait_if_paused():
                    raise RuntimeError('Profile capture stopped.')
                data = self._read_standard_experience_dom()
                if data.get('currentText') and data.get('companyLinks'):
                    data['name'] = profile_name
                    return data
                self.page.evaluate(r"""() => {
                    const marker = document.querySelector('#experience');
                    if (marker && marker.getBoundingClientRect().top > innerHeight) marker.scrollIntoView();
                    else window.scrollBy(0, 600);
                    for (const el of document.querySelectorAll('main, [role="main"]')) {
                        if (el.scrollHeight > el.clientHeight && /auto|scroll/.test(getComputedStyle(el).overflowY)) el.scrollTop += 600;
                    }
                }""")
                self.page.wait_for_timeout(500)
        self.trace('LINK_PROFILE', 'experience_unresolved', url=self.page.url,
                   name=profile_name, data=data)
        raise RuntimeError('Current Experience could not be extracted. Profile was not saved as Dead. '
                           'Check the Experience page in the scraper browser; see experience_unresolved in the activity log.')

''' + s[b:]
s = s.replace('APP_VERSION = "42.2.1"', 'APP_VERSION = "42.3.0"')
p.write_text(s, encoding='utf-8')
