"""Conservative visible search-result extraction, shared by browser and fixtures.

The returned cards contain observations only. Company identity, monetary values,
classification codes, and evidence priority are checked by the Python parser.
"""

SEARCH_RESULT_SCRIPT = r"""provider => {
    const excludedSelector = [
        '#tads', '#bottomads', '[data-text-ad]', '[data-ad]',
        '[aria-label="Ads"]', '[aria-label="Sponsored"]',
        '.related-question-pair', '[data-initq]',
        '[aria-label="People also ask"]', '[data-attrid="People also ask"]',
        '[data-ai-overview]', '[aria-label="AI Overview"]',
        '[aria-label="AI overview"]', '[data-attrid="SGE"]'
    ].join(',');
    const normalize = value => (value || '').replace(/\s+/g, ' ').trim();
    const visible = element => {
        if (!element || !element.isConnected || !element.getClientRects().length) return false;
        for (let item = element; item && item !== document.documentElement; item = item.parentElement) {
            const style = getComputedStyle(item);
            if (item.hidden || item.getAttribute('aria-hidden') === 'true' ||
                style.display === 'none' || style.visibility === 'hidden' ||
                style.visibility === 'collapse' || style.opacity === '0') return false;
        }
        return true;
    };
    const boundary = element => !element || /^(BODY|HTML|MAIN)$/.test(element.tagName) ||
        ['search', 'rso', 'b_results'].includes(element.id);
    const excluded = element => {
        if (element.closest(excludedSelector)) return true;
        for (let item = element; item && !boundary(item); item = item.parentElement) {
            const headings = item.querySelectorAll(':scope > h1, :scope > h2, :scope > [role="heading"]');
            if (Array.from(headings).some(h => visible(h) &&
                /^(?:AI Overview|People also ask|Sponsored|Ads)$/i.test(normalize(h.innerText)))) return true;
        }
        return false;
    };
    const usable = element => visible(element) && !excluded(element);
    const text = element => (element?.innerText || '').trim().slice(0, 6000);
    const href = anchor => {
        if (!anchor || !usable(anchor)) return '';
        try {
            const url = new URL(anchor.href);
            if (!['http:', 'https:'].includes(url.protocol)) return '';
            if (/(^|\.)google\.[a-z.]+$/.test(url.hostname) &&
                ['/search', '/preferences', '/advanced_search'].includes(url.pathname)) return '';
            return url.href;
        } catch (_) { return ''; }
    };
    const rows = [];
    const add = row => {
        if (row.title && row.text && !rows.some(old => old.title === row.title &&
            old.text === row.text && old.url === row.url && old.kind === row.kind)) rows.push(row);
    };
    if (provider === 'bing') {
        for (const card of document.querySelectorAll('li.b_algo')) {
            if (!usable(card)) continue;
            const heading = Array.from(card.querySelectorAll('h2')).find(usable);
            const anchor = heading?.querySelector('a[href]');
            const url = href(anchor);
            if (heading && url) add({title: normalize(heading.innerText), text: text(card), url, kind: 'organic'});
            if (rows.length >= 7) break;
        }
        return rows;
    }
    if (provider !== 'google') return [];

    // Walk to the outermost single-result card, rather than assuming a fixed
    // number of wrapper elements. Stop before another result or another widget.
    for (const heading of document.querySelectorAll('#search h3, #rso h3')) {
        if (!usable(heading)) continue;
        const anchor = heading.closest('a[href]');
        const url = href(anchor);
        if (!url) continue;
        let card = anchor;
        for (let depth = 0; depth < 12; depth++) {
            const parent = card.parentElement;
            if (boundary(parent) || !usable(parent)) break;
            const resultHeadings = Array.from(parent.querySelectorAll('h3')).filter(usable);
            if (resultHeadings.length !== 1 || resultHeadings[0] !== heading) break;
            const otherHeadings = Array.from(parent.querySelectorAll('h1,h2,[role="heading"]'))
                .filter(item => usable(item) && item !== heading && !item.contains(heading));
            if (otherHeadings.length) break;
            // A result card must not absorb a neighboring PAA, advertisement,
            // AI overview, or knowledge panel just because it has no h3.
            if (parent.querySelector(excludedSelector) ||
                parent.querySelector('[data-attrid="title"], .kno-ecr-pt')) break;
            card = parent;
        }
        const featured = card.matches('[data-featured-snippet], [data-attrid="wa:/description"]') ||
            card.querySelector('[data-featured-snippet], [data-attrid="wa:/description"]') ||
            heading.closest('[data-featured-snippet], [data-attrid="wa:/description"]');
        add({title: normalize(heading.innerText), text: text(card), url,
             kind: featured ? 'featured' : 'organic'});
        if (rows.filter(row => row.kind !== 'knowledge').length >= 7) break;
    }

    // Knowledge properties require an actual entity heading in the same
    // bounded panel. Never derive identity from the query or page title.
    const titleSelector = '[data-attrid="title"], .kno-ecr-pt';
    for (const property of document.querySelectorAll('[data-attrid]')) {
        const marker = property.getAttribute('data-attrid') || '';
        if (!/(?:^|[:/])(?:revenue|sic(?:[ _]code)?|naics(?:[ _]code)?)$/i.test(marker) || !usable(property)) continue;
        let panel = property.parentElement;
        let entity = null;
        for (let depth = 0; depth < 8 && !boundary(panel); depth++, panel = panel.parentElement) {
            if (!usable(panel)) break;
            const titles = Array.from(panel.querySelectorAll(titleSelector)).filter(usable);
            if (titles.length > 1) break;
            if (titles.length === 1) { entity = titles[0]; break; }
        }
        if (!entity || !panel || panel.querySelector(excludedSelector)) continue;
        const title = normalize(entity.innerText);
        if (!title || title.length > 180 || /^(?:Search|AI Overview|People also ask)$/i.test(title)) continue;
        const url = Array.from(property.querySelectorAll('a[href]')).map(href).find(Boolean) || '';
        add({title, text: title + '\n' + text(property), url, kind: 'knowledge'});
    }
    return rows;
}"""
