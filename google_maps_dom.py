"""Read and scroll the visible Maps place panel, excluding the results feed."""

PLACE_SCRIPT = r"""({scroll = 0} = {}) => {
    const clean = value => (value || '').replace(/[\uE000-\uF8FF]/g, '').replace(/\s+/g, ' ').trim();
    const visible = element => {
        if (!element || !element.getClientRects().length) return false;
        for (let node = element; node; node = node.parentElement) {
            const style = getComputedStyle(node);
            if (node.hidden || node.getAttribute('aria-hidden') === 'true' ||
                style.display === 'none' || style.visibility === 'hidden') return false;
        }
        return true;
    };
    const eligible = node => visible(node) && !node.closest('[role="feed"]');
    const titles = [...document.querySelectorAll('h1.DUwDvf, [role="main"] h1, main h1, h1')];
    const title = titles.find(node => eligible(node) &&
        !/^(search results|results)$/i.test(clean(node.innerText)));
    if (!title) return {title:'', address:'', phone:'', website:'', url:location.href,
                        addressFound:false, phoneFound:false, hasPanel:false};
    const root = title.closest('[role="main"], main') || document.body;
    const first = selectors => {
        for (const selector of selectors) {
            const match = [...root.querySelectorAll(selector)].find(eligible);
            if (match) return match;
        }
        return null;
    };
    const read = node => node ? clean(node.getAttribute('aria-label') || node.innerText || node.textContent) : '';
    const addressNode = first(['[data-item-id="address"]', '[data-tooltip="Copy address"]',
                               '[aria-label^="Address:"]']);
    const phoneNode = first(['[data-item-id^="phone:tel:"]', 'a[href^="tel:"]',
                             '[data-item-id^="phone"]', '[aria-label^="Phone:"]',
                             '[data-tooltip="Copy phone number"]']);
    const websiteNode = first(['a[data-item-id="authority"]', 'a[aria-label^="Website:"]']);
    let address = read(addressNode).replace(/^Address\s*:\s*/i, '');
    if (/^copy address$/i.test(address)) address = clean(addressNode.innerText);
    let phone = '';
    if (phoneNode) {
        const item = phoneNode.getAttribute('data-item-id') || '';
        const href = phoneNode.getAttribute('href') || '';
        phone = item.startsWith('phone:tel:') ? item.slice(10) :
                href.startsWith('tel:') ? href.slice(4) : read(phoneNode);
        if (!/\d/.test(phone)) phone = clean(phoneNode.innerText);
        phone = phone.replace(/^Phone\s*:\s*/i, '');
    }
    // Pick the details scroller, never the longer neighbouring results feed.
    const panels = [root, ...root.querySelectorAll('div, main, section')].filter(node => {
        if (!eligible(node) || node.matches('[role="feed"]') || node.querySelector('[role="feed"]')) return false;
        return /(auto|scroll)/.test(getComputedStyle(node).overflowY) &&
               node.scrollHeight > node.clientHeight + 5 && node.clientHeight > 100;
    });
    const detailNode = addressNode || phoneNode || title;
    panels.sort((a,b) => Number(!a.contains(detailNode)) - Number(!b.contains(detailNode)) ||
                        a.clientHeight - b.clientHeight);
    const panel = panels[0];
    const before = panel ? panel.scrollTop : window.scrollY;
    if (scroll) {
        if (panel) {
            panel.scrollTop = Math.min(panel.scrollTop + scroll, panel.scrollHeight - panel.clientHeight);
            panel.dispatchEvent(new Event('scroll', {bubbles:true}));
        } else if (root === document.body && !document.querySelector('[role="feed"]')) {
            window.scrollBy(0, scroll);
        }
    }
    const after = panel ? panel.scrollTop : window.scrollY;
    return {title:clean(title.innerText), address, phone,
            website:websiteNode ? websiteNode.href : '', url:location.href,
            addressFound:!!address && !/^copy address$/i.test(address), phoneFound:/\d/.test(phone),
            websiteFound:!!websiteNode, hasPanel:!!panel,
            scrollTop:Math.round(after), clientHeight:panel ? panel.clientHeight : innerHeight,
            scrollHeight:panel ? panel.scrollHeight : document.documentElement.scrollHeight,
            usedPanel:!!panel, before:Math.round(before), after:Math.round(after)};
}"""
