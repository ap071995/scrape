"""Capture a bounded public search response for scraper development."""
import json
from pathlib import Path
from urllib.parse import quote
from playwright.sync_api import sync_playwright

folder = Path('Output/search_result_verification')
folder.mkdir(parents=True, exist_ok=True)
with sync_playwright() as playwright:
    browser = playwright.chromium.launch(channel='chrome', headless=True)
    page = browser.new_page()
    try:
        page.goto('https://www.google.com/search?q=' + quote('UBS revenue'), wait_until='domcontentloaded', timeout=20000)
        page.wait_for_timeout(1500)
        data = {'url':page.url, 'title':page.title(), 'text':page.locator('body').inner_text(timeout=4000)}
        (folder / 'google_initial.json').write_text(json.dumps(data,indent=2,ensure_ascii=False),encoding='utf-8')
        (folder / 'google_initial.html').write_text(page.content(),encoding='utf-8')
        print(json.dumps(data,ensure_ascii=True)[:16000])
    finally:
        browser.close()
