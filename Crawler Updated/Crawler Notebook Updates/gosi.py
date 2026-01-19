#!/usr/bin/env python3
import asyncio
import os
import re
from pathlib import Path

from bs4 import BeautifulSoup
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

OUTPUT_FOLDER = "GOSI_DOCX"
MAIN_URL = "https://www.gosi.gov.sa/ar/SystemsAndRegulations"

def ensure_out_dir():
    Path(OUTPUT_FOLDER).mkdir(parents=True, exist_ok=True)

def make_rtl(paragraph):
    """Force paragraph RTL in python-docx."""
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "true")
    pPr.append(bidi)

def sanitize_filename(name: str) -> str:
    # Keep Arabic, letters, numbers, spaces, underscore, dash
    # Strip anything else, truncate to safe length
    safe = re.sub(r"[^0-9A-Za-z\u0600-\u06FF _\-]+", "", name)
    return safe.strip()[:180] or "document"

async def scrape_and_save_docx_rtl():
    ensure_out_dir()

    async with async_playwright() as p:
        # Headless Chromium is fine for servers/CI. Set headless=False to watch it run.
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await page.goto(MAIN_URL, timeout=60_000)
        await page.wait_for_load_state("domcontentloaded")

        # The list of items to visit
        await page.wait_for_selector("#mediaCenterElements li", timeout=30_000)
        items_all = await page.query_selector_all("#mediaCenterElements li")

        titles = []
        for item in items_all:
            text = (await item.inner_text()).strip()
            if text and text != "كتيبات الأنظمة":
                titles.append(text)

        print(f"✅ Total items to visit: {len(titles)}\n")

        for idx, title in enumerate(titles, start=1):
            print(f"➡ Visiting ({idx}/{len(titles)}): {title}")

            # Reload the main page each loop to avoid stale handles
            await page.goto(MAIN_URL, timeout=60_000)
            await page.wait_for_selector("#mediaCenterElements li", timeout=30_000)

            # More reliable: use a locator that matches the LI with exact text
            # (Playwright's has_text is substring; we filter exact match)
            lis = page.locator("#mediaCenterElements li")
            count = await lis.count()
            target_index = None
            for i in range(count):
                t = (await lis.nth(i).inner_text()).strip()
                if t == title:
                    target_index = i
                    break

            if target_index is None:
                print(f"⚠ Could not find item: {title}")
                continue

            await lis.nth(target_index).click()

            try:
                await page.wait_for_selector("#systemsAndRegulationsPageContent", timeout=30_000)
                # Give the page a moment for dynamic content to finish rendering
                await page.wait_for_load_state("networkidle")
            except PlaywrightTimeoutError:
                print(f"⚠ Timeout waiting for content: {title}")
                continue

            content_div = await page.query_selector("#systemsAndRegulationsPageContent")
            if not content_div:
                print(f"⚠ No content container for: {title}")
                continue

            html_content = await content_div.inner_html()

            # Convert HTML to plain text (lxml parser is faster/better if installed)
            soup = BeautifulSoup(html_content, "lxml")
            plain_text = soup.get_text(separator="\n", strip=True)

            # Write DOCX (set Normal style a bit larger; force RTL on paragraph)
            doc = Document()
            try:
                style = doc.styles["Normal"]
                font = style.font
                font.size = Pt(12)
            except Exception:
                pass

            para = doc.add_paragraph(plain_text)
            make_rtl(para)

            safe_name = sanitize_filename(title)
            output_path = os.path.join(OUTPUT_FOLDER, f"{safe_name}.docx")
            doc.save(output_path)
            print(f"✅ Saved DOCX (RTL): {output_path}\n")

        await browser.close()
        print(f"\n✅ All pages processed. Files are in '{OUTPUT_FOLDER}/'.\n")

def main():
    asyncio.run(scrape_and_save_docx_rtl())

if __name__ == "__main__":
    main()
