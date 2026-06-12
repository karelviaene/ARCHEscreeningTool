from io import BytesIO
import json
import csv
import asyncio
from playwright.async_api import async_playwright
import random
import copy
import pandas as pd
import io

# List of headers to cycle through to avoid detection when scraping
user_agents_list = [
    'Mozilla/5.0 (iPad; CPU OS 12_2 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.83 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36'
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; WOW64; rv:91.0) Gecko/20100101 Firefox/91.0',
    'Mozilla/5.0 (iPhone; CPU iPhone OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15A372 Safari/604.1',
    'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 11_2_3) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0.3 Safari/605.1.15',
    'Mozilla/5.0 (Linux; Android 10; SM-G973F) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.78 Mobile Safari/537.36',
    'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:98.0) Gecko/20100101 Firefox/98.0'
]


async def download_echachem_list(list_url):
    """Click the 'Download full list' button and capture the file into BytesIO."""

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=['--disable-blink-features=AutomationControlled']
        )
        context = await browser.new_context(
            user_agent=random.choice(user_agents_list),
            viewport={'width': 1280, 'height': 720},
            device_scale_factor=1
        )
        page = await context.new_page()

        await page.goto(list_url, wait_until="networkidle")

        # Handle cookie/consent banner if present
        try:
            consent_button = page.locator('button:has-text("I accept the terms")')
            if await consent_button.is_visible():
                await consent_button.click()
                await page.wait_for_load_state("networkidle")
        except Exception as e:
            print(f"[WARN] Consent button not found: {e}")

        # Wait for the download button to appear before attempting click
        download_btn = page.locator('button:has-text("Download full list")')
        try:
            await download_btn.wait_for(state="visible", timeout=15000)
        except Exception:
            await browser.close()
            raise RuntimeError(f"'Download full list' button not found on page: {list_url}")

        # Intercept the download and click the button simultaneously
        try:
            async with page.expect_download(timeout=30000) as download_info:
                await download_btn.click()
            download = await download_info.value
        except Exception as e:
            await browser.close()
            raise RuntimeError(f"Download did not start after clicking the button: {e}")

        # Read the downloaded file directly into BytesIO (never touches disk)
        try:
            stream = await download.path()  # temp path Playwright wrote it to
            if stream is None:
                raise FileNotFoundError("Download path is None — the file may have failed to download.")
            echachem_bytes = BytesIO()
            with open(stream, "rb") as f:
                echachem_bytes.write(f.read())
            echachem_bytes.seek(0)
        except Exception as e:
            await browser.close()
            raise RuntimeError(f"Failed to read downloaded file into BytesIO: {e}")

        await browser.close()

    print(f"[✓] Downloaded '{download.suggested_filename}' into BytesIO.")
    return echachem_bytes

# Run it
echachem_data = asyncio.run(download_echachem_list(list_url = "https://chem.echa.europa.eu/activity-lists/svhcIdentification"))

# Example: load into pandas (adjust based on file type — csv or xlsx)
df = pd.read_excel(echachem_data)

print(df.head())