"""
Power BI Playwright Exporter
Automates the manual "Export data" action from a Power BI visual.
Persists the browser session so login only happens once.
"""

import asyncio
import os
import json
import tempfile
from pathlib import Path
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

SESSION_DIR = str(Path(__file__).parent / "session")
DOWNLOAD_DIR = str(Path(__file__).parent / "downloads")


async def export_visual_as_csv(config: dict) -> str:
    """
    Opens the Power BI report, finds the target visual by title,
    triggers 'Export data → CSV', and returns the path to the downloaded file.

    Args:
        config: dict loaded from config.json

    Returns:
        Path to the downloaded CSV file
    """
    report_url = config["powerbi"]["report_url"]
    visual_title = config["powerbi"]["visual_title"]
    export_type = config["powerbi"].get("export_type", "Underlying data")  # or "Summarized data"

    os.makedirs(SESSION_DIR, exist_ok=True)
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    async with async_playwright() as p:
        # Persistent context keeps cookies/session across runs
        browser = await p.chromium.launch_persistent_context(
            SESSION_DIR,
            headless=False,
            viewport={"width": 1920, "height": 1080},
            accept_downloads=True,
            ignore_https_errors=True,  # Bypass corporate SSL inspection CA
            args=["--disable-blink-features=AutomationControlled"],
        )

        page = browser.pages[0] if browser.pages else await browser.new_page()

        print(f"Navigating to Power BI report...")
        try:
            await page.goto(report_url, wait_until="domcontentloaded", timeout=60_000)
        except Exception:
            # Power BI sometimes triggers navigation errors on redirect — that's OK
            pass

        # Handle Microsoft / Shibboleth login if not already authenticated
        await _handle_login_if_needed(page)

        # Wait for the report to fully render
        print("Waiting for report to load...")
        await _wait_for_report(page)

        # Find and export the target visual
        print(f"Looking for visual: '{visual_title}'...")
        downloaded_path = await _export_visual(page, visual_title, export_type, DOWNLOAD_DIR)

        await browser.close()
        return downloaded_path


async def _handle_login_if_needed(page):
    """
    Wait for user to complete login if needed.
    Power BI first goes to /singleSignOn then redirects to the Microsoft login page,
    so we must check for the actual report URL (contains /groups/), not just powerbi.com.
    """
    # Poll the URL for up to 8 seconds to let all redirects settle
    for _ in range(16):
        await page.wait_for_timeout(500)
        url = page.url
        if "/groups/" in url and "app.powerbi.com" in url:
            print("Session active, already on report.")
            return

    # If we're not on the actual report yet, login is required
    print("\n" + "=" * 60)
    print("LOGIN REQUIRED")
    print("Please log in to Power BI in the browser window.")
    print("The script will continue automatically after you log in.")
    print("=" * 60 + "\n")

    try:
        await page.wait_for_url(
            "**/app.powerbi.com/groups/**",
            timeout=180_000,
        )
        print("Login successful! Waiting for report to load...")
        await page.wait_for_timeout(5000)
    except PlaywrightTimeout:
        raise RuntimeError("Login timed out after 3 minutes. Please run the script again.")


async def _wait_for_report(page):
    """Wait for Power BI report visuals to render."""
    selectors = [
        "[class*='visualContainer']",
        "[class*='visual-container']",
        ".visualContainer",
        "[data-testid*='visual']",
        "div[aria-label*='visual']",
    ]

    for selector in selectors:
        try:
            await page.wait_for_selector(selector, timeout=45_000, state="visible")
            print(f"Report loaded (detected via: {selector})")
            await page.wait_for_timeout(3000)
            return
        except PlaywrightTimeout:
            continue
        except Exception:
            # Page may still be navigating — wait and retry
            await page.wait_for_timeout(2000)
            continue

    # Fallback: just wait for network to settle
    print("Visuals not detected yet, waiting for page to fully load...")
    try:
        await page.wait_for_load_state("networkidle", timeout=45_000)
    except Exception:
        pass
    await page.wait_for_timeout(5000)


async def _export_visual(page, visual_title: str, export_type: str, download_dir: str) -> str:
    """
    Finds the visual with the given title, opens its context menu,
    clicks 'Export data', and handles the download.
    """
    # --- Step 1: Find the visual container by title ---
    visual_container = await _find_visual_by_title(page, visual_title)

    if visual_container is None:
        # Discover mode: list all visible visual titles and raise
        await _print_available_visuals(page)
        raise RuntimeError(
            f"Could not find visual with title containing '{visual_title}'.\n"
            "See available visuals listed above. "
            "Update 'visual_title' in config.json to match one of them."
        )

    print(f"Found visual. Hovering to reveal options...")
    await visual_container.hover()
    await page.wait_for_timeout(800)

    # --- Step 2: Click the "More options" (...) button ---
    options_button = await _find_options_button(visual_container, page)
    if options_button is None:
        raise RuntimeError(
            "Could not find the '...' options button on the visual. "
            "Try increasing wait time or check if the report requires a different interaction."
        )

    print("Clicking '...' options button...")
    await options_button.click()
    await page.wait_for_timeout(500)

    # --- Step 3: Click "Export data" in the context menu ---
    export_menu_selectors = [
        "text=Export data",
        "[aria-label*='Export data']",
        "[data-testid*='export']",
        "li:has-text('Export data')",
        "button:has-text('Export data')",
        "a:has-text('Export data')",
    ]

    export_clicked = False
    for selector in export_menu_selectors:
        try:
            item = page.locator(selector).first
            if await item.is_visible(timeout=3000):
                await item.click()
                export_clicked = True
                print("Clicked 'Export data'")
                break
        except Exception:
            continue

    if not export_clicked:
        raise RuntimeError(
            "Could not find 'Export data' in the context menu. "
            "The visual may not support data export."
        )

    await page.wait_for_timeout(1500)

    # --- Step 4: Handle the Export dialog ---
    downloaded_path = await _handle_export_dialog(page, export_type, download_dir)
    return downloaded_path


async def _find_visual_by_title(page, title: str):
    """
    Searches for a visual container whose title text contains `title` (case-insensitive).
    Returns the container element handle or None.
    """
    title_lower = title.lower()

    # Try multiple title selector patterns Power BI uses
    title_selectors = [
        "[class*='visualTitle']",
        "[class*='visual-title']",
        ".title",
        "[aria-label*='title']",
        "h2",
        "h3",
        "[class*='header'] span",
        "[class*='titleText']",
    ]

    for selector in title_selectors:
        elements = await page.query_selector_all(selector)
        for el in elements:
            text = (await el.inner_text()).strip().lower()
            if title_lower in text:
                # Walk up to the visual container
                container = await _get_visual_container(el)
                if container:
                    return container

    return None


async def _get_visual_container(element):
    """Walk up the DOM tree to find the visual container parent."""
    container_class_hints = [
        "visualContainer", "visual-container", "visualWell",
        "visual-modern", "transform-container"
    ]

    # Try up to 10 levels up
    current = element
    for _ in range(10):
        try:
            class_attr = await current.get_attribute("class") or ""
            if any(hint.lower() in class_attr.lower() for hint in container_class_hints):
                return current
            parent = await current.evaluate_handle("el => el.parentElement")
            if parent:
                current = parent.as_element()
                if current is None:
                    break
        except Exception:
            break

    return element  # Return original element if no container found


async def _find_options_button(visual_container, page):
    """Find the '...' more options button within or near the visual container."""
    option_selectors = [
        "[aria-label*='More options']",
        "[aria-label*='more options']",
        "[title*='More options']",
        "button[aria-haspopup='menu']",
        "[class*='headerAction']",
        "[class*='moreOptions']",
        "[class*='optionsMenu']",
        "button[aria-label*='options']",
    ]

    # Search within the visual container first
    for selector in option_selectors:
        try:
            btn = visual_container.locator(selector).first
            if await btn.is_visible(timeout=2000):
                return btn
        except Exception:
            continue

    # Fallback: search in the full page (in case button is outside container)
    for selector in option_selectors:
        try:
            btn = page.locator(selector).first
            if await btn.is_visible(timeout=2000):
                return btn
        except Exception:
            continue

    return None


async def _handle_export_dialog(page, export_type: str, download_dir: str) -> str:
    """
    Handle the Power BI export dialog and capture the downloaded file.
    Returns path to the downloaded file.
    """
    print("Handling export dialog...")
    await page.wait_for_timeout(1000)

    # Try to select the data type (Underlying data vs Summarized data)
    data_type_selectors = [
        f"label:has-text('{export_type}')",
        f"input[value*='{export_type.lower().replace(' ', '')}']",
        f"[aria-label*='{export_type}']",
        f"text={export_type}",
    ]

    for selector in data_type_selectors:
        try:
            el = page.locator(selector).first
            if await el.is_visible(timeout=2000):
                await el.click()
                print(f"Selected export type: {export_type}")
                break
        except Exception:
            continue

    # Select CSV format if the option is available
    csv_selectors = [
        "label:has-text('.csv')",
        "label:has-text('CSV')",
        "input[value='csv']",
        "input[value='CSV']",
        "[aria-label*='CSV']",
        "text=.csv",
    ]

    for selector in csv_selectors:
        try:
            el = page.locator(selector).first
            if await el.is_visible(timeout=2000):
                await el.click()
                print("Selected CSV format")
                break
        except Exception:
            continue

    # Click the Export button and capture the download
    export_btn_selectors = [
        "button:has-text('Export')",
        "button[aria-label*='Export']",
        "button[type='submit']",
        "button:has-text('Download')",
    ]

    async with page.expect_download(timeout=60_000) as download_info:
        for selector in export_btn_selectors:
            try:
                btn = page.locator(selector).last  # "Export" button (not "Export data" from menu)
                if await btn.is_visible(timeout=2000):
                    await btn.click()
                    print("Clicked Export button, waiting for download...")
                    break
            except Exception:
                continue

    download = await download_info.value
    save_path = os.path.join(download_dir, download.suggested_filename or "powerbi_export.csv")
    await download.save_as(save_path)
    print(f"Downloaded: {save_path}")
    return save_path


async def _print_available_visuals(page):
    """Discovery mode: print all detectable visual titles on the page."""
    print("\n" + "=" * 60)
    print("DISCOVER MODE — Visuals found on this report page:")
    print("=" * 60)

    selectors = [
        "[class*='visualTitle']",
        "[class*='visual-title']",
        "[class*='titleText']",
        ".title",
        "h2",
        "h3",
    ]

    seen = set()
    i = 1
    for selector in selectors:
        elements = await page.query_selector_all(selector)
        for el in elements:
            text = (await el.inner_text()).strip()
            if text and text not in seen and len(text) < 200:
                seen.add(text)
                print(f"  {i}. {text!r}")
                i += 1

    print("=" * 60)
    print("Set 'visual_title' in config.json to one of the above names.")
    print()
