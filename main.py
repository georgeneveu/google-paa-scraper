# ================= INSTALL REQUIREMENTS =================
# pip install playwright
# pip install pandas
# pip install openpyxl
# pip install win10toast
#
# IMPORTANT:
# After installing Playwright, run:
# python -m playwright install
# =======================================================

from playwright.sync_api import sync_playwright
import time
import random
import pandas as pd
import os
from win10toast import ToastNotifier


# ================= CONFIG =================

# Run browser visibly (False) or in background (True)
HEADLESS = False

# Random delay range between actions (helps mimic human behavior)
DELAY_RANGE = (1.5, 3.5)

# Persistent Chrome profile folder
# Stores cookies, sessions, and browser state between runs
PROFILE_DIR = "chrome_profile"

# Input file containing keywords in a "Query" column
INPUT_FILE = "list-to-parse.xlsx"

# Output file where results are continuously saved
OUTPUT_FILE = "list-to-parse-output.xlsx"

# ==========================================

# Windows notification object
toaster = ToastNotifier()


# ================= UTILITIES =================

def human_delay():
    """
    Sleep for a random amount of time.
    Makes scraping behavior less robotic.
    """
    time.sleep(random.uniform(*DELAY_RANGE))


def notify_captcha():
    """
    Display a Windows notification when CAPTCHA is detected.
    Does not block script execution.
    """
    try:
        toaster.show_toast(
            "CAPTCHA Needed 🚨",
            "Solve CAPTCHA in browser. Script will continue automatically.",
            duration=8,
            threaded=True
        )
    except:
        print("⚠️ Notification failed (win10toast issue)")


def is_captcha(page):
    """
    Check page content for common CAPTCHA indicators.
    Returns True if CAPTCHA is likely present.
    """
    content = page.content().lower()
    return (
        "captcha" in content
        or "unusual traffic" in content
    )


def wait_if_captcha(page):
    """
    If CAPTCHA is detected:
    - Notify user
    - Periodically reload page
    - Continue automatically once CAPTCHA disappears
    """

    if not is_captcha(page):
        return

    print("\n🛑 CAPTCHA detected!")
    notify_captcha()

    while True:
        time.sleep(3)

        try:
            page.reload(wait_until="domcontentloaded")
        except:
            pass

        if not is_captcha(page):
            print("✅ CAPTCHA solved, continuing...")
            time.sleep(2)
            return

        print("⏳ Waiting for CAPTCHA to be solved...")


# ================= GOOGLE PAA SCRAPER =================

def get_google_paa(page, keyword):
    """
    Search Google for a keyword and extract
    the first People Also Ask (PAA) questions.

    Returns:
        list[str] -> first 4 PAA questions
    """

    # Build Google search URL
    url = (
        f"https://www.google.com/search?"
        f"q={keyword.replace(' ', '+')}&hl=en&gl=us"
    )

    # Open search results
    page.goto(url, wait_until="domcontentloaded")
    human_delay()

    # Pause if Google triggers CAPTCHA
    wait_if_captcha(page)

    # Attempt to accept cookie popup
    try:
        page.click("button:has-text('Accept')", timeout=3000)
        human_delay()
    except:
        pass

    # Scroll to encourage PAA block loading
    page.mouse.wheel(0, random.randint(2000, 3500))
    human_delay()

    # Wait for PAA section
    try:
        page.wait_for_selector(
            "div[jsname='yEVEwb']",
            timeout=8000
        )
    except:
        print("❌ No PAA block")
        return ["", "", "", ""]

    # Locate all PAA question elements
    elements = page.locator("div[jsname='yEVEwb']")

    questions = []

    # Extract unique questions
    for i in range(elements.count()):
        text = elements.nth(i).inner_text().strip()

        if text and text not in questions:
            questions.append(text)

    # Ensure list always contains 4 items
    questions += [""] * (4 - len(questions))

    return questions[:4]


# ================= LOAD OR RESUME =================

# Resume previous run if output file exists
if os.path.exists(OUTPUT_FILE):

    df = pd.read_excel(OUTPUT_FILE)

    print("🔁 Resuming existing file...")

else:
    # Start fresh from input file
    df = pd.read_excel(INPUT_FILE)

    # Create output columns
    df["PAA1"] = ""
    df["PAA2"] = ""
    df["PAA3"] = ""
    df["PAA4"] = ""

    print("🆕 Starting new run...")


# ================= MAIN EXECUTION =================

with sync_playwright() as p:

    # Launch Chromium using persistent profile
    # Preserves cookies and browser sessions
    context = p.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        headless=HEADLESS,
        viewport={"width": 1280, "height": 900},

        # Browser fingerprint
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0.0.0 Safari/537.36"
        ),
    )

    # Reuse existing page if available
    page = (
        context.pages[0]
        if context.pages
        else context.new_page()
    )

    # Process each keyword
    for i, row in df.iterrows():

        keyword = row["Query"]

        # Skip rows that already contain results
        if str(row["PAA1"]).strip():
            print(f"⏭ Skipping: {keyword}")
            continue

        print(f"\n🔎 {i+1}/{len(df)}: {keyword}")

        try:
            # Retrieve PAA questions
            results = get_google_paa(page, keyword)

            # Save results into dataframe
            df.at[i, "PAA1"] = results[0]
            df.at[i, "PAA2"] = results[1]
            df.at[i, "PAA3"] = results[2]
            df.at[i, "PAA4"] = results[3]

            # Print collected questions
            for idx, q in enumerate(results, 1):
                if q:
                    print(f"{idx}. {q}")

        except Exception as e:
            print(f"⚠️ Error: {e}")

        # Save after every keyword
        # Prevents data loss if script crashes
        df.to_excel(OUTPUT_FILE, index=False)

        print("💾 Saved")

        human_delay()

    # Close browser
    context.close()

print("\n✅ DONE!")
