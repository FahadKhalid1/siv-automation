import os
import asyncio
from dotenv import load_dotenv
from playwright.async_api import async_playwright

# Load environment variables from .env file
load_dotenv()

# Read environment variables
TARGET_URL = os.getenv("TARGET_URL")
FIELDS = {
    "numero_immatriculation": os.getenv("SEL_numero_immatriculation"),
    "date_premiere_immat": os.getenv("SEL_date_premiere_immat"),
    "date_certificat": os.getenv("SEL_date_certificat"),
    "nom_prenom": os.getenv("SEL_nom_prenom"),
    "raison_sociale": os.getenv("SEL_raison_sociale"),
}
SUBMIT_SELECTOR = os.getenv("SUBMIT_SELECTOR")
SUCCESS_SELECTOR = os.getenv("SUCCESS_SELECTOR")
ERROR_SELECTOR = os.getenv("ERROR_SELECTOR")
NAV_TIMEOUT = int(os.getenv("NAVIGATION_TIMEOUT_MS", "60000"))


async def submit_form(data):
    """
    Submits a single form entry on the SIV website.
    Expects a dict like:
      {
        "numero_immatriculation": "GN-124-XM",
        "date_premiere_immat": "24/04/2023",
        "date_certificat": "24/04/2023",
        "nom_prenom": "",
        "raison_sociale": "VOLKSWAGEN BANK GESELLSCHAFT MIT BESCHRAENKTER HAFTUNG"
      }
    """
    async with async_playwright() as p:
        # Launch browser (VISIBLE mode)
        # browser = await p.chromium.launch(headless=False, slow_mo=300)
        # Launch browser (headless mode)
        browser = await p.chromium.launch(headless=True)

        page = await browser.new_page()

        try:
            print(f"➡️  Opening {TARGET_URL}")
            await page.goto(TARGET_URL, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
            print("➡️ Page opened")

            # Identify iframe that holds the form
            frame = None
            for f in page.frames:
                if "csa_retour_dem_certificat" in (f.url or ""):
                    frame = f
                    break

            if not frame:
                raise Exception("❌ Could not locate iframe containing the form.")
            print(f"✅ Found iframe: {frame.url}")

            # Wait for first field to appear
            await frame.wait_for_selector(FIELDS["numero_immatriculation"], timeout=60000)
            print("✅ Form loaded — filling fields")

            # Fill all provided fields
            for key, selector in FIELDS.items():
                if not selector:
                    continue
                value = data.get(key, "")
                if value:
                    print(f"   Filling {key}: {value}")
                    await frame.fill(selector, value)

            print("➡️  Submitting form…")
            await asyncio.sleep(1.5)

            # Try to scroll and ensure button is in view
            try:
                button = await frame.wait_for_selector(SUBMIT_SELECTOR, timeout=20000)
                await button.scroll_into_view_if_needed()
                await button.focus()
                await asyncio.sleep(0.5)
                await button.click(force=True)
                print("✅ Clicked the submit button.")
            except Exception as e:
                print(f"⚠️ Could not click normally, trying JavaScript click… ({e})")
                await frame.evaluate("""
                    (sel) => {
                        const btn = document.querySelector(sel);
                        if (btn) btn.click();
                    }
                """, SUBMIT_SELECTOR)
                print("✅ Clicked with JS fallback.")

            # Wait for response/network completion
            await frame.wait_for_load_state("networkidle", timeout=NAV_TIMEOUT)

            # Check for success or error messages
            success_found = False
            error_found = False
            try:
                if SUCCESS_SELECTOR:
                    await frame.wait_for_selector(SUCCESS_SELECTOR, timeout=5000)
                    success_found = True
            except:
                pass
            try:
                if ERROR_SELECTOR:
                    await frame.wait_for_selector(ERROR_SELECTOR, timeout=5000)
                    error_found = True
            except:
                pass

            # Prioritize error detection first
            if error_found:
                print("❌ Submission error detected.")
                result = {"status": "error"}
            elif success_found:
                print("✅ Submission successful!")
                result = {"status": "ok"}
            else:
                print("⚠️ Could not detect success or error explicitly.")
                result = {"status": "unknown"}

        except Exception as e:
            print(f"💥 Exception: {e}")
            result = {"status": "error", "message": str(e)}
        finally:
            await browser.close()

    return result


# === Example local test ===
if __name__ == "__main__":
    test_data = {
        "numero_immatriculation": "GN-124-XM",
        "date_premiere_immat": "24/04/2023",
        "date_certificat": "24/04/2023",
        "nom_prenom": "",
        "raison_sociale": "VOLKSWAGEN BANK GESELLSCHAFT MIT BESCHRAENKTER HAFTUNG",
    }

    asyncio.run(submit_form(test_data))
