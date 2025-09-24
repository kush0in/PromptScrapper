import os
import time
import io
from datetime import datetime
from urllib.parse import urlparse
import mimetypes
import tempfile
import shutil
import requests
import pandas as pd
from tqdm import tqdm
import sys
import threading
import queue
import getpass

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import SessionNotCreatedException
import shutil as _shutil


# -------------------------
# CONFIG - edit these
# -------------------------
SAVED_PAGE_URL = os.getenv("THREADS_SAVED_URL", "https://www.threads.com/saved")

# Save images to the PARENT folder of this project, inside 'threads_saved_images'
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.abspath(os.path.join(PROJECT_DIR, os.pardir))
IMAGES_DIR = os.path.join(PARENT_DIR, "threads_saved_images")
os.makedirs(IMAGES_DIR, exist_ok=True)

# Selenium profile reuse
USE_EXISTING_PROFILE = True

BROWSER = os.getenv("BROWSER", "chrome").lower()  # 'edge' or 'chrome'
ALLOW_BROWSER_FALLBACK = False

# Optional: attach to an already running Chrome with remote debugging
CHROME_ATTACH = os.getenv("CHROME_ATTACH", "0") in ("1", "true", "True", "YES", "yes")
CHROME_DEBUG_ADDRESS = os.getenv("CHROME_DEBUG_ADDRESS", "127.0.0.1:9222")

def get_default_user_data_dir():
    win = os.path.expandvars(r"%LOCALAPPDATA%\\Google\\Chrome\\User Data")
    linux = os.path.expanduser("~/.config/google-chrome")
    mac = os.path.expanduser("~/Library/Application Support/Google/Chrome")
    if os.name == "nt":
        return win
    return linux if os.path.isdir(linux) else mac

DEFAULT_WIN_USER_DATA = get_default_user_data_dir()
DEFAULT_LINUX_USER_DATA = DEFAULT_WIN_USER_DATA  # not used directly after refactor
DEFAULT_MAC_USER_DATA = DEFAULT_WIN_USER_DATA    # not used directly after refactor

CHROME_USER_DATA_DIR = DEFAULT_WIN_USER_DATA if os.name == "nt" else (DEFAULT_LINUX_USER_DATA if os.path.isdir(DEFAULT_LINUX_USER_DATA) else DEFAULT_MAC_USER_DATA)

CHROME_PROFILE_DIR = os.getenv("CHROME_PROFILE_DIR", "Default")
# Control whether we are allowed to auto-fallback to a fresh temp profile if the
# existing profile fails to launch (e.g., locked). Set to "0" to DISABLE fallback.
ALLOW_FRESH_PROFILE_FALLBACK = os.getenv("ALLOW_FRESH_PROFILE_FALLBACK", "1") in ("1", "true", "True", "YES", "yes")

# Output
OUTPUT_XLSX = os.getenv("OUTPUT_XLSX_LOCAL", "saved_posts_local.xlsx")

# Credentials (fixed defaults; can be overridden by env vars)
THREADS_ID = os.getenv("THREADS_ID", "Killian_kuffen").strip()
THREADS_PASSWORD = os.getenv("THREADS_PASSWORD", "@Nujagrawa14").strip()


# -------------------------
# Helpers
# -------------------------
def make_driver(use_profile=True, user_data_dir=None, profile_dir=None, headless=False):
    def build_options(browser):
        opts = ChromeOptions()
        # If attaching to a running Chrome, do NOT force headless or user-data-dir
        attaching = CHROME_ATTACH
        if headless and not attaching:
            opts.add_argument("--headless=new")
            opts.add_argument("--disable-gpu")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-notifications")
        opts.add_argument("--window-size=1920,1080")
        opts.add_argument("--disable-extensions")
        opts.add_argument("--disable-logging")
        opts.add_argument("--no-first-run")
        opts.add_argument("--no-default-browser-check")
        opts.add_argument("--remote-allow-origins=*")
        # Attach to an existing Chrome instance if requested
        if attaching:
            try:
                opts.add_experimental_option("debuggerAddress", CHROME_DEBUG_ADDRESS)
                print(f"Attaching to existing Chrome at {CHROME_DEBUG_ADDRESS}")
            except Exception:
                pass
        elif use_profile and user_data_dir and os.path.isdir(user_data_dir):
            opts.add_argument(f"--user-data-dir={user_data_dir}")
            if profile_dir:
                opts.add_argument(f"--profile-directory={profile_dir}")
            print(f"Using {browser} profile: user-data-dir='{user_data_dir}', profile='{profile_dir}'")
        else:
            if use_profile:
                print(f"Warning: Provided {browser} user data dir not found. Launching without profile.")
        return opts

    def start_browser(browser):
        # Always choose the appropriate user_data_dir for Chrome
        user_data = get_default_user_data_dir()
        opts = build_options("chrome")
        # ensure the options reflect the correct browser-specific user data dir
        if not CHROME_ATTACH and use_profile and user_data and os.path.isdir(user_data):
            # rebuild opts to include user_data (since build_options used original)
            opts = ChromeOptions()
            if headless:
                opts.add_argument("--headless=new")
                opts.add_argument("--disable-gpu")
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--disable-notifications")
            opts.add_argument("--window-size=1920,1080")
            opts.add_argument("--disable-extensions")
            opts.add_argument("--disable-logging")
            opts.add_argument("--no-first-run")
            opts.add_argument("--no-default-browser-check")
            opts.add_argument("--remote-allow-origins=*")
            opts.add_argument(f"--user-data-dir={user_data}")
            if profile_dir:
                opts.add_argument(f"--profile-directory={profile_dir}")
            print(f"Using {browser} profile: user-data-dir='{user_data}', profile='{profile_dir}'")
        elif CHROME_ATTACH:
            # When attaching, ensure options include debuggerAddress
            opts = ChromeOptions()
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--disable-notifications")
            opts.add_argument("--window-size=1920,1080")
            opts.add_argument("--disable-extensions")
            opts.add_argument("--disable-logging")
            opts.add_argument("--no-first-run")
            opts.add_argument("--no-default-browser-check")
            opts.add_argument("--remote-allow-origins=*")
            try:
                opts.add_experimental_option("debuggerAddress", CHROME_DEBUG_ADDRESS)
            except Exception:
                pass
        # Chrome only: Prefer Selenium Manager; fall back to CHROME_DRIVER_PATH/PATH; else webdriver_manager
        try:
            return webdriver.Chrome(options=opts)
        except Exception:
            chrome_driver_from_env = os.getenv("CHROME_DRIVER_PATH")
            if not chrome_driver_from_env:
                try:
                    which_path = _shutil.which("chromedriver")
                    if which_path and os.path.isfile(which_path):
                        chrome_driver_from_env = which_path
                except Exception:
                    pass
            if chrome_driver_from_env and os.path.isfile(chrome_driver_from_env):
                service = ChromeService(executable_path=chrome_driver_from_env)
                return webdriver.Chrome(service=service, options=opts)
            service = ChromeService(ChromeDriverManager().install())
            return webdriver.Chrome(service=service, options=opts)

    try:
        driver = start_browser(BROWSER)
        return driver
    except SessionNotCreatedException as e:
        if not ALLOW_FRESH_PROFILE_FALLBACK and use_profile:
            raise RuntimeError(
                "Browser failed to start with your existing profile. Close all browser windows and try again, "
                "or set CHROME_PROFILE_DIR to the correct folder (see edge://version or chrome://version -> Profile Path). "
                "To allow temporary fallback, set ALLOW_FRESH_PROFILE_FALLBACK=1."
            ) from e
        # Fallback: use a temporary fresh profile (avoids DevToolsActivePort crash due to locked profile)
        # If we are attaching to an existing Chrome, do not fallback to temp profile; just re-raise
        if BROWSER == "chrome" and CHROME_ATTACH:
            raise
        temp_profile_dir = tempfile.mkdtemp(prefix="chrome_tmp_profile_")
        print("Existing profile failed to launch; falling back to a temporary fresh profile (not logged in).")
        print(f"Temp profile at: {temp_profile_dir}")
        try:
            # Try with temp profile for current browser
            opts = build_options(BROWSER)
            opts.add_argument(f"--user-data-dir={temp_profile_dir}")
            if BROWSER == "edge":
                service = EdgeService(EdgeChromiumDriverManager().install())
                driver = webdriver.Edge(service=service, options=opts)
            else:
                service = ChromeService(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=opts)
            return driver
        except Exception:
            try:
                shutil.rmtree(temp_profile_dir, ignore_errors=True)
            except Exception:
                pass
            raise
    except Exception as e:
        # Optional cross-browser fallback only if explicitly allowed
        raise


def safe_get_text(elem):
    try:
        txt = elem.text
        return txt.strip() if txt else ""
    except Exception:
        return ""


def extract_image_urls_from_element(elem):
    urls = set()
    try:
        imgs = elem.find_elements(By.TAG_NAME, "img")
        for im in imgs:
            try:
                src = im.get_attribute("src")
                if src:
                    urls.add(src)
            except Exception:
                pass

        all_descendants = elem.find_elements(By.XPATH, ".//*")
        for d in all_descendants:
            try:
                style = d.get_attribute("style")
                if style and "background-image" in style:
                    import re
                    m = re.search(r'url\(["\']?(.*?)["\']?\)', style)
                    if m:
                        urls.add(m.group(1))
            except Exception:
                pass
    except Exception:
        pass
    return list(urls)


# NEW: More robust text extraction
from selenium.webdriver.support.ui import WebDriverWait as _WebDriverWait

def extract_text_from_element(driver, elem, timeout=2):
    try:
        # Ensure element is in viewport for virtualized UIs
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
            time.sleep(0.15)
        except Exception:
            pass

        # 1) Expand any "see more" / "more" buttons within the element
        try:
            more_buttons = elem.find_elements(By.XPATH, ".//*[self::button or @role='button'][contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'see more') or contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'more') or contains(., '…')]")
            for btn in more_buttons:
                try:
                    driver.execute_script("arguments[0].click();", btn)
                except Exception:
                    try:
                        btn.click()
                    except Exception:
                        pass
        except Exception:
            pass

        # 2) Prefer the browser-computed innerText which respects visibility and CSS
        try:
            inner_text = driver.execute_script("return arguments[0].innerText;", elem)
            if inner_text and inner_text.strip():
                return " ".join(inner_text.split())
        except Exception:
            pass

        # 3) Fallback to WebElement.text
        txt = safe_get_text(elem)
        if txt:
            return txt

        # 4) Threads-specific: try to capture caption near a Translate button
        try:
            translate_btns = elem.find_elements(By.XPATH, ".//*[self::button or @role='button'][contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'translate')]")
            for tbtn in translate_btns:
                try:
                    container = tbtn
                    for _ in range(3):
                        container = container.find_element(By.XPATH, "..")
                    # Within this container, look for the first span/p/div with non-trivial text
                    candidates = container.find_elements(By.XPATH, ".//span|.//p|.//div")
                    texts = []
                    for c in candidates:
                        s = safe_get_text(c)
                        if not s:
                            continue
                        ss = s.strip()
                        # ignore UI labels and counters
                        if len(ss) < 2:
                            continue
                        if ss.lower() in ("translate", "more", "see more"):
                            continue
                        if ss.isdigit():
                            continue
                        texts.append(ss)
                    if texts:
                        return " ".join(texts)
                except Exception:
                    continue
        except Exception:
            pass

        # 5) Last resort: join span/p/div texts within elem
        try:
            text_nodes = elem.find_elements(By.XPATH, ".//span|.//p|.//div")
            combined = " ".join([safe_get_text(x) for x in text_nodes if safe_get_text(x)])
            return combined.strip()
        except Exception:
            return ""
    except Exception:
        return ""


def guess_extension_from_response(url, response):
    # 1) Try URL path
    path = urlparse(url).path
    base, ext = os.path.splitext(path)
    if ext and len(ext) <= 5:
        return ext
    # 2) Try content-type header
    ctype = response.headers.get("Content-Type", "").split(";")[0].strip()
    if ctype:
        ext = mimetypes.guess_extension(ctype)
        if ext:
            return ext
    # 3) Default
    return ".jpg"


def download_image_to_disk(url, dest_dir, filename_prefix):
    if not (url.startswith("http://") or url.startswith("https://")):
        raise ValueError("Unsupported image URL: " + url)

    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/122 Safari/537.36"}
    with requests.get(url, headers=headers, timeout=25, stream=True) as resp:
        resp.raise_for_status()
        ext = guess_extension_from_response(url, resp)
        fname = f"{filename_prefix}{ext}"
        fpath = os.path.join(dest_dir, fname)
        with open(fpath, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
    return fpath


# -------------------------
# Auth helpers
# -------------------------
def prompt_with_timeout(prompt_text, timeout_seconds):
    """Prompt the user for input, but return None if no input within timeout_seconds."""
    input_queue = queue.Queue()

    def _get_input():
        try:
            user_in = input(prompt_text)
            input_queue.put(user_in)
        except Exception:
            input_queue.put(None)

    t = threading.Thread(target=_get_input, daemon=True)
    t.start()
    try:
        return input_queue.get(timeout=timeout_seconds)
    except queue.Empty:
        return None


def is_login_page(driver):
    url = driver.current_url.lower()
    if "login" in url or "signin" in url:
        return True
    try:
        forms = driver.find_elements(By.TAG_NAME, "form")
        text_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], input[type='email']")
        pass_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='password']")
        if forms and (text_inputs or pass_inputs):
            return True
    except Exception:
        pass
    return False


def try_fill_and_click(driver, wait, selectors, text=None, click=False, clear_first=True):
    for sel in selectors:
        try:
            elem = driver.find_element(By.CSS_SELECTOR, sel)
            if click:
                elem.click()
                return True
            if clear_first:
                try:
                    elem.clear()
                except Exception:
                    pass
            if text is not None:
                elem.send_keys(text)
            return True
        except Exception:
            continue
    return False


def login_if_needed(driver, wait, saved_page_url):
    """Attempt login if we appear to be on a login page. Uses THREADS_ID/PASSWORD and prompts for OTP (5 min timeout)."""
    try:
        # Heuristic: if not on saved page after initial load, or obvious login page markers, try login
        if not is_login_page(driver):
            # Some sites lazy-redirect; small grace period
            time.sleep(2)
        if not is_login_page(driver):
            return  # assume already logged in

        username = THREADS_ID or prompt_with_timeout("Enter Threads ID/Email: ", 300)
        if not username:
            print("No Threads ID provided; skipping login.")
            return
        password = THREADS_PASSWORD or getpass.getpass("Enter Threads Password: ")

        # Try common field selectors
        user_selectors = [
            "input[name='username']",
            "input[name='email']",
            "input[type='email']",
            "input[type='text']",
            "input[autocomplete='username']",
        ]
        pass_selectors = [
            "input[name='password']",
            "input[type='password']",
            "input[autocomplete='current-password']",
        ]
        submit_selectors_css = [
            "button[type='submit']",
            "input[type='submit']",
            "button[data-testid='login']",
            "button[aria-label*='Log in' i]",
            "button[aria-label*='Sign in' i]",
        ]
        submit_selectors_xpath = [
            "//button[@type='submit']",
            "//input[@type='submit']",
            "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'log in')]",
            "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sign in')]",
            "//div[contains(@role,'button') and (contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'log in') or contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sign in'))]",
        ]

        ok_user = try_fill_and_click(driver, wait, user_selectors, text=username)
        ok_pass = try_fill_and_click(driver, wait, pass_selectors, text=password)
        if not (ok_user and ok_pass):
            print("Could not find login inputs; please log in manually in the browser window.")
        else:
            # Click login/submit if we can find it; if not, press Enter by submitting the form
            clicked = try_fill_and_click(driver, wait, submit_selectors_css, click=True)
            if not clicked:
                # Try XPath variants for text-based buttons
                for xp in submit_selectors_xpath:
                    try:
                        btn = driver.find_element(By.XPATH, xp)
                        btn.click()
                        clicked = True
                        break
                    except Exception:
                        continue
            if not clicked:
                try:
                    # Last resort: submit the password field's form
                    driver.find_element(By.CSS_SELECTOR, pass_selectors[0]).submit()
                except Exception:
                    pass

        # Wait a moment for OTP challenge/redirect
        time.sleep(3)

        # Try detect OTP input; if found, prompt for OTP with 5-min timeout
        otp_selectors = [
            "input[name='otp']",
            "input[name='code']",
            "input[type='tel']",
            "input[autocomplete='one-time-code']",
        ]
        otp_field = None
        # Poll up to ~30s for OTP field to render
        otp_deadline = time.time() + 30
        while time.time() < otp_deadline and otp_field is None:
            for sel in otp_selectors:
                try:
                    otp_field = driver.find_element(By.CSS_SELECTOR, sel)
                    break
                except Exception:
                    continue
            if otp_field is None:
                time.sleep(1)

        if otp_field is not None:
            print("Waiting for OTP (up to 5 minutes). Enter the code here in the terminal.")
            otp = prompt_with_timeout("Enter OTP: ", 300)
            if not otp:
                print("OTP entry timed out. You can enter it manually in the browser if the page allows.")
            else:
                try:
                    otp_field.clear()
                except Exception:
                    pass
                otp_field.send_keys(otp)
                # Try submit
                try:
                    otp_field.submit()
                except Exception:
                    # Try a confirm/continue button
                    try_fill_and_click(driver, wait, [
                        "button[type='submit']",
                        "button[aria-label*='Confirm' i]",
                        "button[aria-label*='Continue' i]",
                    ], click=True)

        # After login, navigate back to saved page
        try:
            driver.get(saved_page_url)
            time.sleep(2)
        except Exception:
            pass
    except Exception as e:
        print(f"Login attempt skipped/failed: {e}")


# -------------------------
# Main pipeline
# -------------------------
def run(saved_page_url=SAVED_PAGE_URL, max_posts=None, headless=False):
    driver = make_driver(use_profile=USE_EXISTING_PROFILE, user_data_dir=None,
                         profile_dir=CHROME_PROFILE_DIR, headless=headless)
    wait = WebDriverWait(driver, 20)

    try:
        print("Opening saved page:", saved_page_url)
        driver.get(saved_page_url)

        try:
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        except Exception:
            pass
        time.sleep(3)

        # Attempt login if we're on a login page; supports env vars and 5-min OTP input
        login_if_needed(driver, wait, saved_page_url)

        CANDIDATE_POST_SELECTORS = [
            'article',
            'div[role="article"]',
            'div[data-testid="post"]',
            'div[class*="post"]',
            'div[class*="card"]',
            'div[class*="thread"]',
            'div[class*="item"]'
        ]

        post_elements = []
        for sel in CANDIDATE_POST_SELECTORS:
            try:
                elems = driver.find_elements(By.CSS_SELECTOR, sel)
                if elems:
                    post_elements = elems
                    print(f"Found {len(elems)} elements using selector '{sel}'")
                    break
            except Exception:
                continue

        if not post_elements:
            imgs = driver.find_elements(By.TAG_NAME, "img")
            print(f"No post elements found. Found {len(imgs)} images on page; will try to group them.")
            for im in imgs:
                try:
                    parent = im.find_element(By.XPATH, "..")
                    post_elements.append(parent)
                except Exception:
                    continue

        SCROLL_PAUSE = 1.0
        last_height = driver.execute_script("return document.body.scrollHeight")
        scrolls = 0
        while (max_posts is None or len(post_elements) < max_posts) and scrolls < 20:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(SCROLL_PAUSE)
            new_elems = []
            for sel in CANDIDATE_POST_SELECTORS:
                try:
                    elems = driver.find_elements(By.CSS_SELECTOR, sel)
                    if elems:
                        new_elems = elems
                        break
                except Exception:
                    continue
            if new_elems:
                post_elements = new_elems
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height
            scrolls += 1

        print(f"Total candidate post elements: {len(post_elements)}")
        if max_posts:
            post_elements = post_elements[:max_posts]

        results = []

        for idx, elem in enumerate(tqdm(post_elements, desc="Processing posts")):
            try:
                src_url = ""
                try:
                    a = elem.find_element(By.TAG_NAME, "a")
                    src_url = a.get_attribute("href") or ""
                except Exception:
                    src_url = driver.current_url

                # Scroll into view then use robust text extractor
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
                    time.sleep(0.1)
                except Exception:
                    pass
                text = extract_text_from_element(driver, elem)

                img_urls = extract_image_urls_from_element(elem)
                saved_paths = []
                for i, img_url in enumerate(img_urls):
                    try:
                        prefix = f"{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}_{idx}_{i}"
                        fpath = download_image_to_disk(img_url, IMAGES_DIR, prefix)
                        saved_paths.append(fpath)
                    except Exception as e:
                        print(f"Failed to save image {img_url[:80]}...: {e}")
                        continue

                if not saved_paths:
                    imgs = elem.find_elements(By.TAG_NAME, "img")
                    for i, im in enumerate(imgs):
                        try:
                            src = im.get_attribute("src")
                            if not src:
                                continue
                            prefix = f"{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}_{idx}_{i}"
                            fpath = download_image_to_disk(src, IMAGES_DIR, prefix)
                            saved_paths.append(fpath)
                        except Exception:
                            continue

                results.append({
                    "source_url": src_url,
                    "text": text,
                    "image_paths": saved_paths,
                    "num_images": len(saved_paths),
                    "scraped_at": datetime.utcnow().isoformat()
                })

                time.sleep(0.3)

            except Exception as e:
                print(f"Error processing element #{idx}: {e}")
                continue

        rows = []
        for r in results:
            rows.append({
                "source_url": r["source_url"],
                "text": r["text"],
                "image_paths": ", ".join(r["image_paths"]),
                "num_images": r["num_images"],
                "scraped_at": r["scraped_at"]
            })

        df = pd.DataFrame(rows)
        csv_out = OUTPUT_XLSX.replace(".xlsx", ".csv")
        df.to_csv(csv_out, index=False, encoding="utf-8-sig")
        df.to_excel(OUTPUT_XLSX, index=False)
        print(f"Saved {len(df)} rows to {csv_out} and {OUTPUT_XLSX}")

        print(f"Images saved to: {IMAGES_DIR}")

    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    run(headless=False, max_posts=200)


