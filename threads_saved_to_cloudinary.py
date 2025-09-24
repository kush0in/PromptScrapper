import os
import time
import io
from datetime import datetime
import requests
import pandas as pd
from tqdm import tqdm

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager
import cloudinary
import cloudinary.uploader


# -------------------------
# CONFIG - edit these
# -------------------------
# 1) URL of the saved posts page on the site (Threads or other). If using profile reuse, open the saved page manually
SAVED_PAGE_URL = os.getenv("THREADS_SAVED_URL", "https://www.threads.com/saved")  # <-- replace with actual saved page URL

# 2) Cloudinary credentials (recommended: set these as environment variables)
# PowerShell example:
#   setx CLOUDINARY_CLOUD_NAME "your_cloud_name"
#   setx CLOUDINARY_API_KEY "your_api_key"
#   setx CLOUDINARY_API_SECRET "your_api_secret"
CLOUD_NAME = os.getenv("CLOUDINARY_CLOUD_NAME")
CLOUD_KEY = os.getenv("CLOUDINARY_API_KEY")
CLOUD_SECRET = os.getenv("CLOUDINARY_API_SECRET")

# 3) Selenium options
USE_EXISTING_PROFILE = True

# Windows default Chrome user data dir
DEFAULT_WIN_USER_DATA = os.path.expandvars(r"%LOCALAPPDATA%\\Google\\Chrome\\User Data")
# Linux/mac fallbacks if running elsewhere
DEFAULT_LINUX_USER_DATA = os.path.expanduser("~/.config/google-chrome")
DEFAULT_MAC_USER_DATA = os.path.expanduser("~/Library/Application Support/Google/Chrome")

if os.name == "nt":
    CHROME_USER_DATA_DIR = DEFAULT_WIN_USER_DATA
else:
    # coarse fallback
    CHROME_USER_DATA_DIR = DEFAULT_LINUX_USER_DATA if os.path.isdir(DEFAULT_LINUX_USER_DATA) else DEFAULT_MAC_USER_DATA

# The folder inside user-data-dir that has your logged-in profile (e.g., "Default", "Profile 1")
CHROME_PROFILE_DIR = os.getenv("CHROME_PROFILE_DIR", "Default")

# 4) Output Excel filename
OUTPUT_XLSX = os.getenv("OUTPUT_XLSX", "saved_posts_cloudinary.xlsx")


# -------------------------
# Safety checks
# -------------------------
if not (CLOUD_NAME and CLOUD_KEY and CLOUD_SECRET):
    raise SystemExit(
        "Cloudinary credentials not found in environment. Please set CLOUDINARY_CLOUD_NAME, "
        "CLOUDINARY_API_KEY and CLOUDINARY_API_SECRET, or edit the script to include them."
    )

# Configure Cloudinary
cloudinary.config(
    cloud_name=CLOUD_NAME,
    api_key=CLOUD_KEY,
    api_secret=CLOUD_SECRET,
    secure=True
)


# -------------------------
# Helper functions
# -------------------------
def make_driver(use_profile=True, user_data_dir=None, profile_dir=None, headless=False):
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--window-size=1920,1080")

    if use_profile and user_data_dir and os.path.isdir(user_data_dir):
        chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
        if profile_dir:
            chrome_options.add_argument(f"--profile-directory={profile_dir}")

    # Create driver using webdriver-manager
    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


def safe_get_text(elem):
    try:
        txt = elem.text
        return txt.strip() if txt else ""
    except Exception:
        return ""


# NEW: More robust text extraction
from selenium.webdriver.support.ui import WebDriverWait as _WebDriverWait

# Heuristic caption picker from block text
_UI_NOISE_WORDS = {
    "like", "reply", "repost", "share", "follow", "translate", "more", "see more",
    "followers", "following", "posts", "views", "comments"
}

def _pick_caption_from_text_block(text_block):
    try:
        lines = [" ".join(line.split()) for line in text_block.splitlines()]
        candidates = []
        for line in lines:
            if not line:
                continue
            low = line.lower()
            if any(w in low for w in _UI_NOISE_WORDS):
                if len(low) <= 14:
                    continue
            if low.endswith("h") or low.endswith("m") or low.endswith("d"):
                if len(low) <= 3 and low[:-1].isdigit():
                    continue
            if low.isdigit():
                continue
            if len(line) < 2:
                continue
            candidates.append(line)
        if candidates:
            candidates.sort(key=lambda s: len(s), reverse=True)
            return candidates[0]
        return ""
    except Exception:
        return ""


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
        inner_text = None
        try:
            inner_text = driver.execute_script("return arguments[0].innerText;", elem)
            if inner_text and inner_text.strip():
                picked = _pick_caption_from_text_block(inner_text)
                if picked:
                    return picked
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
                    candidates = container.find_elements(By.XPATH, ".//span|.//p|.//div")
                    texts = []
                    for c in candidates:
                        s = safe_get_text(c)
                        if not s:
                            continue
                        ss = s.strip()
                        if len(ss) < 2:
                            continue
                        if ss.lower() in ("translate", "more", "see more"):
                            continue
                        if ss.isdigit():
                            continue
                        texts.append(ss)
                    if texts:
                        joined = "\n".join(texts)
                        picked = _pick_caption_from_text_block(joined)
                        if picked:
                            return picked
                        return " ".join(joined.split())
                except Exception:
                    continue
        except Exception:
            pass

        # 5) Last resort: join span/p/div texts within elem
        try:
            text_nodes = elem.find_elements(By.XPATH, ".//span|.//p|.//div")
            combined = " ".join([safe_get_text(x) for x in text_nodes if safe_get_text(x)])
            combined = combined.strip()
            if combined:
                picked = _pick_caption_from_text_block(combined)
                if picked:
                    return picked
            return combined
        except Exception:
            return ""
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

        # search for background-image style
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


def download_image_bytes(url, session=None, timeout=20):
    session = session or requests.Session()
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"}
    resp = session.get(url, headers=headers, timeout=timeout, stream=True)
    resp.raise_for_status()
    return resp.content


def upload_to_cloudinary_bytes(image_bytes, public_id_prefix=None, tags=None):
    fobj = io.BytesIO(image_bytes)
    import uuid
    public_id = (public_id_prefix + "_" + uuid.uuid4().hex) if public_id_prefix else uuid.uuid4().hex
    res = cloudinary.uploader.upload(fobj, public_id=public_id, resource_type="image", tags=tags or [])
    return res.get("secure_url")


# -------------------------
# Main pipeline
# -------------------------
def run(saved_page_url=SAVED_PAGE_URL, max_posts=None, headless=False):
    driver = make_driver(use_profile=USE_EXISTING_PROFILE, user_data_dir=CHROME_USER_DATA_DIR,
                         profile_dir=CHROME_PROFILE_DIR, headless=headless)
    wait = WebDriverWait(driver, 20)

    try:
        print("Opening saved page:", saved_page_url)
        driver.get(saved_page_url)

        # If auth is required, allow manual login window
        try:
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        except Exception:
            pass
        time.sleep(3)

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
        session = requests.Session()

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
                uploaded_urls = []
                for i, img_url in enumerate(img_urls):
                    if not (img_url.startswith("http://") or img_url.startswith("https://")):
                        continue
                    try:
                        img_bytes = download_image_bytes(img_url, session=session)
                        cloud_url = upload_to_cloudinary_bytes(img_bytes, public_id_prefix=f"{datetime.utcnow().strftime('%Y%m%d')}_{idx}_{i}")
                        uploaded_urls.append(cloud_url)
                    except Exception as e:
                        print(f"Failed to process image {img_url[:80]}...: {e}")
                        continue

                if not uploaded_urls:
                    imgs = elem.find_elements(By.TAG_NAME, "img")
                    for i, im in enumerate(imgs):
                        try:
                            src = im.get_attribute("src")
                            if not src:
                                continue
                            img_bytes = download_image_bytes(src, session=session)
                            cloud_url = upload_to_cloudinary_bytes(img_bytes, public_id_prefix=f"{datetime.utcnow().strftime('%Y%m%d')}_{idx}_{i}")
                            uploaded_urls.append(cloud_url)
                        except Exception:
                            continue

                results.append({
                    "source_url": src_url,
                    "text": text,
                    "image_urls": uploaded_urls,
                    "num_images": len(uploaded_urls),
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
                "image_urls": ", ".join(r["image_urls"]),
                "num_images": r["num_images"],
                "scraped_at": r["scraped_at"]
            })

        df = pd.DataFrame(rows)
        csv_out = OUTPUT_XLSX.replace(".xlsx", ".csv")
        df.to_csv(csv_out, index=False, encoding="utf-8-sig")
        df.to_excel(OUTPUT_XLSX, index=False)
        print(f"Saved {len(df)} rows to {csv_out} and {OUTPUT_XLSX}")

    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    run(headless=False, max_posts=200)


