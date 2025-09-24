
# Threads Saved -> Cloudinary -> Excel

A Selenium + Chromium pipeline that opens your Threads "Saved" posts page, extracts image(s) and text, uploads images to Cloudinary, and writes an Excel/CSV with the Cloudinary URLs and post text.

## Prerequisites
- Google Chrome installed and logged into your Threads account.
- Python 3.10+ installed.
- A free Cloudinary account (`cloud_name`, `api_key`, `api_secret`).

## Install
```bash
pip install -r requirements.txt
```

## Configure (Windows PowerShell examples)
Set Cloudinary credentials:
```powershell
setx CLOUDINARY_CLOUD_NAME "your_cloud_name"
setx CLOUDINARY_API_KEY "your_api_key"
setx CLOUDINARY_API_SECRET "your_api_secret"
```

Set the Threads Saved URL (optional; you can also navigate manually in the reused Chrome profile):
```powershell
setx THREADS_SAVED_URL "https://www.threads.net/your_saved_page_or_equivalent"
```

If needed, specify Chrome profile folder name (default is `Default`):
```powershell
setx CHROME_PROFILE_DIR "Profile 1"
```

Notes:
- On Windows, Chrome user data dir auto-detected as `%LOCALAPPDATA%\Google\Chrome\User Data`.
- Ensure the selected profile is already logged in to Threads.

## Run
```powershell
python .\threads_saved_to_cloudinary.py
```
- A Chrome window opens to your saved posts page.
- Script scrolls and processes up to 200 posts by default.
- Outputs: `saved_posts_cloudinary.xlsx` and `saved_posts_cloudinary.csv` in this folder.

## Output Columns
- `source_url`: Best-effort link to the post.
- `text`: Combined text content found in the post container.
- `image_urls`: Comma-separated Cloudinary URLs for uploaded images.
- `num_images`: Count of uploaded images.
- `scraped_at`: UTC timestamp.

## Tips
- If you see no posts detected, open your Saved page manually in the spawned Chrome and refresh.
- Some dynamic sites require more scrolling; you can tweak `max_posts` or the selectors inside the script.
- Respect platform Terms of Service and only scrape your own data or data you have rights to.
