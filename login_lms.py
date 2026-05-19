from playwright.sync_api import sync_playwright

LOGIN_URL = "https://lms.kitakerja.id/login/index.php"
STORAGE_FILE = "storage_state.json"

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=False,
        channel="chromium"
    )

    context = browser.new_context()
    page = context.new_page()

    page.goto(LOGIN_URL, wait_until="networkidle")

    print("Silakan login manual di browser yang terbuka.")
    print("Setelah berhasil masuk dashboard LMS, tekan ENTER di terminal ini.")

    input()

    context.storage_state(path=STORAGE_FILE)

    print(f"Session berhasil disimpan ke {STORAGE_FILE}")

    browser.close()