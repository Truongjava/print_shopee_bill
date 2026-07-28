"""Test script v2: dùng headless browser để bypass Chrome print dialog."""
import os, json
from pathlib import Path
from datetime import datetime
from playwright.sync_api import sync_playwright

COOKIE_FILE = r"banhang.shopee.vn_27-07-2026.json"
AWB_URL = "https://banhang.shopee.vn/awbprint?shop_id=397879680&job_id=SDK0001_f60c6af943416333c8fe840eaa1d8ae9&lang=vi"
OUTPUT_DIR = r"outputs"

with open(COOKIE_FILE, 'r', encoding='utf-8') as f:
    cd = json.load(f)
cookies_list = cd.get('cookies', cd if isinstance(cd, list) else [])

os.makedirs(OUTPUT_DIR, exist_ok=True)

with sync_playwright() as p:
    # Dùng Chromium của Playwright (headless)
    browser = p.chromium.launch(
        headless=True,
        args=['--disable-blink-features=AutomationControlled']
    )
    context = browser.new_context(
        viewport={'width': 1366, 'height': 768},
        accept_downloads=True
    )

    # Inject cookies
    pw_cookies = []
    for c in cookies_list:
        if not c.get('name') or not c.get('value'):
            continue
        pw = {
            'name': c['name'], 'value': c['value'],
            'domain': c.get('domain', '.shopee.vn'),
            'path': c.get('path', '/')
        }
        if c.get('expirationDate') and not c.get('session'):
            pw['expires'] = int(c['expirationDate'])
        if 'httpOnly' in c: pw['httpOnly'] = c['httpOnly']
        if 'secure' in c: pw['secure'] = c['secure']
        if c.get('sameSite'):
            pw['sameSite'] = {'strict': 'Strict', 'lax': 'Lax',
                            'no_restriction': 'None', 'unspecified': 'Lax'}.get(c['sameSite'], 'Lax')
        pw_cookies.append(pw)
    context.add_cookies(pw_cookies)

    page = context.new_page()

    # Bắt download nếu có
    downloaded = []
    def on_download(dl):
        ts = datetime.now().strftime("%m-%d_%H-%M-%S")
        bp = str(Path(OUTPUT_DIR) / f'awbprint_dl_{ts}.pdf')
        dl.save_as(bp)
        downloaded.append(bp)
        print(f'💾 Download: {bp}')

    page.on('download', on_download)

    print(f'🌐 Đang mở (headless): {AWB_URL}')
    page.goto(AWB_URL, wait_until='networkidle', timeout=60000)
    page.wait_for_timeout(5000)

    # Log nội dung page hiện tại
    body_text = page.evaluate('() => document.body.innerText')
    print(f'📄 Nội dung page: {body_text[:200]}')

    # Click "In phiếu" — headless không mở dialog
    for btn_text in ['In phiếu', 'In', 'Print']:
        try:
            btn = page.locator(f'button:has-text("{btn_text}")').first
            if btn.count() > 0 and btn.is_visible(timeout=3000):
                btn.click(force=True, timeout=5000)
                print(f'✓ Đã bấm "{btn_text}"')
                break
        except Exception as e:
            print(f'  Thử "{btn_text}": {e}')

    # Đợi sau khi click — nội dung có thể thay đổi
    page.wait_for_timeout(8000)
    body_text2 = page.evaluate('() => document.body.innerText')
    print(f'📄 Nội dung sau click: {body_text2[:200]}')

    # Kiểm tra download
    if downloaded:
        print(f'✅ Đã tải {len(downloaded)} file qua download')
    else:
        # Fallback: page.pdf()
        ts = datetime.now().strftime("%m-%d_%H-%M-%S")
        output_path = str(Path(OUTPUT_DIR) / f'awbprint_pdf_{ts}.pdf')
        page.pdf(path=output_path)
        print(f'✅ Đã lưu page.pdf(): {output_path}')

    browser.close()
    print('Done!')
