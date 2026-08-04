"""Restore Foxit printing + printer error detection + batch=0"""
import re

file_path = r'c:\Users\thanh\Downloads\print_shopee_bill\main_app.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

foxit_code = r'''_foxit_exe_cache = None


def _find_foxit_exe():
    """Tìm Foxit PDF Reader — dùng XPS Print Path, spool nhẹ ~15MB."""
    global _foxit_exe_cache
    if _foxit_exe_cache is not None:
        return _foxit_exe_cache or ''
    foxit_paths = [
        r'C:\Program Files (x86)\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe',
        r'C:\Program Files\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe',
    ]
    for fp in foxit_paths:
        if Path(fp).exists():
            _foxit_exe_cache = fp
            return fp
    _foxit_exe_cache = ''
    return ''


def _check_print_errors(printer_name, doc_name_hint='', log_cb=None, timeout=600):
    """Poll print queue để phát hiện lỗi máy in (kẹt giấy, hết mực...).
    Đợi đến khi job Complete hoặc Error thì trả về."""
    import subprocess as _sp, os as _os
    deadline = __import__('time').time() + timeout
    last_status = ''
    error_reported = False
    while __import__('time').time() < deadline:
        __import__('time').sleep(5)
        result = _sp.run(['powershell', '-NoProfile', '-WindowStyle', 'Hidden', '-Command',
            f"$jobs = Get-PrintJob -PrinterName '{printer_name}' -ErrorAction SilentlyContinue"
            + (f" | Where-Object {{ $_.DocumentName -like '*{doc_name_hint[:30]}*' }}" if doc_name_hint else "")
            + " | Select-Object JobStatus | ConvertTo-Json -Compress"
        ], capture_output=True, text=True,
           creationflags=_sp.CREATE_NO_WINDOW if sys.platform == 'win32' else 0)
        status_raw = result.stdout.strip()
        status = ''
        if status_raw:
            try:
                import json as _json
                job_info = _json.loads(status_raw)
                if isinstance(job_info, list):
                    job_info = job_info[0] if job_info else {}
                status = job_info.get('JobStatus', '')
            except Exception:
                status = status_raw

        if status and status != last_status:
            if log_cb: log_cb(f'  🖨️ Trạng thái: {status}', 'dim')
            last_status = status

        if status:
            if 'Complete' in status and 'Printing' not in status and 'Spooling' not in status:
                return True

            error_msgs = {
                'PaperJam': '🛑 KẸT GIẤY! Hãy gỡ giấy kẹt rồi nhấn nút trên máy in.',
                'PaperOut': '📄 HẾT GIẤY! Hãy nạp thêm giấy vào khay.',
                'TonerLow': '⚠ SẮP HẾT MỰC! Chuẩn bị thay mực.',
                'NoToner': '🖌 HẾT MỰC! Cần thay cartridge mực.',
                'Offline': '🔌 MÁY IN MẤT KẾT NỐI! Kiểm tra cáp/WiFi.',
                'Paused': '⏸ MÁY IN ĐANG TẠM DỪNG! Kiểm tra nút trên máy.',
                'DoorOpen': '🚪 NẮP MÁY IN ĐANG MỞ! Đóng nắp lại.',
                'OutputFull': '📦 KHAY RA ĐẦY! Lấy giấy đã in ra.',
            }
            for code, msg in error_msgs.items():
                if code in status and not error_reported:
                    if log_cb: log_cb(f'  {msg}', 'err')
                    error_reported = True
                    break

            if 'Error' in status and not error_reported:
                if log_cb: log_cb(f'  ⚠ Lỗi máy in: {status}', 'err')
                return False
        else:
            if last_status:
                return True

    return True


'''

# === 1. Replace Edge functions with Foxit ===
start_marker = 'def _set_printer_devmode('
end_marker = '\ndef _wait_print_queue('
start_idx = content.find(start_marker)
end_idx = content.find(end_marker)

if start_idx == -1 or end_idx == -1:
    print(f"ERROR: start={start_idx}, end={end_idx}")
    exit(1)

content = content[:start_idx] + foxit_code + content[end_idx + 1:]

# === 2. Fix _print_file ===
# Replace Edge check
content = content.replace("if not _find_edge_exe():", "if not _find_foxit_exe():")
content = content.replace(
    "            if not _find_foxit_exe():",
    "            foxit_exe = _find_foxit_exe()\n            if not foxit_exe:"
)
content = content.replace(
    "Không tìm thấy Microsoft Edge. ",
    "Không tìm thấy Foxit PDF Reader. "
)
content = content.replace(
    "Edge là trình duyệt có sẵn trên Windows 10/11.",
    "Vui lòng cài Foxit PDF Reader để in file PDF."
)

# Replace batch block (Edge → Foxit)
old_batch = """                    _wait_print_queue(printer_name, max_jobs=0)
                    if not _print_pdf_via_edge(batch_path, printer_name, log_cb=log_cb):
                        raise RuntimeError(f'Edge kiosk batch {batch_num} thất bại')
                    _wait_print_queue(printer_name, max_jobs=0)"""
new_batch = """                    _wait_print_queue(printer_name, max_jobs=0)
                    cmd = [foxit_exe, '/t', batch_path, printer_name]
                    result = subprocess.run(cmd, check=False, timeout=600)
                    if result.returncode != 0:
                        raise RuntimeError(f'Foxit batch {batch_num} exit code: {result.returncode}')
                    _check_print_errors(printer_name, doc_name_hint=os.path.basename(batch_path), log_cb=log_cb)"""
content = content.replace(old_batch, new_batch)

# Replace non-batch block
old_nobatch = """                _wait_print_queue(printer_name, max_jobs=0)
                if not _print_pdf_via_edge(print_path, printer_name, log_cb=log_cb):
                    raise RuntimeError('Edge kiosk printing thất bại')
                _wait_print_queue(printer_name, max_jobs=0)"""
new_nobatch = """                _wait_print_queue(printer_name, max_jobs=0)
                cmd = [foxit_exe, '/t', print_path, printer_name]
                result = subprocess.run(cmd, check=False, timeout=600)
                if result.returncode != 0:
                    raise RuntimeError(f'Foxit exit code: {result.returncode}')
                _check_print_errors(printer_name, doc_name_hint=os.path.basename(print_path), log_cb=log_cb)"""
content = content.replace(old_nobatch, new_nobatch)

# Replace log labels
content = content.replace("Edge kiosk batch", "Foxit batch")
content = content.replace("Edge kiosk printing", "Foxit printing")
content = content.replace("  📦 Edge kiosk: Chia", "  📦 Foxit: Chia")

# === 3. batch_size defaults to 0 ===
content = content.replace(
    "batch_size = config.get('batch_size', 55)",
    "batch_size = config.get('batch_size', 0)"
)

# === 4. Update UI ===
content = content.replace("edge_detected = _find_edge_exe()", "foxit_detected = _find_foxit_exe()")
content = content.replace("edge_status", "foxit_status")
content = content.replace("if foxit_detected:", "if foxit_detected:")  # already replaced by above
content = content.replace(
    'foxit_status.setText("✅ Microsoft Edge — Kiosk Printing (tự động poll queue)")',
    'foxit_status.setText("✅ Foxit PDF Reader — XPS Print Path (spool ~15MB)")'
)
content = content.replace(
    'foxit_status.setText("⚠ Không tìm thấy Microsoft Edge — Edge có sẵn trên Windows 10/11!")',
    'foxit_status.setText("⚠ Chưa cài Foxit PDF Reader — Vui lòng cài để in file PDF!")'
)
content = content.replace(
    "# ── Engine in PDF: Microsoft Edge (kiosk printing qua Playwright) ──",
    "# ── Engine in PDF: Foxit PDF Reader (XPS Print Path, spool ~15MB) ──"
)

# === 5. Remove Selenium top-level imports ===
selenium_block = """
# Selenium import cho PyInstaller detect (dùng trong _print_pdf_via_edge)
import selenium.webdriver.edge.webdriver  # noqa: F401 — required for PyInstaller
import selenium.webdriver.edge.options      # noqa: F401
import selenium.webdriver.edge.service      # noqa: F401
"""
content = content.replace(selenium_block, "")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Done! Foxit + printer errors + batch=0 restored.")
