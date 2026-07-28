"""Apply all ShopeePrint fixes to main_app.py."""
file_path = r'c:\Users\thanh\Downloads\print_shopee_bill\main_app.py'
new_steps_path = r'c:\Users\thanh\Downloads\print_shopee_bill\_new_steps_v2.txt'

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

with open(new_steps_path, 'r', encoding='utf-8') as f:
    new_steps = f.read()

fixes = 0

# Fix 1: Header color
content = content.replace(
    'QWidget#headerBar { background: #065F46; }',
    'QWidget#headerBar { background: #EE4D2D; }')
content = content.replace(
    'QWidget#headerBar QLabel#headerSubtitle { color: #A7F3D0;',
    'QWidget#headerBar QLabel#headerSubtitle { color: #FFD4C4;')
fixes += 1

# Fix 2: Window title
content = content.replace(
    'self.setWindowTitle("TikTokPrint")',
    'self.setWindowTitle("ShopeePrint")')
fixes += 1

# Fix 3: 10s delay before "Đã hiểu"
content = content.replace(
    '# ── 1. Dismiss popup "Đã hiểu" ──\n        try:',
    '# ── 1. Dismiss popup "Đã hiểu" ──\n        page.wait_for_timeout(10000)  # Đợi 10 giây\n        try:')
fixes += 1

# Fix 4: 60s delay after pickup
content = content.replace(
    "log_cb('  ⏳ Đợi 30 giây cho popup hiện ra...', 'info')\n            page.wait_for_timeout(30000)",
    "log_cb('  ⏳ Đợi 1 phút cho popup hiện ra...', 'info')\n            page.wait_for_timeout(60000)")
fixes += 1

# Fix 5: Replace step 7-8 with new flow
old_start = '# ── 7. Chọn loại phiếu: "Phiếu xuất hàng + Phiếu gửi hàng và phiếu đóng gói" ──'
old_end = '# ── 9. Đợi tab mới mở ra ──'

idx_start = content.find(old_start)
idx_end = content.find(old_end)

if idx_start != -1 and idx_end != -1:
    content = content[:idx_start] + new_steps + '\n        ' + content[idx_end:]
    fixes += 1
    print(f'Fix 5: Replaced old step 7-9')
else:
    print(f'Fix 5 FAILED: start={idx_start}, end={idx_end}')

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print(f'Applied {fixes} fixes, saved {len(content)} chars')
