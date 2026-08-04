"""
Test full pipeline: tách PDF → tính toán → merge 2-up → in shipping label
"""
import sys, os
sys.path.insert(0, r'c:\Users\thanh\Downloads\print_shopee_bill')
sys.path.insert(0, r'c:\Users\thanh\Downloads\print_shopee_bill\bill_calculate')

from pathlib import Path

PDF_FILE = r"c:\Users\thanh\Downloads\print_shopee_bill\outputs\2026-08-01\Shopee_Viettel Post_batch1_08-01_07-10-27.pdf"
OUT_DIR = r"c:\Users\thanh\Downloads\print_shopee_bill\outputs\2026-08-01"
MASTER = r"c:\Users\thanh\Downloads\print_shopee_bill\mã combo.xlsx"
RETAIL = r"c:\Users\thanh\Downloads\print_shopee_bill\sp bán lẻ.xlsx"
TEMPLATE = r"c:\Users\thanh\Downloads\print_shopee_bill\Bảng thống kê hàng.xlsx"
PRINTER = "HP LaserJet M402dn (C54621)"

from calculator import detect_pdf_type, split_shopee_pdf, extract_order_counts_shopee, \
    load_master_data, load_retail_data, build_sku_index, calculate_results, fill_template
from main_app import _merge_pdf_2up, _wait_print_queue, _print_file
import pypdf

print("=" * 60)
print(f"FILE: {Path(PDF_FILE).name}")
print("=" * 60)

# ── Bước 1: Detect loại PDF ──
pdf_type = detect_pdf_type(PDF_FILE)
print(f"\n1. Loai PDF: {pdf_type}")

# ── Bước 2: Tách phiếu xuất & shipping label ──
print("\n2. Tach PDF...")
picking_path, shipping_path, shipping_pages = split_shopee_pdf(PDF_FILE, OUT_DIR)
print(f"   Picking list: {Path(picking_path).name if picking_path else 'KHONG CO'}")
print(f"   Shipping label: {Path(shipping_path).name if shipping_path else 'KHONG CO'}")
print(f"   So trang shipping: {shipping_pages}")

# ── Bước 3: Load master + retail data ──
print("\n3. Load du lieu...")
master_data = load_master_data(MASTER)
master_skus, prefix_map, ambiguous_map = build_sku_index(master_data)
retail_lookup = load_retail_data(RETAIL) if os.path.exists(RETAIL) else None
print(f"   Master: {len(master_data)} dong")
print(f"   Retail: {len(retail_lookup) if retail_lookup else 0} SKU")

# ── Bước 4: Trích xuất order counts từ picking list ──
if picking_path:
    print("\n4. Trich xuat order counts...")
    order_counts, header_info = extract_order_counts_shopee(
        picking_path, master_skus, prefix_map, ambiguous_map, retail_lookup)
    print(f"   Order counts: {len(order_counts)} SKU")
    for sku, qty in list(order_counts.items())[:5]:
        print(f"     {sku}: x{qty}")
    if len(order_counts) > 5:
        print(f"     ... (+{len(order_counts)-5} SKU nua)")

    # ── Bước 5: Tính toán ──
    print("\n5. Tinh toan...")
    results = calculate_results(master_data, order_counts, retail_lookup)
    print(f"   Results: {len(results)} dong")
    print(f"   Tong Qty: {sum(r['qty'] for r in results)}")
    print(f"   Tong Sold: {sum(r['qty_sold'] for r in results)}")
    print(f"   Tong Promo: {sum(r['promo_qty'] for r in results)}")

    # ── Bước 6: Fill template ──
    print("\n6. Fill template...")
    from datetime import datetime
    xlsx_path = os.path.join(OUT_DIR, f"Phieu_xuat_hang_Viettel_Post_test_{datetime.now().strftime('%H-%M-%S')}.xlsx")
    fill_template(results, TEMPLATE, xlsx_path, carrier='Viettel Post', order_count=shipping_pages, source_label='Shopee')
    print(f"   Da luu: {Path(xlsx_path).name}")

# ── Bước 7: Merge 2-up & in shipping label ──
if shipping_path:
    print("\n7. Merge 2-up shipping label...")
    reader = pypdf.PdfReader(shipping_path)
    total_pages = len(reader.pages)
    print(f"   So trang shipping: {total_pages}")

    if total_pages <= 1:
        print("   ⚠ Chi co 1 trang — in thang, khong merge")
        merged_path = shipping_path
    else:
        merged_path = _merge_pdf_2up(shipping_path)
        if merged_path:
            merged_reader = pypdf.PdfReader(merged_path)
            print(f"   Merged: {len(merged_reader.pages)} trang (tu {total_pages} trang goc)")
        else:
            print("   ⚠ Merge that bai — in file goc")
            merged_path = shipping_path

    print(f"\n8. In shipping label ({Path(merged_path).name})...")
    _print_file(merged_path, PRINTER, batch_size=55,
                log_cb=lambda m, t='': print(f'   [{t}] {m}'))

print("\n" + "=" * 60)
print("HOAN TAT!")
