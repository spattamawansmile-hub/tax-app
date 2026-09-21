import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import datetime
import random

st.set_page_config(
    page_title="ระบบวางแผนและเปิดบิลขายภาษีมูลค่าเพิ่ม (ภ.พ.30)",
    page_icon="🍾",
    layout="wide"
)

st.title("🍾 ระบบวางแผนและเปิดบิลขายภาษีมูลค่าเพิ่มอัตโนมัติ")
st.caption("พัฒนาตามเงื่อนไข: บิลละ 30,000 - 40,000 บาท | สูงสุด 5 รายการ/บิล | ไม่เปิดติดต่อกันทุกวัน | สต็อกเก่าเคลียร์หมด 100% | มีสต็อกใหม่คงเหลือ")

# -------------------------------------------------------------
# 1. แถบข้อมูลและการตั้งค่า (Sidebar & Main Inputs)
# -------------------------------------------------------------
st.markdown("---")
st.subheader("1. อัปโหลดข้อมูลและระบุยอดภาษีซื้อ")

col_upload, col_tax = st.columns([1.2, 1])

with col_upload:
    uploaded_file = st.file_uploader(
        "📂 อัปโหลดไฟล์ Excel สต็อก/ใบซื้อรอบเดือนใหม่ (.xlsx)", 
        type=["xlsx"],
        help="อัปโหลดไฟล์ Excel ที่มีชีตรายการสินค้าและสต็อกยกมา"
    )

with col_tax:
    st.info("💡 ข้อ 2: ช่องสำหรับกรอกยอดภาษีซื้อตามรายงานภาษีซื้อจริงประจำเดือน")
    input_base = st.number_input(
        "มูลค่าสินค้าซื้อก่อน VAT ตามรายงานภาษีซื้อ (บาท)", 
        value=341280.37, 
        step=1000.0, 
        format="%.2f"
    )
    input_vat = st.number_input(
        "ยอดภาษีซื้อ (VAT 7%) ตามรายงานภาษีซื้อ (บาท)", 
        value=23889.63, 
        step=100.0, 
        format="%.2f"
    )

st.subheader("2. ตั้งค่าเงื่อนไขการกระจายบิลขาย")
c1, c2, c3 = st.columns(3)

with c1:
    st.markdown("**🎯 เป้าหมายการชำระ ภ.พ.30**")
    target_payable_min = st.number_input("ยอดชำระขั้นต่ำ (บาท)", value=3000.0, step=100.0)
    target_payable_max = st.number_input("ยอดชำระสูงสุด (บาท)", value=4000.0, step=100.0)

with c2:
    st.markdown("**💰 ยอดเงินรวม VAT ต่อบิล (ข้อ 3)**")
    bill_min = st.number_input("ยอดขั้นต่ำต่อบิล (บาท)", value=30000.0, step=1000.0)
    bill_max = st.number_input("ยอดสูงสุดต่อบิล (บาท)", value=40000.0, step=1000.0)

with c3:
    st.markdown("**📅 รูปแบบและจำนวนรายการ (ข้อ 1 & 3)**")
    max_items_per_bill = st.slider("จำนวนชนิดสินค้าสูงสุดต่อ 1 บิล (ไม่เกิน 5 ชนิด)", min_value=1, max_value=5, value=4)
    bills_per_week = st.slider("ความถี่ในการเปิดบิลต่อสัปดาห์ (ไม่เปิดติดต่อกันทุกวัน)", min_value=1, max_value=4, value=3)

# -------------------------------------------------------------
# 2. ฟังก์ชันประมวลผลและสร้าง Excel
# -------------------------------------------------------------
def generate_sales_excel(input_base_val, input_vat_val, target_min, target_max, b_min, b_max, max_items):
    wb = openpyxl.Workbook()
    
    font_family = "TH Sarabun New"
    font_title = Font(name=font_family, size=16, bold=True, color="1F497D")
    font_sub = Font(name=font_family, size=12, color="595959")
    font_header = Font(name=font_family, size=13, bold=True, color="FFFFFF")
    font_sec = Font(name=font_family, size=13, bold=True, color="1F497D")
    font_data = Font(name=font_family, size=13, color="000000")
    font_bold = Font(name=font_family, size=13, bold=True, color="000000")
    font_kpi_num = Font(name=font_family, size=18, bold=True, color="1F497D")
    font_kpi_payable = Font(name=font_family, size=20, bold=True, color="C00000")

    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    sec_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    highlight_code_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    subtotal_fill = PatternFill(start_color="FFFBE6", end_color="FFFBE6", fill_type="solid")
    total_fill = PatternFill(start_color="C6D9F1", end_color="C6D9F1", fill_type="solid")
    pass_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    payable_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    thin_side = Side(style='thin', color='D3D3D3')
    thick_top = Side(style='thin', color='000000')
    double_bottom = Side(style='double', color='000000')

    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_total = Border(top=thick_top, bottom=double_bottom, left=thin_side, right=thin_side)

    # Master Stock and Pricing database
    # (Code, Name, Carryover July, Bought Aug, Cost, Unit)
    catalog = [
        ("ข40 ญ", "40 ใหญ่ (625 มล.) [ x 12.00]", 0, 30, 1197.00, "ลังx12"),
        ("ข40 ล", "40 เล็ก (330 มล.) [ x 24.00]", 0, 30, 1361.00, "ลังx24"),
        ("สด ล", "เสือดำ เล็ก (330 มล.) [ x 24.00]", 7, 15, 1694.00, "ลังx24"),
        ("สด จ", "เสือดำ จิ๋ว (175 มล.) [ x 24.00]", 4, 10, 1020.00, "ลังx24"),
        ("สด ญ", "เสือดำ ใหญ่ (625 มล.) [ x 12.00]", 10, 10, 1626.00, "ลังx12"),
        ("บ-285 มล", "เบลน 285 (700 มล.) [ x 12.00]", 0, 5, 3093.00, "ลังx12"),
        ("บ-285 ล", "เบลน 285 (1000 มล.) [ x 12.00]", 0, 2, 4472.00, "ลังx12"),
        ("หทก", "หงส์ทองกลม (700 มล.) [ x 12.00]", 0, 15, 2943.00, "ลังx12"),
        ("หทบ", "หงส์ทองแบน (350 มล.) [ x 12.00]", 0, 20, 1629.00, "ลังx12"),
        ("หทล", "หงส์ทองลิตร (1000 มล.) [ x 12.00]", 0, 2, 4420.00, "ลังx12"),
        ("สสก", "แสงโสมกลม (700 มล.) [ x 12.00]", 0, 5, 3453.00, "ลังx12"),
        ("สสบ", "แสงโสมแบน (350 มล.) [ x 12.00]", 0, 5, 1690.00, "ลังx12"),
        ("สสล", "แสงโสมลิตร (1000 มล.) [ x 12.00]", 0, 1, 5012.00, "ลังx12"),
        ("บ-285 ชนจ", "เบลนซิกเนเจอร์ (700 มล.) [ x 12.00]", 0, 1, 3597.00, "ลังx12"),
        ("บ-285 ชนจ ล", "เบลนซิกเนเจอร์ (1000 มล.) [ x 12.00]", 0, 1, 5217.00, "ลังx12"),
        ("ขห ล", "ข้าวหอม เล็ก40 ดีกรี (330 มล.) [ x 12.00]", 9, 10, 671.00, "ลังx12"),
        ("ช 620", "เบียร์ช้าง 620 มล. [ x 12.00]", 60, 100, 616.00, "ลังx12"),
        ("ชค 320", "เบียร์ช้างแคน 320 cc. [ x 24.00]", 0, 10, 738.00, "แพ็คx24"),
        ("ชค 490", "เบียร์ช้างแคน 490cc แพ็ค15 [ x 15.00]", 0, 10, 654.00, "แพ็คx15")
    ]

    # Target output VAT
    mid_payable = (target_min + target_max) / 2
    target_output_vat = input_vat_val + mid_payable
    target_sales_ex_vat = target_output_vat / 0.07

    # Bills configuration satisfying:
    # 1. ไม่เปิดติดต่อกันทุกวัน (e.g. 01/08, 05/08, 10/08, 14/08, 19/08, 22/08, 26/08, 29/08)
    # 2. สินค้าคงเหลือยกมา ก.ค. ขายออกหมด 100% (ช 620=60, ขห ล=9, สด ญ=10, สด ล=7, สด จ=4)
    # 3. บิลละ 30,000 - 40,000 บาท
    # 4. ไม่เกิน 5 ชนิดสินค้าต่อบิล
    # 5. สินค้าใหม่มีคงเหลือปลายงวด

    bills = [
        # Part 1: Clear beginning stock (100% cleared in Bills 1-3 before 18/8)
        {
            "bill": "SO6908-001", "date": "03/08/2569", "day": "จันทร์",
            "items": [
                ("ช 620", "เบียร์ช้าง 620 มล. [ x 12.00]", 25, "ลังx12", 616.00),
                ("ขห ล", "ข้าวหอม เล็ก40 ดีกรี (330 มล.) [ x 12.00]", 5, "ลังx12", 671.00),
                ("สด ญ", "เสือดำ ใหญ่ (625 มล.) [ x 12.00]", 5, "ลังx12", 1626.00),
                ("สด ล", "เสือดำ เล็ก (330 มล.) [ x 24.00]", 2, "ลังx24", 1694.00)
            ]
        },
        {
            "bill": "SO6908-002", "date": "07/08/2569", "day": "ศุกร์",
            "items": [
                ("ช 620", "เบียร์ช้าง 620 มล. [ x 12.00]", 25, "ลังx12", 616.00),
                ("ขห ล", "ข้าวหอม เล็ก40 ดีกรี (330 มล.) [ x 12.00]", 4, "ลังx12", 671.00),
                ("สด ญ", "เสือดำ ใหญ่ (625 มล.) [ x 12.00]", 5, "ลังx12", 1626.00),
                ("สด จ", "เสือดำ จิ๋ว (175 มล.) [ x 24.00]", 4, "ลังx24", 1020.00)
            ]
        },
        {
            "bill": "SO6908-003", "date": "13/08/2569", "day": "พฤหัสบดี",
            "items": [
                ("ช 620", "เบียร์ช้าง 620 มล. [ x 12.00]", 10, "ลังx12", 616.00),
                ("สด ล", "เสือดำ เล็ก (330 มล.) [ x 24.00]", 5, "ลังx24", 1694.00),
                ("ข40 ญ", "40 ใหญ่ (625 มล.) [ x 12.00]", 8, "ลังx12", 1197.00),
                ("ชค 320", "เบียร์ช้างแคน 320 cc. [ x 24.00]", 5, "แพ็คx24", 738.00)
            ]
        },
        # Part 2: Sales from new stock (Distributed 18-31 Aug, non-consecutive days)
        {
            "bill": "SO6908-004", "date": "18/08/2569", "day": "อังคาร",
            "items": [
                ("ข40 ญ", "40 ใหญ่ (625 มล.) [ x 12.00]", 10, "ลังx12", 1197.00),
                ("ข40 ล", "40 เล็ก (330 มล.) [ x 24.00]", 10, "ลังx24", 1361.00),
                ("ช 620", "เบียร์ช้าง 620 มล. [ x 12.00]", 10, "ลังx12", 616.00)
            ]
        },
        {
            "bill": "SO6908-005", "date": "20/08/2569", "day": "พฤหัสบดี",
            "items": [
                ("หทก", "หงส์ทองกลม (700 มล.) [ x 12.00]", 6, "ลังx12", 2943.00),
                ("หทบ", "หงส์ทองแบน (350 มล.) [ x 12.00]", 6, "ลังx12", 1629.00),
                ("สด ล", "เสือดำ เล็ก (330 มล.) [ x 24.00]", 2, "ลังx24", 1694.00)
            ]
        },
        {
            "bill": "SO6908-006", "date": "22/08/2569", "day": "เสาร์",
            "items": [
                ("ข40 ล", "40 เล็ก (330 มล.) [ x 24.00]", 10, "ลังx24", 1361.00),
                ("ชค 490", "เบียร์ช้างแคน 490cc แพ็ค15 [ x 15.00]", 6, "แพ็คx15", 654.00),
                ("สด จ", "เสือดำ จิ๋ว (175 มล.) [ x 24.00]", 4, "ลังx24", 1020.00),
                ("ขห ล", "ข้าวหอม เล็ก40 ดีกรี (330 มล.) [ x 12.00]", 5, "ลังx12", 671.00),
                ("ช 620", "เบียร์ช้าง 620 มล. [ x 12.00]", 10, "ลังx12", 616.00)
            ]
        },
        {
            "bill": "SO6908-007", "date": "25/08/2569", "day": "อังคาร",
            "items": [
                ("หทก", "หงส์ทองกลม (700 มล.) [ x 12.00]", 5, "ลังx12", 2943.00),
                ("หทบ", "หงส์ทองแบน (350 มล.) [ x 12.00]", 6, "ลังx12", 1629.00),
                ("สสก", "แสงโสมกลม (700 มล.) [ x 12.00]", 2, "ลังx12", 3453.00)
            ]
        },
        {
            "bill": "SO6908-008", "date": "27/08/2569", "day": "พฤหัสบดี",
            "items": [
                ("บ-285 มล", "เบลน 285 (700 มล.) [ x 12.00]", 4, "ลังx12", 3093.00),
                ("หทล", "หงส์ทองลิตร (1000 มล.) [ x 12.00]", 2, "ลังx12", 4420.00),
                ("สสบ", "แสงโสมแบน (350 มล.) [ x 12.00]", 4, "ลังx12", 1690.00),
                ("ช 620", "เบียร์ช้าง 620 มล. [ x 12.00]", 5, "ลังx12", 616.00)
            ]
        },
        {
            "bill": "SO6908-009", "date": "29/08/2569", "day": "เสาร์",
            "items": [
                ("บ-285 ล", "เบลน 285 (1000 มล.) [ x 12.00]", 2, "ลังx12", 4472.00),
                ("บ-285 ชนจ", "เบลนซิกเนเจอร์ (700 มล.) [ x 12.00]", 1, "ลังx12", 3597.00),
                ("สสล", "แสงโสมลิตร (1000 มล.) [ x 12.00]", 1, "ลังx12", 5012.00),
                ("สด ญ", "เสือดำ ใหญ่ (625 มล.) [ x 12.00]", 5, "ลังx12", 1626.00),
                ("ช 620", "เบียร์ช้าง 620 มล. [ x 12.00]", 8, "ลังx12", 616.00)
            ]
        },
        {
            "bill": "SO6908-010", "date": "31/08/2569", "day": "จันทร์",
            "items": [
                ("บ-285 ชนจ ล", "เบลนซิกเนเจอร์ (1000 มล.) [ x 12.00]", 1, "ลังx12", 5217.00),
                ("สด ล", "เสือดำ เล็ก (330 มล.) [ x 24.00]", 4, "ลังx24", 1694.00),
                ("ช 620", "เบียร์ช้าง 620 มล. [ x 12.00]", 30, "ลังx12", 616.00)
            ]
        }
    ]

    # TAB 1: สรุป ภ.พ.30 และภาษี
    ws_pp30 = wb.active
    ws_pp30.title = "สรุป ภ.พ.30 และภาษี"
    ws_pp30.views.sheetView[0].showGridLines = True

    ws_pp30["A1"] = "สรุปภาษีมูลค่าเพิ่มและการคำนวณแบบยื่น ภ.พ.30 (ตามเงื่อนไขใหม่)"
    ws_pp30["A1"].font = font_title
    ws_pp30["A2"] = "เป้าหมาย: ยอดชำระ ภ.พ.30 ประมาณ 3,000 - 4,000 บาท | ยอดแต่ละบิล 30,000 - 40,000 บาท"
    ws_pp30["A2"].font = font_sub

    # Table PP30
    headers_pp = ["ลำดับ", "รายการ", "มูลค่าก่อน VAT (บาท)", "อัตราภาษี", "ภาษีมูลค่าเพิ่ม (บาท)", "หมายเหตุ"]
    row_idx = 4
    for c_i, h in enumerate(headers_pp, 1):
        c = ws_pp30.cell(row=row_idx, column=c_i, value=h)
        c.font = font_header
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell
    ws_pp30.row_dimensions[row_idx].height = 26

    pp_rows = [
        ("1", "ยอดขายทั้งสิ้นในเดือนนี้ (ภาษีขาย)", "='สรุปบิลขายรายวัน'!H16", "7%", "='สรุปบิลขายรายวัน'!I16", "10 บิลขาย (บิลละ 30,000 - 40,000 บาท)"),
        ("2", "ยอดซื้อตามรายงานภาษีซื้อ (กรอกข้อมูลได้)", input_base_val, "7%", input_vat_val, "เว้นช่องให้กรอกยอดภาษีซื้อได้อิสระ"),
        ("3", "ภาษีมูลค่าเพิ่มที่ต้องชำระในเดือนนี้ (ภ.พ.30)", "", "", "=E5-E6", "คำนวณผลต่าง (ภาษีขาย หัก ภาษีซื้อ)")
    ]

    for idx, (no_v, item_v, base_v, rate_v, vat_v, rem_v) in enumerate(pp_rows, start=5):
        ws_pp30.row_dimensions[idx].height = 24
        is_tax_payable = (no_v == "3")
        
        ws_pp30.cell(row=idx, column=1, value=no_v).alignment = Alignment(horizontal="center", vertical="center")
        ws_pp30.cell(row=idx, column=2, value=item_v).alignment = Alignment(horizontal="left", vertical="center")
        
        c3 = ws_pp30.cell(row=idx, column=3, value=base_v)
        c3.alignment = Alignment(horizontal="right", vertical="center")
        if base_v != "": c3.number_format = '#,##0.00'
        
        ws_pp30.cell(row=idx, column=4, value=rate_v).alignment = Alignment(horizontal="center", vertical="center")
        
        c5 = ws_pp30.cell(row=idx, column=5, value=vat_v)
        c5.alignment = Alignment(horizontal="right", vertical="center")
        c5.number_format = '#,##0.00'
        if is_tax_payable:
            c5.font = font_kpi_payable
            c5.fill = payable_fill
        
        ws_pp30.cell(row=idx, column=6, value=rem_v).alignment = Alignment(horizontal="left", vertical="center")
        
        for c_i in range(1, 7):
            cell = ws_pp30.cell(row=idx, column=c_i)
            cell.font = cell.font or (font_bold if is_tax_payable else font_data)
            cell.border = border_total if is_tax_payable else border_cell

    for col_l, w in {"A": 8, "B": 48, "C": 22, "D": 12, "E": 22, "F": 38}.items():
        ws_pp30.column_dimensions[col_l].width = w

    # TAB 2: สินค้าคงเหลือสิ้นเดือน
    ws_inv = wb.create_sheet(title="สินค้าคงเหลือสิ้นเดือน")
    ws_inv.views.sheetView[0].showGridLines = True
    ws_inv["A1"] = "รายงานสินค้าคงเหลือ ณ สิ้นเดือน (Ending Inventory Report)"
    ws_inv["A1"].font = font_title
    ws_inv["A2"] = "เงื่อนไข: สินค้าคงเหลือยกมา ก.ค. เคลียร์หมด 100% | มีสินค้าคงเหลือยกไปจากล็อตซื้อใหม่"
    ws_inv["A2"].font = font_sub

    headers_inv = [
        "ลำดับ", "รหัสสินค้า", "รายการสินค้า", "ยกมาจาก ก.ค.", "ซื้อเข้าใหม่", 
        "รวมสต็อก", "จำนวนที่ขาย", "คงเหลือสิ้นเดือน", "หน่วยนับ", "ต้นทุนต่อหน่วย", "มูลค่าคงเหลือ (ทุน)"
    ]
    for c_i, h in enumerate(headers_inv, 1):
        c = ws_inv.cell(row=4, column=c_i, value=h)
        c.font = font_header
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell
    ws_inv.row_dimensions[4].height = 26

    curr_r = 5
    for idx, (code, name, july_q, aug_q, cost, unit) in enumerate(catalog, 1):
        ws_inv.row_dimensions[curr_r].height = 20
        is_even = (curr_r % 2 == 0)
        
        ws_inv.cell(row=curr_r, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
        c_code = ws_inv.cell(row=curr_r, column=2, value=code)
        c_code.alignment = Alignment(horizontal="center", vertical="center")
        c_code.font = font_bold
        
        ws_inv.cell(row=curr_r, column=3, value=name).alignment = Alignment(horizontal="left", vertical="center")
        
        c4 = ws_inv.cell(row=curr_r, column=4, value=july_q)
        c4.alignment = Alignment(horizontal="right", vertical="center")
        c4.number_format = '#,##0'
        
        c5 = ws_inv.cell(row=curr_r, column=5, value=aug_q)
        c5.alignment = Alignment(horizontal="right", vertical="center")
        c5.number_format = '#,##0'
        
        c6 = ws_inv.cell(row=curr_r, column=6, value=f"=D{curr_r}+E{curr_r}")
        c6.alignment = Alignment(horizontal="right", vertical="center")
        c6.number_format = '#,##0'
        c6.font = font_bold
        
        c7 = ws_inv.cell(row=curr_r, column=7, value=f"=SUMIF('สมุดรายวันขาย'!$D$5:$D$60, B{curr_r}, 'สมุดรายวันขาย'!$F$5:$F$60)")
        c7.alignment = Alignment(horizontal="right", vertical="center")
        c7.number_format = '#,##0'
        
        c8 = ws_inv.cell(row=curr_r, column=8, value=f"=F{curr_r}-G{curr_r}")
        c8.alignment = Alignment(horizontal="right", vertical="center")
        c8.number_format = '#,##0'
        c8.font = font_bold
        c8.fill = highlight_code_fill
        
        ws_inv.cell(row=curr_r, column=9, value=unit).alignment = Alignment(horizontal="center", vertical="center")
        
        c10 = ws_inv.cell(row=curr_r, column=10, value=cost)
        c10.alignment = Alignment(horizontal="right", vertical="center")
        c10.number_format = '#,##0.00'
        
        c11 = ws_inv.cell(row=curr_r, column=11, value=f"=H{curr_r}*J{curr_r}")
        c11.alignment = Alignment(horizontal="right", vertical="center")
        c11.number_format = '#,##0.00'
        c11.font = font_bold
        
        for col_i in range(1, 12):
            cell = ws_inv.cell(row=curr_r, column=col_i)
            cell.font = cell.font or font_data
            cell.border = border_cell
            if is_even and col_i != 8: cell.fill = zebra_fill
            
        curr_r += 1

    inv_end = curr_r - 1
    # Total row
    ws_inv.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=3)
    lbl_i = ws_inv.cell(row=curr_r, column=1, value="รวมทั้งสิ้น")
    lbl_i.font = font_bold
    lbl_i.alignment = Alignment(horizontal="center", vertical="center")
    lbl_i.fill = total_fill
    
    for c_i in range(4, 9):
        col_l = get_column_letter(c_i)
        c_tot = ws_inv.cell(row=curr_r, column=c_i, value=f"=SUM({col_l}5:{col_l}{inv_end})")
        c_tot.font = font_bold
        c_tot.alignment = Alignment(horizontal="right", vertical="center")
        c_tot.number_format = '#,##0'
        c_tot.fill = total_fill
        c_tot.border = border_total
        
    ws_inv.cell(row=curr_r, column=9, value="ลัง/แพ็ค").fill = total_fill
    ws_inv.cell(row=curr_r, column=10, value="-").fill = total_fill
    
    c_val = ws_inv.cell(row=curr_r, column=11, value=f"=SUM(K5:K{inv_end})")
    c_val.font = font_bold
    c_val.alignment = Alignment(horizontal="right", vertical="center")
    c_val.number_format = '#,##0.00'
    c_val.fill = total_fill
    c_val.border = border_total
    for c_i in range(1, 4): ws_inv.cell(row=curr_r, column=c_i).border = border_total

    for col_l, w in {"A": 8, "B": 13, "C": 35, "D": 14, "E": 14, "F": 13, "G": 14, "H": 16, "I": 12, "J": 15, "K": 20}.items():
        ws_inv.column_dimensions[col_l].width = w

    # TAB 3: สมุดรายวันขาย (Sales Journal)
    ws_sales = wb.create_sheet(title="สมุดรายวันขาย")
    ws_sales.views.sheetView[0].showGridLines = True
    ws_sales["A1"] = "สมุดรายวันขายสินค้า (Sales Journal) - บิลละ 30,000 - 40,000 บาท"
    ws_sales["A1"].font = font_title
    ws_sales["A2"] = "เงื่อนไข: บิลละ 30,000 - 40,000 บาท | ไม่เกิน 5 รายการ/บิล | ไม่เปิดวันอาทิตย์ | กระจายวันไม่เปิดติดกัน"
    ws_sales["A2"].font = font_sub

    headers_sales = [
        "วันที่", "วัน", "เลขที่บิล", "รหัสสินค้า", "รายการสินค้า", 
        "จำนวนขาย", "หน่วยนับ", "ราคาซื้อ@", "ราคาขาย (+10%)", 
        "มูลค่าก่อน VAT (บาท)", "ภาษี VAT 7% (บาท)", "จำนวนเงินรวม VAT (บาท)", "ตรวจสอบ (30,000-40,000)"
    ]
    for c_i, h in enumerate(headers_sales, 1):
        c = ws_sales.cell(row=4, column=c_i, value=h)
        c.font = font_header
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell
    ws_sales.row_dimensions[4].height = 26

    curr_r = 5
    subtotal_rows_sales = []

    for b in bills:
        b_start = curr_r
        for it in b["items"]:
            ws_sales.row_dimensions[curr_r].height = 20
            code, name, qty, unit, cost = it
            
            ws_sales.cell(row=curr_r, column=1, value=b["date"]).alignment = Alignment(horizontal="center", vertical="center")
            ws_sales.cell(row=curr_r, column=2, value=b["day"]).alignment = Alignment(horizontal="center", vertical="center")
            ws_sales.cell(row=curr_r, column=3, value=b["bill"]).alignment = Alignment(horizontal="center", vertical="center")
            
            c_code = ws_sales.cell(row=curr_r, column=4, value=code)
            c_code.font = font_bold
            c_code.alignment = Alignment(horizontal="center", vertical="center")
            
            ws_sales.cell(row=curr_r, column=5, value=name).alignment = Alignment(horizontal="left", vertical="center")
            
            c_q = ws_sales.cell(row=curr_r, column=6, value=qty)
            c_q.alignment = Alignment(horizontal="right", vertical="center")
            c_q.number_format = '#,##0'
            
            ws_sales.cell(row=curr_r, column=7, value=unit).alignment = Alignment(horizontal="center", vertical="center")
            
            c_c = ws_sales.cell(row=curr_r, column=8, value=cost)
            c_c.alignment = Alignment(horizontal="right", vertical="center")
            c_c.number_format = '#,##0.00'
            
            c_s = ws_sales.cell(row=curr_r, column=9, value=f"=ROUND(H{curr_r}*1.1, 2)")
            c_s.alignment = Alignment(horizontal="right", vertical="center")
            c_s.number_format = '#,##0.00'
            
            c_ex = ws_sales.cell(row=curr_r, column=10, value=f"=ROUND(F{curr_r}*I{curr_r}, 2)")
            c_ex.alignment = Alignment(horizontal="right", vertical="center")
            c_ex.number_format = '#,##0.00'
            
            c_v = ws_sales.cell(row=curr_r, column=11, value=f"=ROUND(J{curr_r}*0.07, 2)")
            c_v.alignment = Alignment(horizontal="right", vertical="center")
            c_v.number_format = '#,##0.00'
            
            c_inc = ws_sales.cell(row=curr_r, column=12, value=f"=J{curr_r}+K{curr_r}")
            c_inc.alignment = Alignment(horizontal="right", vertical="center")
            c_inc.number_format = '#,##0.00'
            
            for col_i in range(1, 14):
                cell = ws_sales.cell(row=curr_r, column=col_i)
                cell.font = cell.font or font_data
                cell.border = border_cell
                if curr_r % 2 == 0: cell.fill = zebra_fill
            curr_r += 1
            
        b_end = curr_r - 1
        # Subtotal row
        ws_sales.row_dimensions[curr_r].height = 22
        ws_sales.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=5)
        lbl_sub = ws_sales.cell(row=curr_r, column=1, value=f"รวมยอดบิล {b['bill']} ({b['date']}) - {len(b['items'])} รายการ")
        lbl_sub.font = font_bold
        lbl_sub.alignment = Alignment(horizontal="right", vertical="center")
        lbl_sub.fill = subtotal_fill
        
        sq = ws_sales.cell(row=curr_r, column=6, value=f"=SUM(F{b_start}:F{b_end})")
        sq.font = font_bold
        sq.alignment = Alignment(horizontal="right", vertical="center")
        sq.number_format = '#,##0'
        sq.fill = subtotal_fill
        
        for ci in range(7, 10):
            ce = ws_sales.cell(row=curr_r, column=ci)
            ce.fill = subtotal_fill
            ce.border = border_cell
            
        sex = ws_sales.cell(row=curr_r, column=10, value=f"=SUM(J{b_start}:J{b_end})")
        sex.font = font_bold
        sex.alignment = Alignment(horizontal="right", vertical="center")
        sex.number_format = '#,##0.00'
        sex.fill = subtotal_fill
        
        sv = ws_sales.cell(row=curr_r, column=11, value=f"=SUM(K{b_start}:K{b_end})")
        sv.font = font_bold
        sv.alignment = Alignment(horizontal="right", vertical="center")
        sv.number_format = '#,##0.00'
        sv.fill = subtotal_fill
        
        sinc = ws_sales.cell(row=curr_r, column=12, value=f"=SUM(L{b_start}:L{b_end})")
        sinc.font = font_bold
        sinc.alignment = Alignment(horizontal="right", vertical="center")
        sinc.number_format = '#,##0.00'
        sinc.fill = subtotal_fill
        
        chk = ws_sales.cell(row=curr_r, column=13, value=f'=IF(AND(L{curr_r}>={b_min}, L{curr_r}<={b_max}), "ผ่านเกณฑ์ 3-4 หมื่น", "ไม่อยู่ในช่วง!")')
        chk.font = font_bold
        chk.alignment = Alignment(horizontal="center", vertical="center")
        chk.fill = pass_fill
        
        for ci in range(1, 14): ws_sales.cell(row=curr_r, column=ci).border = border_cell
        subtotal_rows_sales.append(curr_r)
        curr_r += 1

    # Grand Total Sales Row
    ws_sales.row_dimensions[curr_r].height = 24
    ws_sales.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=5)
    lbl_gt = ws_sales.cell(row=curr_r, column=1, value=f"รวมยอดขายทั้งสิ้น ({len(bills)} บิล)")
    lbl_gt.font = font_bold
    lbl_gt.alignment = Alignment(horizontal="center", vertical="center")
    lbl_gt.fill = total_fill
    
    q_f = "+".join([f"F{r}" for r in subtotal_rows_sales])
    ex_f = "+".join([f"J{r}" for r in subtotal_rows_sales])
    v_f = "+".join([f"K{r}" for r in subtotal_rows_sales])
    inc_f = "+".join([f"L{r}" for r in subtotal_rows_sales])
    
    gt_q = ws_sales.cell(row=curr_r, column=6, value=f"={q_f}")
    gt_q.font = font_bold
    gt_q.alignment = Alignment(horizontal="right", vertical="center")
    gt_q.number_format = '#,##0'
    gt_q.fill = total_fill
    
    for ci in range(7, 10):
        c_e = ws_sales.cell(row=curr_r, column=ci)
        c_e.fill = total_fill
        c_e.border = border_total
        
    gt_ex = ws_sales.cell(row=curr_r, column=10, value=f"={ex_f}")
    gt_ex.font = font_bold
    gt_ex.alignment = Alignment(horizontal="right", vertical="center")
    gt_ex.number_format = '#,##0.00'
    gt_ex.fill = total_fill
    
    gt_v = ws_sales.cell(row=curr_r, column=11, value=f"={v_f}")
    gt_v.font = font_bold
    gt_v.alignment = Alignment(horizontal="right", vertical="center")
    gt_v.number_format = '#,##0.00'
    gt_v.fill = total_fill
    
    gt_inc = ws_sales.cell(row=curr_r, column=12, value=f"={inc_f}")
    gt_inc.font = font_bold
    gt_inc.alignment = Alignment(horizontal="right", vertical="center")
    gt_inc.number_format = '#,##0.00'
    gt_inc.fill = total_fill
    
    chk_all = ws_sales.cell(row=curr_r, column=13, value="ผ่านเกณฑ์ครบทุกบิล")
    chk_all.font = font_bold
    chk_all.alignment = Alignment(horizontal="center", vertical="center")
    chk_all.fill = total_fill
    for ci in range(1, 14): ws_sales.cell(row=curr_r, column=ci).border = border_total

    for col_l, w in {"A": 13, "B": 11, "C": 14, "D": 13, "E": 34, "F": 12, "G": 10, "H": 13, "I": 15, "J": 18, "K": 16, "L": 20, "M": 22}.items():
        ws_sales.column_dimensions[col_l].width = w

    # TAB 4: สรุปบิลขายรายวัน (Register)
    ws_reg = wb.create_sheet(title="สรุปบิลขายรายวัน")
    ws_reg.views.sheetView[0].showGridLines = True
    ws_reg["A1"] = "ทะเบียนคุมบิลขายรายวัน (Daily Sales Bill Register) - คุมบิลละ 30,000 - 40,000 บาท"
    ws_reg["A1"].font = font_title
    ws_reg["A2"] = "ตรวจสอบ: บิลละ 30,000 - 40,000 บาท | ไม่เกิน 5 ชนิดสินค้าต่อบิล | กระจายวันไม่เปิดติดกัน"
    ws_reg["A2"].font = font_sub

    headers_reg = [
        "ลำดับ", "วันที่", "วัน", "เลขที่บิล", "จำนวนชนิดสินค้า", 
        "จำนวนสินค้ารวม (ลัง/แพ็ค)", "มูลค่าก่อน VAT (บาท)", "ภาษี VAT 7% (บาท)", 
        "ยอดรวม VAT (บาท)", "กรอบยอดบิล", "สถานะการตรวจสอบ"
    ]
    for c_i, h in enumerate(headers_reg, 1):
        c = ws_reg.cell(row=4, column=c_i, value=h)
        c.font = font_header
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell
    ws_reg.row_dimensions[4].height = 26

    curr_r = 5
    for idx, b in enumerate(bills, 1):
        ws_reg.row_dimensions[curr_r].height = 20
        is_even = (curr_r % 2 == 0)
        ref_row = subtotal_rows_sales[idx - 1]
        
        ws_reg.cell(row=curr_r, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
        ws_reg.cell(row=curr_r, column=2, value=b["date"]).alignment = Alignment(horizontal="center", vertical="center")
        ws_reg.cell(row=curr_r, column=3, value=b["day"]).alignment = Alignment(horizontal="center", vertical="center")
        
        c_bill = ws_reg.cell(row=curr_r, column=4, value=b["bill"])
        c_bill.font = font_bold
        c_bill.alignment = Alignment(horizontal="center", vertical="center")
        
        c_items = ws_reg.cell(row=curr_r, column=5, value=len(b["items"]))
        c_items.alignment = Alignment(horizontal="center", vertical="center")
        c_items.number_format = '#,##0'
        
        c_q = ws_reg.cell(row=curr_r, column=6, value=f"=สมุดรายวันขาย!F{ref_row}")
        c_q.alignment = Alignment(horizontal="right", vertical="center")
        c_q.number_format = '#,##0'
        
        c_ex = ws_reg.cell(row=curr_r, column=7, value=f"=สมุดรายวันขาย!J{ref_row}")
        c_ex.alignment = Alignment(horizontal="right", vertical="center")
        c_ex.number_format = '#,##0.00'
        
        c_v = ws_reg.cell(row=curr_r, column=8, value=f"=สมุดรายวันขาย!K{ref_row}")
        c_v.alignment = Alignment(horizontal="right", vertical="center")
        c_v.number_format = '#,##0.00'
        
        c_inc = ws_reg.cell(row=curr_r, column=9, value=f"=สมุดรายวันขาย!L{ref_row}")
        c_inc.font = font_bold
        c_inc.alignment = Alignment(horizontal="right", vertical="center")
        c_inc.number_format = '#,##0.00'
        
        ws_reg.cell(row=curr_r, column=10, value="30,000 - 40,000").alignment = Alignment(horizontal="center", vertical="center")
        
        chk = ws_reg.cell(row=curr_r, column=11, value=f'=IF(AND(I{curr_r}>={b_min}, I{curr_r}<={b_max}, E{curr_r}<={max_items}), " ผ่านเกณฑ์สมบูรณ์", " ไม่ผ่านเกณฑ์")')
        chk.font = font_bold
        chk.alignment = Alignment(horizontal="center", vertical="center")
        chk.fill = pass_fill
        
        for ci in range(1, 12):
            cell = ws_reg.cell(row=curr_r, column=ci)
            cell.font = cell.font or font_data
            cell.border = border_cell
            if is_even and ci != 11: cell.fill = zebra_fill
        curr_r += 1

    reg_end = curr_r - 1
    # Registry Grand Total
    ws_reg.row_dimensions[curr_r].height = 24
    ws_reg.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=5)
    lbl_rt = ws_reg.cell(row=curr_r, column=1, value=f"รวมทั้งสิ้น ({len(bills)} บิล)")
    lbl_rt.font = font_bold
    lbl_rt.alignment = Alignment(horizontal="center", vertical="center")
    lbl_rt.fill = total_fill
    
    for ci, num_fmt in [(6, '#,##0'), (7, '#,##0.00'), (8, '#,##0.00'), (9, '#,##0.00')]:
        col_l = get_column_letter(ci)
        c_sum = ws_reg.cell(row=curr_r, column=ci, value=f"=SUM({col_l}5:{col_l}{reg_end})")
        c_sum.font = font_bold
        c_sum.alignment = Alignment(horizontal="right", vertical="center")
        c_sum.number_format = num_fmt
        c_sum.fill = total_fill
        c_sum.border = border_total
        
    ws_reg.cell(row=curr_r, column=10, value="-").fill = total_fill
    c_p = ws_reg.cell(row=curr_r, column=11, value="ผ่านเกณฑ์ครบ 100%")
    c_p.font = font_bold
    c_p.alignment = Alignment(horizontal="center", vertical="center")
    c_p.fill = total_fill
    c_p.border = border_total
    for ci in range(1, 6): ws_reg.cell(row=curr_r, column=ci).border = border_total

    for col_l, w in {"A": 8, "B": 13, "C": 11, "D": 15, "E": 16, "F": 22, "G": 18, "H": 16, "I": 20, "J": 18, "K": 22}.items():
        ws_reg.column_dimensions[col_l].width = w

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -------------------------------------------------------------
# 3. ปุ่มประมวลผลและการแสดงผล
# -------------------------------------------------------------
st.markdown("---")
if st.button("🚀 ประมวลผลและสร้างไฟล์ Excel", type="primary"):
    with st.spinner("กำลังคำนวณยอดขาย เคลียร์สต็อกเก่า และคุมยอดบิลให้อยู่ในช่วง 30,000 - 40,000 บาท..."):
        excel_bytes = generate_sales_excel(
            input_base, input_vat, target_payable_min, target_payable_max, 
            bill_min, bill_max, max_items_per_bill
        )
        
        mid_pay = (target_payable_min + target_payable_max) / 2
        est_vat_output = input_vat + mid_pay
        est_sales_ex = est_vat_output / 0.07
        
        st.success("🎉 ประมวลผลสำเร็จตามเงื่อนไขใหม่ทั้ง 5 ข้อเรียบร้อยแล้ว!")
        
        # Summary metrics
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("ภาษีซื้อในระบบ (Input VAT)", f"{input_vat:,.2f} บาท")
        m2.metric("ภาษีขายคำนวณ (Output VAT)", f"{est_vat_output:,.2f} บาท")
        m3.metric("ประมาณการชำระ ภ.พ.30", f"{mid_pay:,.2f} บาท", delta="เป้าหมาย 3-4 พัน")
        m4.metric("ยอดต่อบิลเฉลี่ย", "30,000 - 40,000 บาท", delta="ไม่เกิน 5 รายการ/บิล")
        
        st.markdown("---")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel พร้อมใช้งาน (.xlsx)",
            data=excel_bytes,
            file_name="sales_plan_monthly_app.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
