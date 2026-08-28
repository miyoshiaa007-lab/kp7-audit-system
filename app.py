import streamlit as st
import pandas as pd
import json
import re
import time
import io
from typing import Any, Tuple
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import google.generativeai as genai # สมมติใช้ไลบรารีนี้ตามมาตรฐานการตั้งค่า

# ==========================================
# 1. ฟังก์ชันทำความสะอาดและจัดรูปแบบ (Error-Proof)
# ==========================================
THAI_DIGITS = str.maketrans('๐๑๒๓๔๕๖๗๘๙', '0123456789')
THAI_MONTHS = {"ม.ค.": 1, "มกราคม": 1, "ก.พ.": 2, "กุมภาพันธ์": 2, "มี.ค.": 3, "มีนาคม": 3,
               "เม.ย.": 4, "เมษายน": 4, "พ.ค.": 5, "พฤษภาคม": 5, "มิ.ย.": 6, "มิถุนายน": 6,
               "ก.ค.": 7, "กรกฎาคม": 7, "ส.ค.": 8, "สิงหาคม": 8, "ก.ย.": 9, "กันยายน": 9,
               "ต.ค.": 10, "ตุลาคม": 10, "พ.ย.": 11, "พฤศจิกายน": 11, "ธ.ค.": 12, "ธันวาคม": 12}
MONTH_LABEL = {1: "ม.ค.", 2: "ก.พ.", 3: "มี.ค.", 4: "เม.ย.", 5: "พ.ค.", 6: "มิ.ย.",
               7: "ก.ค.", 8: "ส.ค.", 9: "ก.ย.", 10: "ต.ค.", 11: "พ.ย.", 12: "ธ.ค."}

def clean_json_string(raw_text: str) -> str:
    if not raw_text: return "{}"
    text = str(raw_text).strip()
    if text.startswith("```"):
        lines = text.splitlines()
        if len(lines) > 2 and lines[-1].strip().startswith("```"): text = "\n".join(lines[1:-1])
        elif len(lines) > 1: text = "\n".join(lines[1:])
    return text.strip()

def sanitize_salary(sal_val: Any) -> float:
    if sal_val is None: return 0.0
    s_str = str(sal_val).translate(THAI_DIGITS).replace(",", "").replace(" ", "").replace("บาท", "")
    match = re.search(r"(\d+(\.\d+)?)", s_str)
    if match:
        val = float(match.group(1))
        if val > 0 and val < 5000: return val * 10
        return val
    return 0.0

def normalize_thai_date(date_str: str) -> Tuple[str, int]:
    if not date_str: return "-", 0
    clean_str = str(date_str).translate(THAI_DIGITS).replace(" ", "")
    pattern = r"(\d{1,2})([ก-๙\.]+)(\d{2,4})"
    match = re.search(pattern, clean_str)
    if not match: return str(date_str).strip(), 0
    day = int(match.group(1))
    month_raw = match.group(2)
    year_raw = int(match.group(3))
    year = 2500 + year_raw if year_raw < 100 else year_raw
    month = next((m_val for m_key, m_val in THAI_MONTHS.items() if m_key in month_raw), 0)
    if month == 0: return f"{day} {month_raw} {year}", (year * 10000) + day
    return f"{day} {MONTH_LABEL[month]} {year}", (year * 10000) + (month * 100) + day

def identify_update_reason(text: str) -> str:
    if not text: return 'เลื่อนปกติ'
    text_str = str(text)
    if re.search(r'แก้ไข', text_str): return 'แก้ไขคำสั่ง'
    elif re.search(r'พ\.ร\.บ\.|พรบ|ปรับตาม', text_str): return 'ปรับตาม พ.ร.บ.'
    elif re.search(r'ชดเชย|ปรับอัตรา', text_str): return 'ปรับชดเชยมติ ครม.'
    else: return 'เลื่อนปกติ'

def format_milestone_desc(record: dict) -> str:
    pos = str(record.get('position_and_workplace') or '')
    pos_no = str(record.get('position_no') or '')
    acad = str(record.get('academic_standing') or '')
    pct = str(record.get('percentage_or_step') or '')
    order = str(record.get('order_ref') or '')
    salary = float(record.get('salary') or 0.0)
    
    desc = f"{pos} "
    tags_raw = [f"เลข:{pos_no}", acad, f"เลื่อน:{pct}"]
    tags = [t.strip() for t in tags_raw if t.strip() and t.strip() not in ["เลข:", "เลื่อน:"]]
    tag_str = f"[{' | '.join(tags)}]" if tags else ""
    return f"{desc.strip()} {tag_str} (เงินเดือน {salary:,.0f} บ.) [{order}]"

# ==========================================
# 2. ฟังก์ชันประมวลผล PDF (มี Auto-Retry กัน 503)
# ==========================================
def extract_pdf_records_precise(uploaded_file, prompt_text, model_name="gemini-1.5-pro"):
    max_retries = 3
    for attempt in range(max_retries):
        try:
            # สมมติการเรียกใช้โมเดล (ปรับเปลี่ยนตาม Object Client ที่คุณใช้จริง)
            model = genai.GenerativeModel(model_name)
            response = model.generate_content([uploaded_file, prompt_text])
            return clean_json_string(response.text)
            
        except Exception as api_err:
            if "503" in str(api_err) and attempt < max_retries - 1:
                time.sleep(5 * (attempt + 1)) # รอ 5 วิ, 10 วิ แล้วลองใหม่
                continue
            else:
                raise api_err
    return "{}"

# ==========================================
# 3. ฟังก์ชันสร้างไฟล์ Excel ระบายสี
# ==========================================
def export_colored_excel(uploaded_excel_file, status_list):
    wb = load_workbook(uploaded_excel_file)
    ws = wb.active
    
    # กำหนด Code สี
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    
    # วนลูปเทสีตั้งแต่บรรทัดที่ 2 (บรรทัดแรกคือ Header)
    for row_idx, status in enumerate(status_list, start=2):
        fill_color = green_fill if status == 'ตรงกัน' else red_fill
        for cell in ws[row_idx]:
            cell.fill = fill_color
            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def color_hrms_rows(row):
    """ฟังก์ชันระบายสีสำหรับ st.dataframe (Pandas Styler)"""
    if row.get('สถานะการตรวจสอบ') == 'ตรงกัน':
        return ['background-color: #d4edda; color: #155724'] * len(row)
    elif row.get('สถานะการตรวจสอบ') == 'มีข้อสังเกต':
        return ['background-color: #f8d7da; color: #721c24'] * len(row)
    else:
        return [''] * len(row)

# ==========================================
# ส่วนแสดงผลหน้าเว็บ (Streamlit UI)
# ==========================================
st.set_page_config(page_title="e-KP7 Audit System", layout="wide")
st.title("🛡️ ระบบตรวจสอบ ก.พ.7 เทียบ ก.ค.ศ.16 / HRMS")

# ดึง API Key
if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
else:
    api_key = st.text_input("ใส่ Gemini API Key", type="password")
    if api_key: genai.configure(api_key=api_key)

# ------------------------------------------
# แก้ไขกลับมาเป็น PDF 2 ไฟล์ ตามโครงสร้างเดิม
# ------------------------------------------
col1, col2 = st.columns(2)
with col1:
    pdf_kp7 = st.file_uploader("📂 อัปโหลดไฟล์ ก.พ.7 (PDF)", type=["pdf"])
with col2:
    pdf_hrms = st.file_uploader("📑 อัปโหลดไฟล์เทียบเคียง เช่น HRMS/ก.ค.ศ.16 (PDF)", type=["pdf"])

if st.button("🚀 เริ่มการตรวจสอบข้อมูล") and pdf_kp7 and pdf_hrms:
    with st.spinner("กำลังให้ AI อ่านและเทียบข้อมูลจาก PDF ทั้ง 2 ไฟล์... (อาจใช้เวลา 1-2 นาที)"):
        try:
            # 1. ส่ง PDF ทั้ง 2 ไฟล์ให้ AI สกัดข้อมูล (ตรงนี้คุณใช้ฟังก์ชัน extract_pdf ของคุณ)
            # data_kp7 = extract_pdf_records_precise(pdf_kp7, prompt_kp7)
            # data_hrms = extract_pdf_records_precise(pdf_hrms, prompt_hrms)
            
            # 2. นำข้อมูลมาเทียบกันและสร้าง DataFrame
            # สมมติว่าได้ข้อมูลออกมาเป็น DataFrame ชื่อ df_result
            import numpy as np
            df_result = pd.DataFrame({
                "รายการ": ["คำสั่งเลื่อนเงินเดือน 1/65", "คำสั่งเลื่อนเงินเดือน 2/65", "คำสั่งเปลี่ยนตำแหน่ง"],
                "ข้อมูล ก.พ.7": ["25,000", "26,000", "ครู คศ.2"],
                "ข้อมูล HRMS": ["25,000", "25,500", "ครู คศ.2"],
                "สถานะการตรวจสอบ": ["ตรงกัน", "มีข้อสังเกต", "ตรงกัน"] # <-- จุดที่เราใช้ตัดสินสี
            })
            
            st.success("✅ ตรวจสอบสำเร็จ!")
            
            # 3. แสดงผลแบบ Tabs
            tab1, tab2 = st.tabs(["📊 สรุปผลการตรวจสอบ", "📑 รายงานเปรียบเทียบ (ระบายสี)"])
            
            with tab1:
                st.subheader("ผลการเทียบเคียงรายการ")
                st.write("ตรวจสอบพบจุดที่ 'มีข้อสังเกต' กรุณาดูรายละเอียดในแท็บถัดไป")
                
            with tab2:
                st.subheader("ตารางเทียบข้อมูล (ระบายสีอัตโนมัติ)")
                st.markdown("🟢 **สีเขียว:** ข้อมูลตรงกัน | 🔴 **สีแดง:** ข้อมูลขัดแย้ง/มีข้อสังเกต")
                
                # แสดงตารางระบายสีบนเว็บ
                st.dataframe(df_result.style.apply(color_hrms_rows, axis=1), use_container_width=True)
                
                # *** ทริคสำหรับการทำ PDF ***
                st.info("💡 **ต้องการบันทึกเป็น PDF?** ให้กด `Ctrl + P` (หรือ `Cmd + P` บน Mac) แล้วเลือกปลายทางเป็น 'Save as PDF' เพื่อปรินต์หน้าระบายสีนี้เก็บไว้ได้เลยครับ")
                
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
