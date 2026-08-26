import streamlit as st
import pandas as pd
import json
import re
import io
import os
from datetime import datetime
from pydantic import BaseModel, Field
from typing import List, Optional, Tuple, Dict, Any
import google.generativeai as genai
from pdf2image import convert_from_bytes
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ==========================================
# 1. โครงสร้างข้อมูลมาตรฐาน (Pydantic Schema)
# ==========================================
class RecordEntry(BaseModel):
    date_raw: str = Field(default="", description="วันเดือนปี เช่น '1 เม.ย. 54' หรือ '1 เม.ย. 2554'")
    position_and_workplace: str = Field(default="", description="ตำแหน่ง หน่วยงาน วิทยฐานะ หรือการเลื่อนขั้น")
    position_no: Optional[str] = Field(default="", description="เลขที่ตำแหน่ง เช่น '5693', '3332'")
    academic_standing: Optional[str] = Field(default="", description="วิทยฐานะ เช่น 'ชำนาญการ', 'ชำนาญการพิเศษ'")
    salary: Optional[float] = Field(default=0.0, description="อัตราเงินเดือนเป็นตัวเลข เช่น 25190")
    order_ref: Optional[str] = Field(default="", description="เลขที่คำสั่งและวันที่ลงนามคำสั่ง")

class KP7ExtractionResult(BaseModel):
    records: List[RecordEntry] = Field(default=[], description="รายการประวัติทั้งหมดเรียงตามลำดับในหน้าเอกสาร")

# ==========================================
# 2. ฟังก์ชันแปลงข้อมูล (Data Normalization)
# ==========================================
THAI_MONTHS = {
    "ม.ค.": 1, "มค": 1, "มกราคม": 1,
    "ก.พ.": 2, "กพ": 2, "กุมภาพันธ์": 2,
    "มี.ค.": 3, "มีค": 3, "มีนาคม": 3,
    "เม.ย.": 4, "เมย": 4, "เมษายน": 4,
    "พ.ค.": 5, "พค": 5, "พฤษภาคม": 5,
    "มิ.ย.": 6, "มิย": 6, "มิถุนายน": 6,
    "ก.ค.": 7, "กค": 7, "กรกฎาคม": 7,
    "ส.ค.": 8, "สค": 8, "สิงหาคม": 8,
    "ก.ย.": 9, "กย": 9, "กันยายน": 9,
    "ต.ค.": 10, "ตค": 10, "ตุลาคม": 10,
    "พ.ย.": 11, "พย": 11, "พฤศจิกายน": 11,
    "ธ.ค.": 12, "ธค": 12, "ธันวาคม": 12
}

MONTH_LABEL = {
    1: "ม.ค.", 2: "ก.พ.", 3: "มี.ค.", 4: "เม.ย.",
    5: "พ.ค.", 6: "มิ.ย.", 7: "ก.ค.", 8: "ส.ค.",
    9: "ก.ย.", 10: "ต.ค.", 11: "พ.ย.", 12: "ธ.ค."
}

def clean_json_string(raw_text: str) -> str:
    """ตัดบล็อก Markdown ครอบ JSON ออกอย่างปลอดภัย"""
    text = raw_text.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        if len(lines) > 2 and lines[-1].strip().startswith("```"):
            text = "\n".join(lines[1:-1])
        elif len(lines) > 1:
            text = "\n".join(lines[1:])
    return text.strip()

def normalize_thai_date(date_str: str) -> Tuple[str, int]:
    """แปลงวันที่ภาษาไทยให้อยู่ในรูปแบบมาตรฐาน 'D ด.ด. YYYY' และ Sort Key ตัวเลข"""
    if not date_str or not isinstance(date_str, str):
        return "-", 0
    
    clean_str = date_str.replace(" ", "").replace(".", ". ")
    pattern = r"(\d{1,2})\s*([ก-๙\.]+)\s*(\d{2,4})"
    match = re.search(pattern, clean_str)
    
    if not match:
        return date_str.strip(), 0
    
    day = int(match.group(1))
    month_raw = match.group(2).replace(" ", "")
    year_raw = int(match.group(3))
    
    year = 2500 + year_raw if year_raw < 100 else year_raw
    month = 0
    for m_key, m_val in THAI_MONTHS.items():
        if m_key in month_raw:
            month = m_val
            break
            
    if month == 0:
        return f"{day} {month_raw} {year}", (year * 10000) + day
        
    formatted = f"{day} {MONTH_LABEL[month]} {year}"
    sort_key = (year * 10000) + (month * 100) + day
    return formatted, sort_key

# ==========================================
# 3. VLM Data Extractor Engine
# ==========================================
def extract_pdf_records(pdf_bytes: bytes, api_key: str, model_name: str, hint: str) -> List[Dict[str, Any]]:
    genai.configure(api_key=api_key)
    
    try:
        model = genai.GenerativeModel(
            model_name=model_name,
            generation_config={
                "response_mime_type": "application/json",
                "response_schema": KP7ExtractionResult,
                "temperature": 0.0
            }
        )
    except Exception:
        model = genai.GenerativeModel(
            model_name=model_name,
            generation_config={"temperature": 0.0}
        )
    
    images = convert_from_bytes(pdf_bytes, dpi=200)
    extracted_rows = []
    
    prompt = f"""
    คุณคือผู้เชี่ยวชาญการอ่านทะเบียนประวัติ ก.พ.7 / ก.ค.ศ.16 สพป.มหาสารคาม เขต 2
    ประเภทเอกสาร: {hint}
    
    คำสั่ง:
    1. สกัดข้อมูลประวัติการดำรงตำแหน่งและเงินเดือนทุกบรรทัด 100% ห้ามข้ามแม้แต่แถวเดียว
    2. อ่านลายมือภาษาไทย ตัวเลขไทย/อารบิก และคำย่อให้เที่ยงตรง
    3. ส่งผลลัพธ์เป็น JSON โครงสร้างนี้เท่านั้น:
    {{
      "records": [
        {{
          "date_raw": "1 เม.ย. 54",
          "position_and_workplace": "ครู ชำนาญการ รร.บ้านโคกสูงหนองเสียวหนอง (เลื่อน 1 ขั้น)",
          "position_no": "5693",
          "academic_standing": "ชำนาญการ",
          "salary": 25190,
          "order_ref": "คส.สพป.มค.2 ที่ 199/54 ลว. 18 เม.ย. 54"
        }}
      ]
    }}
    """
    
    for page_idx, img in enumerate(images, start=1):
        resp = model.generate_content([prompt, img])
        cleaned_str = clean_json_string(resp.text)
        
        try:
            data = json.loads(cleaned_str)
            records = data.get("records", [])
        except Exception:
            records = []
            
        for r in records:
            norm_date, s_key = normalize_thai_date(r.get("date_raw", ""))
            r["normalized_date"] = norm_date
            r["sort_key"] = s_key
            r["page_no"] = page_idx
            # แปลงเงินเดือนเป็น Float เสมอ
            try:
                r["salary"] = float(r.get("salary", 0))
            except Exception:
                r["salary"] = 0.0
            extracted_rows.append(r)
            
    return extracted_rows

# ==========================================
# 4. อัลกอริทึม Two-Way Full Outer Matching
# ==========================================
def run_two_way_full_outer_matching(records_a: List[Dict[str, Any]], records_b: List[Dict[str, Any]]):
    """
    อัลกอริทึมจับคู่ 2 ทิศทาง (Two-Way Full Outer Match):
    records_a = ไฟล์ ก.พ.7 อิเล็กทรอนิกส์ (HRMS)
    records_b = ไฟล์ ก.ค.ศ.16 เล่มเขียนมือ
    """
    # 1. ตรวจสอบการเขียนวันที่ย้อนหลัง (Timeline Inversions)
    def detect_inversions(records, name):
        inv_list = []
        for i in range(1, len(records)):
            prev = records[i - 1]
            curr = records[i]
            if curr["sort_key"] > 0 and prev["sort_key"] > 0:
                if curr["sort_key"] < prev["sort_key"]:
                    inv_list.append({
                        "row": i + 1,
                        "date_curr": curr["date_raw"],
                        "date_prev": prev["date_raw"],
                        "msg": f"แถวที่ {i+1}: ลงวันที่ '{curr['date_raw']}' อยู่ถัดจาก '{prev['date_raw']}' (ลำดับเวลาย้อนกลับ)"
                    })
        return inv_list

    inversions_a = detect_inversions(records_a, "ก.พ.7 อิเล็กทรอนิกส์")
    inversions_b = detect_inversions(records_b, "ก.ค.ศ.16 เขียนมือ")

    # 2. จัดกลุ่ม Record ตาม Composite Key (Normalized Date + Salary)
    # ใช้ Composite Key เพื่อรองรับการปรับเงินเดือนหลายคำสั่งในวันเดียวกัน
    grouped_a = {}
    for r in records_a:
        key = (r["normalized_date"], r["salary"], r.get("position_no", ""))
        grouped_a.setdefault(key, []).append(r)

    grouped_b = {}
    for r in records_b:
        key = (r["normalized_date"], r["salary"], r.get("position_no", ""))
        grouped_b.setdefault(key, []).append(r)

    # 3. รวม Key ทั้งหมดเข้าด้วยกัน (Full Outer Set of Keys)
    all_keys = set(grouped_a.keys()).union(set(grouped_b.keys()))

    # จัดเรียง Key ตามลำดับเวลาจริง
    def key_sort_function(k):
        norm_date, _, _ = k
        _, sort_val = normalize_thai_date(norm_date)
        return sort_val

    sorted_keys = sorted(list(all_keys), key=key_sort_function)

    # 4. Two-Way Outer Reconciliation
    matched_rows = []
    stats = {
        "total_unique_events": len(sorted_keys),
        "perfect_match": 0,
        "duplicate_in_hrms": 0,
        "duplicate_in_manual": 0,
        "missing_in_manual": 0,
        "missing_in_hrms": 0
    }

    for key in sorted_keys:
        norm_date, sal, pos_no = key
        list_a = grouped_a.get(key, [])
        list_b = grouped_b.get(key, [])
        
        len_a = len(list_a)
        len_b = len(list_b)
        
        # กรณี 1: มีข้อมูลทั้ง 2 ฝั่ง (Matched)
        if len_a > 0 and len_b > 0:
            if len_a == 1 and len_b == 1:
                status = "✅ ตรงกันสมบูรณ์"
                action = "-"
                stats["perfect_match"] += 1
            elif len_a > 1 and len_b == 1:
                status = "⚠️ ตรงกัน แต่ระบบอิเล็กทรอนิกส์มีรายการซ้ำ"
                action = f"พบข้อมูลซ้ำในระบบ ก.พ.7 จำนวน {len_a} แถว (แนะนำลบออก {len_a - 1} แถว)"
                stats["duplicate_in_hrms"] += 1
            elif len_a == 1 and len_b > 1:
                status = "⚠️ ตรงกัน แต่เล่มเขียนมือมีรายการซ้ำ"
                action = f"พบเขียนซ้ำในเล่ม ก.ค.ศ.16 จำนวน {len_b} แถว"
                stats["duplicate_in_manual"] += 1
            else:
                status = "⚠️ ตรงกัน แต่มีรายการซ้ำทั้งสองระบบ"
                action = f"ก.พ.7 มี {len_a} แถว, ก.ค.ศ.16 มี {len_b} แถว"
                stats["duplicate_in_hrms"] += 1

            desc_a = f"{list_a[0]['position_and_workplace']} | เงินเดือน {list_a[0]['salary']:,.0f} | เลขตำแหน่ง: {list_a[0]['position_no']} | คำสั่ง: {list_a[0]['order_ref']}"
            desc_b = f"{list_b[0]['position_and_workplace']} | เงินเดือน {list_b[0]['salary']:,.0f} | เลขตำแหน่ง: {list_b[0]['position_no']} | คำสั่ง: {list_b[0]['order_ref']}"

        # กรณี 2: มีเฉพาะในระบบอิเล็กทรอนิกส์ (Missing in Manual)
        elif len_a > 0 and len_b == 0:
            status = "❌ ขาดในเล่มเขียนมือ (ต้องเพิ่ม)"
            action = "เพิ่มรายการคำสั่งนี้ลงในสมุด ก.ค.ศ.16 เขียนมือ"
            stats["missing_in_manual"] += 1
            desc_a = f"{list_a[0]['position_and_workplace']} | เงินเดือน {list_a[0]['salary']:,.0f} | เลขตำแหน่ง: {list_a[0]['position_no']} | คำสั่ง: {list_a[0]['order_ref']}"
            desc_b = "-"

        # กรณี 3: มีเฉพาะในเล่มเขียนมือ (Missing in HRMS)
        elif len_a == 0 and len_b > 0:
            status = "❌ ขาดในระบบอิเล็กทรอนิกส์ (ต้องบันทึก)"
            action = "นำเข้าข้อมูลคำสั่งนี้เข้าสู่ระบบ ก.พ.7 อิเล็กทรอนิกส์"
            stats["missing_in_hrms"] += 1
            desc_a = "-"
            desc_b = f"{list_b[0]['position_and_workplace']} | เงินเดือน {list_b[0]['salary']:,.0f} | เลขตำแหน่ง: {list_b[0]['position_no']} | คำสั่ง: {list_b[0]['order_ref']}"

        matched_rows.append({
            "วัน เดือน ปี (พ.ศ.)": norm_date,
            "อัตราเงินเดือน (บาท)": f"{sal:,.0f}" if sal > 0 else "-",
            "ก.พ.7 อิเล็กทรอนิกส์ (HRMS)": desc_a,
            "ก.ค.ศ.16 (เขียนมือ)": desc_b,
            "สถานะการตรวจสอบ": status,
            "สิ่งที่ต้องดำเนินการแก้ไข": action
        })

    return matched_rows, stats, inversions_a, inversions_b

# ==========================================
# 5. ฟังก์ชันสร้างรายงาน Excel (.xlsx)
# ==========================================
def generate_audit_excel(table_rows: List[Dict[str, Any]], stats: Dict[str, Any], inv_b: List[Dict[str, Any]]) -> io.BytesIO:
    wb = openpyxl.Workbook()
    
    # Sheet 1: ผลการเปรียบเทียบ
    ws = wb.active
    ws.title = "ผลการเทียบเคียง กพ7"
    
    headers = [
        "ลำดับ", "วัน เดือน ปี (พ.ศ.)", "อัตราเงินเดือน (บาท)", 
        "ก.พ.7 อิเล็กทรอนิกส์ (HRMS)", "ก.ค.ศ.16 (เขียนมือ)", 
        "สถานะการตรวจสอบ", "สิ่งที่ต้องดำเนินการแก้ไข"
    ]
    ws.append(headers)
    
    # Header Styling
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="TH Sarabun New", size=14, bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    fill_pass = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_warn = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_error = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    row_font = Font(name="TH Sarabun New", size=13)
    
    for idx, r in enumerate(table_rows, start=1):
        ws.append([
            idx,
            r["วัน เดือน ปี (พ.ศ.)"],
            r["อัตราเงินเดือน (บาท)"],
            r["ก.พ.7 อิเล็กทรอนิกส์ (HRMS)"],
            r["ก.ค.ศ.16 (เขียนมือ)"],
            r["สถานะการตรวจสอบ"],
            r["สิ่งที่ต้องดำเนินการแก้ไข"]
        ])
        
        row_cells = ws[idx + 1]
        for cell in row_cells:
            cell.font = row_font
            if "ตรงกันสมบูรณ์" in r["สถานะการตรวจสอบ"]:
                cell.fill = fill_pass
            elif "ซ้ำ" in r["สถานะการตรวจสอบ"]:
                cell.fill = fill_warn
            else:
                cell.fill = fill_error

    # Auto-adjust column widths
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 45
    ws.column_dimensions['E'].width = 45
    ws.column_dimensions['F'].width = 30
    ws.column_dimensions['G'].width = 35

    # Sheet 2: สรุปข้อผิดปกติและ Timeline Inversion
    ws_summary = wb.create_sheet(title="สรุปข้อบกพร่องและลำดับเวลา")
    ws_summary.append(["หัวข้อการตรวจสอบ", "จำนวนที่พบ", "หน่วย"])
    for cell in ws_summary[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    summary_data = [
        ["รายการที่ตรงกันสมบูรณ์", stats["perfect_match"], "รายการ"],
        ["รายการที่ขาดในเล่มเขียนมือ (ต้องเพิ่ม)", stats["missing_in_manual"], "รายการ"],
        ["รายการที่ขาดในระบบ ก.พ.7 (ต้องนำเข้า)", stats["missing_in_hrms"], "รายการ"],
        ["รายการซ้ำซ้อนในระบบ ก.พ.7", stats["duplicate_in_hrms"], "รายการ"],
        ["รายการวันที่ลงสลับลำดับในเล่มเขียนมือ", len(inv_b), "จุด"]
    ]
    
    for s_row in summary_data:
        ws_summary.append(s_row)
        for cell in ws_summary[ws_summary.max_row]:
            cell.font = row_font

    if inv_b:
        ws_summary.append([])
        ws_summary.append(["รายละเอียดจุดที่วันที่ลงสลับลำดับ (Timeline Inversion)"])
        ws_summary[ws_summary.max_row][0].font = Font(name="TH Sarabun New", size=14, bold=True, color="C00000")
        for item in inv_b:
            ws_summary.append([item["msg"]])
            ws_summary[ws_summary.max_row][0].font = row_font

    ws_summary.column_dimensions['A'].width = 50
    ws_summary.column_dimensions['B'].width = 15
    ws_summary.column_dimensions['C'].width = 15

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

# ==========================================
# 6. หน้าจอส่วนต่อประสาน (Streamlit UI)
# ==========================================
st.set_page_config(page_title="ระบบ Two-Way Full Outer Matching ก.พ.7", layout="wide")

st.title("⚖️ ระบบตรวจสอบและเทียบเคียง ก.พ.7 ด้วย Two-Way Full Outer Matching")
st.caption("ระบบวิเคราะห์ความสอดคล้องของทะเบียนประวัติข้าราชการครู สพป.มหาสารคาม เขต 2")

with st.sidebar:
    st.header("⚙️ การตั้งค่าระบบ")
    api_key_input = st.text_input("ใส่ Google Gemini API Key:", type="password")
    st.markdown("[👉 ขอรับ API Key ฟรีจาก Google AI Studio](https://aistudio.google.com/)")
    
    active_model = None
    if api_key_input:
        try:
            genai.configure(api_key=api_key_input)
            model_list = [
                m.name for m in genai.list_models() 
                if 'generateContent' in m.supported_generation_methods and 'gemini' in m.name.lower()
            ]
            if model_list:
                st.success(f"เชื่อมต่อสำเร็จ (พบ {len(model_list)} โมเดล)")
                default_index = 0
                for i, m_name in enumerate(model_list):
                    if "flash" in m_name.lower():
                        default_index = i
                        break
                active_model = st.selectbox("เลือกโมเดล VLM:", model_list, index=default_index)
            else:
                st.warning("ไม่พบโมเดล generateContent ในบัญชีนี้")
        except Exception as err:
            st.error(f"API Key ไม่ถูกต้อง: {str(err)}")
            
    st.divider()
    st.info("💡 **หลักการทำงานของอัลกอริทึม:**\n- ทำ Full Outer Join ด้วย Composite Key `(วันที่ + เงินเดือน)`\n- ดักจับแถวที่ซ้ำซ้อนในระบบ HRMS\n- ตรวจจับวันที่เขียนกระโดดย้อนหลังในเล่ม ก.ค.ศ.16")

col_left, col_right = st.columns(2)
with col_left:
    st.subheader("📄 ไฟล์ที่ 1: ก.พ.7 อิเล็กทรอนิกส์ (HRMS)")
    file_hrms = st.file_uploader("อัปโหลด PDF ทะเบียนประวัติจากระบบอิเล็กทรอนิกส์", type=["pdf"], key="file_hrms")

with col_right:
    st.subheader("✍️ ไฟล์ที่ 2: ก.ค.ศ.16 (เล่มเขียนมือ)")
    file_manual = st.file_uploader("อัปโหลด PDF สแกนสมุดทะเบียนประวัติเขียนมือ", type=["pdf"], key="file_manual")

if st.button("🚀 เริ่มการตรวจสอบเปรียบเทียบแบบ Two-Way Matching", type="primary"):
    if not api_key_input:
        st.error("กรุณาระบุ Google Gemini API Key ที่แถบด้านซ้าย")
    elif not active_model:
        st.error("กรุณาเลือกโมเดล AI ที่พร้อมใช้งาน")
    elif not file_hrms or not file_manual:
        st.error("กรุณาอัปโหลดไฟล์ PDF ให้ครบทั้ง 2 ไฟล์")
    else:
        with st.spinner(f"🔍 กำลังสกัดข้อมูลและรันอัลกอริทึม Two-Way Outer Match ด้วยโมเดล {active_model}..."):
            try:
                # 1. สกัดข้อมูลทั้ง 2 ไฟล์
                records_hrms = extract_pdf_records(file_hrms.read(), api_key_input, active_model, "ก.พ.7 อิเล็กทรอนิกส์")
                records_man = extract_pdf_records(file_manual.read(), api_key_input, active_model, "ก.ค.ศ.16 เขียนมือ")
                
                # 2. รันอัลกอริทึม Two-Way Full Outer Matching
                comp_results, stats_data, inv_hrms, inv_man = run_two_way_full_outer_matching(records_hrms, records_man)
                
                st.success("✅ การประมวลผลและการจับคู่ข้อมูลเสร็จสมบูรณ์!")
                
                # 3. แสดง Metrics การตรวจเช็ค
                m1, m2, m3, m4, m5 = st.columns(5)
                m1.metric("รายการตรงกันสมบูรณ์", f"{stats_data['perfect_match']} รายการ")
                m2.metric("ขาดในเล่มเขียนมือ", f"{stats_data['missing_in_manual']} รายการ", delta_color="inverse")
                m3.metric("ขาดในระบบ ก.พ.7", f"{stats_data['missing_in_hrms']} รายการ", delta_color="inverse")
                m4.metric("รายการซ้ำซ้อนในระบบ", f"{stats_data['duplicate_in_hrms']} รายการ", delta_color="inverse")
                m5.metric("วันที่เขียนสลับลำดับ", f"{len(inv_man)} จุด", delta_color="inverse")
                
                st.divider()

                # 4. แจ้งเตือนข้อผิดปกติพิเศษ
                if inv_man:
                    st.error(f"⚠️ **ตรวจพบลำดับวันที่สลับที่กัน (Timeline Inversion) ในไฟล์เขียนมือ {len(inv_man)} จุด:**")
                    for inv in inv_man:
                        st.write(f"- {inv['msg']}")
                        
                if stats_data['duplicate_in_hrms'] > 0:
                    st.warning(f"⚠️ **ตรวจพบรายการบันทึกซ้ำซ้อนในระบบ ก.พ.7 อิเล็กทรอนิกส์ {stats_data['duplicate_in_hrms']} เหตุการณ์** (แนะนำให้ตรวจสอบและลบรายการเบิ้ลซ้ำในระบบ HRMS)")

                # 5. แสดงผลตาราง Two-Way Full Outer Matching
                st.subheader("📊 ตารางแสดงผลการเปรียบเทียบข้อมูล (Reconciliation Table)")
                df_out = pd.DataFrame(comp_results)
                st.dataframe(df_out, use_container_width=True)
                
                # 6. ปุ่มดาวน์โหลดไฟล์รายงาน Excel
                excel_file = generate_audit_excel(comp_results, stats_data, inv_man)
                st.download_button(
                    label="📥 ดาวน์โหลดรายงานผลการตรวจสอบ (.xlsx)",
                    data=excel_file,
                    file_name=f"KP7_OuterMatch_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดระหว่างการประมวลผล: {str(e)}")
