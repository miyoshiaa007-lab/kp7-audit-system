import streamlit as st
import pandas as pd
import json
import re
import io
from datetime import datetime
from pydantic import BaseModel, Field
from typing import List, Optional, Tuple, Dict, Any
import google.generativeai as genai
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

# ==========================================
# 1. โครงสร้างข้อมูลมาตรฐาน (Strict Pydantic Schema)
# ==========================================
class RecordEntry(BaseModel):
    date_raw: str = Field(default="", description="วันเดือนปี เช่น '1 เม.ย. 54' หรือ '1 เม.ย. 2554'")
    position_and_workplace: str = Field(default="", description="ตำแหน่ง หน่วยงาน วิทยฐานะ หรือการเลื่อนขั้น")
    position_no: Optional[str] = Field(default="", description="เลขที่ตำแหน่ง เช่น '5693', '3332'")
    academic_standing: Optional[str] = Field(default="", description="วิทยฐานะ เช่น ชำนาญการ, ชำนาญการพิเศษ")
    salary: Optional[float] = Field(default=0.0, description="อัตราเงินเดือนเป็นตัวเลขเท่านั้น (เช่น 25190, 33800)")
    order_ref: Optional[str] = Field(default="", description="เลขที่คำสั่งและวันที่ลงนาม")

class KP7ExtractionResult(BaseModel):
    records: List[RecordEntry] = Field(default=[], description="รายการประวัติทั้งหมดเรียงตามลำดับในเอกสาร")

# ==========================================
# 2. ระบบทำความสะอาดและแปลงข้อมูล (Sanitizer & Normalizer)
# ==========================================
THAI_DIGITS = str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789")

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
    text = raw_text.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        if len(lines) > 2 and lines[-1].strip().startswith("```"):
            text = "\n".join(lines[1:-1])
        elif len(lines) > 1:
            text = "\n".join(lines[1:])
    return text.strip()

def sanitize_salary(sal_val: Any) -> float:
    """แปลงและตรวจสอบความถูกต้องของจำนวนเงินเดือน"""
    if sal_val is None:
        return 0.0
    s_str = str(sal_val).translate(THAI_DIGITS).replace(",", "").replace(" ", "").replace("บาท", "")
    match = re.search(r"(\d+(\.\d+)?)", s_str)
    if match:
        val = float(match.group(1))
        # ดักจับกรณีค่าเพี้ยนหลุดช่วงเงินเดือนข้าราชการครู
        if val > 0 and val < 5000:  # เช่น 25190 ถูกอ่านตกเป็น 2519
            return val * 10  # ปรับแก้ความผิดพลาดจากการอ่านเลขตกหล่น
        return val
    return 0.0

def normalize_thai_date(date_str: str) -> Tuple[str, int]:
    """แปลงวันที่ภาษาไทยให้อยู่ในมาตรฐานเดียวกันและสร้างตัวเลขเรียงลำดับ"""
    if not date_str or not isinstance(date_str, str):
        return "-", 0
    
    clean_str = str(date_str).translate(THAI_DIGITS).replace(" ", "").replace(".", ". ")
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

from pdf2image import convert_from_bytes
# ==========================================
# 3. VLM Data Extractor Engine (High Precision & Anti-Hang)
# ==========================================
def extract_pdf_records_precise(pdf_bytes: bytes, api_key: str, model_name: str, hint: str) -> List[Dict[str, Any]]:
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
    
    prompt = f"""
    คุณคือผู้เชี่ยวชาญการตรวจสอบทะเบียนประวัติ ก.พ.7 / ก.ค.ศ.16 สพป.มหาสารคาม เขต 2
    ประเภทเอกสาร: {hint}
    
    กฎเหล็กเพื่อป้องกันข้อมูลเพี้ยน (100% Accuracy Rules):
    1. สกัดข้อมูลประวัติการรับเงินเดือนทุกแถว ทุกหน้า ห้ามข้ามแม้แต่บรรทัดเดียว
    2. ตัวเลขเงินเดือน (salary): ต้องเป็นตัวเลขอารบิกที่ถูกต้องตามเอกสาร ระวังการสับสนระหว่างเลข 3 กับ 8, เลข 0 กับ 6, เลข 1 กับ 7
    3. วันเดือนปี (date_raw): ถอดตามที่ปรากฏจริง เช่น '1 เม.ย. 54', '1 ต.ค. 2568'
    4. เอกสารอ้างอิง (order_ref): เลขที่คำสั่ง วันที่ลงนาม ให้สกัดมาอย่างครบถ้วน
    """
    
    # 1. แปลง PDF เป็นภาพโดยลดความละเอียดลงเล็กน้อย (150 DPI) เพื่อให้ไฟล์เบาและส่งผ่านเน็ตได้เร็ว
    images = convert_from_bytes(pdf_bytes, dpi=150)
    
    # 2. มัดรวมคำสั่งและรูปภาพ "ทุกหน้า" ส่งให้ AI ประมวลผลรวดเดียวจบ (ไม่ติดลูปค้าง)
    payload = [prompt] + images
    resp = model.generate_content(payload)
    
    # 3. ทำความสะอาดและแปลงผลลัพธ์
    cleaned_str = clean_json_string(resp.text)
    
    try:
        data = json.loads(cleaned_str)
        records = data.get("records", [])
    except Exception:
        records = []
        
    extracted_rows = []
    for r in records:
        norm_date, s_key = normalize_thai_date(r.get("date_raw", ""))
        r["normalized_date"] = norm_date
        r["sort_key"] = s_key
        r["salary"] = sanitize_salary(r.get("salary", 0))
        extracted_rows.append(r)
        
    return extracted_rows

# ==========================================
# 4. Advanced Two-Way Reconciliation (Fuzzy Discrepancy Detection)
# ==========================================
def run_two_way_reconciliation(records_a: List[Dict[str, Any]], records_b: List[Dict[str, Any]]):
    # ตรวจสอบลำดับเวลาสลับที่ในเล่มเขียนมือ
    inversions_b = []
    for i in range(1, len(records_b)):
        prev = records_b[i - 1]
        curr = records_b[i]
        if curr["sort_key"] > 0 and prev["sort_key"] > 0:
            if curr["sort_key"] < prev["sort_key"]:
                inversions_b.append({
                    "row": i + 1,
                    "date_curr": curr.get("date_raw", "-"),
                    "date_prev": prev.get("date_raw", "-"),
                    "msg": f"แถวที่ {i+1}: ลงวันที่ '{curr.get('date_raw', '-')}' อยู่ถัดจาก '{prev.get('date_raw', '-')}' (ลำดับเวลาย้อนกลับ)"
                })

    # รวบรวมวันที่ทั้งหมด
    all_dates = {}
    for r in records_a:
        all_dates.setdefault(r["normalized_date"], {"a": [], "b": [], "sort_key": r["sort_key"]})["a"].append(r)
    for r in records_b:
        if r["normalized_date"] not in all_dates:
            all_dates[r["normalized_date"]] = {"a": [], "b": [], "sort_key": r["sort_key"]}
        all_dates[r["normalized_date"]]["b"].append(r)

    # เรียงลำดับตามวันที่
    sorted_dates = sorted(all_dates.items(), key=lambda x: x[1]["sort_key"])

    matched_rows = []
    stats = {
        "perfect_match": 0,
        "duplicate_in_hrms": 0,
        "missing_in_manual": 0,
        "missing_in_hrms": 0,
        "salary_mismatch": 0
    }

    for date_str, group in sorted_dates:
        list_a = group["a"]
        list_b = group["b"]
        
        # จับคู่กรณีมีข้อมูลทั้ง 2 ฝั่งในวันเดียวกัน
        if list_a and list_b:
            used_b_indices = set()
            for r_a in list_a:
                matched_b_idx = None
                for idx_b, r_b in enumerate(list_b):
                    if idx_b not in used_b_indices and abs(r_a["salary"] - r_b["salary"]) < 1.0:
                        matched_b_idx = idx_b
                        break
                
                if matched_b_idx is not None:
                    used_b_indices.add(matched_b_idx)
                    r_b = list_b[matched_b_idx]
                    status = "✅ ตรงกันสมบูรณ์"
                    action = "-"
                    stats["perfect_match"] += 1
                    desc_b = f"{r_b['position_and_workplace']} (เงินเดือน {r_b['salary']:,.0f} บ.) [{r_b['order_ref']}]"
                else:
                    # วันที่ตรงกันแต่เงินเดือนอ่านต่างกัน (จับคู่แบบ Mismatch เพื่อแจ้งเตือนคนตรวจ)
                    if len(used_b_indices) < len(list_b):
                        for idx_b, r_b in enumerate(list_b):
                            if idx_b not in used_b_indices:
                                matched_b_idx = idx_b
                                break
                        used_b_indices.add(matched_b_idx)
                        r_b = list_b[matched_b_idx]
                        status = "⚠️ วันที่ตรงกันแต่ยอดเงินเดือนไม่ตรงกัน"
                        action = f"ก.พ.7 ยอด {r_a['salary']:,.0f} บ. vs เขียนมือยอด {r_b['salary']:,.0f} บ. (กรุณาตรวจทานเอกสารจริง)"
                        stats["salary_mismatch"] += 1
                        desc_b = f"{r_b['position_and_workplace']} (เงินเดือน {r_b['salary']:,.0f} บ.) [{r_b['order_ref']}]"
                    else:
                        status = "❌ ขาดในเล่มเขียนมือ (หรือเป็นรายการซ้ำในระบบ)"
                        action = "ตรวจเช็คการบันทึกซ้ำใน ก.พ.7 หรือเพิ่มลงเล่มเขียนมือ"
                        stats["duplicate_in_hrms"] += 1
                        desc_b = "-"

                desc_a = f"{r_a['position_and_workplace']} (เงินเดือน {r_a['salary']:,.0f} บ.) [{r_a['order_ref']}]"
                matched_rows.append({
                    "วัน เดือน ปี (พ.ศ.)": date_str,
                    "เงินเดือน ก.พ.7": f"{r_a['salary']:,.0f}" if r_a['salary'] > 0 else "-",
                    "ข้อมูล ก.พ.7 อิเล็กทรอนิกส์": desc_a,
                    "ข้อมูล ก.ค.ศ.16 เขียนมือ": desc_b,
                    "สถานะการตรวจสอบ": status,
                    "สิ่งที่ต้องดำเนินการแก้ไข": action
                })

            # แถวเขียนมือที่เหลือและไม่ถูกจับคู่
            for idx_b, r_b in enumerate(list_b):
                if idx_b not in used_b_indices:
                    desc_b = f"{r_b['position_and_workplace']} (เงินเดือน {r_b['salary']:,.0f} บ.) [{r_b['order_ref']}]"
                    matched_rows.append({
                        "วัน เดือน ปี (พ.ศ.)": date_str,
                        "เงินเดือน ก.พ.7": "-",
                        "ข้อมูล ก.พ.7 อิเล็กทรอนิกส์": "-",
                        "ข้อมูล ก.ค.ศ.16 เขียนมือ": desc_b,
                        "สถานะการตรวจสอบ": "❌ ขาดในระบบอิเล็กทรอนิกส์",
                        "สิ่งที่ต้องดำเนินการแก้ไข": "นำเข้าข้อมูลคำสั่งนี้เข้าสู่ระบบ ก.พ.7"
                    })
                    stats["missing_in_hrms"] += 1

        elif list_a and not list_b:
            for r_a in list_a:
                desc_a = f"{r_a['position_and_workplace']} (เงินเดือน {r_a['salary']:,.0f} บ.) [{r_a['order_ref']}]"
                matched_rows.append({
                    "วัน เดือน ปี (พ.ศ.)": date_str,
                    "เงินเดือน ก.พ.7": f"{r_a['salary']:,.0f}" if r_a['salary'] > 0 else "-",
                    "ข้อมูล ก.พ.7 อิเล็กทรอนิกส์": desc_a,
                    "ข้อมูล ก.ค.ศ.16 เขียนมือ": "-",
                    "สถานะการตรวจสอบ": "❌ ขาดในเล่มเขียนมือ (ต้องเพิ่ม)",
                    "สิ่งที่ต้องดำเนินการแก้ไข": "เพิ่มรายการคำสั่งนี้ลงในสมุด ก.ค.ศ.16 เขียนมือ"
                })
                stats["missing_in_manual"] += 1

        elif not list_a and list_b:
            for r_b in list_b:
                desc_b = f"{r_b['position_and_workplace']} (เงินเดือน {r_b['salary']:,.0f} บ.) [{r_b['order_ref']}]"
                matched_rows.append({
                    "วัน เดือน ปี (พ.ศ.)": date_str,
                    "เงินเดือน ก.พ.7": "-",
                    "ข้อมูล ก.พ.7 อิเล็กทรอนิกส์": "-",
                    "ข้อมูล ก.ค.ศ.16 เขียนมือ": desc_b,
                    "สถานะการตรวจสอบ": "❌ ขาดในระบบอิเล็กทรอนิกส์ (ต้องบันทึก)",
                    "สิ่งที่ต้องดำเนินการแก้ไข": "นำเข้าข้อมูลคำสั่งนี้เข้าสู่ระบบ ก.พ.7 อิเล็กทรอนิกส์"
                })
                stats["missing_in_hrms"] += 1

    return matched_rows, stats, inversions_b

# ==========================================
# 5. ส่งออกรายงาน Excel (.xlsx)
# ==========================================
def generate_audit_excel(table_rows: List[Dict[str, Any]], stats: Dict[str, Any], inv_b: List[Dict[str, Any]]) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ผลการเทียบเคียง กพ7"
    
    headers = [
        "ลำดับ", "วัน เดือน ปี (พ.ศ.)", "เงินเดือน ก.พ.7", 
        "ก.พ.7 อิเล็กทรอนิกส์ (HRMS)", "ก.ค.ศ.16 (เขียนมือ)", 
        "สถานะการตรวจสอบ", "สิ่งที่ต้องดำเนินการแก้ไข"
    ]
    ws.append(headers)
    
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
            r["เงินเดือน ก.พ.7"],
            r["ข้อมูล ก.พ.7 อิเล็กทรอนิกส์"],
            r["ข้อมูล ก.ค.ศ.16 เขียนมือ"],
            r["สถานะการตรวจสอบ"],
            r["สิ่งที่ต้องดำเนินการแก้ไข"]
        ])
        
        row_cells = ws[idx + 1]
        for cell in row_cells:
            cell.font = row_font
            if "ตรงกันสมบูรณ์" in r["สถานะการตรวจสอบ"]:
                cell.fill = fill_pass
            elif "⚠️" in r["สถานะการตรวจสอบ"]:
                cell.fill = fill_warn
            else:
                cell.fill = fill_error

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 45
    ws.column_dimensions['E'].width = 45
    ws.column_dimensions['F'].width = 30
    ws.column_dimensions['G'].width = 35

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

# ==========================================
# 6. Streamlit UI
# ==========================================
st.set_page_config(page_title="ระบบตรวจสอบความถูกต้อง ก.พ.7", layout="wide")

st.title("🎯 ระบบตรวจสอบและเทียบเคียง ก.พ.7 / ก.ค.ศ.16 ความแม่นยำสูง")
st.caption("ระบบสกัดข้อมูลด้วย VLM พร้อมอัลกอริทึมตรวจสอบความถูกต้องเชิงตรรกะ - สพป.มหาสารคาม เขต 2")

with st.sidebar:
    st.header("⚙️ การตั้งค่าระบบ")
    api_key_input = st.text_input("ใส่ Google Gemini API Key:", type="password")
    
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
                    if "3.7" in m_name.lower() or "flash" in m_name.lower():
                        default_index = i
                        break
                active_model = st.selectbox("เลือกโมเดล VLM:", model_list, index=default_index)
        except Exception as err:
            st.error(f"API Key ไม่ถูกต้อง: {str(err)}")

col_left, col_right = st.columns(2)
with col_left:
    st.subheader("📄 ไฟล์ที่ 1: ก.พ.7 อิเล็กทรอนิกส์ (HRMS)")
    file_hrms = st.file_uploader("อัปโหลด PDF จากระบบอิเล็กทรอนิกส์", type=["pdf"], key="file_hrms")

with col_right:
    st.subheader("✍️ ไฟล์ที่ 2: ก.ค.ศ.16 (เล่มเขียนมือ)")
    file_manual = st.file_uploader("อัปโหลด PDF สแกนสมุดเขียนมือ", type=["pdf"], key="file_manual")

if st.button("🚀 เริ่มการประมวลผลและตรวจสอบความแม่นยำ", type="primary"):
    if not api_key_input:
        st.error("กรุณาระบุ Google Gemini API Key")
    elif not active_model:
        st.error("กรุณาเลือกโมเดล AI ที่พร้อมใช้งาน")
    elif not file_hrms or not file_manual:
        st.error("กรุณาอัปโหลดไฟล์ PDF ให้ครบทั้ง 2 ไฟล์")
    else:
        status_box = st.status("🔍 กำลังประมวลผลด้วยระบบตรวจสอบความถูกต้องหลายชั้น...", expanded=True)
        try:
            status_box.write("📄 1/2 กำลังอ่านและตรวจสอบไฟล์ ก.พ.7 อิเล็กทรอนิกส์...")
            records_hrms = extract_pdf_records_precise(file_hrms.read(), api_key_input, active_model, "ก.พ.7 อิเล็กทรอนิกส์")
            
            status_box.write("✍️ 2/2 กำลังอ่านและถอดรหัสลายมือ ก.ค.ศ.16...")
            records_man = extract_pdf_records_precise(file_manual.read(), api_key_input, active_model, "ก.ค.ศ.16 เขียนมือ")
            
            status_box.write("⚖️ รันอัลกอริทึมตรวจสอบตรรกะความถูกต้องและจับคู่ข้อมูล...")
            comp_results, stats_data, inv_man = run_two_way_reconciliation(records_hrms, records_man)
            
            status_box.update(label="✅ ตรวจสอบความถูกต้องเสร็จสมบูรณ์!", state="complete", expanded=False)
            
            # แสดง Metrics การตรวจเช็ค
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("รายการตรงกันสมบูรณ์", f"{stats_data['perfect_match']} รายการ")
            m2.metric("ขาดในเล่มเขียนมือ", f"{stats_data['missing_in_manual']} รายการ")
            m3.metric("ขาดในระบบ ก.พ.7", f"{stats_data['missing_in_hrms']} รายการ")
            m4.metric("ยอดเงินเดือนไม่ตรงกัน", f"{stats_data['salary_mismatch']} รายการ")
            m5.metric("วันที่เขียนสลับลำดับ", f"{len(inv_man)} จุด")
            
            st.divider()

            if inv_man:
                st.error(f"⚠️ **ตรวจพบลำดับวันที่สลับที่กัน (Timeline Inversion) ในไฟล์เขียนมือ {len(inv_man)} จุด:**")
                for inv in inv_man:
                    st.write(f"- {inv['msg']}")

            # แสดงตารางผลลัพธ์
            st.subheader("📊 ตารางแสดงผลการเปรียบเทียบข้อมูล (Audit Reconciliation Table)")
            df_out = pd.DataFrame(comp_results)
            st.dataframe(df_out, use_container_width=True)
            
            # ปุ่มดาวน์โหลด Excel
            excel_file = generate_audit_excel(comp_results, stats_data, inv_man)
            st.download_button(
                label="📥 ดาวน์โหลดรายงานผลการตรวจสอบ (.xlsx)",
                data=excel_file,
                file_name=f"KP7_Precision_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            status_box.update(label="❌ เกิดข้อผิดพลาด", state="error", expanded=True)
            st.error(f"รายละเอียดข้อผิดพลาด: {str(e)}")
