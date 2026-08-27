import streamlit as st
import pandas as pd
import json
import re
import io
import os
import tempfile
from datetime import datetime
from pydantic import BaseModel, Field
from typing import List, Optional, Tuple, Dict, Any

# นำเข้า SDK ตัวใหม่ล่าสุดของ Google
from google import genai
from google.genai import types
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

# ==========================================
# 1. โครงสร้างข้อมูลมาตรฐาน (Strict Schema)
# ==========================================
class RecordEntry(BaseModel):
    date_raw: str = Field(default="", description="วันเดือนปี เช่น '1 เม.ย. 54'")
    position_and_workplace: str = Field(default="", description="ตำแหน่ง หน่วยงาน")
    position_no: Optional[str] = Field(default="", description="เลขที่ตำแหน่ง เช่น 5693, 3332 (สำคัญมาก ต้องดึงมาให้ได้)")
    academic_standing: Optional[str] = Field(default="", description="วิทยฐานะ เช่น คศ.1, คศ.2, คศ.3, ชำนาญการ, ชำนาญการพิเศษ")
    salary: Optional[float] = Field(default=0.0, description="อัตราเงินเดือนเป็นตัวเลขเท่านั้น")
    order_ref: Optional[str] = Field(default="", description="เลขที่คำสั่งและวันที่ลงนาม")

class KP7ExtractionResult(BaseModel):
    records: list[RecordEntry] = Field(default=[], description="รายการประวัติทั้งหมดเรียงตามลำดับในเอกสาร")

# ==========================================
# 2. ระบบทำความสะอาดและแปลงข้อมูล
# ==========================================
THAI_DIGITS = str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789")
THAI_MONTHS = {
    "ม.ค.": 1, "มค": 1, "มกราคม": 1, "ก.พ.": 2, "กพ": 2, "กุมภาพันธ์": 2,
    "มี.ค.": 3, "มีค": 3, "มีนาคม": 3, "เม.ย.": 4, "เมย": 4, "เมษายน": 4,
    "พ.ค.": 5, "พค": 5, "พฤษภาคม": 5, "มิ.ย.": 6, "มิย": 6, "มิถุนายน": 6,
    "ก.ค.": 7, "กค": 7, "กรกฎาคม": 7, "ส.ค.": 8, "สค": 8, "สิงหาคม": 8,
    "ก.ย.": 9, "กย": 9, "กันยายน": 9, "ต.ค.": 10, "ตค": 10, "ตุลาคม": 10,
    "พ.ย.": 11, "พย": 11, "พฤศจิกายน": 11, "ธ.ค.": 12, "ธค": 12, "ธันวาคม": 12
}
MONTH_LABEL = {
    1: "ม.ค.", 2: "ก.พ.", 3: "มี.ค.", 4: "เม.ย.", 5: "พ.ค.", 6: "มิ.ย.",
    7: "ก.ค.", 8: "ส.ค.", 9: "ก.ย.", 10: "ต.ค.", 11: "พ.ย.", 12: "ธ.ค."
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
    if sal_val is None: return 0.0
    s_str = str(sal_val).translate(THAI_DIGITS).replace(",", "").replace(" ", "").replace("บาท", "")
    match = re.search(r"(\d+(\.\d+)?)", s_str)
    if match:
        val = float(match.group(1))
        if val > 0 and val < 5000: return val * 10 
        return val
    return 0.0

def normalize_thai_date(date_str: str) -> Tuple[str, int]:
    if not date_str or not isinstance(date_str, str): return "-", 0
    clean_str = str(date_str).translate(THAI_DIGITS).replace(" ", "")
    pattern = r"(\d{1,2})([ก-๙\.]+)(\d{2,4})"
    match = re.search(pattern, clean_str)
    if not match: return date_str.strip(), 0
    day = int(match.group(1))
    month_raw = match.group(2)
    year_raw = int(match.group(3))
    year = 2500 + year_raw if year_raw < 100 else year_raw
    month = next((m_val for m_key, m_val in THAI_MONTHS.items() if m_key in month_raw), 0)
    if month == 0: return f"{day} {month_raw} {year}", (year * 10000) + day
    return f"{day} {MONTH_LABEL[month]} {year}", (year * 10000) + (month * 100) + day

# ==========================================
# 3. VLM Data Extractor (SDK ใหม่)
# ==========================================
def extract_pdf_records_precise(pdf_bytes: bytes, api_key: str, model_name: str, hint: str) -> List[Dict[str, Any]]:
    client = genai.Client(api_key=api_key)
    
    prompt = f"""
    คุณคือผู้เชี่ยวชาญการตรวจสอบทะเบียนประวัติ ก.พ.7 / ก.ค.ศ.16 สพป.มหาสารคาม เขต 2
    ประเภทเอกสาร: {hint}
    กฎเหล็ก: 
    1. สกัดข้อมูลประวัติรับเงินเดือนทุกแถว ห้ามข้าม 
    2. หา "เลขที่ตำแหน่ง" (มักเป็นเลข 4-6 หลัก เช่น 5693, 3332) ให้เจอและแยกไว้
    3. หา "วิทยฐานะ" (เช่น ชำนาญการ, ชำนาญการพิเศษ, คศ.1, คศ.2, คศ.3)
    4. ตัวเลขเงินเดือนอารบิก วันที่ตามจริง เลขคำสั่งครบถ้วน
    5. **ข้อควรระวังปี 2567-2568:** จะมีรายการ "ปรับชดเชยผู้ได้รับผลกระทบ" ในวันที่ "1 พ.ค. 67" และ "1 พ.ค. 68" แทรกเข้ามา ห้ามข้าม และให้อ่านวันที่เป็น 1 พ.ค. ตามจริง
    """
    
    temp_pdf_path = ""
    uploaded_file = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf.write(pdf_bytes)
            temp_pdf_path = temp_pdf.name
            
        uploaded_file = client.files.upload(file=temp_pdf_path)
        
        response = client.models.generate_content(
            model=model_name,
            contents=[uploaded_file, prompt],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=KP7ExtractionResult,
                temperature=0.0
            )
        )
        cleaned_str = clean_json_string(response.text)
        
    finally:
        if uploaded_file:
            try: client.files.delete(name=uploaded_file.name)
            except: pass
        if temp_pdf_path and os.path.exists(temp_pdf_path):
            try: os.remove(temp_pdf_path)
            except: pass
            
    try:
        data = json.loads(cleaned_str)
        records = data.get("records", [])
    except:
        records = []
        
    extracted_rows = []
    for idx, r in enumerate(records):
        norm_date, s_key = normalize_thai_date(r.get("date_raw", ""))
        r["normalized_date"] = norm_date
        r["sort_key"] = s_key
        r["salary"] = sanitize_salary(r.get("salary", 0))
        r["original_index"] = idx # เก็บลำดับดั้งเดิมไว้เช็ค Milestone
        extracted_rows.append(r)
        
    return extracted_rows

# ==========================================
# 4. Smart Reconciliation (ยึดเงินเดือน และตรวจ Milestone)
# ==========================================
def format_milestone_desc(record: dict) -> str:
    desc = f"{record.get('position_and_workplace', '')} "
    pos_no = record.get('position_no', '')
    acad = record.get('academic_standing', '')
    
    tags = []
    if pos_no: tags.append(f"เลขตำแหน่ง: {pos_no}")
    if acad: tags.append(f"วิทยฐานะ: {acad}")
    
    tag_str = f"[{' | '.join(tags)}]" if tags else ""
    return f"{desc.strip()} {tag_str} (เงินเดือน {record['salary']:,.0f} บ.) [{record.get('order_ref', '')}]"

def run_two_way_reconciliation(records_a: List[Dict[str, Any]], records_b: List[Dict[str, Any]]):
    # ตรวจหา Milestone ใน HRMS (เป็น Ground Truth ด้านลำดับเวลา)
    for i in range(1, len(records_a)):
        prev, curr = records_a[i-1], records_a[i]
        curr["is_transfer"] = (curr.get("position_no") and prev.get("position_no") and curr.get("position_no") != prev.get("position_no"))
        curr["is_promotion"] = (curr.get("academic_standing") and prev.get("academic_standing") and curr.get("academic_standing") != prev.get("academic_standing"))

    matched_rows = []
    stats = {"perfect_match": 0, "duplicate_in_hrms": 0, "missing_in_manual": 0, "missing_in_hrms": 0, "salary_mismatch": 0}
    used_b_indices = set()

    # 1. จับคู่โดยใช้ "ยอดเงินเดือน" เป็น Anchor หลัก (แก้ปัญหาเขียนมือวันที่สลับ)
    for r_a in records_a:
        date_str = r_a["normalized_date"]
        matched_b_idx = None
        
        # ค้นหาในฝั่ง B (เขียนมือ) ที่เงินเดือนตรงกัน และยังไม่ถูกจับคู่
        for idx_b, r_b in enumerate(records_b):
            if idx_b not in used_b_indices and abs(r_a["salary"] - r_b["salary"]) < 1.0:
                matched_b_idx = idx_b
                break
                
        desc_a = format_milestone_desc(r_a)
        flag_msg = ""
        if r_a.get("is_transfer"): flag_msg += f" 🚩 ย้าย (เลข {r_a.get('position_no')})"
        if r_a.get("is_promotion"): flag_msg += f" 🌟 ปรับวิทยฐานะ ({r_a.get('academic_standing')})"
        
        # ดักจับคำสั่งปรับชดเชยมติ ครม. 67-68 (แก้ไขการจัดหน้าบรรทัดแล้ว)
        desc_text = r_a.get("position_and_workplace", "")
        is_compensation = "ชดเชย" in desc_text or "ปรับอัตรา" in desc_text
        if is_compensation or ("1 พ.ค. 2567" in date_str or "1 พ.ค. 2568" in date_str):
            flag_msg += " 💰 ปรับชดเชยมติ ครม."
        
        if matched_b_idx is not None:
            used_b_indices.add(matched_b_idx)
            r_b = records_b[matched_b_idx]
            desc_b = format_milestone_desc(r_b)
            
            # ถ้ายอดเงินเดือนตรง แต่วันที่ต่างกัน (เกิดจากการจดย้อนหลัง)
            if r_a["normalized_date"] != r_b["normalized_date"]:
                status = "⚠️ เงินเดือนตรง แต่วันที่เขียนมือคลาดเคลื่อน"
                action = f"ยึดวันที่ HRMS ({r_a['normalized_date']}) เป็นหลัก"
            else:
                status = "✅ ตรงกันสมบูรณ์" + flag_msg
                action = "-"
                stats["perfect_match"] += 1
        else:
            desc_b = "-"
            status = "❌ ขาดในเล่มเขียนมือ (หรือ AI อ่านเงินเดือนไม่ออก)" + flag_msg
            action = "เพิ่มรายการลงสมุด ก.ค.ศ.16 หรือตรวจสอบเลขคำสั่ง"
            stats["missing_in_manual"] += 1

        matched_rows.append({
            "วัน เดือน ปี (HRMS เป็นหลัก)": date_str, 
            "เงินเดือน": f"{r_a['salary']:,.0f}", 
            "ข้อมูล ก.พ.7 อิเล็กทรอนิกส์": desc_a, 
            "ข้อมูล ก.ค.ศ.16 เขียนมือ": desc_b, 
            "สถานะการตรวจสอบ": status, 
            "สิ่งที่ต้องดำเนินการแก้ไข": action
        })

    # 2. กวาดรายการที่เหลือในเขียนมือ (รายการก่อนปี 54 หรือรายการที่ตกหล่นใน HRMS)
    for idx_b, r_b in enumerate(records_b):
        if idx_b not in used_b_indices:
            matched_rows.append({
                "วัน เดือน ปี (HRMS เป็นหลัก)": r_b["normalized_date"], 
                "เงินเดือน": f"{r_b['salary']:,.0f}", 
                "ข้อมูล ก.พ.7 อิเล็กทรอนิกส์": "-", 
                "ข้อมูล ก.ค.ศ.16 เขียนมือ": format_milestone_desc(r_b), 
                "สถานะการตรวจสอบ": "❌ ขาดในระบบอิเล็กทรอนิกส์ (HRMS)", 
                "สิ่งที่ต้องดำเนินการแก้ไข": "คีย์ข้อมูลคำสั่งย้อนหลังเข้าสู่ระบบ e-KP7"
            })
            stats["missing_in_hrms"] += 1

    # เรียงลำดับผลลัพธ์ตามวันที่ (คร่าวๆ) เพื่อให้อ่านง่าย
    def get_sort_val(row):
        try:
            pts = row["วัน เดือน ปี (HRMS เป็นหลัก)"].split()
            return int(pts[2]) * 10000 + THAI_MONTHS.get(pts[1], 0) * 100 + int(pts[0])
        except:
            return 99999999
            
    matched_rows = sorted(matched_rows, key=get_sort_val)
    return matched_rows, stats, []

def generate_audit_excel(table_rows, stats, inv_b) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "ผลการเทียบเคียง กพ7"
    ws.append(["ลำดับ", "วัน เดือน ปี (HRMS เป็นหลัก)", "เงินเดือน", "ก.พ.7 อิเล็กทรอนิกส์ (HRMS)", "ก.ค.ศ.16 (เขียนมือ)", "สถานะการตรวจสอบ", "สิ่งที่ต้องดำเนินการแก้ไข"])
    
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="TH Sarabun New", size=14, bold=True, color="FFFFFF")
    for cell in ws[1]: cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal="center", vertical="center")
        
    fill_pass = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_warn = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_error = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    row_font = Font(name="TH Sarabun New", size=13)
    
    for idx, r in enumerate(table_rows, 1):
        ws.append([idx, r["วัน เดือน ปี (HRMS เป็นหลัก)"], r["เงินเดือน"], r["ข้อมูล ก.พ.7 อิเล็กทรอนิกส์"], r["ข้อมูล ก.ค.ศ.16 เขียนมือ"], r["สถานะการตรวจสอบ"], r["สิ่งที่ต้องดำเนินการแก้ไข"]])
        for cell in ws[idx + 1]:
            cell.font = row_font
            cell.fill = fill_pass if "ตรงกันสมบูรณ์" in r["สถานะการตรวจสอบ"] else (fill_warn if "⚠️" in r["สถานะการตรวจสอบ"] else fill_error)

    ws.column_dimensions['B'].width, ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width, ws.column_dimensions['E'].width = 45
    ws.column_dimensions['F'].width, ws.column_dimensions['G'].width = 30

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

# ==========================================
# 5. Streamlit UI
# ==========================================
st.set_page_config(page_title="ระบบตรวจสอบความถูกต้อง ก.พ.7", layout="wide")
st.title("🎯 ระบบตรวจสอบและเทียบเคียง ก.พ.7 / ก.ค.ศ.16 (Gemini API ตัวใหม่ล่าสุด)")
st.caption("ประมวลผลความเร็วสูงด้วย Native PDF SDK (google-genai) - สพป.มหาสารคาม เขต 2")

with st.sidebar:
    st.header("⚙️ การตั้งค่าระบบ")
    api_key_input = st.text_input("ใส่ Google Gemini API Key:").strip()
    
    model_list = ["gemini-3.7-flash", "gemini-3.6-flash"]
    active_model = st.selectbox("เลือกโมเดล VLM:", model_list, index=0)
    st.info("แนะนำให้สร้าง API Key จาก Gmail บัญชีใหม่เพื่อหลีกเลี่ยงปัญหา Error 400 ครับ")

col1, col2 = st.columns(2)
with col1:
    st.subheader("📄 ไฟล์ 1: ก.พ.7 อิเล็กทรอนิกส์ (HRMS)")
    file_hrms = st.file_uploader("อัปโหลด PDF จากระบบ", type=["pdf"], key="file_hrms")
with col2:
    st.subheader("✍️ ไฟล์ 2: ก.ค.ศ.16 (เขียนมือ)")
    file_manual = st.file_uploader("อัปโหลด PDF สแกน", type=["pdf"], key="file_manual")

if st.button("🚀 เริ่มประมวลผล (ความเร็วสูง)", type="primary"):
    if not api_key_input or not active_model or not file_hrms or not file_manual:
        st.error("กรุณาใส่ API Key, เลือกโมเดล และอัปโหลดไฟล์ให้ครบ")
    else:
        status_box = st.status("🔍 กำลังประมวลผลผ่าน Google GenAI SDK...", expanded=True)
        try:
            status_box.write("📄 1/2 อ่าน ก.พ.7 อิเล็กทรอนิกส์...")
            records_hrms = extract_pdf_records_precise(file_hrms.read(), api_key_input, active_model, "ก.พ.7 อิเล็กทรอนิกส์")
            
            status_box.write("✍️ 2/2 อ่าน ก.ค.ศ.16 เขียนมือ...")
            records_man = extract_pdf_records_precise(file_manual.read(), api_key_input, active_model, "ก.ค.ศ.16 เขียนมือ")
            
            status_box.write("⚖️ กำลังเทียบเคียงข้อมูลและดักจับ Milestone...")
            comp_results, stats_data, inv_man = run_two_way_reconciliation(records_hrms, records_man)
            
            status_box.update(label="✅ ตรวจสอบเสร็จสมบูรณ์!", state="complete", expanded=False)
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("ตรงกันสมบูรณ์", f"{stats_data['perfect_match']} รายการ")
            m2.metric("ขาดในเล่มเขียนมือ", f"{stats_data['missing_in_manual']} รายการ")
            m3.metric("ขาดในระบบ HRMS", f"{stats_data['missing_in_hrms']} รายการ")
            m4.metric("รอการตรวจสอบเพิ่มเติม", f"{len(comp_results) - stats_data['perfect_match']} จุด")

            st.subheader("📊 ตารางเปรียบเทียบข้อมูล (Smart Reconciliation)")
            st.dataframe(pd.DataFrame(comp_results), use_container_width=True)
            
            excel_file = generate_audit_excel(comp_results, stats_data, inv_man)
            st.download_button(
                label="📥 ดาวน์โหลดรายงานผล (.xlsx)",
                data=excel_file,
                file_name=f"KP7_Audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            status_box.update(label="❌ เกิดข้อผิดพลาด", state="error", expanded=True)
            st.error(f"รายละเอียด: {str(e)}")
