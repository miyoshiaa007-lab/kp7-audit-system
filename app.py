import streamlit as st
import pandas as pd
import json
import re
import io
import os
from datetime import datetime
from pydantic import BaseModel, Field
from typing import List, Optional
import google.generativeai as genai
from PIL import Image
from pdf2image import convert_from_bytes
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ==========================================
# 1. โครงสร้างข้อมูลมาตรฐาน (Pydantic Schema)
# ==========================================
class RecordEntry(BaseModel):
    date_raw: str = Field(description="วันเดือนปีที่ปรากฏในเอกสาร เช่น '1 เม.ย. 54' หรือ '1 เม.ย. 2554'")
    position_and_workplace: str = Field(description="ตำแหน่ง/หน่วยงาน/วิทยฐานะ/การเลื่อนขั้น")
    position_no: Optional[str] = Field(default="", description="เลขที่ตำแหน่ง เช่น '5693', '3332'")
    academic_standing: Optional[str] = Field(default="", description="วิทยฐานะ เช่น 'ชำนาญการ', 'ชำนาญการพิเศษ'")
    salary: Optional[float] = Field(default=0.0, description="อัตราเงินเดือน (ตัวเลขเท่านั้น ไม่ใส่ลูกน้ำ) เช่น 25190, 33800")
    order_ref: Optional[str] = Field(default="", description="เอกสารอ้างอิง/คำสั่ง เช่น 'คส.สพป.มค.2 ที่ 199/54 ลว. 18 เม.ย. 54'")

class KP7ExtractionResult(BaseModel):
    records: List[RecordEntry] = Field(description="รายการประวัติการรับเงินเดือนและตำแหน่งทั้งหมด เรียงตามลำดับที่ปรากฏในหน้าเอกสาร")

# ==========================================
# 2. ฟังก์ชันช่วยแปลงและเตรียมข้อมูล (Normalization Engine)
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

def normalize_thai_date(date_str: str):
    if not date_str or not isinstance(date_str, str):
        return "-", 0
    
    clean_str = date_str.replace(" ", "").replace(".", ". ")
    pattern = r"(\d{1,2})\s*([ก-๙\.]+)\s*(\d{2,4})"
    match = re.search(pattern, clean_str)
    
    if not match:
        return date_str, 0
    
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
        return date_str, year * 10000 + day
        
    date_formatted = f"{day} {list(THAI_MONTHS.keys())[(month-1)*3]} {year}"
    sort_key = (year * 10000) + (month * 100) + day
    return date_formatted, sort_key

# ==========================================
# 3. VLM Data Extraction Engine (ระบบสกัดข้อมูลยืดหยุ่นสูง)
# ==========================================
def extract_data_from_pdf(pdf_bytes: bytes, api_key: str, model_name: str, file_type_hint: str) -> List[dict]:
    genai.configure(api_key=api_key)
    
    # กำหนดค่าโมเดล
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
        # Fallback กรณีโมเดลบางรุ่นไม่รองรับ response_schema
        model = genai.GenerativeModel(
            model_name=model_name,
            generation_config={"temperature": 0.0}
        )
    
    images = convert_from_bytes(pdf_bytes, dpi=200)
    all_records = []
    
    prompt = f"""
    คุณคือผู้เชี่ยวชาญการอ่านและวิเคราะห์ทะเบียนประวัติ ก.พ.7 / ก.ค.ศ.16 ของข้าราชการครูไทย (สพป.มหาสารคาม เขต 2)
    ประเภทเอกสาร: {file_type_hint}
    
    คำสั่งการทำงาน:
    1. สกัดข้อมูลประวัติการดำรงตำแหน่งและอัตราเงินเดือนทุกแถว (Row) ให้ครบถ้วน 100% ห้ามข้ามแม้แต่แถวเดียว
    2. สำหรับเอกสารเขียนมือ ให้อ่านลายมือภาษาไทย ตัวเลขไทย/อารบิก และสัญลักษณ์ย่ออย่างแม่นยำ
    3. ส่งผลลัพธ์เป็น JSON ในรูปแบบนี้เท่านั้น:
    {{
      "records": [
        {{
          "date_raw": "1 เม.ย. 54",
          "position_and_workplace": "ครู/ชำนาญการ รร.บ้านโคกสูงหนองเสียวหนอง อ.วาปีปทุม สพป.มค.2 (เลื่อนเงินเดือน 1 ขั้น)",
          "position_no": "5693",
          "academic_standing": "ชำนาญการ",
          "salary": 25190,
          "order_ref": "คส.สพป.มค.2 ที่ 199/54 ลว. 18 เม.ย. 54"
        }}
      ]
    }}
    """
    
    for idx, img in enumerate(images):
        response = model.generate_content([prompt, img])
        text_resp = response.text.strip()
        
        # ตัด Markdown code block ออก (ถ้ามี)
        if "```json" in text_resp:
            text_resp = re.search(r"```json\s*(.*?)\s*```", text_resp, re.DOTALL).group(1)
        elif "```" in text_resp:
            text_resp = re.search(r"```\s*(.*?)\s*```", text_resp, re.DOTALL).group(1)
            
        res_json = json.loads(text_resp)
        for row in res_json.get("records", []):
            norm_date, sort_key = normalize_thai_date(row.get("date_raw", ""))
            row["normalized_date"] = norm_date
            row["sort_key"] = sort_key
            row["page_no"] = idx + 1
            all_records.append(row)
            
    return all_records

# ==========================================
# 4. Reconciliation Engine (5-Step Algorithm)
# ==========================================
def reconcile_records(records_a: List[dict], records_b: List[dict]):
    df_a = pd.DataFrame(records_a)
    df_b = pd.DataFrame(records_b)
    
    # 1. ตรวจสอบการเรียงลำดับเวลา (Timeline Inversion)
    def check_timeline(df, name):
        warnings = []
        if not df.empty and 'sort_key' in df.columns:
            for i in range(1, len(df)):
                if df.iloc[i]['sort_key'] < df.iloc[i-1]['sort_key'] and df.iloc[i]['sort_key'] != 0:
                    warnings.append({
                        "row": i + 1,
                        "date": df.iloc[i]['date_raw'],
                        "prev_date": df.iloc[i-1]['date_raw'],
                        "msg": f"พบวันที่เรียงย้อนหลัง ({df.iloc[i]['date_raw']} อยู่หลัง {df.iloc[i-1]['date_raw']})"
                    })
        return warnings

    inversions_a = check_timeline(df_a, "ไฟล์ 1 (ก.พ.7 อิเล็กทรอนิกส์)")
    inversions_b = check_timeline(df_b, "ไฟล์ 2 (ก.ค.ศ.16 เขียนมือ)")

    # 2. ตรวจหารายการซ้ำซ้อนในตัวเอง (Duplicate Records)
    dup_a = pd.DataFrame()
    if not df_a.empty and 'normalized_date' in df_a.columns and 'salary' in df_a.columns:
        dup_a = df_a[df_a.duplicated(subset=['normalized_date', 'salary'], keep=False)]
    
    # 3. Two-Way Comparison Matching
    all_keys = set()
    for r in records_a:
        all_keys.add((r['normalized_date'], r.get('salary', 0)))
    for r in records_b:
        all_keys.add((r['normalized_date'], r.get('salary', 0)))

    def get_sort_key_from_norm(norm_date):
        _, key = normalize_thai_date(norm_date)
        return key

    sorted_keys = sorted(list(all_keys), key=lambda x: get_sort_key_from_norm(x[0]))
    comparison_table = []
    
    for norm_date, salary in sorted_keys:
        matches_a = [r for r in records_a if r['normalized_date'] == norm_date and r.get('salary', 0) == salary]
        matches_b = [r for r in records_b if r['normalized_date'] == norm_date and r.get('salary', 0) == salary]
        
        status = ""
        action_note = ""
        
        if matches_a and matches_b:
            if len(matches_a) > 1:
                status = "⚠️ ตรงกัน แต่ระบบอิเล็กทรอนิกส์มีรายการเบิ้ลซ้ำ"
                action_note = f"พบรายการซ้ำในระบบอิเล็กทรอนิกส์ {len(matches_a)} แถว (ควรลบออก 1 แถว)"
            else:
                status = "✅ ตรงกันสมบูรณ์"
                action_note = "-"
            desc_a = f"{matches_a[0]['position_and_workplace']} | เงินเดือน {matches_a[0]['salary']:,.0f} | คำสั่ง: {matches_a[0]['order_ref']}"
            desc_b = f"{matches_b[0]['position_and_workplace']} | เงินเดือน {matches_b[0]['salary']:,.0f} | คำสั่ง: {matches_b[0]['order_ref']}"
        elif matches_a and not matches_b:
            status = "❌ ขาดหายในเอกสารเขียนมือ"
            action_note = "ต้องเพิ่มรายการนี้ลงในสมุดประวัติเขียนมือ"
            desc_a = f"{matches_a[0]['position_and_workplace']} | เงินเดือน {matches_a[0]['salary']:,.0f} | คำสั่ง: {matches_a[0]['order_ref']}"
            desc_b = "-"
        elif not matches_a and matches_b:
            status = "❌ ขาดหายในระบบอิเล็กทรอนิกส์"
            action_note = "ต้องนำเข้ารายการนี้เข้าสู่ระบบ ก.พ.7 อิเล็กทรอนิกส์"
            desc_a = "-"
            desc_b = f"{matches_b[0]['position_and_workplace']} | เงินเดือน {matches_b[0]['salary']:,.0f} | คำสั่ง: {matches_b[0]['order_ref']}"
            
        comparison_table.append({
            "วัน เดือน ปี (พ.ศ.)": norm_date,
            "อัตราเงินเดือน": salary,
            "ไฟล์ 1: ก.พ.7 อิเล็กทรอนิกส์": desc_a,
            "ไฟล์ 2: ก.ค.ศ.16 เขียนมือ": desc_b,
            "สถานะการตรวจสอบ": status,
            "ข้อเสนอแนะในการแก้ไข": action_note
        })

    return comparison_table, inversions_a, inversions_b, dup_a

# ==========================================
# 5. ฟังก์ชันส่งออกรายงาน Excel สวยงาม
# ==========================================
def export_to_excel(comparison_table):
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "ผลการเปรียบเทียบ กพ7"
    
    headers = ["ลำดับ", "วัน เดือน ปี (พ.ศ.)", "อัตราเงินเดือน (บาท)", "ข้อมูลไฟล์ 1 (อิเล็กทรอนิกส์)", "ข้อมูลไฟล์ 2 (เขียนมือ)", "สถานะการตรวจสอบ", "ข้อเสนอแนะในการแก้ไข"]
    ws1.append(headers)
    
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="TH Sarabun New", size=14, bold=True, color="FFFFFF")
    for cell in ws1[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    fill_pass = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_warn = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_error = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    regular_font = Font(name="TH Sarabun New", size=13)
    
    for idx, row in enumerate(comparison_table, start=1):
        ws1.append([
            idx,
            row["วัน เดือน ปี (พ.ศ.)"],
            f"{row['อัตราเงินเดือน']:,.0f}" if row['อัตราเงินเดือน'] else "-",
            row["ไฟล์ 1: ก.พ.7 อิเล็กทรอนิกส์"],
            row["ไฟล์ 2: ก.ค.ศ.16 เขียนมือ"],
            row["สถานะการตรวจสอบ"],
            row["ข้อเสนอแนะในการแก้ไข"]
        ])
        
        current_row = ws1[idx + 1]
        for cell in current_row:
            cell.font = regular_font
            if "ตรงกันสมบูรณ์" in str(row["สถานะการตรวจสอบ"]):
                cell.fill = fill_pass
            elif "รายการเบิ้ลซ้ำ" in str(row["สถานะการตรวจสอบ"]):
                cell.fill = fill_warn
            else:
                cell.fill = fill_error
                
    for col in ws1.columns:
        col_letter = col[0].column_letter
        ws1.column_dimensions[col_letter].width = 25
    ws1.column_dimensions['D'].width = 45
    ws1.column_dimensions['E'].width = 45
    
    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io

# ==========================================
# 6. ส่วนแสดงผล Streamlit Dashboard UI
# ==========================================
st.set_page_config(page_title="ระบบตรวจสอบ ก.พ.7 - สพป.มหาสารคาม เขต 2", layout="wide")

st.title("📋 ระบบตรวจสอบและเทียบเคียงทะเบียนประวัติ ก.พ.7 / ก.ค.ศ.16 อัจฉริยะ")
st.caption("กลุ่มบริหารงานบุคคล สำนักงานเขตพื้นที่การศึกษาประถมศึกษามหาสารคาม เขต 2")

with st.sidebar:
    st.header("⚙️ การตั้งค่าระบบ")
    gemini_api_key = st.text_input("ใส่ Google Gemini API Key:", type="password")
    st.markdown("[👉 คลิกที่นี่เพื่อขอรับ API Key ฟรีจาก Google AI Studio](https://aistudio.google.com/)")
    
    selected_model = None
    if gemini_api_key:
        try:
            genai.configure(api_key=gemini_api_key)
            # ดึงรายชื่อโมเดลที่ใช้งานได้จริงจาก API Key
            valid_models = [
                m.name for m in genai.list_models() 
                if 'generateContent' in m.supported_generation_methods
            ]
            if valid_models:
                st.success(f"เชื่อมต่อ API สำเร็จ (พบ {len(valid_models)} โมเดล)")
                selected_model = st.selectbox("เลือกโมเดล AI ที่ต้องการใช้:", valid_models, index=0)
            else:
                st.warning("ไม่พบโมเดล generateContent ในบัญชีนี้")
        except Exception as e:
            st.error(f"API Key ไม่ถูกต้อง: {str(e)}")
            
    st.divider()
    st.info("💡 **ขั้นตอนการใช้งาน:**\n1. ใส่ API Key\n2. อัปโหลดไฟล์ PDF ทั้ง 2 ฝั่ง\n3. กดปุ่ม 'เริ่มการประมวลผล'\n4. ดาวน์โหลดรายงานผล Excel")

col1, col2 = st.columns(2)
with col1:
    st.subheader("📄 ไฟล์ที่ 1: ก.พ.7 อิเล็กทรอนิกส์")
    file_e = st.file_uploader("อัปโหลดไฟล์ PDF จากระบบ HRMS", type=["pdf"], key="file_e")

with col2:
    st.subheader("✍️ ไฟล์ที่ 2: ก.ค.ศ.16 (เขียนมือ)")
    file_m = st.file_uploader("อัปโหลดไฟล์ PDF สแกนเอกสารเขียนมือ", type=["pdf"], key="file_m")

if st.button("🚀 เริ่มต้นการสกัดข้อมูลและตรวจสอบความถูกต้อง", type="primary"):
    if not gemini_api_key:
        st.error("กรุณากรอก Gemini API Key ในแถบด้านข้างก่อนเริ่มทำงาน")
    elif not selected_model:
        st.error("กรุณาเลือกโมเดล AI ในแถบด้านข้าง")
    elif not file_e or not file_m:
        st.error("กรุณาอัปโหลดไฟล์ PDF ให้ครบทั้ง 2 ไฟล์")
    else:
        with st.spinner(f"🔍 กำลังประมวลผลด้วยโมเดล {selected_model}..."):
            try:
                # 1. สกัดข้อมูลไฟล์ 1
                bytes_e = file_e.read()
                records_e = extract_data_from_pdf(bytes_e, gemini_api_key, selected_model, "ไฟล์อิเล็กทรอนิกส์ กพ.7")
                
                # 2. สกัดข้อมูลไฟล์ 2
                bytes_m = file_m.read()
                records_m = extract_data_from_pdf(bytes_m, gemini_api_key, selected_model, "ไฟล์เขียนมือ กคศ.16")
                
                # 3. ตรวจสอบเปรียบเทียบ
                comp_table, inv_a, inv_b, dups_a = reconcile_records(records_e, records_m)
                
                st.success("✅ ประมวลผลและตรวจสอบข้อมูลเสร็จสมบูรณ์!")
                
                # แสดงกล่องแจ้งเตือนความผิดปกติ
                if inv_b:
                    st.error(f"⚠️ **ตรวจพบลำดับวันที่สลับที่กัน (Timeline Inversion) ในไฟล์เขียนมือ {len(inv_b)} จุด:**")
                    for w in inv_b:
                        st.write(f"- แถวที่ {w['row']}: {w['msg']}")
                        
                if not dups_a.empty:
                    st.warning(f"⚠️ **ตรวจพบรายการบันทึกซ้ำซ้อนในระบบอิเล็กทรอนิกส์:** พบ {len(dups_a)} รายการที่มีวันที่และเงินเดือนซ้ำกัน")

                # แสดงตารางผลลัพธ์
                df_results = pd.DataFrame(comp_table)
                st.subheader("📊 ตารางเปรียบเทียบความถูกต้องของ วัน เดือน ปี พ.ศ. และรายการข้อมูล")
                st.dataframe(df_results, use_container_width=True)
                
                # ปุ่มดาวน์โหลด Excel
                excel_data = export_to_excel(comp_table)
                st.download_button(
                    label="📥 ดาวน์โหลดรายงานผลการตรวจสอบ (.xlsx)",
                    data=excel_data,
                    file_name=f"KP7_Audit_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {str(e)}")
