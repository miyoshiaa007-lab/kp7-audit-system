import streamlit as st
import pandas as pd
import json
import re
import io
import base64
from datetime import datetime
from pydantic import BaseModel, Field
from typing import List, Optional, Tuple, Dict, Any
import anthropic
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

# ==========================================
# 1. โครงสร้างข้อมูลมาตรฐาน (Strict Schema)
# ==========================================
class RecordEntry(BaseModel):
    date_raw: str = Field(default="", description="วันเดือนปี เช่น '1 เม.ย. 54'")
    position_and_workplace: str = Field(default="", description="ตำแหน่ง หน่วยงาน วิทยฐานะ หรือการเลื่อนขั้น")
    position_no: Optional[str] = Field(default="", description="เลขที่ตำแหน่ง")
    academic_standing: Optional[str] = Field(default="", description="วิทยฐานะ")
    salary: Optional[float] = Field(default=0.0, description="อัตราเงินเดือนเป็นตัวเลขเท่านั้น")
    order_ref: Optional[str] = Field(default="", description="เลขที่คำสั่งและวันที่ลงนาม")

# ==========================================
# 2. ระบบทำความสะอาดและแปลงข้อมูล (Accuracy Guardrails)
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
    if "```json" in text:
        text = text.split("```json")[1].split("```")[0].strip()
    elif "```" in text:
        text = text.split("```")[1].strip()
    return text

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
    clean_str = str(date_str).translate(THAI_DIGITS).replace(" ", "").replace(".", ". ")
    pattern = r"(\d{1,2})\s*([ก-๙\.]+)\s*(\d{2,4})"
    match = re.search(pattern, clean_str)
    if not match: return date_str.strip(), 0
    day = int(match.group(1))
    month_raw = match.group(2).replace(" ", "")
    year_raw = int(match.group(3))
    year = 2500 + year_raw if year_raw < 100 else year_raw
    month = next((m_val for m_key, m_val in THAI_MONTHS.items() if m_key in month_raw), 0)
    if month == 0: return f"{day} {month_raw} {year}", (year * 10000) + day
    return f"{day} {MONTH_LABEL[month]} {year}", (year * 10000) + (month * 100) + day

# ==========================================
# 3. VLM Data Extractor (Claude Native PDF API)
# ==========================================
def extract_pdf_records_claude(pdf_bytes: bytes, api_key: str, model_name: str, hint: str) -> List[Dict[str, Any]]:
    client = anthropic.Anthropic(api_key=api_key)
    pdf_base64 = base64.b64encode(pdf_bytes).decode("utf-8")
    
    prompt = f"""
    คุณคือผู้เชี่ยวชาญการตรวจสอบทะเบียนประวัติ ก.พ.7 / ก.ค.ศ.16 สพป.มหาสารคาม เขต 2
    ประเภทเอกสาร: {hint}
    
    กฎเหล็กเพื่อความแม่นยำ 100%:
    1. สกัดข้อมูลประวัติการรับเงินเดือนทุกแถว ทุกหน้า ห้ามข้ามเด็ดขาด
    2. ตัวเลขเงินเดือน (salary) ต้องเป็นตัวเลขอารบิกที่ถูกต้อง ระวังการสับสนเลข 3 กับ 8, 0 กับ 6
    3. วันเดือนปี (date_raw) ถอดตามที่ปรากฏจริง
    4. เอกสารอ้างอิง (order_ref) ให้สกัดเลขที่คำสั่งและวันที่ลงนามมาให้ครบถ้วน
    
    ส่งผลลัพธ์เป็นโครงสร้าง JSON ล้วนๆ ในรูปแบบนี้เท่านั้น (ไม่ต้องมีคำอธิบายอื่น):
    {{
      "records": [
        {{
          "date_raw": "1 เม.ย. 54",
          "position_and_workplace": "ครู รร.บ้านโคกสูงหนองเสียวหนอง (เลื่อนเงินเดือน 1 ขั้น)",
          "position_no": "5693",
          "academic_standing": "ชำนาญการ",
          "salary": 25190,
          "order_ref": "คส.สพป.มค.2 ที่ 199/54 ลว. 18 เม.ย. 54"
        }}
      ]
    }}
    """
    
    response = client.messages.create(
        model=model_name,
        max_tokens=8192,
        temperature=0.0,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_base64
                        }
                    },
                    {
                        "type": "text",
                        "text": prompt
                    }
                ]
            }
        ]
    )
    
    cleaned_str = clean_json_string(response.content[0].text)
    
    try:
        data = json.loads(cleaned_str)
        records = data.get("records", [])
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการแปล JSON: {str(e)}")
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
# 4. Two-Way Reconciliation (Fuzzy Match & Logic Check)
# ==========================================
def run_two_way_reconciliation(records_a: List[Dict[str, Any]], records_b: List[Dict[str, Any]]):
    inversions_b = []
    for i in range(1, len(records_b)):
        prev, curr = records_b[i - 1], records_b[i]
        if curr["sort_key"] > 0 and prev["sort_key"] > 0 and curr["sort_key"] < prev["sort_key"]:
            inversions_b.append({
                "row": i + 1, "msg": f"แถวที่ {i+1}: ลงวันที่ '{curr.get('date_raw', '-')}' อยู่ถัดจาก '{prev.get('date_raw', '-')}' (ลำดับเวลาย้อนกลับ)"
            })

    all_dates = {}
    for r in records_a: all_dates.setdefault(r["normalized_date"], {"a": [], "b": [], "sort_key": r["sort_key"]})["a"].append(r)
    for r in records_b: all_dates.setdefault(r["normalized_date"], {"a": [], "b": [], "sort_key": r["sort_key"]})["b"].append(r)

    sorted_dates = sorted(all_dates.items(), key=lambda x: x[1]["sort_key"])
    matched_rows = []
    stats = {"perfect_match": 0, "duplicate_in_hrms": 0, "missing_in_manual": 0, "missing_in_hrms": 0, "salary_mismatch": 0}

    for date_str, group in sorted_dates:
        list_a, list_b = group["a"], group["b"]
        
        if list_a and list_b:
            used_b = set()
            for r_a in list_a:
                matched_b_idx = None
                for idx_b, r_b in enumerate(list_b):
                    if idx_b not in used_b and abs(r_a["salary"] - r_b["salary"]) < 1.0:
                        matched_b_idx = idx_b
                        break
                
                if matched_b_idx is not None:
                    used_b.add(matched_b_idx)
                    r_b = list_b[matched_b_idx]
                    status, action = "✅ ตรงกันสมบูรณ์", "-"
                    stats["perfect_match"] += 1
                    desc_b = f"{r_b['position_and_workplace']} (เงินเดือน {r_b['salary']:,.0f} บ.) [{r_b['order_ref']}]"
                else:
                    if len(used_b) < len(list_b):
                        matched_b_idx = next(i for i in range(len(list_b)) if i not in used_b)
                        used_b.add(matched_b_idx)
                        r_b = list_b[matched_b_idx]
                        status, action = "⚠️ วันที่ตรงกันแต่ยอดเงินเดือนไม่ตรง", f"ก.พ.7 ยอด {r_a['salary']:,.0f} vs เขียนมือ {r_b['salary']:,.0f} (โปรดตรวจสอบ)"
                        stats["salary_mismatch"] += 1
                        desc_b = f"{r_b['position_and_workplace']} (เงินเดือน {r_b['salary']:,.0f} บ.) [{r_b['order_ref']}]"
                    else:
                        status, action = "❌ ขาดในเล่มเขียนมือ (หรือซ้ำในระบบ)", "เช็คการบันทึกซ้ำใน ก.พ.7 หรือเพิ่มลงเล่มเขียนมือ"
                        stats["duplicate_in_hrms"] += 1
                        desc_b = "-"

                desc_a = f"{r_a['position_and_workplace']} (เงินเดือน {r_a['salary']:,.0f} บ.) [{r_a['order_ref']}]"
                matched_rows.append({"วัน เดือน ปี (พ.ศ.)": date_str, "เงินเดือน ก.พ.7": f"{r_a['salary']:,.0f}" if r_a['salary'] > 0 else "-", "ข้อมูล ก.พ.7 อิเล็กทรอนิกส์": desc_a, "ข้อมูล ก.ค.ศ.16 เขียนมือ": desc_b, "สถานะการตรวจสอบ": status, "สิ่งที่ต้องดำเนินการแก้ไข": action})

            for idx_b, r_b in enumerate(list_b):
                if idx_b not in used_b:
                    matched_rows.append({"วัน เดือน ปี (พ.ศ.)": date_str, "เงินเดือน ก.พ.7": "-", "ข้อมูล ก.พ.7 อิเล็กทรอนิกส์": "-", "ข้อมูล ก.ค.ศ.16 เขียนมือ": f"{r_b['position_and_workplace']} (เงินเดือน {r_b['salary']:,.0f} บ.) [{r_b['order_ref']}]", "สถานะการตรวจสอบ": "❌ ขาดในระบบอิเล็กทรอนิกส์", "สิ่งที่ต้องดำเนินการแก้ไข": "นำเข้าข้อมูลคำสั่งนี้เข้าสู่ระบบ ก.พ.7"})
                    stats["missing_in_hrms"] += 1

        elif list_a and not list_b:
            for r_a in list_a:
                matched_rows.append({"วัน เดือน ปี (พ.ศ.)": date_str, "เงินเดือน ก.พ.7": f"{r_a['salary']:,.0f}" if r_a['salary'] > 0 else "-", "ข้อมูล ก.พ.7 อิเล็กทรอนิกส์": f"{r_a['position_and_workplace']} (เงินเดือน {r_a['salary']:,.0f} บ.) [{r_a['order_ref']}]", "ข้อมูล ก.ค.ศ.16 เขียนมือ": "-", "สถานะการตรวจสอบ": "❌ ขาดในเล่มเขียนมือ (ต้องเพิ่ม)", "สิ่งที่ต้องดำเนินการแก้ไข": "เพิ่มรายการนี้ลงในสมุด ก.ค.ศ.16"})
                stats["missing_in_manual"] += 1

        elif not list_a and list_b:
            for r_b in list_b:
                matched_rows.append({"วัน เดือน ปี (พ.ศ.)": date_str, "เงินเดือน ก.พ.7": "-", "ข้อมูล ก.พ.7 อิเล็กทรอนิกส์": "-", "ข้อมูล ก.ค.ศ.16 เขียนมือ": f"{r_b['position_and_workplace']} (เงินเดือน {r_b['salary']:,.0f} บ.) [{r_b['order_ref']}]", "สถานะการตรวจสอบ": "❌ ขาดในระบบอิเล็กทรอนิกส์ (ต้องบันทึก)", "สิ่งที่ต้องดำเนินการแก้ไข": "นำเข้าข้อมูลนี้เข้าสู่ระบบ ก.พ.7"})
                stats["missing_in_hrms"] += 1

    return matched_rows, stats, inversions_b

# ==========================================
# 5. สร้างรายงาน Excel
# ==========================================
def generate_audit_excel(table_rows, stats, inv_b) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "ผลการเทียบเคียง กพ7"
    ws.append(["ลำดับ", "วัน เดือน ปี (พ.ศ.)", "เงินเดือน ก.พ.7", "ก.พ.7 อิเล็กทรอนิกส์ (HRMS)", "ก.ค.ศ.16 (เขียนมือ)", "สถานะการตรวจสอบ", "สิ่งที่ต้องดำเนินการแก้ไข"])
    
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="TH Sarabun New", size=14, bold=True, color="FFFFFF")
    for cell in ws[1]: cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal="center", vertical="center")
        
    fill_pass = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_warn = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_error = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    row_font = Font(name="TH Sarabun New", size=13)
    
    for idx, r in enumerate(table_rows, 1):
        ws.append([idx, r["วัน เดือน ปี (พ.ศ.)"], r["เงินเดือน ก.พ.7"], r["ข้อมูล ก.พ.7 อิเล็กทรอนิกส์"], r["ข้อมูล ก.ค.ศ.16 เขียนมือ"], r["สถานะการตรวจสอบ"], r["สิ่งที่ต้องดำเนินการแก้ไข"]])
        for cell in ws[idx + 1]:
            cell.font = row_font
            cell.fill = fill_pass if "ตรงกันสมบูรณ์" in r["สถานะการตรวจสอบ"] else (fill_warn if "⚠️" in r["สถานะการตรวจสอบ"] else fill_error)

    ws.column_dimensions['B'].width, ws.column_dimensions['C'].width = 18, 16
    ws.column_dimensions['D'].width, ws.column_dimensions['E'].width = 45, 45
    ws.column_dimensions['F'].width, ws.column_dimensions['G'].width = 30, 35

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

# ==========================================
# 6. Streamlit UI
# ==========================================
st.set_page_config(page_title="ระบบตรวจสอบความถูกต้อง ก.พ.7", layout="wide")
st.title("🎯 ระบบตรวจสอบและเทียบเคียง ก.พ.7 / ก.ค.ศ.16 ด้วย Claude API")
st.caption("ประมวลผลความเร็วสูงผ่าน Native PDF - สพป.มหาสารคาม เขต 2")

with st.sidebar:
    st.header("⚙️ การตั้งค่าระบบ")
    api_key_input = st.text_input("ใส่ Anthropic API Key (sk-ant-...):").strip()
    
    # รายชื่อโมเดลล่าสุดของ Claude
    model_list = [
        "claude-3-5-sonnet-latest",
        "claude-3-5-haiku-latest",
        "claude-5-sonnet-202608",
        "claude-4-5-haiku-202608",
        "claude-5-mythos-202608",
        "claude-5-fable-202608"
    ]
    active_model = st.selectbox("เลือกโมเดล VLM:", model_list, index=0)

col1, col2 = st.columns(2)
with col1:
    st.subheader("📄 ไฟล์ 1: ก.พ.7 อิเล็กทรอนิกส์ (HRMS)")
    file_hrms = st.file_uploader("อัปโหลด PDF จากระบบ", type=["pdf"], key="file_hrms")
with col2:
    st.subheader("✍️ ไฟล์ 2: ก.ค.ศ.16 (เขียนมือ)")
    file_manual = st.file_uploader("อัปโหลด PDF สแกน", type=["pdf"], key="file_manual")

if st.button("🚀 เริ่มประมวลผล (ความเร็วสูง)", type="primary"):
    if not api_key_input or not active_model or not file_hrms or not file_manual:
        st.error("กรุณาใส่ Anthropic API Key และอัปโหลดไฟล์ให้ครบ")
    elif not api_key_input.startswith("sk-ant-"):
        st.error("API Key ของ Claude จะต้องขึ้นต้นด้วย 'sk-ant-' กรุณาตรวจสอบอีกครั้ง")
    else:
        status_box = st.status("🔍 กำลังประมวลผล Native PDF ผ่าน Claude API...", expanded=True)
        try:
            status_box.write(f"📄 1/2 อ่าน ก.พ.7 อิเล็กทรอนิกส์ด้วยโมเดล {active_model}...")
            records_hrms = extract_pdf_records_claude(file_hrms.read(), api_key_input, active_model, "ก.พ.7 อิเล็กทรอนิกส์")
            
            status_box.write("✍️ 2/2 อ่าน ก.ค.ศ.16 เขียนมือ และถอดรหัสลายมือ...")
            records_man = extract_pdf_records_claude(file_manual.read(), api_key_input, active_model, "ก.ค.ศ.16 เขียนมือ")
            
            status_box.write("⚖️ กำลังเทียบเคียงข้อมูลและตรวจสอบตรรกะ...")
            comp_results, stats_data, inv_man = run_two_way_reconciliation(records_hrms, records_man)
            
            status_box.update(label="✅ ตรวจสอบเสร็จสมบูรณ์!", state="complete", expanded=False)
            
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("ตรงกันสมบูรณ์", f"{stats_data['perfect_match']} รายการ")
            m2.metric("ขาดในเล่มเขียนมือ", f"{stats_data['missing_in_manual']} รายการ")
            m3.metric("ขาดในระบบ ก.พ.7", f"{stats_data['missing_in_hrms']} รายการ")
            m4.metric("ยอดเงินเดือนเพี้ยน/ซ้ำ", f"{stats_data['salary_mismatch'] + stats_data['duplicate_in_hrms']} รายการ")
            m5.metric("วันที่เขียนสลับลำดับ", f"{len(inv_man)} จุด")
            
            st.divider()
            if inv_man:
                st.error(f"⚠️ **ตรวจพบลำดับวันที่สลับที่กัน (Timeline Inversion) {len(inv_man)} จุด:**")
                for inv in inv_man: st.write(f"- {inv['msg']}")

            st.subheader("📊 ตารางเปรียบเทียบข้อมูล (Reconciliation Table)")
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
