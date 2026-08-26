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
            text_resp = re.search(r"
