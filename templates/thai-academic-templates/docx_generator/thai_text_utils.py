# -*- coding: utf-8 -*-
"""
thai_text_utils.py
-------------------
ฟังก์ชันช่วยจัดการข้อความภาษาไทยสำหรับ python-docx

ปัญหา: ภาษาไทยไม่มีการเว้นวรรคระหว่างคำ ทำให้ Word/LibreOffice บางเวอร์ชัน
(โดยเฉพาะเมื่อสร้างเอกสารด้วยสคริปต์ ไม่ได้พิมพ์ในโปรแกรม Word โดยตรง)
ไม่สามารถตัดคำท้ายบรรทัดได้อย่างถูกต้อง ทำให้ข้อความล้นขอบกระดาษ

วิธีแก้: แทรกอักขระ Zero-Width Space (U+200B) ระหว่างคำภาษาไทย
เพื่อบอกโปรแกรมว่า "ตัดบรรทัดตรงนี้ได้" อักขระนี้มองไม่เห็นและไม่มีความกว้าง
จึงไม่กระทบรูปลักษณ์ของข้อความเมื่อพิมพ์ออกมา

อ้างอิงแนวทางจาก skill: thai-docx (ใช้ pythainlp ตัดคำ)
"""
import re

try:
    from pythainlp.tokenize import word_tokenize
    _HAS_PYTHAINLP = True
except ImportError:  # pragma: no cover
    _HAS_PYTHAINLP = False

ZWS = "​"

_THAI_RUN_RE = re.compile(r"[ก-๙]+")


def insert_zwsp(text: str) -> str:
    """แทรก Zero-Width Space ระหว่างคำไทยในข้อความ (ข้อความอังกฤษ/ตัวเลข/URL ไม่ถูกแตะต้อง)"""
    if not text:
        return text
    if not _HAS_PYTHAINLP:
        return text

    def _tokenize_thai_chunk(m: re.Match) -> str:
        chunk = m.group(0)
        words = word_tokenize(chunk, engine="newmm")
        return ZWS.join(words)

    return _THAI_RUN_RE.sub(_tokenize_thai_chunk, text)


def clean_for_docx(text: str) -> str:
    """เตรียมข้อความก่อนใส่ใน docx: normalize ช่องว่าง + แทรก ZWS"""
    text = re.sub(r"[ \t]+", " ", text.strip())
    return insert_zwsp(text)
