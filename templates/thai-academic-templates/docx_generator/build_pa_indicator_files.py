# -*- coding: utf-8 -*-
"""
build_pa_indicator_files.py
-----------------------------
สร้างไฟล์ Word แยกไฟล์ต่างหาก "1 ไฟล์ต่อ 1 ตัวชี้วัด" สำหรับผู้ที่ต้องการเก็บ
แต่ละตัวชี้วัดเป็นคนละไฟล์ถาวร (ไม่รวมเล่ม) แทนการคัดลอกวางบล็อกด้วยมือใน Word
ซึ่งเสี่ยงทำให้รูปแบบ (ฟอนต์/ตาราง/ระยะขอบ) เพี้ยนได้ง่าย

แต่ละไฟล์ที่ได้:
- มีเฉพาะ "เนื้อหาตัวชี้วัด" ล้วน ๆ (ไม่มีหน้าปก ไม่มีสารบัญ)
- มีเลขหน้าอารบิก (1, 2, 3, ...) ที่ท้ายกระดาษ เริ่มนับใหม่จาก 1 ทุกไฟล์
- ใช้ฟอนต์/ระยะขอบ/รูปแบบตารางชุดเดียวกับเทมเพลตหลัก (ผ่าน docx_common.py)
- ตั้งชื่อไฟล์ตามรูปแบบ PA4_ด้านที่{X}_ตัวชี้วัดที่{Y}.docx ให้ตรงกับที่ผู้ใช้ตั้งไว้เอง

แก้จำนวนตัวชี้วัดต่อด้าน หรือชื่อด้าน ได้ที่ตัวแปร DOMAINS ด้านล่าง

รัน:  python build_pa_indicator_files.py
ผลลัพธ์: ../output/pa_indicator_files/PA4_ด้านที่1_ตัวชี้วัดที่1.docx ... เป็นต้น
"""
import os
from docx import Document

from docx_common import set_document_defaults, set_page_setup, set_page_number_format, add_page_number_footer
from build_pa_template import build_indicator_block

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output", "pa_indicator_files")

# ---------------------------------------------------------------------------
# แก้ไขจำนวนตัวชี้วัด/ชื่อด้านที่นี่ ให้ตรงกับแบบประเมินตำแหน่ง/วิทยฐานะของผู้จัดทำ
# ---------------------------------------------------------------------------
DOMAINS = [
    {
        "domain_no": 1,
        "part_label": "ส่วนที่ 2",
        "domain_title": "ผลการดำเนินงาน ด้านที่ 1  [ชื่อด้านที่ 1 ตามแบบประเมินตำแหน่ง/วิทยฐานะของผู้จัดทำ]",
        "num_indicators": 8,
    },
    {
        "domain_no": 2,
        "part_label": "ส่วนที่ 3",
        "domain_title": "ผลการดำเนินงาน ด้านที่ 2  [ชื่อด้านที่ 2 ตามแบบประเมินตำแหน่ง/วิทยฐานะของผู้จัดทำ]",
        "num_indicators": 4,
    },
]


def build_one_indicator_file(part_label, domain_title, domain_no, indicator_no):
    doc = Document()
    set_document_defaults(doc)

    section = doc.sections[0]
    set_page_setup(section)
    set_page_number_format(section, fmt="decimal", start=1)
    add_page_number_footer(section)

    build_indicator_block(
        doc,
        part_label=part_label,
        domain_title=domain_title,
        indicator_no=indicator_no,
        indicator_title=f"[ชื่อตัวชี้วัดที่ {indicator_no}]",
        leading_pagebreak=False,   # ไฟล์เดี่ยว ไม่ต้องมีหน้าว่างนำหน้า
    )

    filename = f"PA4_ด้านที่{domain_no}_ตัวชี้วัดที่{indicator_no}.docx"
    out_path = os.path.join(OUT_DIR, filename)
    doc.save(out_path)
    return out_path


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    saved = []
    for domain in DOMAINS:
        for i in range(1, domain["num_indicators"] + 1):
            path = build_one_indicator_file(
                part_label=domain["part_label"],
                domain_title=domain["domain_title"],
                domain_no=domain["domain_no"],
                indicator_no=i,
            )
            saved.append(path)
            print(f"บันทึกไฟล์แล้ว: {os.path.abspath(path)}")
    print(f"\nสร้างไฟล์ทั้งหมด {len(saved)} ไฟล์ ในโฟลเดอร์ {os.path.abspath(OUT_DIR)}")


if __name__ == "__main__":
    main()
