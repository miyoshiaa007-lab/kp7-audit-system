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
