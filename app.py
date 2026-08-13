import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF to Excel Converter", page_icon="📑", layout="wide")

st.title("📑 ระบบแปลงข้อมูลตารางจาก PDF เป็น Excel")
st.write("ระบบจะอ่านข้อมูลเฉพาะส่วน **'วันที่ใช้บัตร'** ไปจนถึง **'สรุปยอดงวดนี้'**")

# 1. ส่วนการอัปโหลดไฟล์ PDF
uploaded_files = st.file_uploader(
    "เลือกหรือลากไฟล์ PDF มาวางที่นี่ (รองรับหลายไฟล์พร้อมกัน)", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"📁 จำนวนไฟล์ที่อัปโหลด: **{len(uploaded_files)}** ไฟล์")
    
    all_extracted_rows = []
    
    with st.spinner("กำลังประมวลผลและค้นหาช่วงข้อมูล..."):
        for uploaded_file in uploaded_files:
            recording = False  # ตัวแปรสถานะบันทึกข้อมูล (เริ่มเปิดเมื่อเจอคำที่กำหนด)
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    # ดึงบรรทัดข้อความทั้งหมดในแต่ละหน้า
                    text = page.extract_text()
                    if not text:
                        continue
                    
                    lines = text.split("\n")
                    
                    for line in lines:
                        clean_line = line.strip()
                        
                        # ตรวจสอบจุดเริ่มต้น: เจอคำว่า "วันที่ใช้บัตร"
                        if "วันที่ใช้บัตร" in clean_line:
                            recording = True
                        
                        # ตรวจสอบจุดสิ้นสุด: เจอคำว่า "สรุปยอดงวดนี้"
                        if "สรุปยอดงวดนี้" in clean_line:
                            recording = False
                            break  # หยุดอ่านไฟล์นี้ทันที
                        
                        # หากอยู่ในช่วงที่กำหนด ให้บันทึกบรรทัดข้อมูลไว้
                        if recording and clean_line:
                            # แยกคอลัมน์ด้วยช่องว่าง หรือเว้นวรรคหลายช่อง
                            row_data = clean_line.split()
                            
                            # เพิ่มชื่อไฟล์และหน้าไว้เป็นอ้างอิง
                            all_extracted_rows.append({
                                "ชื่อไฟล์ PDF": uploaded_file.name,
                                "หน้า PDF": page_num,
                                "ข้อมูลบรรทัด": clean_line,
                                "รายละเอียด": row_data
                            })

    if all_extracted_rows:
        st.success("✅ สกัดข้อมูลสำเร็จ!")
        
        # 2. จัดโครงสร้างข้อมูลให้อยู่ในรูปแบบ DataFrame
        # แยกข้อมูลข้อความดั้งเดิมและคอลัมน์ย่อย
        df_raw = pd.DataFrame(all_extracted_rows)
        
        # ขยายรายการข้อมูลย่อยออกเป็นคอลัมน์แยก (Col_1, Col_2, ...)
        details_df = pd.DataFrame(df_raw["รายละเอียด"].tolist())
        details_df.columns = [f"Col_{i+1}" for i in range(details_df.shape[1])]
        
        # รวมข้อมูลอ้างอิงและคอลัมน์ย่อยเข้าด้วยกัน
        final_df = pd.concat([df_raw[["ชื่อไฟล์ PDF", "หน้า PDF", "ข้อมูลบรรทัด"]], details_df], axis=1)
        
        # แสดงตัวอย่างตารางบนหน้าเว็บ
        st.subheader("📊 ข้อมูลตารางที่ดึงได้:")
        st.dataframe(final_df, use_container_width=True)
        
        # 3. นำข้อมูลลงตาราง Excel
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Extracted_Data')
        
        # ปุ่มสำหรับดาวน์โหลด
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name="statement_extracted_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ ไม่พบคำว่า 'วันที่ใช้บัตร' ในไฟล์ PDF ที่อัปโหลด กรุณาตรวจสอบเนื้อหาในไฟล์")
