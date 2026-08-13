import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF to Excel Converter", page_icon="📑", layout="wide")

st.title("📑 ระบบแปลงข้อมูลตารางจาก PDF เป็น Excel")
st.write("อัปโหลดไฟล์ PDF เพื่อทำการสกัดตารางข้อมูลและดาวน์โหลดเป็นไฟล์ Excel (.xlsx)")

# 1. ส่วนการอัปโหลดไฟล์ PDF ขึ้นระบบ
uploaded_files = st.file_uploader(
    "เลือกหรือลากไฟล์ PDF มาวางที่นี่ (รองรับหลายไฟล์พร้อมกัน)", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"📁 จำนวนไฟล์ที่อัปโหลด: **{len(uploaded_files)}** ไฟล์")
    
    all_extracted_tables = []
    
    with st.spinner("กำลังอ่านและดึงข้อมูลตารางจากไฟล์ PDF..."):
        for uploaded_file in uploaded_files:
            with pdfplumber.open(uploaded_file) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    tables = page.extract_tables()
                    for t in tables:
                        if t and len(t) > 0:
                            # สร้าง DataFrame โดยไม่กำหนด columns ทันที เพื่อป้องกันชื่อซ้ำ
                            df_table = pd.DataFrame(t)
                            
                            # ลบแถวหรือคอลัมน์ที่เป็นค่าว่างทั้งหมดออกไปก่อน
                            df_table = df_table.dropna(how='all').dropna(axis=1, how='all')
                            
                            if not df_table.empty:
                                # ตั้งชื่อคอลัมน์ชั่วคราวเป็น Column_0, Column_1, ... ป้องกัน Index ซ้ำ
                                df_table.columns = [f"Col_{i+1}" for i in range(df_table.shape[1])]
                                
                                # แทรกคอลัมน์ชื่อไฟล์และหน้าเพื่อระบุที่มา
                                df_table.insert(0, "หน้า PDF", page_num)
                                df_table.insert(0, "ชื่อไฟล์ PDF", uploaded_file.name)
                                
                                all_extracted_tables.append(df_table)

    if all_extracted_tables:
        st.success("✅ สกัดข้อมูลตารางสำเร็จ!")
        
        # รวมตารางทั้งหมดเข้าด้วยกันอย่างปลอดภัย
        combined_df = pd.concat(all_extracted_tables, ignore_index=True)
        
        # แสดงตัวอย่างตารางบนหน้าเว็บ
        st.subheader("📊 ตารางข้อมูลที่ดึงได้:")
        st.dataframe(combined_df, use_container_width=True)
        
        # 2. นำข้อมูลลงตาราง Excel
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            combined_df.to_excel(writer, index=False, sheet_name='Extracted_Data')
        
        # ปุ่มสำหรับให้ผู้ใช้ดาวน์โหลดไฟล์ Excel
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name="extracted_pdf_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ ไม่พบโครงสร้างตารางในไฟล์ PDF ที่อัปโหลด (หากเป็นภาพสแกนหรือรูปถ่าย จำเป็นต้องใช้ระบบ OCR เพิ่มเติม)")
