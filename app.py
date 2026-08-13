import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF to Excel Converter", page_icon="📑", layout="wide")

st.title("📑 ระบบแปลงข้อมูลตารางจาก PDF เป็น Excel")
st.write("อัปโหลดไฟล์ PDF เพื่อทำการสกัดตารางข้อมูลและดาวน์โหลดเป็นไฟล์ Excel (.xlsx)")

# 1. ส่วนการอัปโหลดไฟล์ PDF
uploaded_files = st.file_uploader(
    "เลือกไฟล์ PDF (รองรับการอัปโหลดหลายไฟล์พร้อมกัน)", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"📁 อัปโหลดทั้งหมด **{len(uploaded_files)}** ไฟล์")
    
    all_extracted_tables = []
    
    with st.spinner("กำลังอ่านและดึงข้อมูลจากไฟล์ PDF..."):
        for uploaded_file in uploaded_files:
            with pdfplumber.open(uploaded_file) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    tables = page.extract_tables()
                    for t in tables:
                        if t and len(t) > 1:
                            # แปลงเป็น DataFrame และใช้แถวแรกของตารางเป็น Header
                            df_table = pd.DataFrame(t[1:], columns=t[0])
                            # ใส่ข้อมูลชื่อไฟล์และหน้าไว้เป็นอ้างอิง
                            df_table.insert(0, "หน้า PDF", page_num)
                            df_table.insert(0, "ชื่อไฟล์ PDF", uploaded_file.name)
                            all_extracted_tables.append(df_table)

    if all_extracted_tables:
        st.success("✅ สกัดข้อมูลตารางสำเร็จ!")
        
        # รวมตารางทั้งหมดเข้าด้วยกัน
        combined_df = pd.concat(all_extracted_tables, ignore_index=True)
        
        # แสดงตัวอย่างข้อมูล
        st.subheader("📊 ตารางข้อมูลที่ดึงได้:")
        st.dataframe(combined_df, use_container_width=True)
        
        # 2. แปลงลงไฟล์ Excel (.xlsx)
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            combined_df.to_excel(writer, index=False, sheet_name='Extracted_Data')
        
        # ปุ่มสำหรับดาวน์โหลด
        st.download_button(
            label="📥 ดาวน์โหลดข้อมูลเป็นไฟล์ Excel (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name="extracted_pdf_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ ไม่พบตารางข้อมูลในไฟล์ PDF ที่อัปโหลด (หากเป็นไฟล์ภาพสแกน ต้องใช้ระบบ OCR ดึงข้อมูลแทน)")
