import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF Statement Extractor", page_icon="💳", layout="wide")

st.title("💳 ระบบดึงข้อมูลรายการใช้บัตรจาก PDF")
st.write("ดึงเฉพาะคอลัมน์: **วันที่ใช้บัตร**, **วันที่บันทึกรายการ**, **รายการ**, และ **จำนวนเงิน** (รวมบรรทัดสรุปยอด)")

# 1. ส่วนการอัปโหลดไฟล์ PDF
uploaded_files = st.file_uploader(
    "เลือกหรือลากไฟล์ PDF ใบแจ้งยอดมาวางที่นี่ (รองรับหลายไฟล์พร้อมกัน)", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"📁 จำนวนไฟล์ที่อัปโหลด: **{len(uploaded_files)}** ไฟล์")
    
    extracted_records = []
    
    with st.spinner("กำลังประมวลผลและดึงข้อมูลรายการ..."):
        for uploaded_file in uploaded_files:
            recording = False
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    text = page.extract_text()
                    if not text:
                        continue
                    
                    lines = text.split("\n")
                    
                    for line in lines:
                        clean_line = line.strip()
                        
                        # เริ่มอ่านเมื่อพบข้อความ "วันที่ใช้บัตร"
                        if "วันที่ใช้บัตร" in clean_line:
                            recording = True
                            continue  # ข้ามบรรทัดหัวข้อ
                        
                        if recording and clean_line:
                            parts = clean_line.split()
                            
                            # เช็กเมื่อเจอคำว่า "สรุปยอด" -> เก็บบรรทัดสรุปยอด แล้วค่อยจบการอ่าน
                            if "สรุปยอด" in clean_line:
                                if len(parts) >= 2:
                                    summary_label = " ".join([p for p in parts if not p.replace(',', '').replace('.', '').replace('-', '').isdigit()])
                                    summary_amount = parts[-1]
                                    
                                    extracted_records.append({
                                        "วันที่ใช้บัตร": "-",
                                        "วันที่บันทึกรายการ": "-",
                                        "รายการ": summary_label if summary_label else "สรุปยอดงวดนี้",
                                        "จำนวนเงิน": summary_amount,
                                        "ชื่อไฟล์": uploaded_file.name
                                    })
                                
                                recording = False
                                break  # หยุดอ่านหลังจากเก็บบรรทัดสรุปยอดเรียบร้อยแล้ว
                            
                            # เก็บบรรทัดรายการปกติ
                            if len(parts) >= 4:
                                txn_date = parts[0]       # วันที่ใช้บัตร
                                post_date = parts[1]      # วันที่บันทึกรายการ
                                amount = parts[-1]        # จำนวนเงิน (ท้ายสุด)
                                description = " ".join(parts[2:-1]) # รายการสินค้า/บริการ
                                
                                extracted_records.append({
                                    "วันที่ใช้บัตร": txn_date,
                                    "วันที่บันทึกรายการ": post_date,
                                    "รายการ": description,
                                    "จำนวนเงิน": amount,
                                    "ชื่อไฟล์": uploaded_file.name
                                })

    if extracted_records:
        st.success("✅ ดึงข้อมูลสำเร็จ!")
        
        df = pd.DataFrame(extracted_records)
        
        # แสดงผลตารางบนหน้าเว็บ
        st.subheader("📊 ตารางรายการใช้บัตรเครดิต (จนถึงบรรทัดสรุปยอด):")
        st.dataframe(df[["วันที่ใช้บัตร", "วันที่บันทึกรายการ", "รายการ", "จำนวนเงิน", "ชื่อไฟล์"]], use_container_width=True)
        
        # 2. นำข้อมูลลงไฟล์ Excel
        excel_buffer = io.BytesIO()
        export_df = df[["วันที่ใช้บัตร", "วันที่บันทึกรายการ", "รายการ", "จำนวนเงิน"]]
        
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            export_df.to_excel(writer, index=False, sheet_name='Credit_Card_Statement')
        
        # ปุ่มดาวน์โหลด
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name="credit_card_statement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ ไม่พบข้อมูลรายการใช้บัตรในช่วงที่กำหนด กรุณาตรวจสอบรูปแบบไฟล์ PDF")
