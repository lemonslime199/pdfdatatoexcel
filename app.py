import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="PDF Statement Extractor to Excel", page_icon="💳", layout="wide")

st.title("💳 ระบบดึงข้อมูลรายการใช้บัตรจาก PDF ลง Excel")
st.write("ดึงเฉพาะคอลัมน์: **วันที่ใช้บัตร**, **วันที่บันทึกรายการ**, **รายการ**, และ **จำนวนเงิน** (อ่านถึงบรรทัดสรุปยอด พร้อมแยกแท็บตาม **วันที่ครบกำหนดชำระ**)")

uploaded_files = st.file_uploader(
    "เลือกหรือลากไฟล์ PDF ใบแจ้งยอดมาวางที่นี่ (รองรับการอัปโหลดหลายไฟล์พร้อมกัน)", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"📁 จำนวนไฟล์ที่อัปโหลด: **{len(uploaded_files)}** ไฟล์")
    
    file_data_map = {}
    
    with st.spinner("กำลังประมวลผลและจัดรูปแบบตาราง..."):
        for uploaded_file in uploaded_files:
            recording = False
            extracted_records = []
            due_date = "Sheet_Data"  # ค่าเริ่มต้นถ้าหาไม่เจอ
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if not text:
                        continue
                    
                    lines = text.split("\n")
                    
                    for line in lines:
                        clean_line = line.strip()
                        
                        # ค้นหาวันที่ครบกำหนดชำระ
                        if ("ครบกำหนด" in clean_line or "Due Date" in clean_line) and due_date == "Sheet_Data":
                            date_match = re.search(r'(\d{2}[/.-]\d{2}[/.-]\d{2,4})', clean_line)
                            if date_match:
                                raw_date = date_match.group(1).replace('/', '-').replace('.', '-')
                                due_date = f"ชำระ_{raw_date}"
                        
                        # เริ่มอ่านเมื่อพบข้อความ "วันที่ใช้บัตร"
                        if "วันที่ใช้บัตร" in clean_line:
                            recording = True
                            continue
                        
                        if recording and clean_line:
                            parts = clean_line.split()
                            
                            # เจอคำว่า "สรุปยอด" -> เก็บบรรทัดสรุปยอด แล้วค่อยจบการอ่านไฟล์นี้
                            if "สรุปยอด" in clean_line:
                                if len(parts) >= 2:
                                    summary_label = " ".join([p for p in parts if not p.replace(',', '').replace('.', '').replace('-', '').isdigit()])
                                    summary_amount = parts[-1]
                                    
                                    extracted_records.append({
                                        "วันที่ใช้บัตร": "-",
                                        "วันที่บันทึกรายการ": "-",
                                        "รายการ": summary_label if summary_label else "สรุปยอดงวดนี้",
                                        "จำนวนเงิน": summary_amount
                                    })
                                recording = False
                                break
                            
                            # เก็บบรรทัดรายการปกติ
                            if len(parts) >= 4:
                                txn_date = parts[0]
                                post_date = parts[1]
                                amount = parts[-1]
                                description = " ".join(parts[2:-1])
                                
                                extracted_records.append({
                                    "วันที่ใช้บัตร": txn_date,
                                    "วันที่บันทึกรายการ": post_date,
                                    "รายการ": description,
                                    "จำนวนเงิน": amount
                                })
            
            if extracted_records:
                # ป้องกันชื่อ sheet ซ้ำและจำกัดความยาวไม่เกิน 31 ตัวอักษรตามข้อกำหนดของ Excel
                sheet_name = due_date[:31]
                base_name = sheet_name
                counter = 1
                while sheet_name in file_data_map:
                    sheet_name = f"{base_name}_{counter}"[:31]
                    counter += 1
                
                file_data_map[sheet_name] = pd.DataFrame(extracted_records)

    if file_data_map:
        st.success("✅ ดึงและจัดกลุ่มข้อมูลสำเร็จ!")
        
        # แสดงตัวอย่างข้อมูลตามแท็บหน้าเว็บ Streamlit
        tabs = st.tabs(list(file_data_map.keys()))
        for tab, (s_name, df_data) in zip(tabs, file_data_map.items()):
            with tab:
                st.write(f"### 📑 รายการสำหรับแท็บ: **{s_name}**")
                st.dataframe(df_data, use_container_width=True)
        
        # สร้างไฟล์ Excel และจัดสไตล์อย่างเป็นมืออาชีพ
        excel_buffer = io.BytesIO()
        
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            for s_name, df_data in file_data_map.items():
                df_data.to_excel(writer, index=False, sheet_name=s_name, startrow=4)
                
                worksheet = writer.sheets[s_name]
                worksheet.views.sheetView[0].showGridLines = True
                
                # Title Header
                worksheet.merge_cells("A1:D1")
                title_cell = worksheet["A1"]
                title_cell.value = "💳 รายงานสรุปรายการใช้บัตรเครดิต"
                title_cell.font = Font(name="Calibri", size=15, bold=True, color="1F4E78")
                title_cell.alignment = Alignment(horizontal="left", vertical="center")
                
                worksheet.merge_cells("A2:D2")
                sub_cell = worksheet["A2"]
                sub_cell.value = f"กำหนดชำระ / แท็บอ้างอิง: {s_name}"
                sub_cell.font = Font(name="Calibri", size=10, italic=True, color="595959")
                sub_cell.alignment = Alignment(horizontal="left", vertical="center")
                
                # Style Header (Row 5)
                header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                header_align = Alignment(horizontal="center", vertical="center")
                
                thin_border = Border(
                    left=Side(style='thin', color='D9D9D9'),
                    right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'),
                    bottom=Side(style='thin', color='D9D9D9')
                )
                
                worksheet.row_dimensions[5].height = 26
                headers = ["วันที่ใช้บัตร", "วันที่บันทึกรายการ", "รายการ", "จำนวนเงิน"]
                for col_idx, text in enumerate(headers, start=1):
                    cell = worksheet.cell(row=5, column=col_idx, value=text)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = header_align
                    cell.border = thin_border
                
                # Style Data Rows
                zebra_fill = PatternFill(start_color="F2F7FA", end_color="F2F7FA", fill_type="solid")
                total_row_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                
                max_r = 5 + len(df_data)
                for r_idx in range(6, max_r + 1):
                    worksheet.row_dimensions[r_idx].height = 20
                    is_last_row = (r_idx == max_r)
                    
                    worksheet.cell(row=r_idx, column=1).alignment = Alignment(horizontal="center", vertical="center")
                    worksheet.cell(row=r_idx, column=2).alignment = Alignment(horizontal="center", vertical="center")
                    worksheet.cell(row=r_idx, column=3).alignment = Alignment(horizontal="left", vertical="center")
                    
                    amt_cell = worksheet.cell(row=r_idx, column=4)
                    amt_cell.alignment = Alignment(horizontal="right", vertical="center")
                    
                    # แปลงข้อความตัวเลขให้เป็น Number Format สำหรับคำนวณสูตรต่อได้
                    try:
                        clean_val = str(amt_cell.value).replace(',', '').replace('฿', '').strip()
                        amt_cell.value = float(clean_val)
                        amt_cell.number_format = '#,##0.00'
                    except:
                        pass
                    
                    for c_idx in range(1, 5):
                        cell = worksheet.cell(row=r_idx, column=c_idx)
                        cell.border = thin_border
                        if is_last_row:
                            cell.fill = total_row_fill
                            cell.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
                        elif r_idx % 2 == 1:
                            cell.fill = zebra_fill
                
                # Width Adjustment
                worksheet.column_dimensions['A'].width = 16
                worksheet.column_dimensions['B'].width = 18
                worksheet.column_dimensions['C'].width = 45
                worksheet.column_dimensions['D'].width = 20

        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel ตกแต่งสวยงาม (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name="statement_formatted_by_due_date.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ ไม่พบข้อมูลรายการใช้บัตรในไฟล์ PDF ที่อัปโหลด กรุณาตรวจสอบไฟล์อีกครั้ง")
