import streamlit as st
import pandas as pd
import io
import zipfile

st.title("📊 ระบบตรวจสอบสถานะร้านค้า (Visit + Leave)")

# อัปโหลดไฟล์
master_file = st.file_uploader("1. อัปโหลดไฟล์ Master (.xlsx)", type=["xlsx"])
leave_file = st.file_uploader("2. อัปโหลดไฟล์ Leave (.xlsx)", type=["xlsx"])
visit_files = st.file_uploader("3. อัปโหลดไฟล์ Visit (.csv) หลายไฟล์", type=["csv"], accept_multiple_files=True)

if master_file and leave_file and visit_files:
    if st.button("🔁 ประมวลผลและรวมไฟล์"):
        with st.spinner("กำลังประมวลผลข้อมูล..."):
            master_df = pd.read_excel(master_file)
            leave_df = pd.read_excel(leave_file)

            # รวมไฟล์ Visit
            visit_all = [pd.read_csv(f) for f in visit_files]
            visit_df = pd.concat(visit_all, ignore_index=True)
            st.write(\"Master ตัวอย่าง:\", master_df.head())
            st.write(\"Leave ตัวอย่าง:\", leave_df.head())
            st.write(\"Visit ตัวอย่าง:\", visit_df.head())

            # วิเคราะห์สถานะ
            weeks = sorted(visit_df['Week'].unique())
            customer_list = master_df['Customer_code'].unique()
            cancelled_codes = visit_df[visit_df['VisitDetail'].str.contains("ยกเลิกโครงการ", na=False)]['Customer_code'].unique()

            result = []
            for code in customer_list:
                for week in weeks:
                    visit_check = visit_df[(visit_df['Customer_code'] == code) & (visit_df['Week'] == week)]
                    leave_check = leave_df[(leave_df['Customer_code'] == code) & (leave_df['Week'] == week)]

                    if not visit_check.empty:
                        status = "เยี่ยมแล้ว"
                    elif code in cancelled_codes:
                        status = "ยกเลิกโครงการ"
                    elif not leave_check.empty:
                        status = "ลา"
                    else:
                        status = "ขาดเยี่ยม"

                    result.append({"Customer_code": code, "Week": week, "Status": status})

            status_df = pd.DataFrame(result)
            pivot_df = status_df.pivot(index="Customer_code", columns="Week", values="Status").reset_index()

            # บันทึกไฟล์เป็น zip
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                visit_bytes = io.BytesIO()
                visit_df.to_csv(visit_bytes, index=False)
                zip_file.writestr("visit_merged.csv", visit_bytes.getvalue())

                pivot_bytes = io.BytesIO()
                pivot_df.to_csv(pivot_bytes, index=False)
                zip_file.writestr("status_summary.csv", pivot_bytes.getvalue())

            zip_buffer.seek(0)
            st.success("✅ ประมวลผลเสร็จสิ้น!")
            st.download_button("📥 ดาวน์โหลดผลลัพธ์ (ZIP)", data=zip_buffer, file_name="visit_status_output.zip")
