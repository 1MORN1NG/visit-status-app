import streamlit as st
import pandas as pd
import io
import zipfile
import re
import itertools
import os
from datetime import datetime

st.title("📊 ระบบตรวจสอบสถานะร้านค้า (Visit + Leave + Sell In)")

# ตัวเลือกโหมดการใช้งาน
mode = st.radio("📌 เลือกรูปแบบการใช้งาน", [
    "รวมเฉพาะ Visit",
    "รวม Visit + สรุปสถานะร้าน",
    "รวมไฟล์ Sell In Total (Excel)"
])

# โหมด: รวม Visit อย่างเดียว
if mode == "รวมเฉพาะ Visit":
    st.markdown("### 🔁 (ถ้ามี) อัปโหลดไฟล์ Visit Master ที่เคยรวมไว้")
    previous_file = st.file_uploader("📥 visit_merged.csv", type=["csv"])
    visit_files = st.file_uploader("1. อัปโหลดไฟล์ Visit (.csv) ใหม่เพิ่มเติม", type=["csv"], accept_multiple_files=True)

    if visit_files:
        if st.button("🔁 รวม Visit ใหม่"):
            with st.spinner("🚀 กำลังประมวลผล..."):
                visit_columns = [
                    "Id", "Number", "DATE", "UserName", "FirstName", "CustomerCOde", "Customer_Name",
                    "Customer_Location", "survey_updated_at", "เช็คอินหน้าร้าน (เซลฟี่หน้าร้าน)",
                    "ตรวจสอบตำแหน่งที่ตั้งร้าน", "ถ่ายรูปหน้าร้าน", "สถานะร้านค้า",
                    "กรณียกเลิกโครงการ โปรดระบุเหตุผลทุกครั้ง", "เช็คเอ้าหน้าร้าน (เซลฟี่หน้าร้าน)"
                ]

                all_visit_data = pd.DataFrame()
                for upload in visit_files:
                    filename = upload.name
                    df = pd.read_csv(upload, skiprows=2, usecols=range(15), encoding='utf-8-sig')
                    df.columns = visit_columns
                    df["source_file"] = filename

                    match = re.search(r'wk(\d{1,2})', filename.lower())
                    week_num = int(match.group(1)) if match else None
                    df["week"] = week_num

                    all_visit_data = pd.concat([all_visit_data, df], ignore_index=True)

                all_visit_data = all_visit_data.rename(columns={"CustomerCOde": "Customer_COde"})

                if previous_file:
                    previous = pd.read_csv(previous_file, encoding='utf-8-sig')
                    previous = previous.rename(columns={"CustomerCOde": "Customer_COde"})
                else:
                    previous = pd.DataFrame()

                visit_data = pd.concat([previous, all_visit_data], ignore_index=True).drop_duplicates()

                timestamp = datetime.now().strftime("%Y-%m-%d")
                filename = f"visit_merged_{timestamp}.csv"

                buffer = io.BytesIO()
                visit_data.to_csv(buffer, index=False, encoding="utf-8-sig")
                buffer.seek(0)

                st.success("✅ รวมไฟล์ Visit สำเร็จ!")
                st.download_button("📥 ดาวน์โหลด Visit รวมทั้งหมด", data=buffer, file_name=filename)

# โหมด:สรุปสถานะร้าน
if mode == "สรุปสถานะร้าน":
    visit_file = st.file_uploader("📥 visit_merged.csv (ที่รวมไว้แล้ว)", type=["csv"])
    master_file = st.file_uploader("📥 Master.xlsx", type=["xlsx"])
    leave_file = st.file_uploader("📥 Leave.xlsx", type=["xlsx"])

    if visit_file and master_file and leave_file:
        if st.button("📊 สรุปสถานะร้าน"):
            with st.spinner("🔍 กำลังประมวลผล..."):
                visit_data = pd.read_csv(visit_file, encoding='utf-8-sig')
                visit_data = visit_data.rename(columns={"CustomerCOde": "Customer_COde"})

                master_bkk = pd.read_excel(master_file, sheet_name="BKK")
                master_cnx = pd.read_excel(master_file, sheet_name="CNX")
                master_df = pd.concat([master_bkk, master_cnx], ignore_index=True)

                week_ref_preview = pd.read_excel(master_file, sheet_name="Week")
                week_ref = week_ref_preview[1:].copy()
                week_ref.columns = ["Year", "week", "Start_Date", "End_Date", "Monthnum", "Month", "Index"]
                week_ref["Start_Date"] = pd.to_datetime(week_ref["Start_Date"])
                week_ref["End_Date"] = pd.to_datetime(week_ref["End_Date"])
                week_ref["week"] = week_ref["week"].astype(int)

                leave_data = pd.read_excel(leave_file)
                leave_data = leave_data.rename(columns={"user": "User"})
                leave_data["Date"] = pd.to_datetime(leave_data["Date"], errors='coerce')

                def map_week_from_date(date):
                    matched = week_ref[(week_ref["Start_Date"] <= date) & (week_ref["End_Date"] >= date)]
                    if not matched.empty:
                        return int(matched["week"].values[0])
                    return None

                leave_data["mapped_week"] = leave_data["Date"].apply(map_week_from_date)

                user_to_store = master_df[["USER DE", "StoreCode1"]].dropna()
                user_to_store.columns = ["User", "Customer_COde"]
                leave_data = leave_data.merge(user_to_store, on="User", how="left")

                store_list = visit_data["Customer_COde"].dropna().unique()
                week_list = sorted(visit_data["week"].dropna().unique())
                base = pd.DataFrame(itertools.product(store_list, week_list), columns=["Customer_COde", "week"])

                visit_flag = visit_data[["Customer_COde", "week", "สถานะร้านค้า"]].drop_duplicates()
                visit_flag = visit_flag.rename(columns={"สถานะร้านค้า": "has_visit"})
                base = base.merge(visit_flag, on=["Customer_COde", "week"], how="left")

                leave_flag = leave_data[["Customer_COde", "mapped_week", "การลา"]].drop_duplicates()
                leave_flag = leave_flag.rename(columns={"mapped_week": "week"})
                base = base.merge(leave_flag, on=["Customer_COde", "week"], how="left")

                base_sorted = base.sort_values(by=["Customer_COde", "week"])

                def flag_cancel(row):
                    return "ยกเลิกโครงการ" if row["has_visit"] == "ยกเลิกโครงการ" else None

                base_sorted["cancel_flag"] = base_sorted.apply(flag_cancel, axis=1)

                def carry_cancel(df):
                    df = df.sort_values(by="week")
                    df["cancel_carried"] = df["cancel_flag"].eq("ยกเลิกโครงการ").cummax().replace({False: None, True: "ยกเลิกโครงการ"})
                    return df

                base_sorted = base_sorted.groupby("Customer_COde").apply(carry_cancel).reset_index(drop=True)

                def determine_status(row):
                    if row["has_visit"] == "ร้านเปิด":
                        return "ร้านเปิด"
                    elif row["cancel_carried"] == "ยกเลิกโครงการ":
                        return "ยกเลิกโครงการ"
                    elif pd.notna(row["การลา"]):
                        return row["การลา"]
                    else:
                        return "ขาดเยี่ยม"

                base_unique = base_sorted.drop_duplicates(subset=["Customer_COde", "week"])
                base_unique["status"] = base_unique.apply(determine_status, axis=1)

                pivot_df = base_unique.pivot(index="Customer_COde", columns="week", values="status")
                pivot_df.columns = [f"WK{int(c)}" for c in pivot_df.columns]
                pivot_df.reset_index(inplace=True)

                buffer = io.BytesIO()
                timestamp = datetime.now().strftime("%Y-%m-%d")
                pivot_df.to_csv(buffer, index=False, encoding="utf-8-sig")
                buffer.seek(0)

                filename = f"status_summary_{timestamp}.csv"
                st.success("✅ สรุปสถานะร้านเรียบร้อยแล้ว!")
                st.download_button("📥 ดาวน์โหลด status_summary.csv", data=buffer, file_name=filename)

# โหมด: รวม Sell In
if mode == "รวมไฟล์ Sell In Total (Excel)":
    st.markdown("### 🔁 (ถ้ามี) อัปโหลด Sell In Master ที่เคยรวมไว้")
    sellin_master_file = st.file_uploader("📥 sellin_total_master.xlsx", type=["xlsx"])

    sellin_files = st.file_uploader("📦 อัปโหลดไฟล์ Sell In (.xlsx) ใหม่เพิ่มเติม", type=["xlsx"], accept_multiple_files=True)

    if sellin_files and st.button("🔁 รวม Sell In"):
        with st.spinner("🧾 กำลังรวมข้อมูล..."):
            all_sheets = pd.DataFrame()

            if sellin_master_file:
                df_master = pd.read_excel(sellin_master_file)
                all_sheets = pd.concat([all_sheets, df_master], ignore_index=True)

            for f in sellin_files:
                df = pd.read_excel(f)
                all_sheets = pd.concat([all_sheets, df], ignore_index=True)

            buffer = io.BytesIO()
            all_sheets.to_excel(buffer, index=False, engine='xlsxwriter')
            buffer.seek(0)

            timestamp = datetime.now().strftime("%Y-%m-%d")
            filename = f"sell_in_total_{timestamp}.xlsx"

            st.success("✅ รวม Sell In สำเร็จ!")
            st.download_button("📥 ดาวน์โหลด Sell In รวมทั้งหมด", data=buffer, file_name=filename)
