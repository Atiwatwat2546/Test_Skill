#RUN APP > streamlit run app.py


import warnings
import streamlit as st  # ใช้สร้าง Web App แบบง่ายด้วย Python
import pandas as pd     # ใช้สำหรับจัดการข้อมูลตาราง (DataFrame)
import os               # ใช้สำหรับจัดการไฟล์และ path ต่างๆ ในระบบ


# ------------------------------
# อ่านไฟล์ Daily Report (.xlsx)
# ------------------------------
def read_daily_reports(folder):
    all_data = []  # สร้าง list เพื่อเก็บข้อมูลจากทุกไฟล์รวมกัน
    for filename in os.listdir(folder):  # วนลูปทุกไฟล์ในโฟลเดอร์
        if filename.startswith("Daily report") and filename.endswith(".xlsx"):
            path = os.path.join(folder, filename)  # สร้าง path เต็มของไฟล์
            try:
                df = pd.read_excel(path, engine='openpyxl')  # อ่านไฟล์ Excel
                # ดึงชื่อทีมจากชื่อไฟล์ เช่น Daily report_TeamA_John_Doe.xlsx → John Doe
                parts = filename.replace(".xlsx", "").split("_")
                team_member = f"{parts[-2]} {parts[-1]}"
                df["Team Member"] = team_member  # เพิ่มคอลัมน์ชื่อทีมในข้อมูล

                df.columns = [col.strip() for col in df.columns]  # ลบช่องว่างหน้าหลังหัวตาราง
                all_data.append(df)  # เพิ่มข้อมูลเข้า list
            except Exception as e:
                st.error(f"อ่านไฟล์ {filename} ไม่ได้: {e}")  # แจ้งเตือนถ้าอ่านไฟล์ไม่ได้
    if all_data:
        return pd.concat(all_data, ignore_index=True)  # รวมข้อมูลทั้งหมดเป็น DataFrame เดียว
    return pd.DataFrame()  # ถ้าไม่มีข้อมูล ให้คืน DataFrame เปล่า




# ------------------------------
# อ่านไฟล์ New Employee (.xlsx)
# ------------------------------
def read_new_employee(filepath):
    try:
        df = pd.read_excel(filepath, engine='openpyxl')  # อ่านไฟล์ Excel
        df.columns = [col.strip() for col in df.columns]  # ลบช่องว่างชื่อคอลัมน์
        return df
    except Exception as e:
        st.error(f"อ่านไฟล์พนักงานใหม่ไม่สำเร็จ: {e}")
        return pd.DataFrame()  # ถ้าอ่านไม่ได้ คืน DataFrame เปล่า




# ------------------------------
# ดึงเฉพาะคนที่สัมภาษณ์ผ่าน + Role จาก new_df
# ------------------------------
def get_passed_candidates_with_roles(daily_df, new_df):
    # กรองเฉพาะคนที่ Interview = "yes" และ Status = "pass"
    passed = daily_df[
        (daily_df["Interview"].str.strip().str.lower() == "yes") &
        (daily_df["Status"].str.strip().str.lower() == "pass")
    ].copy()

    passed["Candidate Name"] = passed["Candidate Name"].str.strip()  # ลบช่องว่างชื่อ

    # เตรียม lookup key เป็นชื่อแบบ lowercase เพื่อจับคู่แม่นยำขึ้น
    new_df["Employee Name"] = new_df["Employee Name"].str.strip()
    new_df["lookup_key"] = new_df["Employee Name"].str.lower()
    passed["lookup_key"] = passed["Candidate Name"].str.lower()

    # สร้าง dictionary สำหรับ map ชื่อ → Role และ Join Date
    name_to_role = dict(zip(new_df["lookup_key"], new_df["Role"]))
    name_to_join = dict(zip(new_df["lookup_key"], pd.to_datetime(new_df["Join Date"], errors='coerce')))

    # เพิ่มคอลัมน์ที่ต้องการในผลลัพธ์
    passed["Employee Name"] = passed["Candidate Name"]
    passed["Role"] = passed["lookup_key"].map(name_to_role)
    passed["Join Date"] = passed["lookup_key"].map(name_to_join)

    # คืนเฉพาะคอลัมน์ที่เราต้องการ
    return passed[["Employee Name", "Join Date", "Role", "Team Member"]]




# ------------------------------
# ฟังก์ชันแปลงคอลัมน์วันที่เป็นรูปแบบ dd-MMM-YYYY
# ------------------------------
def format_dates(df):
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%d-%b-%Y")
        elif df[col].dtype == 'object':
            try:
                # ปิด warning ชั่วคราวระหว่างการแปลง
                with warnings.catch_warnings():
                    warnings.simplefilter(action='ignore', category=UserWarning)
                    parsed = pd.to_datetime(df[col], errors='coerce')  # หรือเพิ่ม format=...
                if parsed.notna().sum() > 0:
                    df[col] = parsed.dt.strftime("%d-%b-%Y")
            except:
                continue
    return df





# ------------------------------
# เริ่มต้น Streamlit App
# ------------------------------
st.title("📋 Dashboard")  # หัวเรื่องหลักของแอป

# ชื่อโฟลเดอร์และไฟล์
daily_folder = "daily_reports"  # โฟลเดอร์ที่เก็บ Daily Report
new_employee_file = "new_employee/New Employee_YYYYMM.xlsx"  # โฟลเดอร์ไฟล์ข้อมูลพนักงานใหม่


# โหลดข้อมูล
daily_df = read_daily_reports(daily_folder)  # โหลดรายงานจากหลายทีม
new_df = read_new_employee(new_employee_file)  # โหลดข้อมูลพนักงานใหม่

# แสดงข้อมูลดิบ
with st.expander("ข้อมูลจาก Daily Reports"):  # ขยาย/ย่อได้
    daily_df_display = format_dates(daily_df.copy())
    daily_df_display.index = range(1, len(daily_df_display) + 1)
    st.dataframe(daily_df_display)  # แสดงตาราง

with st.expander("ข้อมูลจาก New Employee"):
    new_df_display = format_dates(new_df.copy())
    new_df_display.index = range(1, len(new_df_display) + 1)
    st.dataframe(new_df_display)


# สร้างตารางผลลัพธ์
passed_df = get_passed_candidates_with_roles(daily_df, new_df)

st.header("รายชื่อพนักงานที่ผ่านสัมภาษณ์")
if not passed_df.empty:
    passed_df_display = format_dates(passed_df.copy())
    passed_df_display.index = range(1, len(passed_df_display) + 1)
    st.dataframe(passed_df_display)  # แสดงตารางผลลัพธ์
else:
    st.warning("ไม่พบพนักงานที่ผ่านสัมภาษณ์ในระบบ")  # ถ้าไม่มีคนผ่าน




