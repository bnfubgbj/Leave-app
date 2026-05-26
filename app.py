import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import date
from io import BytesIO
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

st.set_page_config(page_title="ระบบใบลา", page_icon="🌿", layout="centered")

SHEET_ID = "1KR8adBRkkrEhjgY1lf4eTQ3rovLmcLYJdNCSFN827no"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ===============================
# Email Notification
# ===============================
def send_email_notification(leave, receiver=None):
    try:
        cfg = st.secrets["email"]
        sender   = cfg["sender"]
        password = cfg["password"].replace(" ", "")
        if not receiver:
            receiver = cfg["receiver"]

        type_map = {"ลาพักร้อน":"ลาพักร้อน","ลาป่วย":"ลาป่วย","ลากิจ":"ลากิจ","อื่นๆ":"อื่นๆ"}
        subject = f"[ระบบใบลา] {leave['ชื่อ']} ขอ{leave['ประเภท']} {leave['จำนวนวัน']} วัน รอการอนุมัติ"

        body = f"""
        <h2>📋 มีคำขอลาใหม่รอการอนุมัติ</h2>
        <table border="1" cellpadding="8" style="border-collapse:collapse;">
            <tr><td><b>พนักงาน</b></td><td>{leave['ชื่อ']}</td></tr>
            <tr><td><b>แผนก</b></td><td>{leave['แผนก']}</td></tr>
            <tr><td><b>ตำแหน่ง</b></td><td>{leave['ตำแหน่ง']}</td></tr>
            <tr><td><b>ประเภทการลา</b></td><td>{leave['ประเภท']}</td></tr>
            <tr><td><b>วันที่</b></td><td>{leave['วันเริ่ม']} ถึง {leave['วันสิ้นสุด']}</td></tr>
            <tr><td><b>จำนวนวัน</b></td><td>{leave['จำนวนวัน']} วัน</td></tr>
            <tr><td><b>เหตุผล</b></td><td>{leave['เหตุผล']}</td></tr>
        </table>
        <br>
        <p>กรุณาเข้าระบบเพื่ออนุมัติหรือปฏิเสธคำขอ</p>
        """

        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = sender
        msg["To"]      = receiver
        msg.attach(MIMEText(body, "html", "utf-8"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, receiver, msg.as_string())
        return True
    except Exception as e:
        st.warning(f"ส่ง Email ไม่สำเร็จ: {e}")
        return False

# ===============================
# Google Sheets Connection
# ===============================
@st.cache_resource
def get_client():
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def get_sheet(name):
    client = get_client()
    return client.open_by_key(SHEET_ID).worksheet(name)

# ===============================
# Employees
# ===============================
def load_employees():
    try:
        ws = get_sheet("employees")
        data = ws.get_all_records(numericise_ignore=["all"])
        for e in data:
            if "รหัส" in e:
                e["รหัส"] = str(e["รหัส"]).strip().zfill(4)
        return data
    except:
        return []

def save_employee(emp):
    ws = get_sheet("employees")
    employees = load_employees()
    # check if exists
    records = ws.get_all_values()
    if len(records) == 0:
        ws.append_row(["รหัส","ชื่อ","ตำแหน่ง","แผนก","วันเริ่มงาน","ลาพักร้อน","ลากิจ","ลาป่วย"])
    ws.append_row([
        emp["รหัส"], emp["ชื่อ"], emp["ตำแหน่ง"], emp["แผนก"],
        emp["วันเริ่มงาน"], emp["ลาพักร้อน"], emp["ลากิจ"], emp["ลาป่วย"],
        emp.get("รหัสหัวหน้า",""), emp.get("อีเมลหัวหน้า","")
    ])

def update_employee(emp_id, updated):
    ws = get_sheet("employees")
    records = ws.get_all_values()
    for i, row in enumerate(records):
        if row[0] == emp_id:
            ws.update(f"A{i+1}:H{i+1}", [[
                emp_id, updated["ชื่อ"], updated["ตำแหน่ง"], updated["แผนก"],
                row[4], updated["ลาพักร้อน"], updated["ลากิจ"], updated["ลาป่วย"]
            ]])
            break

def delete_employee(emp_id):
    ws = get_sheet("employees")
    records = ws.get_all_values()
    for i, row in enumerate(records):
        if row[0] == emp_id:
            ws.delete_rows(i + 1)
            break

def normalize_id(emp_id):
    return str(emp_id).strip().zfill(4)

def get_employee(emp_id):
    nid = normalize_id(emp_id)
    for e in load_employees():
        if normalize_id(str(e.get("รหัส",""))) == nid:
            return e
    return None

# ===============================
# Leaves
# ===============================
def load_leaves():
    try:
        ws = get_sheet("leaves")
        data = ws.get_all_records(numericise_ignore=["all"])
        return data
    except:
        return []

def save_leave(leave):
    ws = get_sheet("leaves")
    records = ws.get_all_values()
    if len(records) == 0:
        ws.append_row(["id","รหัส","ชื่อ","แผนก","ตำแหน่ง","ประเภท","วันเริ่ม","วันสิ้นสุด","จำนวนวัน","เหตุผล","ผู้อนุมัติ","สถานะ","หมายเหตุ"])
    ws.append_row([
        leave["id"], leave["รหัส"], leave["ชื่อ"], leave["แผนก"], leave["ตำแหน่ง"],
        leave["ประเภท"], leave["วันเริ่ม"], leave["วันสิ้นสุด"], leave["จำนวนวัน"],
        leave["เหตุผล"], leave["ผู้อนุมัติ"], leave["สถานะ"], leave["หมายเหตุ"]
    ])

def update_leave_status(leave_id, status, note):
    ws = get_sheet("leaves")
    records = ws.get_all_values()
    for i, row in enumerate(records):
        if i == 0:
            continue  # skip header
        if str(row[0]).strip() == str(leave_id).strip():
            ws.update_cell(i + 1, 12, status)
            ws.update_cell(i + 1, 13, note)
            break

def get_used_days(emp_id, leave_type=None):
    leaves = load_leaves()
    result = [l for l in leaves if str(l["รหัส"]) == str(emp_id) and l["สถานะ"] == "อนุมัติแล้ว"]
    if leave_type:
        result = [l for l in result if l["ประเภท"] == leave_type]
    return sum(int(l["จำนวนวัน"]) for l in result)

# ===============================
# Config
# ===============================
def load_config():
    try:
        ws = get_sheet("config")
        data = ws.get_all_records()
        config = {row["key"]: row["value"] for row in data}
        if not config:
            raise Exception("empty")
        return config
    except:
        return {
            "admin_password": "admin1234",
            "secret_question": "ชื่อบริษัทของคุณคืออะไร?",
            "secret_answer": "abc"
        }

def save_config(config):
    ws = get_sheet("config")
    ws.clear()
    ws.append_row(["key", "value"])
    for k, v in config.items():
        ws.append_row([k, v])

# ===============================
# Excel Export
# ===============================
def export_summary_excel(employees):
    rows = []
    for e in employees:
        used_annual   = get_used_days(str(e.get("รหัส","")), "ลาพักร้อน")
        used_personal = get_used_days(e["รหัส"], "ลากิจ")
        used_sick     = get_used_days(e["รหัส"], "ลาป่วย")
        rows.append({
            "รหัส": e["รหัส"], "ชื่อ-นามสกุล": e["ชื่อ"],
            "ตำแหน่ง": e["ตำแหน่ง"], "แผนก": e["แผนก"],
            "สิทธิ์ลาพักร้อน": e["ลาพักร้อน"],
            "ลาพักร้อนที่ใช้": used_annual,
            "ลาพักร้อนคงเหลือ": int(e["ลาพักร้อน"]) - used_annual,
            "สิทธิ์ลากิจ": e["ลากิจ"],
            "ลากิจที่ใช้": used_personal,
            "ลากิจคงเหลือ": int(e["ลากิจ"]) - used_personal,
            "สิทธิ์ลาป่วย": e["ลาป่วย"],
            "ลาป่วยที่ใช้": used_sick,
            "ลาป่วยคงเหลือ": int(e["ลาป่วย"]) - used_sick,
        })
    df = pd.DataFrame(rows)
    df_total = pd.DataFrame([{
        "รหัส": "",
        "ชื่อ-นามสกุล": f"รวมทั้งหมด {len(employees)} คน",
        "ตำแหน่ง": "",
        "แผนก": "",
        "สิทธิ์ลาพักร้อน": int(df["สิทธิ์ลาพักร้อน"].sum()),
        "ลาพักร้อนที่ใช้": int(df["ลาพักร้อนที่ใช้"].sum()),
        "ลาพักร้อนคงเหลือ": int(df["ลาพักร้อนคงเหลือ"].sum()),
        "สิทธิ์ลากิจ": int(df["สิทธิ์ลากิจ"].sum()),
        "ลากิจที่ใช้": int(df["ลากิจที่ใช้"].sum()),
        "ลากิจคงเหลือ": int(df["ลากิจคงเหลือ"].sum()),
        "สิทธิ์ลาป่วย": int(df["สิทธิ์ลาป่วย"].sum()),
        "ลาป่วยที่ใช้": int(df["ลาป่วยที่ใช้"].sum()),
        "ลาป่วยคงเหลือ": int(df["ลาป่วยคงเหลือ"].sum()),
    }])
    df_final = pd.concat([df, df_total], ignore_index=True)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="สรุปการลา")
    return output.getvalue()

# ===============================
# MENU
# ===============================
menu = st.sidebar.radio("เมนู", [
    "📝 ยื่นคำขอลา",
    "✅ อนุมัติใบลา",
    "📋 ประวัติการลา",
    "👥 จัดการพนักงาน (Admin)",
])

# ===============================
if menu == "📝 ยื่นคำขอลา":
    st.header("ยื่นคำขอลา")
    emp_id = st.text_input("กรอกรหัสพนักงาน เช่น 0001")
    emp = None
    if emp_id:
        emp = get_employee(emp_id.strip())
        if emp:
            carry = int(emp.get("วันสะสม", 0) or 0)
            total_annual = int(emp.get("ลาพักร้อน", 0)) + carry
            st.success(f"👤 {emp['ชื่อ']} — {emp['ตำแหน่ง']} ({emp['แผนก']})")
            st.info(f"สิทธิ์ลาพักร้อน: {emp['ลาพักร้อน']} วัน + สะสม {carry} วัน = **{total_annual} วัน** | ลากิจ {emp['ลากิจ']} วัน | ลาป่วย {emp['ลาป่วย']} วัน")
        else:
            st.error("ไม่พบรหัสพนักงานนี้ในระบบ")

    if emp:
        with st.form("leave_form"):
            leave_type = st.selectbox("ประเภทการลา", ["ลาพักร้อน", "ลาป่วย", "ลากิจ", "อื่นๆ"])
            col1, col2 = st.columns(2)
            with col1:
                start = st.date_input("วันที่เริ่มลา", date.today())
            with col2:
                end = st.date_input("วันที่สิ้นสุด", date.today())
            days = (end - start).days + 1
            st.info(f"จำนวนวันลา: {days} วัน")
            reason = st.text_area("เหตุผล")
            approver = st.selectbox("ผู้อนุมัติ", ["ผู้จัดการ", "HR"])
            submitted = st.form_submit_button("📤 ส่งคำขอลา")
            if submitted:
                if not reason:
                    st.error("กรุณาระบุเหตุผล")
                elif days <= 0:
                    st.error("วันที่ไม่ถูกต้อง")
                else:
                    leaves = load_leaves()
                    new_leave = {
                        "id": len(leaves) + 1,
                        "รหัส": emp["รหัส"], "ชื่อ": emp["ชื่อ"],
                        "แผนก": emp["แผนก"], "ตำแหน่ง": emp["ตำแหน่ง"],
                        "ประเภท": leave_type, "วันเริ่ม": str(start),
                        "วันสิ้นสุด": str(end), "จำนวนวัน": days,
                        "เหตุผล": reason, "ผู้อนุมัติ": approver,
                        "สถานะ": "รออนุมัติ", "หมายเหตุ": ""
                    }
                    save_leave(new_leave)
                    boss_email = emp.get("อีเมลหัวหน้า", "")
                    send_email_notification(new_leave, receiver=boss_email if boss_email else None)
                    st.success("✅ ยื่นคำขอลาเรียบร้อยแล้ว! แจ้งเตือนผู้บังคับบัญชาแล้ว 📧")

# ===============================
elif menu == "✅ อนุมัติใบลา":
    st.header("อนุมัติใบลา")

    # ให้ผู้อนุมัติกรอกรหัสพนักงานก่อน
    approver_id = st.text_input("กรอกรหัสพนักงานของคุณ เช่น 0001")
    approver_emp = None
    if approver_id:
        approver_emp = get_employee(approver_id.strip())
        if approver_emp:
            st.success(f"👤 {approver_emp['ชื่อ']} — {approver_emp['ตำแหน่ง']} ({approver_emp['แผนก']})")
        else:
            st.error("ไม่พบรหัสพนักงานนี้ในระบบ")

    if not approver_emp:
        st.stop()

    leaves = load_leaves()
    # แสดงเฉพาะรายการที่ส่งให้ผู้อนุมัตินี้
    pending = [l for l in leaves if l.get("สถานะ","") == "รออนุมัติ" and 
               (l.get("ผู้อนุมัติ","") == approver_emp.get("ตำแหน่ง","") or 
                l.get("รหัส","") != approver_id.strip())]
    
    # กรองไม่ให้เห็นใบลาของตัวเอง
    pending = [l for l in pending if normalize_id(str(l.get("รหัส",""))) != normalize_id(approver_id)]

    if not pending:
        st.info("ไม่มีรายการรออนุมัติ")
    for i, l in enumerate(pending):
        with st.expander(f"📋 {l['ชื่อ']} — {l['ประเภท']} ({l['จำนวนวัน']} วัน)"):
            st.write(f"**แผนก:** {l['แผนก']} | **ตำแหน่ง:** {l['ตำแหน่ง']}")
            st.write(f"**วันที่:** {l['วันเริ่ม']} ถึง {l['วันสิ้นสุด']}")
            st.write(f"**เหตุผล:** {l['เหตุผล']}")
            note = st.text_input("หมายเหตุ", key=f"note_{i}")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ อนุมัติ", key=f"approve_{i}"):
                    update_leave_status(l.get("id", l.get("ID", l.get("Id",""))), "อนุมัติแล้ว", note)
                    st.success("อนุมัติเรียบร้อย!")
                    st.rerun()
            with col2:
                if st.button("❌ ปฏิเสธ", key=f"reject_{i}"):
                    update_leave_status(l.get("id", l.get("ID", l.get("Id",""))), "ถูกปฏิเสธ", note)
                    st.error("ปฏิเสธแล้ว")
                    st.rerun()

# ===============================
elif menu == "📋 ประวัติการลา":
    st.header("ประวัติการลา")
    leaves = load_leaves()
    if not leaves:
        st.info("ยังไม่มีประวัติการลา")
    else:
        df = pd.DataFrame(leaves)
        df = df[["ชื่อ","แผนก","ประเภท","วันเริ่ม","วันสิ้นสุด","จำนวนวัน","สถานะ","เหตุผล"]]
        status_filter = st.selectbox("กรองตามสถานะ", ["ทั้งหมด","รออนุมัติ","อนุมัติแล้ว","ถูกปฏิเสธ"])
        if status_filter != "ทั้งหมด":
            df = df[df["สถานะ"] == status_filter]
        st.dataframe(df, use_container_width=True, hide_index=True)

# ===============================
elif menu == "👥 จัดการพนักงาน (Admin)":
    st.header("👥 จัดการพนักงาน")

    if "admin_logged_in" not in st.session_state:
        st.session_state.admin_logged_in = False

    if not st.session_state.admin_logged_in:
        config = load_config()
        tab1, tab2 = st.tabs(["🔒 เข้าสู่ระบบ", "🔑 ลืมรหัสผ่าน"])
        with tab1:
            with st.form("admin_login"):
                password = st.text_input("รหัสผ่าน", type="password")
                if st.form_submit_button("เข้าสู่ระบบ"):
                    if password == config["admin_password"]:
                        st.session_state.admin_logged_in = True
                        st.rerun()
                    else:
                        st.error("❌ รหัสผ่านไม่ถูกต้อง")
        with tab2:
            st.info(f"❓ {config['secret_question']}")
            with st.form("forgot_password"):
                answer  = st.text_input("คำตอบ")
                new_pw  = st.text_input("รหัสผ่านใหม่", type="password")
                new_pw2 = st.text_input("ยืนยันรหัสผ่านใหม่", type="password")
                if st.form_submit_button("รีเซ็ตรหัสผ่าน"):
                    if answer.strip().lower() != config["secret_answer"].strip().lower():
                        st.error("❌ คำตอบไม่ถูกต้อง")
                    elif new_pw != new_pw2:
                        st.error("❌ รหัสผ่านใหม่ไม่ตรงกัน")
                    elif len(new_pw) < 6:
                        st.error("❌ รหัสผ่านต้องมีอย่างน้อย 6 ตัว")
                    else:
                        config["admin_password"] = new_pw
                        save_config(config)
                        st.success("✅ รีเซ็ตรหัสผ่านเรียบร้อย!")
        st.stop()

    col_title, col_logout = st.columns([4, 1])
    with col_logout:
        if st.button("🚪 ออกจากระบบ"):
            st.session_state.admin_logged_in = False
            st.rerun()

    employees = load_employees()

    if employees:
        st.subheader("รายชื่อพนักงานในระบบ")
        st.dataframe(pd.DataFrame(employees), use_container_width=True, hide_index=True)
    else:
        st.info("ยังไม่มีพนักงานในระบบ")

    if employees:
        st.divider()
        st.subheader("📊 สรุปการลา")
        st.download_button(
            label="📥 ดาวน์โหลดสรุปการลา (.xlsx)",
            data=export_summary_excel(employees),
            file_name=f"สรุปการลา_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()
    st.subheader("➕ เพิ่มพนักงานใหม่")
    with st.form("add_employee"):
        col1, col2 = st.columns(2)
        with col1:
            emp_id   = st.text_input("รหัสพนักงาน เช่น 0001")
            fullname = st.text_input("ชื่อ-นามสกุล")
            dept     = st.text_input("แผนก")
        with col2:
            position   = st.text_input("ตำแหน่ง")
            start_date = st.date_input("วันเริ่มงาน")
            annual     = st.number_input("สิทธิ์ลาพักร้อน (วัน)", min_value=0, value=6)
            personal   = st.number_input("สิทธิ์ลากิจ (วัน)", min_value=0, value=3)
            sick       = st.number_input("สิทธิ์ลาป่วย (วัน)", min_value=0, value=30)
        boss_id          = st.text_input("รหัสหัวหน้า (ถ้ามี)")
        boss_email_input = st.text_input("อีเมลหัวหน้า (สำหรับแจ้งเตือน)")
        if st.form_submit_button("➕ เพิ่มพนักงาน"):
            if not emp_id or not fullname:
                st.error("กรุณากรอกรหัสและชื่อ-นามสกุล")
            elif get_employee(emp_id.strip()):
                st.error("❌ รหัสพนักงานนี้มีอยู่แล้ว")
            else:
                save_employee({
                    "รหัส": emp_id.strip().zfill(4), "ชื่อ": fullname.strip(),
                    "ตำแหน่ง": position, "แผนก": dept,
                    "วันเริ่มงาน": str(start_date),
                    "ลาพักร้อน": annual, "ลากิจ": personal, "ลาป่วย": sick,
                    "รหัสหัวหน้า": boss_id.strip().zfill(4) if boss_id.strip() else "",
                    "อีเมลหัวหน้า": boss_email_input.strip(),
                })
                st.success(f"✅ เพิ่ม {fullname} เรียบร้อยแล้ว!")
                st.rerun()

    if employees:
        st.divider()
        st.subheader("✏️ แก้ไขข้อมูลพนักงาน")
        emp_list = [f"{e['รหัส']} — {e['ชื่อ']}" for e in employees]
        selected_edit = st.selectbox("เลือกพนักงาน", emp_list, key="edit_select")
        edit_id = selected_edit.split(" — ")[0]
        edit_emp = get_employee(edit_id)
        if edit_emp:
            with st.form("edit_employee"):
                col1, col2 = st.columns(2)
                with col1:
                    new_name = st.text_input("ชื่อ-นามสกุล", value=edit_emp["ชื่อ"])
                    new_dept = st.text_input("แผนก", value=edit_emp["แผนก"])
                with col2:
                    new_position = st.text_input("ตำแหน่ง", value=edit_emp["ตำแหน่ง"])
                    new_annual   = st.number_input("สิทธิ์ลาพักร้อน", min_value=0, value=int(edit_emp["ลาพักร้อน"]))
                    new_personal = st.number_input("สิทธิ์ลากิจ", min_value=0, value=int(edit_emp["ลากิจ"]))
                    new_sick     = st.number_input("สิทธิ์ลาป่วย", min_value=0, value=int(edit_emp["ลาป่วย"]))
                if st.form_submit_button("💾 บันทึก"):
                    update_employee(edit_id, {
                        "ชื่อ": new_name, "แผนก": new_dept, "ตำแหน่ง": new_position,
                        "ลาพักร้อน": new_annual, "ลากิจ": new_personal, "ลาป่วย": new_sick,
                    })
                    st.success("✅ บันทึกเรียบร้อย!")
                    st.rerun()

    if employees:
        st.divider()
        st.subheader("🗑 ลบพนักงาน")
        emp_list2 = [f"{e['รหัส']} — {e['ชื่อ']}" for e in employees]
        selected_del = st.selectbox("เลือกพนักงาน", emp_list2, key="del_select")
        if st.button("🗑 ลบ"):
            del_id = selected_del.split(" — ")[0]
            delete_employee(del_id)
            st.success("ลบเรียบร้อย!")
            st.rerun()

    st.divider()
    st.subheader("🔑 เปลี่ยนรหัสผ่าน")
    with st.form("change_password"):
        old_pw  = st.text_input("รหัสผ่านเดิม", type="password")
        new_pw  = st.text_input("รหัสผ่านใหม่", type="password")
        new_pw2 = st.text_input("ยืนยันรหัสผ่านใหม่", type="password")
        if st.form_submit_button("🔑 เปลี่ยนรหัสผ่าน"):
            config = load_config()
            if old_pw != config["admin_password"]:
                st.error("❌ รหัสผ่านเดิมไม่ถูกต้อง")
            elif new_pw != new_pw2:
                st.error("❌ รหัสผ่านใหม่ไม่ตรงกัน")
            elif len(new_pw) < 6:
                st.error("❌ รหัสผ่านต้องมีอย่างน้อย 6 ตัว")
            else:
                config["admin_password"] = new_pw
                save_config(config)
                st.success("✅ เปลี่ยนรหัสผ่านเรียบร้อย!")

    st.divider()
    st.subheader("❓ ตั้งคำถามลับ")
    config = load_config()
    with st.form("change_secret"):
        new_question = st.text_input("คำถามลับ", value=config["secret_question"])
        new_answer   = st.text_input("คำตอบ", value=config["secret_answer"])
        if st.form_submit_button("💾 บันทึกคำถามลับ"):
            if not new_question or not new_answer:
                st.error("กรุณากรอกให้ครบ")
            else:
                config["secret_question"] = new_question
                config["secret_answer"]   = new_answer
                save_config(config)
                st.success("✅ บันทึกคำถามลับเรียบร้อย!")
