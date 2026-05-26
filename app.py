import streamlit as st
import pandas as pd
import json
import os
from datetime import date

st.set_page_config(page_title="ระบบใบลา", page_icon="🌿", layout="centered")

# ===============================
# DATA FUNCTIONS
# ===============================
LEAVES_FILE = "leaves.json"
EMPLOYEES_FILE = "employees.json"

def load_leaves():
    if os.path.exists(LEAVES_FILE):
        with open(LEAVES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_leaves(data):
    with open(LEAVES_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_employees():
    if os.path.exists(EMPLOYEES_FILE):
        with open(EMPLOYEES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_employees(data):
    with open(EMPLOYEES_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def get_employee(emp_id):
    employees = load_employees()
    for e in employees:
        if e["รหัส"] == emp_id:
            return e
    return None

# ===============================
# SIDEBAR MENU
# ===============================
menu = st.sidebar.radio("เมนู", [
    "📝 ยื่นคำขอลา",
    "✅ อนุมัติใบลา",
    "📋 ประวัติการลา",
    "👥 จัดการพนักงาน (Admin)",
])

# ===============================
# PAGE: ยื่นคำขอลา
# ===============================
if menu == "📝 ยื่นคำขอลา":
    st.header("ยื่นคำขอลา")

    emp_id = st.text_input("กรอกรหัสพนักงาน")
    emp = None

    if emp_id:
        emp = get_employee(emp_id.strip())
        if emp:
            st.success(f"👤 {emp['ชื่อ']} {emp['นามสกุล']} — {emp['ตำแหน่ง']} ({emp['แผนก']})")
            st.info(f"สิทธิ์ลา: ลาพักร้อน {emp['ลาพักร้อน']} วัน | ลากิจ {emp['ลากิจ']} วัน | ลาป่วย {emp['ลาป่วย']} วัน")
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
                    leaves.append({
                        "id": len(leaves) + 1,
                        "รหัส": emp["รหัส"],
                        "ชื่อ": emp["ชื่อ"] + " " + emp["นามสกุล"],
                        "แผนก": emp["แผนก"],
                        "ตำแหน่ง": emp["ตำแหน่ง"],
                        "ประเภท": leave_type,
                        "วันเริ่ม": str(start),
                        "วันสิ้นสุด": str(end),
                        "จำนวนวัน": days,
                        "เหตุผล": reason,
                        "ผู้อนุมัติ": approver,
                        "สถานะ": "รออนุมัติ",
                        "หมายเหตุ": ""
                    })
                    save_leaves(leaves)
                    st.success("✅ ยื่นคำขอลาเรียบร้อยแล้ว!")

# ===============================
# PAGE: อนุมัติใบลา
# ===============================
elif menu == "✅ อนุมัติใบลา":
    st.header("อนุมัติใบลา")
    leaves = load_leaves()
    pending = [l for l in leaves if l["สถานะ"] == "รออนุมัติ"]

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
                    idx = leaves.index(l)
                    leaves[idx]["สถานะ"] = "อนุมัติแล้ว"
                    leaves[idx]["หมายเหตุ"] = note
                    save_leaves(leaves)
                    st.success("อนุมัติเรียบร้อย!")
                    st.rerun()
            with col2:
                if st.button("❌ ปฏิเสธ", key=f"reject_{i}"):
                    idx = leaves.index(l)
                    leaves[idx]["สถานะ"] = "ถูกปฏิเสธ"
                    leaves[idx]["หมายเหตุ"] = note
                    save_leaves(leaves)
                    st.error("ปฏิเสธแล้ว")
                    st.rerun()

# ===============================
# PAGE: ประวัติการลา
# ===============================
elif menu == "📋 ประวัติการลา":
    st.header("ประวัติการลา")
    leaves = load_leaves()

    if not leaves:
        st.info("ยังไม่มีประวัติการลา")
    else:
        df = pd.DataFrame(leaves)
        df = df[["ชื่อ", "แผนก", "ประเภท", "วันเริ่ม", "วันสิ้นสุด", "จำนวนวัน", "สถานะ", "เหตุผล"]]
        status_filter = st.selectbox("กรองตามสถานะ", ["ทั้งหมด", "รออนุมัติ", "อนุมัติแล้ว", "ถูกปฏิเสธ"])
        if status_filter != "ทั้งหมด":
            df = df[df["สถานะ"] == status_filter]
        st.dataframe(df, use_container_width=True)

# ===============================
# PAGE: จัดการพนักงาน (Admin)
# ===============================
elif menu == "👥 จัดการพนักงาน (Admin)":
    st.header("จัดการพนักงาน")
    employees = load_employees()

    # แสดงรายชื่อพนักงานทั้งหมด
    if employees:
        st.subheader("รายชื่อพนักงานในระบบ")
        df = pd.DataFrame(employees)
        st.dataframe(df, use_container_width=True)
    else:
        st.info("ยังไม่มีพนักงานในระบบ")

    st.divider()

    # เพิ่มพนักงานใหม่
    st.subheader("➕ เพิ่มพนักงานใหม่")
    with st.form("add_employee"):
        col1, col2 = st.columns(2)
        with col1:
            emp_id = st.text_input("รหัสพนักงาน เช่น 001")
            first_name = st.text_input("ชื่อ")
            position = st.text_input("ตำแหน่ง")
            start_date = st.date_input("วันเริ่มงาน")
        with col2:
            last_name = st.text_input("นามสกุล")
            dept = st.text_input("แผนก")
            annual = st.number_input("สิทธิ์ลาพักร้อน (วัน)", min_value=0, value=6)
            personal = st.number_input("สิทธิ์ลากิจ (วัน)", min_value=0, value=3)
            sick = st.number_input("สิทธิ์ลาป่วย (วัน)", min_value=0, value=30)

        add_btn = st.form_submit_button("➕ เพิ่มพนักงาน")
        if add_btn:
            if not emp_id or not first_name or not last_name:
                st.error("กรุณากรอกข้อมูลให้ครบ")
            elif get_employee(emp_id.strip()):
                st.error("รหัสพนักงานนี้มีอยู่แล้ว")
            else:
                employees.append({
                    "รหัส": emp_id.strip(),
                    "ชื่อ": first_name,
                    "นามสกุล": last_name,
                    "ตำแหน่ง": position,
                    "แผนก": dept,
                    "วันเริ่มงาน": str(start_date),
                    "ลาพักร้อน": annual,
                    "ลากิจ": personal,
                    "ลาป่วย": sick,
                })
                save_employees(employees)
                st.success(f"✅ เพิ่ม {first_name} {last_name} เรียบร้อยแล้ว!")
                st.rerun()

    # ลบพนักงาน
    if employees:
        st.divider()
        st.subheader("🗑 ลบพนักงาน")
        emp_list = [f"{e['รหัส']} — {e['ชื่อ']} {e['นามสกุล']}" for e in employees]
        selected = st.selectbox("เลือกพนักงานที่ต้องการลบ", emp_list)
        if st.button("🗑 ลบ"):
            del_id = selected.split(" — ")[0]
            employees = [e for e in employees if e["รหัส"] != del_id]
            save_employees(employees)
            st.success("ลบเรียบร้อยแล้ว!")
            st.rerun()
