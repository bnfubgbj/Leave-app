import streamlit as st
import pandas as pd
import json
import os
from datetime import date
from io import BytesIO

st.set_page_config(page_title="ระบบใบลา", page_icon="🌿", layout="centered")

LEAVES_FILE = "leaves.json"
EMPLOYEES_FILE = "employees.json"
ADMIN_PASSWORD = "admin1234"

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
    for e in load_employees():
        if e["รหัส"] == emp_id:
            return e
    return None

def get_used_days(emp_id, leave_type=None):
    leaves = load_leaves()
    result = [l for l in leaves if l["รหัส"] == emp_id and l["สถานะ"] == "อนุมัติแล้ว"]
    if leave_type:
        result = [l for l in result if l["ประเภท"] == leave_type]
    return sum(l["จำนวนวัน"] for l in result)

def export_summary_excel(employees, leaves):
    rows = []
    for e in employees:
        used_annual   = get_used_days(e["รหัส"], "ลาพักร้อน")
        used_personal = get_used_days(e["รหัส"], "ลากิจ")
        used_sick     = get_used_days(e["รหัส"], "ลาป่วย")
        rows.append({
            "รหัส": e["รหัส"],
            "ชื่อ-นามสกุล": e["ชื่อ"],
            "ตำแหน่ง": e["ตำแหน่ง"],
            "แผนก": e["แผนก"],
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

    df_summary = pd.DataFrame(rows)

    # สรุปรวม
    df_total = pd.DataFrame([{
        "รหัส": "",
        "ชื่อ-นามสกุล": f"รวมทั้งหมด {len(employees)} คน",
        "ตำแหน่ง": "",
        "แผนก": "",
        "สิทธิ์ลาพักร้อน": df_summary["สิทธิ์ลาพักร้อน"].sum(),
        "ลาพักร้อนที่ใช้": df_summary["ลาพักร้อนที่ใช้"].sum(),
        "ลาพักร้อนคงเหลือ": df_summary["ลาพักร้อนคงเหลือ"].sum(),
        "สิทธิ์ลากิจ": df_summary["สิทธิ์ลากิจ"].sum(),
        "ลากิจที่ใช้": df_summary["ลากิจที่ใช้"].sum(),
        "ลากิจคงเหลือ": df_summary["ลากิจคงเหลือ"].sum(),
        "สิทธิ์ลาป่วย": df_summary["สิทธิ์ลาป่วย"].sum(),
        "ลาป่วยที่ใช้": df_summary["ลาป่วยที่ใช้"].sum(),
        "ลาป่วยคงเหลือ": df_summary["ลาป่วยคงเหลือ"].sum(),
    }])

    df_final = pd.concat([df_summary, df_total], ignore_index=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="สรุปการลา")
    return output.getvalue()

menu = st.sidebar.radio("เมนู", [
    "📝 ยื่นคำขอลา",
    "✅ อนุมัติใบลา",
    "📋 ประวัติการลา",
    "👥 จัดการพนักงาน (Admin)",
])

# ===============================
if menu == "📝 ยื่นคำขอลา":
    st.header("ยื่นคำขอลา")
    emp_id = st.text_input("กรอกรหัสพนักงาน")
    emp = None
    if emp_id:
        emp = get_employee(emp_id.strip())
        if emp:
            st.success(f"👤 {emp['ชื่อ']} — {emp['ตำแหน่ง']} ({emp['แผนก']})")
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
                        "ชื่อ": emp["ชื่อ"],
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
elif menu == "👥 จัดการพนักงาน (Admin)":
    st.header("👥 จัดการพนักงาน")

    if "admin_logged_in" not in st.session_state:
        st.session_state.admin_logged_in = False

    if not st.session_state.admin_logged_in:
        st.subheader("🔒 กรุณาใส่รหัสผ่าน Admin")
        with st.form("admin_login"):
            password = st.text_input("รหัสผ่าน", type="password")
            if st.form_submit_button("เข้าสู่ระบบ"):
                if password == ADMIN_PASSWORD:
                    st.session_state.admin_logged_in = True
                    st.rerun()
                else:
                    st.error("❌ รหัสผ่านไม่ถูกต้อง")
        st.stop()

    col_title, col_logout = st.columns([4, 1])
    with col_logout:
        if st.button("🚪 ออกจากระบบ"):
            st.session_state.admin_logged_in = False
            st.rerun()

    employees = load_employees()
    leaves = load_leaves()

    # รายชื่อพนักงาน
    if employees:
        st.subheader("รายชื่อพนักงานในระบบ")
        st.dataframe(pd.DataFrame(employees), use_container_width=True)
    else:
        st.info("ยังไม่มีพนักงานในระบบ")

    # ปุ่มสรุป Excel
    if employees:
        st.divider()
        st.subheader("📊 สรุปการลา")
        excel_data = export_summary_excel(employees, leaves)
        st.download_button(
            label="📥 ดาวน์โหลดสรุปการลา (.xlsx)",
            data=excel_data,
            file_name=f"สรุปการลา_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()

    # เพิ่มพนักงาน
    st.subheader("➕ เพิ่มพนักงานใหม่")
    with st.form("add_employee"):
        col1, col2 = st.columns(2)
        with col1:
            emp_id   = st.text_input("รหัสพนักงาน เช่น 001")
            fullname = st.text_input("ชื่อ-นามสกุล เช่น นายไก่ ขันแต่เช้า")
            dept     = st.text_input("แผนก")
        with col2:
            position   = st.text_input("ตำแหน่ง")
            start_date = st.date_input("วันเริ่มงาน")
            annual     = st.number_input("สิทธิ์ลาพักร้อน (วัน)", min_value=0, value=6)
            personal   = st.number_input("สิทธิ์ลากิจ (วัน)", min_value=0, value=3)
            sick       = st.number_input("สิทธิ์ลาป่วย (วัน)", min_value=0, value=30)

        if st.form_submit_button("➕ เพิ่มพนักงาน"):
            if not emp_id or not fullname:
                st.error("กรุณากรอกรหัสและชื่อ-นามสกุล")
            elif get_employee(emp_id.strip()):
                st.error("รหัสพนักงานนี้มีอยู่แล้ว")
            else:
                employees.append({
                    "รหัส": emp_id.strip(),
                    "ชื่อ": fullname.strip(),
                    "ตำแหน่ง": position,
                    "แผนก": dept,
                    "วันเริ่มงาน": str(start_date),
                    "ลาพักร้อน": annual,
                    "ลากิจ": personal,
                    "ลาป่วย": sick,
                })
                save_employees(employees)
                st.success(f"✅ เพิ่ม {fullname} เรียบร้อยแล้ว!")
                st.rerun()

    # แก้ไขพนักงาน
    if employees:
        st.divider()
        st.subheader("✏️ แก้ไขข้อมูลพนักงาน")
        emp_list = [f"{e['รหัส']} — {e['ชื่อ']}" for e in employees]
        selected_edit = st.selectbox("เลือกพนักงานที่ต้องการแก้ไข", emp_list, key="edit_select")
        edit_id = selected_edit.split(" — ")[0]
        edit_emp = get_employee(edit_id)

        if edit_emp:
            with st.form("edit_employee"):
                col1, col2 = st.columns(2)
                with col1:
                    new_name     = st.text_input("ชื่อ-นามสกุล", value=edit_emp["ชื่อ"])
                    new_dept     = st.text_input("แผนก", value=edit_emp["แผนก"])
                with col2:
                    new_position = st.text_input("ตำแหน่ง", value=edit_emp["ตำแหน่ง"])
                    new_annual   = st.number_input("สิทธิ์ลาพักร้อน (วัน)", min_value=0, value=int(edit_emp["ลาพักร้อน"]))
                    new_personal = st.number_input("สิทธิ์ลากิจ (วัน)", min_value=0, value=int(edit_emp["ลากิจ"]))
                    new_sick     = st.number_input("สิทธิ์ลาป่วย (วัน)", min_value=0, value=int(edit_emp["ลาป่วย"]))

                if st.form_submit_button("💾 บันทึกการแก้ไข"):
                    for e in employees:
                        if e["รหัส"] == edit_id:
                            e["ชื่อ"]      = new_name
                            e["แผนก"]      = new_dept
                            e["ตำแหน่ง"]   = new_position
                            e["ลาพักร้อน"] = new_annual
                            e["ลากิจ"]     = new_personal
                            e["ลาป่วย"]    = new_sick
                    save_employees(employees)
                    st.success("✅ บันทึกเรียบร้อยแล้ว!")
                    st.rerun()

    # ลบพนักงาน
    if employees:
        st.divider()
        st.subheader("🗑 ลบพนักงาน")
        emp_list2 = [f"{e['รหัส']} — {e['ชื่อ']}" for e in employees]
        selected_del = st.selectbox("เลือกพนักงานที่ต้องการลบ", emp_list2, key="del_select")
        if st.button("🗑 ลบ"):
            del_id = selected_del.split(" — ")[0]
            employees = [e for e in employees if e["รหัส"] != del_id]
            save_employees(employees)
            st.success("ลบเรียบร้อยแล้ว!")
            st.rerun()
