import streamlit as st
import pandas as pd
import json
import os
from datetime import date, timedelta

st.set_page_config(page_title="ระบบใบลา", page_icon="🌿", layout="centered")

DATA_FILE = "leaves.json"

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

leaves = load_data()

st.title("🌿 ระบบจัดการใบลา")

menu = st.sidebar.radio("เมนู", ["📝 ยื่นคำขอลา", "✅ อนุมัติใบลา", "📋 ประวัติการลา"])

if menu == "📝 ยื่นคำขอลา":
    st.header("ยื่นคำขอลา")
    with st.form("leave_form"):
        name = st.text_input("ชื่อ-นามสกุล")
        dept = st.text_input("แผนก")
        leave_type = st.selectbox("ประเภทการลา", ["ลาพักร้อน", "ลาป่วย", "ลากิจ", "อื่นๆ"])
        col1, col2 = st.columns(2)
        with col1:
            start = st.date_input("วันที่เริ่มลา", date.today())
        with col2:
            end = st.date_input("วันที่สิ้นสุด", date.today())
        days = (end - start).days + 1
        st.info(f"จำนวนวันลา: {days} วัน")
        reason = st.text_area("เหตุผล")
        approver = st.selectbox("ผู้อนุมัติ", ["วิชัย มั่นคง (ผู้จัดการ)", "พรทิพย์ สุขใจ (HR)"])
        submitted = st.form_submit_button("📤 ส่งคำขอลา")
        if submitted:
            if not name or not reason:
                st.error("กรุณากรอกข้อมูลให้ครบ")
            elif days <= 0:
                st.error("วันที่ไม่ถูกต้อง")
            else:
                leaves.append({
                    "id": len(leaves) + 1,
                    "name": name, "dept": dept,
                    "type": leave_type,
                    "start": str(start), "end": str(end), "days": days,
                    "reason": reason, "approver": approver,
                    "status": "รออนุมัติ", "note": ""
                })
                save_data(leaves)
                st.success("✅ ยื่นคำขอลาเรียบร้อยแล้ว!")

elif menu == "✅ อนุมัติใบลา":
    st.header("อนุมัติใบลา")
    pending = [l for l in leaves if l["status"] == "รออนุมัติ"]
    if not pending:
        st.info("ไม่มีรายการรออนุมัติ")
    for i, l in enumerate(pending):
        with st.expander(f"📋 {l['name']} — {l['type']} ({l['days']} วัน)"):
            st.write(f"**แผนก:** {l['dept']}")
            st.write(f"**วันที่:** {l['start']} ถึง {l['end']}")
            st.write(f"**เหตุผล:** {l['reason']}")
            note = st.text_input("หมายเหตุ", key=f"note_{i}")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ อนุมัติ", key=f"approve_{i}"):
                    idx = leaves.index(l)
                    leaves[idx]["status"] = "อนุมัติแล้ว"
                    leaves[idx]["note"] = note
                    save_data(leaves)
                    st.success("อนุมัติเรียบร้อย!")
                    st.rerun()
            with col2:
                if st.button("❌ ปฏิเสธ", key=f"reject_{i}"):
                    idx = leaves.index(l)
                    leaves[idx]["status"] = "ถูกปฏิเสธ"
                    leaves[idx]["note"] = note
                    save_data(leaves)
                    st.error("ปฏิเสธแล้ว")
                    st.rerun()

elif menu == "📋 ประวัติการลา":
    st.header("ประวัติการลา")
    if not leaves:
        st.info("ยังไม่มีประวัติการลา")
    else:
        df = pd.DataFrame(leaves)
        df = df[["name", "dept", "type", "start", "end", "days", "status", "reason"]]
        df.columns = ["ชื่อ", "แผนก", "ประเภท", "วันเริ่ม", "วันสิ้นสุด", "จำนวนวัน", "สถานะ", "เหตุผล"]
        status_filter = st.selectbox("กรองตามสถานะ", ["ทั้งหมด", "รออนุมัติ", "อนุมัติแล้ว", "ถูกปฏิเสธ"])
        if status_filter != "ทั้งหมด":
            df = df[df["สถานะ"] == status_filter]
        st.dataframe(df, use_container_width=True)
