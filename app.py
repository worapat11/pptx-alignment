import os
import tempfile
import streamlit as st
import align_like_reference  # ดึงฟังก์ชันมาจากไฟล์ align_like_reference.py

st.set_page_config(page_title="PPTX Auto Alignment Tool", layout="centered")

st.title("PPTX Auto Alignment Tool")
st.write("อัปโหลดไฟล์ PowerPoint เพื่อจัดตำแหน่งรูป/ชื่อ/รหัส/ช่องสี ให้ตรงกับหน้าต้นแบบโดยอัตโนมัติ")

# 1. รับไฟล์ PPTX จากผู้ใช้
uploaded_file = st.file_uploader("เลือกไฟล์ PowerPoint (.pptx)", type=["pptx"])

# 2. ตั้งค่าเฉพาะตัวหลักที่จำเป็น
col1, col2 = st.columns(2)
with col1:
    ref_page = st.number_input("เลขหน้าต้นแบบ (Ref Page)", min_value=1, value=1, step=1, help="นับจาก 1")
with col2:
    targets_input = st.text_input("หน้าปลายทางที่ต้องการจัด (เช่น 2-5 หรือ 2,4,6)", value="", help="ถ้าเว้นว่างไว้จะทำทุกหน้าที่เหลือ")

# 3. ปุ่มประมวลผล
if uploaded_file is not None and st.button("เริ่มจัดตำแหน่ง PPTX", type="primary"):
    with st.spinner("กำลังประมวลผลไฟล์ PowerPoint..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, uploaded_file.name)
            output_path = os.path.join(tmpdir, "aligned_" + uploaded_file.name)
            
            # บันทึกไฟล์อัปโหลดลงโฟลเดอร์ชั่วคราว
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # กำหนดค่า default สำหรับฟังก์ชันที่ถูกซ่อนไว้
            class Args:
                ref = ref_page
                targets = targets_input.strip() if targets_input.strip() else None
                out = output_path
                alias = None                  # ซ่อนการใส่ CSV
                code_regex = None             # ซ่อนการปรับ Regex
                no_frame = False              # ปรับกรอบนอกตามปกติ
                remove_extra_frames = False   # ไม่ลบกรอบเกิน
                keep_name_wrap = False        # ปรับชื่อเป็นบรรทัดเดียว
                dry_run = False               # สร้างไฟล์จริงเสมอ

            try:
                # เรียกใช้ฟังก์ชันประมวลผล
                summary = align_like_reference.run_file(input_path, Args(), output_path)
                
                st.success(f"ประมวลผลเสร็จสิ้น! ทำไปทั้งหมด {summary['pages']} หน้า, พบ {summary['found']} ชิ้น, จับคู่สำเร็จ {summary['matched']} ชิ้น")
                
                if summary['unmatched'] > 0:
                    st.warning(f"มีชิ้นที่ไม่สามารถจับคู่ได้ {summary['unmatched']} ชิ้น")

                # ปุ่มดาวน์โหลดไฟล์
                if os.path.exists(output_path):
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="ดาวน์โหลดไฟล์ที่จัดเรียบร้อยแล้ว",
                            data=f,
                            file_name="aligned_" + uploaded_file.name,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")