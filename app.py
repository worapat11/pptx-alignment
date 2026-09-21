import os
import tempfile
import streamlit as st
import align_like_reference  # ดึงฟังก์ชันมาจากไฟล์สคริปต์ของคุณ

st.set_page_config(page_title="PPTX Auto Alignment Tool", layout="wide")

st.title("🧩 PPTX Auto Alignment Tool")
st.write("อัปโหลดไฟล์ PowerPoint เพื่อจัดตำแหน่งรูป/ชื่อ/รหัส/ช่องสี ให้ตรงกับหน้าต้นแบบโดยอัตโนมัติ")

# 1. รับไฟล์จากผู้ใช้
uploaded_file = st.file_uploader("เลือกไฟล์ PowerPoint (.pptx)", type=["pptx"])

# 2. ตั้งค่าการทำงาน
col1, col2 = st.columns(2)
with col1:
    ref_page = st.number_input("เลขหน้าต้นแบบ (Ref Page)", min_value=1, value=1, step=1, help="นับจาก 1")
    targets_input = st.text_input("หน้าปลายทางที่ต้องการจัด (เช่น 2-5 หรือ 2,4,6)", value="", help="ถ้าเว้นว่างไว้จะทำทุกหน้าที่เหลือ")
    alias_file = st.file_uploader("ไฟล์ CSV สำหรับจับคู่รหัสเปลี่ยนเลข (Optional)", type=["csv"])

with col2:
    code_regex = st.text_input("Custom Code Regex (Optional)", value="", help="เว้นว่างไว้ถ้าใช้รูปแบบปกติ (เช่น 53207-K12-V00)")
    no_frame = st.checkbox("ไม่ปรับกรอบนอกให้เท่าต้นแบบ", value=False)
    remove_extra_frames = st.checkbox("ลบกรอบใหญ่ที่เกินมา", value=False)
    keep_name_wrap = st.checkbox("คงการขึ้น 2 บรรทัดของชื่ออะไหล่ไว้", value=False)
    dry_run = st.checkbox("Dry-run (ประมวลผลทดลอง ไม่สร้างไฟล์จริง)", value=False)

# 3. ปุ่มประมวลผล
if uploaded_file is not None and st.button("เริ่มจัดตำแหน่ง PPTX", type="primary"):
    with st.spinner("กำลังประมวลผลไฟล์ PowerPoint..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, uploaded_file.name)
            output_path = os.path.join(tmpdir, "aligned_" + uploaded_file.name)
            
            # เซฟไฟล์ที่อัปโหลดลง temp
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # เซฟไฟล์ Alias CSV ถ้ามี
            alias_path = None
            if alias_file is not None:
                alias_path = os.path.join(tmpdir, alias_file.name)
                with open(alias_path, "wb") as f:
                    f.write(alias_file.getbuffer())

            # จำลอง Class Arguments ให้ส่งเข้าไปใน run_file ได้
            class Args:
                ref = ref_page
                targets = targets_input.strip() if targets_input.strip() else None
                out = output_path
                alias = alias_path
                code_regex = code_regex.strip() if code_regex.strip() else None
                no_frame = no_frame
                remove_extra_frames = remove_extra_frames
                keep_name_wrap = keep_name_wrap
                dry_run = dry_run

            try:
                # เรียกใช้ฟังก์ชันประมวลผลจากไฟล์ align_like_reference.py
                summary = align_like_reference.run_file(input_path, Args(), output_path)
                
                st.success(f"ประมวลผลเสร็จสิ้น! ทำไปทั้งหมด {summary['pages']} หน้า, พบ {summary['found']} ชิ้น, จับคู่สำเร็จ {summary['matched']} ชิ้น")
                
                if summary['unmatched'] > 0:
                    st.warning(f"มีชิ้นที่ไม่สามารถจับคู่ได้ {summary['unmatched']} ชิ้น")

                # ปุ่มดาวน์โหลดไฟล์
                if not dry_run and os.path.exists(output_path):
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="ดาวน์โหลดไฟล์ที่จัดเรียบร้อยแล้ว",
                            data=f,
                            file_name="aligned_" + uploaded_file.name,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")