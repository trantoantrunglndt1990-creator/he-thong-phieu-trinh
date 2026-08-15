import docx
import streamlit as st
import google.generativeai as genai
from docxtpl import DocxTemplate
import pandas as pd
import io
import json
from datetime import datetime

st.set_page_config(page_title="Hệ thống Sản xuất Phiếu trình", layout="wide")
st.title("📄 Trợ lý AI - Tự động hóa Phiếu trình & Danh mục công việc")

# 1. Cấu hình API
api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    st.warning("Vui lòng cấu hình GEMINI_API_KEY trong phần cài đặt của Streamlit.")
    st.stop()

genai.configure(api_key=api_key)
model = genai.GenerativeModel("models/gemini-1.5-flash")
# 2. Khởi tạo Bảng Danh mục công việc
if 'danh_muc' not in st.session_state:
    st.session_state.danh_muc = pd.DataFrame(columns=["Ngày", "Số Phiếu", "Nội dung công việc", "Thời hạn", "Trạng thái"])

col1, col2 = st.columns([1, 1])

with col1:
    st.header("1. Tải tài liệu nguồn")
    uploaded_file = st.file_uploader("Tải công văn, tờ trình... (PDF, TXT, DOCX)", type=["txt", "pdf", "docx"])

if uploaded_file and st.button("🚀 Phân tích & Trích xuất"):
    with st.spinner("Đang đọc tài liệu..."):
        prompt = """
        Đóng vai một chuyên viên hành chính mẫn cán. Đọc tài liệu đính kèm và trích xuất thông tin:
        1. noi_dung_trinh: Tên sự việc chính cần trình Lãnh đạo.
        2. can_cu_phap_ly: Liệt kê số hiệu, ngày, tên cơ quan của các văn bản liên quan.
        3. tom_tat_nhiem_vu: Tóm tắt ngắn gọn yêu cầu, nhiệm vụ (1-2 câu).
        4. thoi_han: Thời hạn hoàn thành (nếu có, không có ghi "Không quy định").
        5. de_xuat_giai_quyet: Tên văn bản đề xuất Lãnh đạo ký.
        Trả về DUY NHẤT một chuỗi JSON hợp lệ: 
        {"noi_dung_trinh": "...", "can_cu_phap_ly": "...", "tom_tat_nhiem_vu": "...", "thoi_han": "...", "de_xuat_giai_quyet": "..."}
        """

        # Xử lý riêng nếu file tải lên là Word (.docx)
        if uploaded_file.name.endswith(".docx"):
               doc_reader = docx.Document(uploaded_file)
               
               # Lấy chữ từ các đoạn văn bình thường
               text_data = [para.text for para in doc_reader.paragraphs if para.text.strip()]
               
               # Lấy thêm chữ từ các bảng biểu (nếu có)
               for table in doc_reader.tables:
                   for row in table.rows:
                       for cell in row.cells:
                           if cell.text.strip():
                               text_data.append(cell.text.strip())
                               
               file_content = "\n".join(text_data)
               
               # Cảnh báo nếu file Word trống
               if not file_content:
                   st.error("⚠️ File Word của bạn đang trống hoặc hệ thống không thể đọc được chữ bên trong!")
               else:
                   response = model.generate_content([file_content, prompt])
with col2:
    st.header("2. Hoàn thiện Phiếu trình")
    if 'temp_data' in st.session_state:
        so_phieu = st.text_input("Số Phiếu trình:", value="")
        noi_dung = st.text_area("Nội dung trình:", value=st.session_state.temp_data.get("noi_dung_trinh", ""))
        can_cu = st.text_area("Căn cứ pháp lý:", value=st.session_state.temp_data.get("can_cu_phap_ly", ""))
        tom_tat = st.text_area("Tóm tắt nhiệm vụ:", value=st.session_state.temp_data.get("tom_tat_nhiem_vu", ""))
        thoi_han = st.text_input("Thời hạn xử lý:", value=st.session_state.temp_data.get("thoi_han", ""))
        de_xuat = st.text_area("Đề xuất giải quyết:", value=st.session_state.temp_data.get("de_xuat_giai_quyet", ""))

        if st.button("💾 Xuất Phiếu Trình & Lưu Báo Cáo"):
            now = datetime.now()
            
            # Điền vào file Word
            doc = DocxTemplate("template.docx")
            doc.render({
                "so_phieu": so_phieu,
                "ngay": now.strftime("%d"),
                "thang": now.strftime("%m"),
                "nam": now.strftime("%Y"),
                "noi_dung_trinh": noi_dung,
                "can_cu_phap_ly": can_cu,
                "tom_tat_nhiem_vu": tom_tat,
                "thoi_han": thoi_han,
                "de_xuat_giai_quyet": de_xuat
            })
            
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            
            # Cập nhật danh mục
            new_row = {
                "Ngày": now.strftime("%d/%m/%Y"), 
                "Số Phiếu": so_phieu, 
                "Nội dung công việc": noi_dung, 
                "Thời hạn": thoi_han, 
                "Trạng thái": "Đang trình ký"
            }
            st.session_state.danh_muc = pd.concat([st.session_state.danh_muc, pd.DataFrame([new_row])], ignore_index=True)
            
            st.download_button("📥 Tải Phiếu Trình (.docx)", data=doc_io, file_name=f"Phieu_Trinh_{so_phieu}.docx")

st.divider()
st.header("3. Danh mục công việc thực hiện (Phục vụ báo cáo tuần/tháng)")
st.dataframe(st.session_state.danh_muc, use_container_width=True)

if not st.session_state.danh_muc.empty:
    excel_io = io.BytesIO()
    st.session_state.danh_muc.to_excel(excel_io, index=False)
    excel_io.seek(0)
    st.download_button("📊 Tải file Danh mục (Excel)", data=excel_io, file_name="Danh_Muc_Cong_Viec.xlsx")
