import streamlit as st
import google.generativeai as genai
from docxtpl import DocxTemplate
import pandas as pd
import io
import json
from datetime import datetime
import docx

st.set_page_config(page_title="Hệ thống Sản xuất Phiếu trình", layout="wide")
st.title("📄 Trợ lý AI - Tự động hóa Phiếu trình & Danh mục công việc")

# 1. Khởi tạo API
api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    st.error("⚠️ Chưa cấu hình GEMINI_API_KEY trong Streamlit Secrets.")
    st.stop()

genai.configure(api_key=api_key)

# 2. GIẢI QUYẾT TRIỆT ĐỂ LỖI 404: Tự động dò tìm mô hình AI phù hợp nhất
@st.cache_resource
def get_best_model():
    try:
        models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        if 'models/gemini-1.5-flash' in models: return 'models/gemini-1.5-flash'
        if 'models/gemini-1.5-pro' in models: return 'models/gemini-1.5-pro'
        if 'models/gemini-pro' in models: return 'models/gemini-pro'
        return 'gemini-1.5-flash' # Mặc định nếu không dò được
    except Exception:
        return 'gemini-1.5-flash'

best_model_name = get_best_model()
model = genai.GenerativeModel(best_model_name)

# 3. Khởi tạo bảng dữ liệu
if 'danh_muc' not in st.session_state:
    st.session_state.danh_muc = pd.DataFrame(columns=["Ngày", "Số Phiếu", "Nội dung công việc", "Thời hạn", "Trạng thái"])

col1, col2 = st.columns([1, 1])

with col1:
    st.header("1. Tải tài liệu nguồn")
    st.info(f"🤖 Đang kết nối với AI: {best_model_name}")
    uploaded_file = st.file_uploader("Tải công văn, tờ trình... (DOCX, PDF, TXT)", type=["txt", "pdf", "docx"])

    if uploaded_file and st.button("🚀 Phân tích & Trích xuất"):
        with st.spinner("Hệ thống đang đọc tài liệu, vui lòng đợi vài giây..."):
            try:
                prompt = """
                Đọc tài liệu đính kèm và trích xuất thông tin:
                1. noi_dung_trinh: Tên sự việc chính cần trình Lãnh đạo.
                2. can_cu_phap_ly: Số hiệu, ngày, tên cơ quan của các văn bản liên quan.
                3. tom_tat_nhiem_vu: Tóm tắt ngắn gọn yêu cầu, nhiệm vụ (1-2 câu).
                4. thoi_han: Thời hạn hoàn thành (nếu có, không có ghi "Không quy định").
                5. de_xuat_giai_quyet: Tên văn bản đề xuất Lãnh đạo ký.
                Trả về DUY NHẤT chuỗi JSON hợp lệ theo format:
                {"noi_dung_trinh": "...", "can_cu_phap_ly": "...", "tom_tat_nhiem_vu": "...", "thoi_han": "...", "de_xuat_giai_quyet": "..."}
                """

                contents = []
                if uploaded_file.name.endswith(".docx"):
                    doc_reader = docx.Document(uploaded_file)
                    text_data = [para.text for para in doc_reader.paragraphs if para.text.strip()]
                    for table in doc_reader.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                if cell.text.strip(): text_data.append(cell.text.strip())
                    
                    file_content = "\n".join(text_data)
                    if not file_content.strip():
                        st.warning("⚠️ File Word này không chứa chữ (có thể là file scan dạng ảnh).")
                        st.stop()
                    contents = [file_content, prompt]

                elif uploaded_file.name.endswith(".txt"):
                    contents = [uploaded_file.read().decode("utf-8"), prompt]
                else:
                    contents = [{"mime_type": "application/pdf", "data": uploaded_file.read()}, prompt]

                # Xử lý kết quả AI
                response = model.generate_content(contents)
                raw_text = response.text
                if "```json" in raw_text:
                    raw_text = raw_text.split("```json")[1].split("```")[0]
                elif "```" in raw_text:
                    raw_text = raw_text.split("```")[1].split("```")[0]
                
                st.session_state.temp_data = json.loads(raw_text.strip())
                st.success("✅ Phân tích thành công! Xem kết quả ở bảng bên cạnh.")

            except Exception as e:
                st.error(f"❌ Có lỗi kỹ thuật: {str(e)}")

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
            try:
                now = datetime.now()
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
                
                new_row = {
                    "Ngày": now.strftime("%d/%m/%Y"), 
                    "Số Phiếu": so_phieu, 
                    "Nội dung công việc": noi_dung, 
                    "Thời hạn": thoi_han, 
                    "Trạng thái": "Đang trình ký"
                }
                st.session_state.danh_muc = pd.concat([st.session_state.danh_muc, pd.DataFrame([new_row])], ignore_index=True)
                
                st.success("✅ Đã tạo Phiếu trình thành công!")
                st.download_button("📥 Tải Phiếu Trình (.docx)", data=doc_io, file_name=f"Phieu_Trinh_{so_phieu}.docx")
            except Exception as e:
                st.error(f"❌ Lỗi khi xuất file: {str(e)}")

st.divider()
st.header("3. Danh mục công việc")
st.dataframe(st.session_state.danh_muc, use_container_width=True)

if not st.session_state.danh_muc.empty:
    excel_io = io.BytesIO()
    st.session_state.danh_muc.to_excel(excel_io, index=False)
    excel_io.seek(0)
    st.download_button("📊 Tải Báo cáo (Excel)", data=excel_io, file_name="Bao_Cao_Cong_Viec.xlsx")
