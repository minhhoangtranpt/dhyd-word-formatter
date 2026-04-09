import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re

# ==========================================
# HÀM LÀM SẠCH SỐ THỨ TỰ CŨ (TRÁNH LẶP SỐ)
# ==========================================
def clean_heading_text(text):
    """Xóa các đánh số thủ công cũ ở đầu dòng (VD: 'Chương 1.', '1.1 ', '1.1.2. ')"""
    # Xóa chữ "Chương X" hoặc các cụm số như "1.1.", "1.2.3" nằm ở đầu câu
    cleaned = re.sub(r'^(Chương\s*\d+\.?|\d+(\.\d+)*)\s*[\.\-\:]?\s*', '', text, flags=re.IGNORECASE)
    return cleaned.strip()

# ==========================================
# HÀM XỬ LÝ ĐỊNH DẠNG THEO CHUẨN ĐHYD TP.HCM
# ==========================================
def format_thesis_dhyd(doc):
    # 1. Định dạng Lề trang giấy (3.5cm trên/trái, 3.0cm dưới, 2.0cm phải)
    for section in doc.sections:
        section.top_margin = Cm(3.5)
        section.bottom_margin = Cm(3.0)
        section.left_margin = Cm(3.5)
        section.right_margin = Cm(2.0)
        
    # 2. Khởi tạo các biến đếm để đánh số tự động
    h1_num = 0 # Đếm số Chương
    h2_num = 0 # Đếm số Mục (1.1, 1.2)
    h3_num = 0 # Đếm số Tiểu mục (1.1.1, 1.1.2)
    
    # 3. Quét qua toàn bộ đoạn văn trong file
    for paragraph in doc.paragraphs:
        style_name = paragraph.style.name
        
        # --- ÉP 100% VỀ FONT TIMES NEW ROMAN ---
        # Quét qua từng đoạn chữ nhỏ để ép font (Xóa bỏ lỗi sai font do copy-paste)
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            
        # Lấy toàn bộ nội dung chữ của dòng hiện tại
        full_text = "".join([run.text for run in paragraph.runs])
        if not full_text.strip():
            continue # Bỏ qua các dòng trống không có chữ

        # --- XỬ LÝ TIÊU ĐỀ CHƯƠNG (HEADING 1) ---
        if style_name == 'Heading 1':
            h1_num += 1
            h2_num = 0 # Chuyển sang chương mới thì reset đếm mục con về 0
            h3_num = 0
            
            clean_text = clean_heading_text(full_text)
            new_text = f"Chương {h1_num}. {clean_text}".upper() # VD: CHƯƠNG 1. TỔNG QUAN
            
            # Xóa các khối chữ cũ, thay bằng chữ mới đã đánh số
            for run in paragraph.runs: run.text = "" 
            new_run = paragraph.add_run(new_text)
            new_run.font.bold = True
            new_run.font.name = 'Times New Roman'
            new_run.font.size = Pt(14)
            
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Cm(0)

        # --- XỬ LÝ TIỂU MỤC CẤP 1 (HEADING 2) ---
        elif style_name == 'Heading 2':
            if h1_num == 0: h1_num = 1 # Đề phòng file gốc quên chọn Heading 1 cho chương đầu
            h2_num += 1
            h3_num = 0 # Reset tiểu mục nhỏ
            
            clean_text = clean_heading_text(full_text)
            new_text = f"{h1_num}.{h2_num}. {clean_text}" # VD: 1.1. Đặt vấn đề
            
            for run in paragraph.runs: run.text = ""
            new_run = paragraph.add_run(new_text)
            new_run.font.bold = True
            new_run.font.name = 'Times New Roman'
            new_run.font.size = Pt(13)
            
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.first_line_indent = Cm(0)
            
        # --- XỬ LÝ TIỂU MỤC CẤP 2 (HEADING 3) ---
        elif style_name == 'Heading 3':
            if h1_num == 0: h1_num = 1
            if h2_num == 0: h2_num = 1
            h3_num += 1
            
            clean_text = clean_heading_text(full_text)
            new_text = f"{h1_num}.{h2_num}.{h3_num}. {clean_text}" # VD: 1.1.1. Bối cảnh
            
            for run in paragraph.runs: run.text = ""
            new_run = paragraph.add_run(new_text)
            new_run.font.bold = True
            new_run.font.name = 'Times New Roman'
            new_run.font.size = Pt(13)
            
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.first_line_indent = Cm(0)

        # --- XỬ LÝ VĂN BẢN THƯỜNG (NORMAL) ---
        elif style_name == 'Normal':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.line_spacing = 1.5
            paragraph.paragraph_format.first_line_indent = Cm(1.27)
            # Ép luôn kích cỡ 13 cho toàn bộ chữ thường
            for run in paragraph.runs:
                run.font.size = Pt(13)
                
    return doc

# ==========================================
# GIAO DIỆN WEB BẰNG STREAMLIT
# ==========================================
st.set_page_config(page_title="Định dạng Luận văn ĐHYD", page_icon="🎓")

st.title("🎓 Công cụ Hỗ trợ Định dạng Luận văn")
st.write("**Tự động căn chỉnh file Word, ép font Times New Roman và đánh số mục tự động.**")

uploaded_file = st.file_uploader("Vui lòng tải file Word (.docx) của bạn lên đây:", type=["docx"])

if uploaded_file is not None:
    st.info("Đã ghi nhận file. Bấm nút bên dưới để hệ thống bắt đầu tự động căn chỉnh!")
    
    if st.button("✨ Bắt đầu chuẩn hóa định dạng", type="primary"):
        with st.spinner('Hệ thống đang phân tích và định dạng lại nội dung...'):
            try:
                doc = docx.Document(uploaded_file)
                formatted_doc = format_thesis_dhyd(doc)
                
                bio = io.BytesIO()
                formatted_doc.save(bio)
                
                st.success("🎉 Đã định dạng và đánh số tự động thành công!")
                
                st.download_button(
                    label="⬇️ TẢI FILE ĐÃ ĐỊNH DẠNG VỀ MÁY",
                    data=bio.getvalue(),
                    file_name="Luan_Van_Da_Chuan_Hoa.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Có lỗi xảy ra trong quá trình xử lý: {e}")
