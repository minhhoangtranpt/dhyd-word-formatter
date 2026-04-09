import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re

# ==========================================
# HÀM LÀM SẠCH VÀ LẤY TÊN TIÊU ĐỀ GỐC
# ==========================================
def extract_title_name(text):
    """Xóa bỏ các chữ 'Chương 1', '1.1.' gõ tay, chỉ giữ lại nội dung chính."""
    # Xóa "Chương X:", "Chương X.", "Chương X "
    text = re.sub(r'^Chương\s+[0-9IVX]+\s*[\.\-\:]?\s*', '', text, flags=re.IGNORECASE)
    # Xóa các cụm số như "1.1.", "1.2.3", "1.1 "
    text = re.sub(r'^\d+(\.\d+)*\s*[\.\-\:]?\s*', '', text)
    return text.strip()

# ==========================================
# HÀM XỬ LÝ ĐỊNH DẠNG THEO CHUẨN ĐHYD TP.HCM
# ==========================================
def format_thesis_dhyd(doc):
    # 1. Căn lề khổ giấy
    for section in doc.sections:
        section.top_margin = Cm(3.5)
        section.bottom_margin = Cm(3.0)
        section.left_margin = Cm(3.5)
        section.right_margin = Cm(2.0)
        
    h1_num = 0
    h2_num = 0
    h3_num = 0
    
    # 2. Xử lý phần văn bản chính (Các đoạn văn)
    for paragraph in doc.paragraphs:
        full_text = paragraph.text.strip()
        if not full_text:
            continue
            
        style_name = paragraph.style.name

        # --- ĐIỀU KIỆN NHẬN DIỆN THÔNG MINH ---
        is_chapter = style_name == 'Heading 1' or re.match(r'^chương\s+[0-9ivx]+', full_text, re.IGNORECASE)
        is_h2 = style_name == 'Heading 2' or re.match(r'^\d+\.\d+[\.\s]', full_text)
        is_h3 = style_name == 'Heading 3' or re.match(r'^\d+\.\d+\.\d+[\.\s]', full_text)
        # Nhận diện tên Bảng hoặc Hình (Ví dụ: "Bảng 3.1.1.", "Hình 1.1")
        is_table_fig_title = re.match(r'^(Bảng|Hình|Biểu đồ)\s+\d+', full_text, re.IGNORECASE)

        # --- XỬ LÝ CHƯƠNG ---
        if is_chapter:
            h1_num += 1
            h2_num = 0
            h3_num = 0
            
            clean_title = extract_title_name(full_text)
            # Mẫu khóa luận dùng dấu hai chấm (:) sau số Chương
            new_text = f"CHƯƠNG {h1_num}: {clean_title}".upper()
            
            paragraph.style = 'Heading 1'
            paragraph.text = new_text
            
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Times New Roman'
                run.font.size = Pt(14)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Cm(0)

        # --- XỬ LÝ MỤC 1.1 ---
        elif is_h2:
            if h1_num == 0: h1_num = 1 
            h2_num += 1
            h3_num = 0
            
            clean_title = extract_title_name(full_text)
            new_text = f"{h1_num}.{h2_num}. {clean_title}"
            
            paragraph.style = 'Heading 2'
            paragraph.text = new_text
            
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Times New Roman'
                run.font.size = Pt(13)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.first_line_indent = Cm(0)
            
        # --- XỬ LÝ MỤC 1.1.1 ---
        elif is_h3:
            if h1_num == 0: h1_num = 1
            if h2_num == 0: h2_num = 1
            h3_num += 1
            
            clean_title = extract_title_name(full_text)
            new_text = f"{h1_num}.{h2_num}.{h3_num}. {clean_title}"
            
            paragraph.style = 'Heading 3'
            paragraph.text = new_text
            
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Times New Roman'
                run.font.size = Pt(13)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.first_line_indent = Cm(0)

        # --- XỬ LÝ TÊN BẢNG / HÌNH ẢNH ---
        elif is_table_fig_title:
            paragraph.style = 'Normal'
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Cm(0) # Không thụt đầu dòng
            paragraph.paragraph_format.line_spacing = 1.5
            for run in paragraph.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(13)
                run.font.bold = True # In đậm tên bảng

        # --- XỬ LÝ CHỮ THƯỜNG (CÒN LẠI) ---
        else:
            paragraph.style = 'Normal'
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.line_spacing = 1.5
            paragraph.paragraph_format.first_line_indent = Cm(1.27)
            
            for run in paragraph.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(13)

    # 3. XỬ LÝ CHỮ BÊN TRONG CÁC BẢNG BIỂU (TABLES)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER # Có thể đổi thành LEFT tùy ý
                    paragraph.paragraph_format.first_line_indent = Cm(0) # Tuyệt đối không thụt đầu dòng trong bảng
                    paragraph.paragraph_format.line_spacing = 1.0 # Giãn dòng đơn cho bảng đỡ bị phình to
                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12) # Chữ trong bảng nên nhỏ hơn một chút (size 11 hoặc 12)

    return doc

# ==========================================
# GIAO DIỆN WEB BẰNG STREAMLIT
# ==========================================
st.set_page_config(page_title="Định dạng Luận văn ĐHYD", page_icon="🎓")

st.title("🎓 Công cụ Hỗ trợ Định dạng Luận văn")
st.write("**Hệ thống tự động canh lề, ép font Times New Roman, xử lý bảng biểu và đánh số mục thông minh.**")

uploaded_file = st.file_uploader("Vui lòng tải file Word (.docx) của bạn lên đây:", type=["docx"])

if uploaded_file is not None:
    if st.button("✨ Bắt đầu chuẩn hóa định dạng", type="primary"):
        with st.spinner('Hệ thống đang phân tích và định dạng lại nội dung...'):
            try:
                doc = docx.Document(uploaded_file)
                formatted_doc = format_thesis_dhyd(doc)
                
                bio = io.BytesIO()
                formatted_doc.save(bio)
                
                st.success("🎉 Đã định dạng và đồng bộ phong cách thành công!")
                
                st.download_button(
                    label="⬇️ TẢI FILE ĐÃ ĐỊNH DẠNG VỀ MÁY",
                    data=bio.getvalue(),
                    file_name="Luan_Van_Da_Chuan_Hoa.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Có lỗi xảy ra trong quá trình xử lý: {e}")
