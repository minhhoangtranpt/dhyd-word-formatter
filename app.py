import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

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
        
    # 2. Định dạng Font Normal mặc định toàn bài
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)
    
    # 3. Quét qua các đoạn văn để ép chuẩn Giãn dòng và Canh lề
    for paragraph in doc.paragraphs:
        style_name = paragraph.style.name
        
        # Xử lý đoạn văn bản thường (Normal)
        if style_name == 'Normal':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # Căn lề hai bên
            paragraph.paragraph_format.line_spacing = 1.5    # Giãn dòng 1.5
            paragraph.paragraph_format.first_line_indent = Cm(1.27) # Thụt đầu dòng 1 tab
            
        # Xử lý Tiêu đề Chương (Heading 1)
        elif style_name == 'Heading 1':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER # Căn giữa khổ giấy
            paragraph.paragraph_format.first_line_indent = Cm(0) # Không thụt đầu dòng
            for run in paragraph.runs:
                run.font.bold = True
                run.text = run.text.upper() # Ép viết hoa
                
        # Xử lý Tiểu mục (Heading 2, 3...)
        elif style_name.startswith('Heading') and style_name != 'Heading 1':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # Căn lề đều 2 bên
            paragraph.paragraph_format.first_line_indent = Cm(0) # Không thụt dòng
            for run in paragraph.runs:
                run.font.bold = True # In đậm
                run.font.size = Pt(13) # Cỡ chữ 13
                
    return doc

# ==========================================
# GIAO DIỆN WEB BẰNG STREAMLIT
# ==========================================
st.set_page_config(page_title="Định dạng Luận văn ĐHYD", page_icon="🎓")

st.title("🎓 Công cụ Hỗ trợ Định dạng Luận văn")
st.write("**Tự động căn chỉnh file Word theo thể thức quy định của Đại học Y Dược TP.HCM.**")

# Tạo hộp thoại tải file lên
uploaded_file = st.file_uploader("Vui lòng tải file Word (.docx) của bạn lên đây:", type=["docx"])

if uploaded_file is not None:
    st.info("Đã ghi nhận file. Bấm nút bên dưới để hệ thống bắt đầu tự động căn chỉnh!")
    
    # Nút bấm để chạy lệnh
    if st.button("✨ Bắt đầu chuẩn hóa định dạng", type="primary"):
        with st.spinner('Hệ thống đang xử lý, vui lòng đợi...'):
            try:
                # Đọc file Word được tải lên
                doc = docx.Document(uploaded_file)
                
                # Gọi hàm xử lý định dạng
                formatted_doc = format_thesis_dhyd(doc)
                
                # Lưu file đã xử lý vào bộ nhớ tạm (RAM) để chuẩn bị tải xuống
                bio = io.BytesIO()
                formatted_doc.save(bio)
                
                st.success("🎉 Đã định dạng thành công!")
                
                # Hiển thị nút tải file về máy
                st.download_button(
                    label="⬇️ TẢI FILE ĐÃ ĐỊNH DẠNG VỀ MÁY",
                    data=bio.getvalue(),
                    file_name="Luan_Van_Da_Chuan_Hoa.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Có lỗi xảy ra trong quá trình xử lý: {e}")
