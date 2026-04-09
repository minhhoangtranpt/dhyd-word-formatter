import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# ==========================================
# CẤU HÌNH TRANG WEB
# ==========================================
st.set_page_config(page_title="Tạo Luận văn Chuẩn ĐHYD", page_icon="🎓", layout="centered")

st.title("🎓 Trình Tạo Luận Văn Chuẩn")
st.write("Điền nội dung vào các phần được cố định sẵn bên dưới. Hệ thống sẽ tự động tạo ra file Word hoàn chỉnh chuẩn form.")
st.divider()

# ==========================================
# GIAO DIỆN NHẬP LIỆU THEO KHUNG CỐ ĐỊNH
# ==========================================
st.header("I. THÔNG TIN CHUNG")
thesis_title = st.text_input("Tên Đề tài Luận văn:", placeholder="Ví dụ: SO SÁNH ĐỘNG LỰC HỌC...")
author_name = st.text_input("Họ và tên tác giả:", placeholder="Ví dụ: TRẦN MINH HOÀNG")
st.divider()

st.header("II. NỘI DUNG LUẬN VĂN")

# 1. ĐẶT VẤN ĐỀ
st.subheader("ĐẶT VẤN ĐỀ")
dat_van_de_content = st.text_area("Nội dung phần Đặt vấn đề:", height=200, key="dvd")

# 2. BA CHƯƠNG CHÍNH (CÓ THỂ THÊM TIỂU MỤC)
fixed_chapters = [
    "TỔNG QUAN TÀI LIỆU",
    "ĐỐI TƯỢNG VÀ PHƯƠNG PHÁP NGHIÊN CỨU",
    "KẾT QUẢ"
]

chapters_data = []

for i, chap_name in enumerate(fixed_chapters):
    st.markdown(f"### Chương {i+1}. {chap_name}")
    num_sections = st.number_input(f"Số lượng mục con trong Chương {i+1}:", min_value=0, max_value=20, value=2, step=1, key=f"num_sec_{i}")
    
    sections_data = []
    for j in range(num_sections):
        with st.expander(f"Mục {i+1}.{j+1}", expanded=True):
            sec_title = st.text_input(f"Tên mục {i+1}.{j+1}:", key=f"sec_title_{i}_{j}")
            sec_content = st.text_area(f"Nội dung mục {i+1}.{j+1}:", height=150, key=f"sec_content_{i}_{j}")
            sections_data.append({"title": sec_title, "content": sec_content})
            
    chapters_data.append({"title": chap_name, "sections": sections_data})
    st.write("---")

# 3. CÁC PHẦN CỐ ĐỊNH CUỐI CÙNG (KHÔNG TIỂU MỤC)
st.subheader("KẾT LUẬN VÀ KIẾN NGHỊ")
ket_luan_content = st.text_area("Nội dung Kết luận và Kiến nghị:", height=200, key="kl")

st.subheader("DANH MỤC CÁC CÔNG TRÌNH CÔNG BỐ CÓ LIÊN QUAN")
danh_muc_content = st.text_area("Nội dung (Nếu không có hãy để trống):", height=150, key="dm")

st.subheader("TÀI LIỆU THAM KHẢO")
tai_lieu_content = st.text_area("Danh sách tài liệu tham khảo:", height=200, key="tl")

st.subheader("PHỤ LỤC")
phu_luc_content = st.text_area("Nội dung Phụ lục (Nếu có):", height=200, key="pl")

st.divider()

# ==========================================
# HÀM HỖ TRỢ XUẤT FILE WORD
# ==========================================
def add_main_heading(doc, text):
    """Hàm hỗ trợ thêm tiêu đề lớn (Căn giữa, In đậm, Size 14)"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text.upper())
    r.bold = True
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)

def add_normal_text(doc, text_content):
    """Hàm hỗ trợ thêm văn bản thường (Căn đều 2 bên, thụt đầu dòng, Size 13, giãn 1.5)"""
    if not text_content.strip():
        return
    paragraphs = text_content.split('\n')
    for para_text in paragraphs:
        if para_text.strip():
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.first_line_indent = Cm(1.27)
            r = p.add_run(para_text.strip())
            r.font.name = 'Times New Roman'
            r.font.size = Pt(13)

# ==========================================
# NÚT XỬ LÝ VÀ TẠO FILE WORD
# ==========================================
if st.button("✨ TẠO FILE WORD HOÀN CHỈNH", type="primary", use_container_width=True):
    if not thesis_title:
        st.warning("⚠️ Vui lòng nhập Tên Đề tài ở phần Thông tin chung trước khi tạo file!")
    else:
        with st.spinner("Hệ thống đang biên dịch và dàn trang file Word..."):
            doc = docx.Document()
            
            # Căn lề khổ giấy
            for section in doc.sections:
                section.top_margin = Cm(3.5)
                section.bottom_margin = Cm(3.0)
                section.left_margin = Cm(3.5)
                section.right_margin = Cm(2.0)

            # Ép chuẩn Normal Style mặc định
            style_normal = doc.styles['Normal']
            style_normal.font.name = 'Times New Roman'
            style_normal.font.size = Pt(13)

            # --- TRANG BÌA ẢO ---
            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_title = p_title.add_run(thesis_title.upper())
            r_title.bold = True
            r_title.font.name = 'Times New Roman'
            r_title.font.size = Pt(16)
            
            p_author = doc.add_paragraph()
            p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_author = p_author.add_run(author_name.upper())
            r_author.bold = True
            r_author.font.name = 'Times New Roman'
            
            doc.add_page_break()

            # --- ĐẶT VẤN ĐỀ ---
            if dat_van_de_content.strip():
                add_main_heading(doc, "ĐẶT VẤN ĐỀ")
                add_normal_text(doc, dat_van_de_content)
                doc.add_page_break()

            # --- 3 CHƯƠNG CHÍNH ---
            for i, chap in enumerate(chapters_data):
                # Tên Chương
                add_main_heading(doc, f"CHƯƠNG {i+1}. {chap['title']}")
                
                # Các Mục con
                for j, sec in enumerate(chap['sections']):
                    if sec['title'].strip() or sec['content'].strip():
                        # Tiêu đề mục (Ví dụ: 1.1. Tác vụ ngồi sang đứng)
                        p_sec = doc.add_paragraph()
                        p_sec.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        r_sec = p_sec.add_run(f"{i+1}.{j+1}. {sec['title']}")
                        r_sec.bold = True
                        r_sec.font.name = 'Times New Roman'
                        r_sec.font.size = Pt(13)
                        
                        # Nội dung mục
                        add_normal_text(doc, sec['content'])
                
                doc.add_page_break()

            # --- KẾT LUẬN VÀ KIẾN NGHỊ ---
            if ket_luan_content.strip():
                add_main_heading(doc, "KẾT LUẬN VÀ KIẾN NGHỊ")
                add_normal_text(doc, ket_luan_content)
                doc.add_page_break()

            # --- DANH MỤC CÔNG TRÌNH ---
            if danh_muc_content.strip():
                add_main_heading(doc, "DANH MỤC CÁC CÔNG TRÌNH CÔNG BỐ CÓ LIÊN QUAN")
                add_normal_text(doc, danh_muc_content)
                doc.add_page_break()

            # --- TÀI LIỆU THAM KHẢO ---
            if tai_lieu_content.strip():
                add_main_heading(doc, "TÀI LIỆU THAM KHẢO")
                add_normal_text(doc, tai_lieu_content)
                doc.add_page_break()

            # --- PHỤ LỤC ---
            if phu_luc_content.strip():
                add_main_heading(doc, "PHỤ LỤC")
                add_normal_text(doc, phu_luc_content)

            # Lưu và tải file
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("🎉 Đã xuất file thành công!")
            st.download_button(
                label="⬇️ TẢI FILE LUẬN VĂN (.docx)",
                data=bio.getvalue(),
                file_name="Luan_Van_Chuan_Format.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
