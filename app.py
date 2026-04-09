import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# ==========================================
# CẤU HÌNH TRANG WEB
# ==========================================
st.set_page_config(page_title="Tạo Luận văn Chuẩn", page_icon="📝", layout="centered")

st.title("📝 Trình Tạo Luận Văn Chuẩn")
st.write("Nhập hoặc dán nội dung vào các ô bên dưới. Hệ thống sẽ tự động ghép lại và tạo ra một file Word hoàn chỉnh, chuẩn lề, chuẩn font 100%.")
st.divider()

# ==========================================
# GIAO DIỆN NHẬP LIỆU
# ==========================================
st.header("1. Thông tin chung")
thesis_title = st.text_input("Tên Đề tài Luận văn:", placeholder="Ví dụ: SO SÁNH ĐỘNG LỰC HỌC...")
author_name = st.text_input("Họ và tên tác giả:", placeholder="Ví dụ: TRẦN MINH HOÀNG")

st.header("2. Cấu trúc Nội dung")

# --- BỔ SUNG PHẦN ĐẶT VẤN ĐỀ ---
st.subheader("ĐẶT VẤN ĐỀ")
dat_van_de_content = st.text_area("Nội dung phần Đặt vấn đề (Dán văn bản vào đây):", height=250, placeholder="Nêu lý do chọn đề tài, tính cấp thiết, mục tiêu nghiên cứu...")

st.divider()

# --- PHẦN CÁC CHƯƠNG ---
st.info("💡 Chọn số lượng Chương. Hệ thống sẽ tạo ra các ô nhập liệu tương ứng bên dưới.")
num_chapters = st.number_input("Số lượng Chương (Từ Chương 1 trở đi):", min_value=1, max_value=10, value=3, step=1)

chapters_data = []

# Vòng lặp tạo giao diện nhập liệu động cho từng Chương và Mục
for i in range(num_chapters):
    st.subheader(f"CHƯƠNG {i+1}")
    chap_title = st.text_input(f"Tên CHƯƠNG {i+1}:", key=f"chap_title_{i}", placeholder="Ví dụ: TỔNG QUAN TÀI LIỆU")
    
    num_sections = st.number_input(f"Số lượng mục con trong CHƯƠNG {i+1}:", min_value=1, max_value=20, value=2, step=1, key=f"num_sec_{i}")
    
    sections_data = []
    for j in range(num_sections):
        with st.expander(f"Mục {i+1}.{j+1}", expanded=True):
            sec_title = st.text_input(f"Tên mục {i+1}.{j+1}:", key=f"sec_title_{i}_{j}", placeholder="Ví dụ: Tác vụ ngồi sang đứng")
            sec_content = st.text_area(f"Nội dung mục {i+1}.{j+1} (Dán văn bản vào đây):", height=200, key=f"sec_content_{i}_{j}")
            
            sections_data.append({
                "title": sec_title,
                "content": sec_content
            })
            
    chapters_data.append({
        "title": chap_title,
        "sections": sections_data
    })

st.divider()

# ==========================================
# NÚT XỬ LÝ VÀ TẠO FILE WORD
# ==========================================
if st.button("✨ TẠO FILE WORD HOÀN CHỈNH", type="primary", use_container_width=True):
    if not thesis_title:
        st.warning("⚠️ Vui lòng nhập Tên Đề tài trước khi tạo file!")
    else:
        with st.spinner("Hệ thống đang biên dịch và dàn trang file Word..."):
            # 1. Khởi tạo file Word trắng
            doc = docx.Document()
            
            # 2. Căn lề khổ giấy chuẩn
            for section in doc.sections:
                section.top_margin = Cm(3.5)
                section.bottom_margin = Cm(3.0)
                section.left_margin = Cm(3.5)
                section.right_margin = Cm(2.0)

            # 3. Ép chuẩn Normal Style
            style_normal = doc.styles['Normal']
            style_normal.font.name = 'Times New Roman'
            style_normal.font.size = Pt(13)

            # 4. Ghi tiêu đề ảo lên đầu file (Sinh viên có thể copy trang bìa thật vào sau)
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
            
            doc.add_page_break() # Sang trang mới bắt đầu nội dung

            # --- 4.5. DÀN TRANG PHẦN ĐẶT VẤN ĐỀ ---
            if dat_van_de_content.strip():
                # Tiêu đề ĐẶT VẤN ĐỀ
                p_dvd_title = doc.add_paragraph()
                p_dvd_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_dvd_title = p_dvd_title.add_run("ĐẶT VẤN ĐỀ")
                r_dvd_title.bold = True
                r_dvd_title.font.name = 'Times New Roman'
                r_dvd_title.font.size = Pt(14)
                
                # Nội dung Đặt vấn đề
                paragraphs_dvd = dat_van_de_content.split('\n')
                for para_text in paragraphs_dvd:
                    if para_text.strip():
                        p_text = doc.add_paragraph()
                        p_text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        p_text.paragraph_format.line_spacing = 1.5
                        p_text.paragraph_format.first_line_indent = Cm(1.27)
                        
                        r_text = p_text.add_run(para_text.strip())
                        r_text.font.name = 'Times New Roman'
                        r_text.font.size = Pt(13)
                
                # Ngắt trang sau khi xong Đặt vấn đề để Chương 1 nhảy sang trang mới
                doc.add_page_break()

            # 5. Dàn trang từng Chương
            for i, chap in enumerate(chapters_data):
                # Tên Chương (In đậm, viết hoa, ra giữa)
                p_chap = doc.add_paragraph()
                p_chap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_chap = p_chap.add_run(f"CHƯƠNG {i+1}: {chap['title']}".upper())
                r_chap.bold = True
                r_chap.font.name = 'Times New Roman'
                r_chap.font.size = Pt(14)
                
                # Các Mục con
                for j, sec in enumerate(chap['sections']):
                    # Tên Mục (Ví dụ: 1.1. Tác vụ ngồi sang đứng)
                    p_sec = doc.add_paragraph()
                    p_sec.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    r_sec = p_sec.add_run(f"{i+1}.{j+1}. {sec['title']}")
                    r_sec.bold = True
                    r_sec.font.name = 'Times New Roman'
                    r_sec.font.size = Pt(13)
                    
                    # Nội dung Mục (Cắt theo dấu Enter để phân đoạn)
                    paragraphs = sec['content'].split('\n')
                    for para_text in paragraphs:
                        if para_text.strip(): # Bỏ qua các dòng trống
                            p_text = doc.add_paragraph()
                            p_text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            p_text.paragraph_format.line_spacing = 1.5
                            p_text.paragraph_format.first_line_indent = Cm(1.27)
                            
                            r_text = p_text.add_run(para_text.strip())
                            r_text.font.name = 'Times New Roman'
                            r_text.font.size = Pt(13)
                
                # Hết một chương thì qua trang mới (Trừ chương cuối cùng)
                if i < len(chapters_data) - 1:
                    doc.add_page_break()

            # 6. Lưu file vào bộ nhớ và cho tải về
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("🎉 Đã tạo file thành công! Bạn có thể tải về ngay bên dưới.")
            st.download_button(
                label="⬇️ TẢI FILE LUẬN VĂN (.docx)",
                data=bio.getvalue(),
                file_name="Luan_Van_Hoan_Chinh.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
