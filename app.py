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
st.write("Hệ thống tự động ép chuẩn lề, font, và tự động xử lý lỗi cấu trúc (xóa tiểu mục nếu mục cha chỉ có 1 mục con).")
st.divider()

# ==========================================
# HÀM ĐỆ QUY TẠO GIAO DIỆN NHẬP LIỆU (TỐI ĐA 4 CẤP)
# ==========================================
def render_section(level, prefix, key_prefix):
    """
    Tạo các ô nhập liệu lồng nhau. 
    level: Cấp độ hiện tại (2, 3, 4 tương ứng 1.1, 1.1.1, 1.1.1.1)
    """
    with st.container(border=True):
        st.markdown(f"**Mục {prefix}**")
        title = st.text_input("Tên mục:", key=f"title_{key_prefix}", label_visibility="collapsed", placeholder=f"Tên mục {prefix}")
        content = st.text_area("Nội dung:", key=f"content_{key_prefix}", height=100, label_visibility="collapsed", placeholder=f"Nội dung mục {prefix}")
        
        children = []
        if level < 4: # Chỉ cho phép đi sâu tối đa 4 cấp
            num_children = st.number_input(f"Số tiểu mục con (cấp {level+1}) trong {prefix}:", min_value=0, max_value=15, value=0, step=1, key=f"num_{key_prefix}")
            for k in range(int(num_children)):
                child_prefix = f"{prefix}.{k+1}"
                children.append(render_section(level+1, child_prefix, f"{key_prefix}_{k}"))
                
    return {"title": title, "content": content, "children": children}

# ==========================================
# THUẬT TOÁN TỈA CÂY (ÉP LUẬT TỐI THIỂU 2 TIỂU MỤC)
# ==========================================
def apply_academic_rules(node):
    """
    Rà soát cây cấu trúc. Nếu một mục chỉ có 1 tiểu mục con:
    - Xóa tiêu đề của mục con, biến nó thành đoạn văn bản.
    - Đưa đoạn văn bản đó lên gộp vào mục cha.
    - Đẩy các tiểu mục cháu (nếu có) lên một cấp.
    """
    if node.get("children"):
        # 1. Đệ quy xử lý từ dưới lên trên (Bottom-up)
        pruned_children = [apply_academic_rules(c) for c in node["children"]]
        
        # 2. Áp dụng luật "Tối thiểu 2 tiểu mục"
        if len(pruned_children) == 1:
            single_child = pruned_children[0]
            
            # Chuyển đổi child thành văn bản và gộp vào cha
            merged_text = single_child["title"]
            if single_child["content"].strip():
                merged_text += "\n" + single_child["content"]
                
            if node.get("content", "").strip():
                node["content"] += "\n\n" + merged_text
            else:
                node["content"] = merged_text
                
            # Cha nhận nuôi luôn các cháu (thăng cấp cháu lên làm con)
            node["children"] = single_child["children"]
        else:
            node["children"] = pruned_children
            
    return node

# ==========================================
# GIAO DIỆN NHẬP LIỆU CHÍNH
# ==========================================
st.header("I. THÔNG TIN CHUNG")
thesis_title = st.text_input("Tên Đề tài Luận văn:", placeholder="Ví dụ: SO SÁNH ĐỘNG LỰC HỌC...")
author_name = st.text_input("Họ và tên tác giả:", placeholder="Ví dụ: TRẦN MINH HOÀNG")
st.divider()

st.header("II. NỘI DUNG LUẬN VĂN")

# 1. ĐẶT VẤN ĐỀ
st.subheader("ĐẶT VẤN ĐỀ")
dat_van_de_content = st.text_area("Nội dung phần Đặt vấn đề:", height=200, key="dvd")

# 2. BA CHƯƠNG CHÍNH
fixed_chapters = [
    "TỔNG QUAN TÀI LIỆU",
    "ĐỐI TƯỢNG VÀ PHƯƠNG PHÁP NGHIÊN CỨU",
    "KẾT QUẢ"
]

chapters_data = []

for i, chap_name in enumerate(fixed_chapters):
    st.markdown(f"### Chương {i+1}. {chap_name}")
    chap_content = st.text_area(f"Nội dung dẫn nhập Chương {i+1} (nếu có):", height=100, key=f"chap_intro_{i}")
    
    num_sections = st.number_input(f"Số lượng mục cấp 2 (ví dụ: {i+1}.1, {i+1}.2):", min_value=0, max_value=20, value=2, step=1, key=f"num_sec_l2_{i}")
    
    l2_children = []
    for j in range(num_sections):
        prefix = f"{i+1}.{j+1}"
        l2_children.append(render_section(2, prefix, f"chap_{i}_sec_{j}"))
            
    chapters_data.append({
        "title": chap_name, 
        "content": chap_content, 
        "children": l2_children
    })
    st.write("---")

# 3. CÁC PHẦN CỐ ĐỊNH CUỐI CÙNG
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
# CÁC HÀM HỖ TRỢ XUẤT FILE WORD
# ==========================================
def add_main_heading(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text.upper())
    r.bold = True
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)

def add_normal_text(doc, text_content):
    if not text_content.strip(): return
    for para_text in text_content.split('\n'):
        if para_text.strip():
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.first_line_indent = Cm(1.27)
            r = p.add_run(para_text.strip())
            r.font.name = 'Times New Roman'
            r.font.size = Pt(13)

def write_sections_to_word(doc, children_list, prefix_list):
    """Đệ quy ghi các mục vào Word, tự động đánh số lại sau khi đã tỉa cành"""
    for i, child in enumerate(children_list):
        current_prefix = prefix_list + [str(i + 1)]
        prefix_str = ".".join(current_prefix)
        
        if child["title"].strip() or child["content"].strip() or child["children"]:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(0) 
            
            title_text = f"{prefix_str}. {child['title']}" if child['title'].strip() else f"{prefix_str}."
            r = p.add_run(title_text)
            r.bold = True
            r.font.name = 'Times New Roman'
            r.font.size = Pt(13)
            
            add_normal_text(doc, child["content"])
            
            if child.get("children"):
                write_sections_to_word(doc, child["children"], current_prefix)

# ==========================================
# NÚT XỬ LÝ VÀ TẠO FILE WORD
# ==========================================
if st.button("✨ TẠO FILE WORD HOÀN CHỈNH", type="primary", use_container_width=True):
    if not thesis_title:
        st.warning("⚠️ Vui lòng nhập Tên Đề tài ở phần Thông tin chung!")
    else:
        with st.spinner("Đang áp dụng luật học thuật và dàn trang..."):
            # Áp dụng thuật toán tỉa cành cho 3 chương chính
            processed_chapters = [apply_academic_rules(chap) for chap in chapters_data]

            doc = docx.Document()
            for section in doc.sections:
                section.top_margin, section.bottom_margin = Cm(3.5), Cm(3.0)
                section.left_margin, section.right_margin = Cm(3.5), Cm(2.0)

            style_normal = doc.styles['Normal']
            style_normal.font.name, style_normal.font.size = 'Times New Roman', Pt(13)

            # Bìa ảo
            doc.add_paragraph().add_run(thesis_title.upper()).bold = True
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.paragraphs[-1].runs[0].font.size = Pt(16)
            
            doc.add_paragraph().add_run(author_name.upper()).bold = True
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_page_break()

            # Đặt vấn đề
            if dat_van_de_content.strip():
                add_main_heading(doc, "ĐẶT VẤN ĐỀ")
                add_normal_text(doc, dat_van_de_content)
                doc.add_page_break()

            # 3 Chương Chính
            for i, chap in enumerate(processed_chapters):
                add_main_heading(doc, f"CHƯƠNG {i+1}: {chap['title']}")
                add_normal_text(doc, chap['content'])
                write_sections_to_word(doc, chap['children'], [str(i+1)])
                doc.add_page_break()

            # Các phần cuối
            if ket_luan_content.strip():
                add_main_heading(doc, "KẾT LUẬN VÀ KIẾN NGHỊ")
                add_normal_text(doc, ket_luan_content)
                doc.add_page_break()

            if danh_muc_content.strip():
                add_main_heading(doc, "DANH MỤC CÁC CÔNG TRÌNH CÔNG BỐ CÓ LIÊN QUAN")
                add_normal_text(doc, danh_muc_content)
                doc.add_page_break()

            if tai_lieu_content.strip():
                add_main_heading(doc, "TÀI LIỆU THAM KHẢO")
                add_normal_text(doc, tai_lieu_content)
                doc.add_page_break()

            if phu_luc_content.strip():
                add_main_heading(doc, "PHỤ LỤC")
                add_normal_text(doc, phu_luc_content)

            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("🎉 Đã xuất file thành công!")
            st.download_button("⬇️ TẢI FILE LUẬN VĂN (.docx)", bio.getvalue(), "Luan_Van_Chuan_Format.docx", 
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
