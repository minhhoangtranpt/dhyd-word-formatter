import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml.xmlchemy import OxmlElement
from docx.oxml.ns import qn  # <--- Dòng này đã được sửa lại cho chuẩn
import io

# ==========================================
# CẤU HÌNH TRANG WEB
# ==========================================
st.set_page_config(page_title="Tạo Đề Cương Luận Văn", page_icon="🎓", layout="centered")

st.title("🎓 Trình Tạo Đề Cương Luận Văn Chuẩn")
st.write("Hệ thống tự động dàn trang bìa có khung viền và ép chuẩn font toàn bộ nội dung.")
st.divider()

# ==========================================
# CÁC HÀM CAN THIỆP XML (ĐÓNG KHUNG TRANG BÌA)
# ==========================================
def add_page_border(sect_pr):
    """Vẽ khung viền cho trang"""
    borders = OxmlElement('w:pgBorders')
    borders.set(qn('w:offsetFrom'), 'text')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '12') # Độ dày viền
        border.set(qn('w:space'), '24')
        border.set(qn('w:color'), 'auto')
        borders.append(border)
    sect_pr.append(borders)

def clear_page_border(sect_pr):
    """Xóa khung viền ở các trang sau"""
    for borders in sect_pr.xpath('./w:pgBorders'):
        sect_pr.remove(borders)

# ==========================================
# HÀM ĐỆ QUY TẠO GIAO DIỆN NHẬP LIỆU (TỐI ĐA 4 CẤP)
# ==========================================
def render_section(level, prefix, key_prefix):
    with st.container(border=True):
        st.markdown(f"**Mục {prefix}**")
        title = st.text_input("Tên mục:", key=f"title_{key_prefix}", label_visibility="collapsed", placeholder=f"Tên mục {prefix}")
        content = st.text_area("Nội dung:", key=f"content_{key_prefix}", height=100, label_visibility="collapsed", placeholder=f"Nội dung mục {prefix}")
        
        children = []
        if level < 4:
            num_children = st.number_input(f"Số tiểu mục con trong {prefix}:", min_value=0, max_value=15, value=0, step=1, key=f"num_{key_prefix}")
            for k in range(int(num_children)):
                child_prefix = f"{prefix}.{k+1}"
                children.append(render_section(level+1, child_prefix, f"{key_prefix}_{k}"))
                
    return {"title": title, "content": content, "children": children}

# ==========================================
# THUẬT TOÁN TỈA CÂY (ÉP LUẬT TỐI THIỂU 2 TIỂU MỤC)
# ==========================================
def apply_academic_rules(node):
    if node.get("children"):
        pruned_children = [apply_academic_rules(c) for c in node["children"]]
        
        if len(pruned_children) == 1:
            single_child = pruned_children[0]
            merged_text = single_child["title"]
            if single_child["content"].strip():
                merged_text += "\n" + single_child["content"]
                
            if node.get("content", "").strip():
                node["content"] += "\n\n" + merged_text
            else:
                node["content"] = merged_text
                
            node["children"] = single_child["children"]
        else:
            node["children"] = pruned_children
            
    return node

# ==========================================
# GIAO DIỆN NHẬP LIỆU CHÍNH
# ==========================================
st.header("I. THÔNG TIN BÌA")
thesis_title = st.text_input("Tên Đề tài Luận văn:", placeholder="Ví dụ: HIỆU QUẢ CỦA CHƯƠNG TRÌNH CAN THIỆP HABIT-ILE...")
author_name = st.text_input("Họ và tên tác giả:", placeholder="Ví dụ: TRẦN MINH HOÀNG")
st.divider()

st.header("II. NỘI DUNG")

st.subheader("ĐẶT VẤN ĐỀ")
dat_van_de_content = st.text_area("Nội dung phần Đặt vấn đề:", height=200, key="dvd")

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
def add_empty_lines(doc, num_lines, size=16):
    """Hỗ trợ tạo khoảng cách dòng trống với cỡ chữ chuẩn"""
    for _ in range(num_lines):
        p = doc.add_paragraph()
        r = p.add_run()
        r.font.size = Pt(size)

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
        st.warning("⚠️ Vui lòng nhập Tên Đề tài ở phần Thông tin bìa!")
    else:
        with st.spinner("Đang biên dịch trang bìa và dàn trang nội dung..."):
            processed_chapters = [apply_academic_rules(chap) for chap in chapters_data]

            doc = docx.Document()
            
            # Khởi tạo kích thước lề cho Section 1 (Trang bìa)
            sec_0 = doc.sections[0]
            sec_0.top_margin, sec_0.bottom_margin = Cm(3.5), Cm(3.0)
            sec_0.left_margin, sec_0.right_margin = Cm(3.5), Cm(2.0)
            add_page_border(sec_0._sectPr) # Đóng khung trang bìa

            style_normal = doc.styles['Normal']
            style_normal.font.name, style_normal.font.size = 'Times New Roman', Pt(13)

            # =====================================
            # THIẾT KẾ TRANG BÌA THEO YÊU CẦU
            # =====================================
            
            # Hàng 1: BỘ GD&ĐT (Trái) - BỘ Y TẾ (Phải) dùng Table ẩn
            table = doc.add_table(rows=1, cols=2)
            p_left = table.cell(0, 0).paragraphs[0]
            p_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r_left = p_left.add_run("BỘ GIÁO DỤC VÀ ĐÀO TẠO")
            r_left.font.name, r_left.font.size = 'Times New Roman', Pt(16)
            
            p_right = table.cell(0, 1).paragraphs[0]
            p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r_right = p_right.add_run("BỘ Y TẾ")
            r_right.font.name, r_right.font.size = 'Times New Roman', Pt(16)
            
            # Xuống 2 hàng
            add_empty_lines(doc, 2)
            
            # Tên trường
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("ĐẠI HỌC Y DƯỢC THÀNH PHỐ HỒ CHÍ MINH")
            r.bold = True
            r.font.name, r.font.size = 'Times New Roman', Pt(16)
            
            # Xuống 10 hàng
            add_empty_lines(doc, 10)
            
            # Tên tác giả
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(author_name.upper())
            r.bold = True
            r.font.name, r.font.size = 'Times New Roman', Pt(16)
            
            # Xuống 5 hàng
            add_empty_lines(doc, 5)
            
            # Tên đề tài (Size 20)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(thesis_title.upper())
            r.bold = True
            r.font.name, r.font.size = 'Times New Roman', Pt(20)
            
            # Xuống 5 hàng
            add_empty_lines(doc, 5)
            
            # Loại luận văn
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("ĐỀ CƯƠNG LUẬN VĂN THẠC SĨ")
            r.bold = True
            r.font.name, r.font.size = 'Times New Roman', Pt(16)
            
            # Đẩy phần "TP HCM - NĂM" xuống sát lề dưới (Khoảng 7-8 dòng)
            add_empty_lines(doc, 7)
            
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("THÀNH PHỐ HỒ CHÍ MINH - NĂM 2026")
            r.bold = True
            r.font.name, r.font.size = 'Times New Roman', Pt(16)

            # =====================================
            # NGẮT TRANG - BẮT ĐẦU PHẦN NỘI DUNG
            # =====================================
            new_section = doc.add_section(WD_SECTION.NEW_PAGE)
            clear_page_border(new_section._sectPr) # Gỡ khung viền cho các trang sau

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
            st.download_button("⬇️ TẢI FILE LUẬN VĂN (.docx)", bio.getvalue(), "De_Cuong_Luan_Van_Hoan_Chinh.docx", 
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
