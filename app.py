import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml.shared import OxmlElement, qn
import io

# ==========================================
# CẤU HÌNH TRANG WEB
# ==========================================
st.set_page_config(page_title="Tạo Đề Cương Luận Văn", page_icon="🎓", layout="centered")

st.title("🎓 Trình Tạo Đề Cương Luận Văn Chuẩn")
st.write("Hệ thống tự động dàn trang, ép chuẩn font và cấu trúc 4 chương, đánh số trang tự động.")
st.divider()

# ==========================================
# CÁC HÀM CAN THIỆP XML 
# ==========================================
def add_page_border(sect_pr):
    borders = OxmlElement('w:pgBorders')
    borders.set(qn('w:offsetFrom'), 'text')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '12') 
        border.set(qn('w:space'), '24')
        border.set(qn('w:color'), 'auto')
        borders.append(border)
    sect_pr.append(borders)

def clear_page_border(sect_pr):
    for borders in sect_pr.xpath('./w:pgBorders'):
        sect_pr.remove(borders)

def add_page_number(paragraph):
    """Chèn số trang động vào đoạn văn (dùng cho Header)"""
    p = paragraph._p
    run = OxmlElement('w:r')
    
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    run.append(fldChar1)
    run.append(instrText)
    run.append(fldChar2)
    run.append(fldChar3)
    p.append(run)

# ==========================================
# HÀM ĐỆ QUY TẠO GIAO DIỆN NHẬP LIỆU
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

st.markdown("##### Người hướng dẫn khoa học")
supervisor_1 = st.text_input("1. Họ và tên người hướng dẫn 1:", placeholder="Ví dụ: PGS.TS. NGUYỄN VĂN A")
supervisor_2 = st.text_input("2. Họ và tên người hướng dẫn 2 (Bỏ trống nếu không có):", placeholder="Ví dụ: TS. TRẦN THỊ B")

st.divider()

st.header("II. NỘI DUNG")

st.subheader("ĐẶT VẤN ĐỀ")
dat_van_de_content = st.text_area("Nội dung phần Đặt vấn đề:", height=200, key="dvd")

# --- CHƯƠNG 1 ---
st.markdown("### Chương 1. TỔNG QUAN TÀI LIỆU")
c1_intro = st.text_area("Nội dung dẫn nhập Chương 1 (nếu có):", height=100, key="c1_intro")
c1_num = st.number_input("Số lượng mục cấp 2 (ví dụ: 1.1, 1.2):", min_value=0, max_value=20, value=2, step=1, key="c1_num")
c1_children = [render_section(2, f"1.{j+1}", f"c1_sec_{j}") for j in range(c1_num)]
st.write("---")

# --- CHƯƠNG 2 (CỐ ĐỊNH 9 MỤC) ---
st.markdown("### Chương 2. PHƯƠNG PHÁP NGHIÊN CỨU")
c2_intro = st.text_area("Nội dung dẫn nhập Chương 2 (nếu có):", height=100, key="c2_intro")
c2_fixed_titles = [
    "Thiết kế nghiên cứu", "Thời gian và địa điểm nghiên cứu", "Đối tượng nghiên cứu", 
    "Cỡ mẫu của nghiên cứu", "Xác định các biến số độc lập và phụ thuộc", 
    "Phương pháp và công cụ đo lường, thu thập số liệu", "Quy trình nghiên cứu", 
    "Phương pháp phân tích dữ liệu", "Đạo đức trong nghiên cứu"
]
c2_children = []
for j, title in enumerate(c2_fixed_titles):
    with st.expander(f"Mục 2.{j+1}. {title}", expanded=True):
        c2_content = st.text_area(f"Nội dung mục 2.{j+1}:", height=150, key=f"c2_sec_{j}")
        c2_children.append({"title": title, "content": c2_content, "children": []}) # C2 không tạo cây con để giữ chuẩn khung
st.write("---")

# --- CHƯƠNG 3 ---
st.markdown("### Chương 3. DỰ KIẾN KẾT QUẢ")
c3_intro = st.text_area("Nội dung dẫn nhập Chương 3 (nếu có):", height=100, key="c3_intro")
c3_num = st.number_input("Số lượng mục cấp 2 (ví dụ: 3.1, 3.2):", min_value=0, max_value=20, value=2, step=1, key="c3_num")
c3_children = [render_section(2, f"3.{j+1}", f"c3_sec_{j}") for j in range(c3_num)]
st.write("---")

# --- CHƯƠNG 4 ---
st.markdown("### Chương 4. KẾ HOẠCH THỰC HIỆN")
c4_intro = st.text_area("Nội dung dẫn nhập Chương 4 (nếu có):", height=100, key="c4_intro")
c4_num = st.number_input("Số lượng mục cấp 2 (ví dụ: 4.1, 4.2):", min_value=0, max_value=20, value=1, step=1, key="c4_num")
c4_children = [render_section(2, f"4.{j+1}", f"c4_sec_{j}") for j in range(c4_num)]
st.write("---")

# Các phần cuối
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
    for _ in range(int(num_lines)):
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

def render_cover_header_and_title(doc, author_name, thesis_title, author_space=4):
    table = doc.add_table(rows=1, cols=2)
    p_left = table.cell(0, 0).paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_left = p_left.add_run("BỘ GIÁO DỤC VÀ ĐÀO TẠO")
    r_left.font.name, r_left.font.size = 'Times New Roman', Pt(16)
    
    p_right = table.cell(0, 1).paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_right = p_right.add_run("BỘ Y TẾ")
    r_right.font.name, r_right.font.size = 'Times New Roman', Pt(16)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ĐẠI HỌC Y DƯỢC THÀNH PHỐ HỒ CHÍ MINH")
    r.bold, r.font.name, r.font.size = True, 'Times New Roman', Pt(16)
    
    add_empty_lines(doc, author_space)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(author_name.upper())
    r.bold, r.font.name, r.font.size = True, 'Times New Roman', Pt(16)
    
    add_empty_lines(doc, 1)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(thesis_title.upper())
    r.bold, r.font.name, r.font.size = True, 'Times New Roman', Pt(20)


# ==========================================
# NÚT XỬ LÝ VÀ TẠO FILE WORD
# ==========================================
if st.button("✨ TẠO FILE WORD HOÀN CHỈNH", type="primary", use_container_width=True):
    if not thesis_title or not supervisor_1:
        st.warning("⚠️ Vui lòng nhập Tên Đề tài và ít nhất 1 Người hướng dẫn!")
    else:
        with st.spinner("Đang dàn trang, đánh số tự động và thiết lập Header..."):
            # Lọc nội dung các chương 1, 3, 4 theo chuẩn học thuật
            c1_processed = {"title": "TỔNG QUAN TÀI LIỆU", "content": c1_intro, "children": [apply_academic_rules(c) for c in c1_children]}
            c2_processed = {"title": "PHƯƠNG PHÁP NGHIÊN CỨU", "content": c2_intro, "children": c2_children} # C2 không cần lọc
            c3_processed = {"title": "DỰ KIẾN KẾT QUẢ", "content": c3_intro, "children": [apply_academic_rules(c) for c in c3_children]}
            c4_processed = {"title": "KẾ HOẠCH THỰC HIỆN", "content": c4_intro, "children": [apply_academic_rules(c) for c in c4_children]}
            all_chapters = [c1_processed, c2_processed, c3_processed, c4_processed]

            doc = docx.Document()
            style_normal = doc.styles['Normal']
            style_normal.font.name, style_normal.font.size = 'Times New Roman', Pt(13)

            title_lines = (len(thesis_title) // 40) + 1

            # =====================================
            # SECTION 1: TRANG BÌA CHÍNH (KHÔNG SỐ TRANG)
            # =====================================
            sec_0 = doc.sections[0]
            sec_0.top_margin, sec_0.bottom_margin = Cm(3.5), Cm(3.0)
            sec_0.left_margin, sec_0.right_margin = Cm(3.5), Cm(2.0)
            add_page_border(sec_0._sectPr)

            render_cover_header_and_title(doc, author_name, thesis_title, author_space=4)
            add_empty_lines(doc, 3)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("ĐỀ CƯƠNG LUẬN VĂN THẠC SĨ")
            r.bold, r.font.name, r.font.size = True, 'Times New Roman', Pt(16)
            
            empty_lines_to_bottom = 8 - title_lines
            if empty_lines_to_bottom < 1: empty_lines_to_bottom = 1
            add_empty_lines(doc, empty_lines_to_bottom)
            
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("THÀNH PHỐ HỒ CHÍ MINH - NĂM 2026")
            r.bold, r.font.name, r.font.size = True, 'Times New Roman', Pt(16)

            # =====================================
            # SECTION 2: TRANG BÌA PHỤ (KHÔNG SỐ TRANG)
            # =====================================
            new_section_cover_2 = doc.add_section(WD_SECTION.NEW_PAGE)
            new_section_cover_2.top_margin, new_section_cover_2.bottom_margin = Cm(3.5), Cm(3.0)
            new_section_cover_2.left_margin, new_section_cover_2.right_margin = Cm(3.5), Cm(2.0)
            add_page_border(new_section_cover_2._sectPr)

            render_cover_header_and_title(doc, author_name, thesis_title, author_space=2)
            add_empty_lines(doc, 1)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("NGÀNH: KỸ THUẬT PHỤC HỒI CHỨC NĂNG")
            r.bold, r.font.name, r.font.size = True, 'Times New Roman', Pt(16)

            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("MÃ SỐ: 8720603")
            r.bold, r.font.name, r.font.size = True, 'Times New Roman', Pt(16)

            add_empty_lines(doc, 1)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("ĐỀ CƯƠNG LUẬN VĂN THẠC SĨ")
            r.bold, r.font.name, r.font.size = True, 'Times New Roman', Pt(16)

            add_empty_lines(doc, 1)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("NGƯỜI DỰ KIẾN HƯỚNG DẪN KHOA HỌC:")
            r.bold, r.font.name, r.font.size = True, 'Times New Roman', Pt(16)

            if not supervisor_2.strip():
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(f"{supervisor_1.upper()}")
                r.bold, r.font.name, r.font.size = True, 'Times New Roman', Pt(16)
            else:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(f"1. {supervisor_1.upper()}")
                r.bold, r.font.name, r.font.size = True, 'Times New Roman', Pt(16)
                
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(f"2. {supervisor_2.upper()}")
                r.bold, r.font.name, r.font.size = True, 'Times New Roman', Pt(16)

            empty_lines_inner = 6 - title_lines - (1 if supervisor_2.strip() else 0)
            if empty_lines_inner < 1: empty_lines_inner = 1
            add_empty_lines(doc, empty_lines_inner)

            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run("THÀNH PHỐ HỒ CHÍ MINH - NĂM 2026")
            r.bold, r.font.name, r.font.size = True, 'Times New Roman', Pt(16)

            # =====================================
            # SECTION 3: BẮT ĐẦU NỘI DUNG (ĐÁNH SỐ TRANG TỪ SỐ 1)
            # =====================================
            new_section_content = doc.add_section(WD_SECTION.NEW_PAGE)
            clear_page_border(new_section_content._sectPr)
            
            # --- TÁCH HEADER VÀ GẮN SỐ TRANG ---
            new_section_content.header.is_linked_to_previous = False
            header_para = new_section_content.header.paragraphs[0]
            header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_page_number(header_para) # Chèn "PAGE" vào header
            
            # Ép Word đếm số trang từ 1
            sectPr = new_section_content._sectPr
            pgNumType = OxmlElement('w:pgNumType')
            pgNumType.set(qn('w:start'), '1')
            sectPr.append(pgNumType)

            if dat_van_de_content.strip():
                add_main_heading(doc, "ĐẶT VẤN ĐỀ")
                add_normal_text(doc, dat_van_de_content)
                doc.add_page_break()

            for i, chap in enumerate(all_chapters):
                add_main_heading(doc, f"CHƯƠNG {i+1}: {chap['title']}")
                add_normal_text(doc, chap['content'])
                write_sections_to_word(doc, chap['children'], [str(i+1)])
                doc.add_page_break()

            if danh_muc_content.strip():
                add_main_heading(doc, "DANH MỤC CÁC CÔNG TRÌNH CÔNG BỐ CÓ LIÊN QUAN")
                add_normal_text(doc, danh_muc_content)
                doc.add_page_break()

            # =====================================
            # SECTION 4: TÀI LIỆU THAM KHẢO & PHỤ LỤC (KHÔNG SỐ TRANG)
            # =====================================
            new_section_end = doc.add_section(WD_SECTION.CONTINUOUS)
            
            # Ngắt liên kết Header để xóa số trang
            new_section_end.header.is_linked_to_previous = False
            for hp in new_section_end.header.paragraphs:
                hp.text = "" 

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
