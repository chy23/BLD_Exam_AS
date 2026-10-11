import io
import os
import re
import zipfile
import streamlit as st
import fitz  # PyMuPDF
import numpy as np
import cv2
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

# ==========================================
# 1. Docx 表格邊框與儲存格樣式（強化雙重保險）
# ==========================================

def apply_table_borders(table, color="000000", sz="6", val="single"):
    """
    在表格的 tblPr 中寫入完整且標準的 w:tblBorders，
    確保在 Microsoft Word、WPS Office、LibreOffice、Google Docs 都能 100% 顯示清晰實線黑邊框。
    """
    tblPr = table._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith("tblBorders"):
            tblPr.remove(child)

    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def apply_cell_borders(cell, color="000000", sz="6", val="single"):
    """為單一儲存格寫入 tcBorders，防止跨平台檢視器忽略表格層級邊框"""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith("tcBorders"):
            tcPr.remove(child)
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)

def set_cell_background(cell, fill_hex):
    """設定儲存格背景色 (例如 'F2F2F2')"""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def format_cell_text(cell, text, bold=False, size_pt=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, font_name='微軟正黑體'):
    """統一格式化儲存格內的文字"""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(text)
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size_pt)
    run.bold = bold
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def add_header_table(doc, exam_title):
    """
    建立標準台灣試卷答案卷標頭表格：
    包含大標題（忠實複製原考卷標題）、班級、年班、座號、姓名、得分格
    """
    table = doc.add_table(rows=2, cols=8)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # 第 1 列：大標題合併
    header_cell = table.cell(0, 0)
    for c in range(1, 8):
        header_cell.merge(table.cell(0, c))
    
    format_cell_text(header_cell, exam_title, bold=True, size_pt=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    table.rows[0].height = Pt(32)

    # 第 2 列：個人資訊格
    col_configs = [
        ("班級", True, "F2F2F2", Inches(0.8)),
        ("    年    班", False, "FFFFFF", Inches(1.5)),
        ("座號", True, "F2F2F2", Inches(0.8)),
        ("    號", False, "FFFFFF", Inches(1.0)),
        ("姓名", True, "F2F2F2", Inches(0.8)),
        ("          ", False, "FFFFFF", Inches(1.5)),
        ("得分", True, "F2F2F2", Inches(0.8)),
        (" ", False, "FFFFFF", Inches(1.0))
    ]

    for col_idx, (text, is_bold, bg_color, width_in) in enumerate(col_configs):
        cell = table.cell(1, col_idx)
        cell.width = width_in
        format_cell_text(cell, text, bold=is_bold, size_pt=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        if bg_color != "FFFFFF":
            set_cell_background(cell, bg_color)
    
    table.rows[1].height = Pt(26)

    # 強制加入全黑實線邊框
    apply_table_borders(table, color="000000", sz="6", val="single")
    for row in table.rows:
        for cell in row.cells:
            apply_cell_borders(cell, color="000000", sz="6", val="single")

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(4)
    p_spacer.paragraph_format.space_after = Pt(4)

# ==========================================
# 2. 題目卷：消除紅色答案 (純題目卷)
# ==========================================

def remove_red_from_pdf(pdf_bytes):
    """
    透過精確 HSV 濾色消除紅字答案，保留原始版面、尺寸 (A4/A3/B4) 與方向 (直式/橫式)。
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    processed_images = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=200)

        img_np = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)

        if pix.n == 4:
            img_np = cv2.cvtColor(img_np, cv2.COLOR_RGBA2RGB)
        elif pix.n == 1:
            img_np = cv2.cvtColor(img_np, cv2.COLOR_GRAY2RGB)

        img_cv = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
        hsv = cv2.cvtColor(img_cv, cv2.COLOR_BGR2HSV)

        # 紅色 HSV 範圍
        lower_red1 = np.array([0, 45, 45])
        upper_red1 = np.array([12, 255, 255])
        lower_red2 = np.array([168, 45, 45])
        upper_red2 = np.array([180, 255, 255])

        mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
        mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
        mask = cv2.bitwise_or(mask1, mask2)

        # 膨脹以消除紅色邊緣反鋸齒
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
        mask_dilated = cv2.dilate(mask, kernel, iterations=1)

        # 覆蓋為純白色背景
        img_cv[mask_dilated > 0] = [255, 255, 255]

        result_image = cv2.cvtColor(img_cv, cv2.COLOR_BGR2RGB)
        processed_images.append(Image.fromarray(result_image))

    output_io = io.BytesIO()
    if processed_images:
        processed_images[0].save(
            output_io,
            format='PDF',
            save_all=True,
            append_images=processed_images[1:],
            resolution=200.0
        )
    return output_io.getvalue()

# ==========================================
# 3. 試卷結構深度解析 (精準標題抓取、計分推估與題號解析)
# ==========================================

def extract_pdf_structure(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    lines_in_order = []

    for page in doc:
        rect = page.rect
        page_width = rect.width
        page_height = rect.height
        blocks = page.get_text("blocks")
        text_blocks = [b for b in blocks if b[6] == 0 and b[4].strip()]

        # 頂部標題閥值
        header_thresh = page_height * 0.15
        header_blocks = [b for b in text_blocks if b[1] < header_thresh]
        body_blocks = [b for b in text_blocks if b[1] >= header_thresh]

        # 智慧判斷左右雙欄
        mid_x = page_width / 2.0
        left_col = [b for b in body_blocks if (b[0] + b[2]) / 2.0 < mid_x]
        right_col = [b for b in body_blocks if (b[0] + b[2]) / 2.0 >= mid_x]

        if len(left_col) > 0 and len(right_col) > 0:
            left_col.sort(key=lambda b: (b[1], b[0]))
            right_col.sort(key=lambda b: (b[1], b[0]))
            sorted_body = left_col + right_col
        else:
            sorted_body = sorted(body_blocks, key=lambda b: (b[1], b[0]))

        header_blocks.sort(key=lambda b: (b[1], b[0]))
        all_blocks = header_blocks + sorted_body

        for b in all_blocks:
            for line in b[4].split('\n'):
                line_str = line.strip()
                if line_str:
                    lines_in_order.append(line_str)

    # 1. 忠實擷取原考卷標題 (抓取大題「一、」之前的所有標題文字)
    section_regex = re.compile(r'^[（\(]?[一二三四五六七八九十壹貳參肆伍陸柒捌玖拾]+[）\)]?[、.．\s]+(.*)')
    
    first_sec_idx = len(lines_in_order)
    for idx, line in enumerate(lines_in_order):
        if section_regex.match(line):
            first_sec_idx = idx
            break

    # 收集第一大題之前的所有純標題文字（過濾個人資訊欄）
    raw_title_lines = []
    for line in lines_in_order[:first_sec_idx]:
        # 過濾掉學生填寫格欄位
        if any(k in line for k in ["姓名：", "姓名:", "座號：", "座號:", "得分", "年 班", "年級"]):
            continue
        if re.search(r'[\d一二三四五六]\s*年.*?班.*?號', line):
            continue
        raw_title_lines.append(line)

    if raw_title_lines:
        exam_title = " ".join(raw_title_lines)
    else:
        exam_title = "國小 評量試卷"

    # 2. 分割大題與解析
    score_regexes = [
        re.compile(r'(?:每|一)[小題題格份個問]+[：:\s]*(\d+)\s*分.*?共\s*(\d+)\s*分'),
        re.compile(r'共\s*(\d+)\s*分.*?(?:每|一)[小題題格份個問]+[：:\s]*(\d+)\s*分'),
        re.compile(r'共\s*(\d+)\s*[題格小題份]')
    ]

    sections = []
    current_sec = None

    for line in lines_in_order[first_sec_idx:]:
        sec_m = section_regex.match(line)
        if sec_m:
            if current_sec:
                sections.append(current_sec)
            
            title = line
            # 判斷題型
            sec_type = "OTHER"
            if any(k in title for k in ["是非"]):
                sec_type = "TRUE_FALSE"
            elif any(k in title for k in ["選擇", "選出", "選一選"]):
                sec_type = "CHOICE"
            elif any(k in title for k in ["算算看", "直式算算看", "用直式", "計算"]):
                sec_type = "CALCULATION"
            elif any(k in title for k in ["做法", "應用", "問答", "記下來", "先圈", "圈一圈"]):
                sec_type = "APPLICATION"
            elif any(k in title for k in ["填填看", "填一填", "填入", "填寫", "填充", "填空", "四捨五入", "求商", "不用計算"]):
                sec_type = "FILL_BLANK"

            current_sec = {
                "title": title,
                "type": sec_type,
                "raw_lines": [],
                "expected_count": 0,
                "questions": []
            }
        else:
            if current_sec:
                current_sec["raw_lines"].append(line)

    if current_sec:
        sections.append(current_sec)

    # 3. 逐一大題分析計分公式與抽取題號
    for sec in sections:
        # 搜尋計分資訊：檢查標題與前 3 行文字（例如第二行「(每題 8 分，共 40 分)」）
        lines_to_search = [sec["title"]] + sec["raw_lines"][:3]
        expected_cnt = 0
        for l in lines_to_search:
            m1 = score_regexes[0].search(l)
            if m1:
                each_pt = int(m1.group(1))
                total_pt = int(m1.group(2))
                if each_pt > 0:
                    expected_cnt = total_pt // each_pt
                    break
            m2 = score_regexes[1].search(l)
            if m2:
                total_pt = int(m2.group(1))
                each_pt = int(m2.group(2))
                if each_pt > 0:
                    expected_cnt = total_pt // each_pt
                    break
            m3 = score_regexes[2].search(l)
            if m3:
                expected_cnt = int(m3.group(1))
                break

        sec["expected_count"] = expected_cnt

        # 抽取題號：尋找所有的 (1), (2) 或 1., 2. 或 3-a 等
        detected_q = []
        for line in sec["raw_lines"]:
            # 搜尋行內所有括弧題號 (1), (2), (3)...
            p_matches = re.findall(r'[\(（](\d{1,2})[\)）]', line)
            for pm in p_matches:
                val = int(pm)
                if 1 <= val <= 40:
                    val_str = str(val)
                    if val_str not in detected_q:
                        detected_q.append(val_str)

            # 搜尋行首題號 1., 2., 1、, 2、
            m_dot = re.match(r'^\s*(\d{1,2})[\.．、]', line)
            if m_dot:
                val = int(m_dot.group(1))
                if 1 <= val <= 40:
                    val_str = str(val)
                    if val_str not in detected_q:
                        detected_q.append(val_str)

            # 搜尋子題號 3-a, 7-1
            m_sub = re.match(r'^\s*(\d{1,2}[-_][a-zA-Z0-9]+)', line)
            if m_sub:
                val_str = m_sub.group(1)
                if val_str not in detected_q:
                    detected_q.append(val_str)

        # 檢驗抓取到的題號是否合乎序列
        is_valid_sequence = False
        if detected_q:
            # 第一題必須是 1
            if detected_q[0] == "1":
                is_valid_sequence = True

        final_questions = []
        if is_valid_sequence:
            final_questions = detected_q
            # 若期望題數大於抓到的題數，補齊至期望題數
            if expected_cnt > len(final_questions) and expected_cnt <= 40:
                last_n = int(re.match(r'^(\d+)', final_questions[-1]).group(1))
                while len(final_questions) < expected_cnt:
                    last_n += 1
                    final_questions.append(str(last_n))
            # 若抓到的題數多於期望題數，且期望題數有效，截取至期望題數
            elif expected_cnt > 0 and len(final_questions) > expected_cnt:
                final_questions = final_questions[:expected_cnt]
        else:
            # 若未抓到或序列不合，直接使用計分推估的題目數
            if expected_cnt > 0 and expected_cnt <= 40:
                final_questions = [str(i) for i in range(1, expected_cnt + 1)]
            else:
                s_type = sec["type"]
                if s_type in ["CHOICE", "TRUE_FALSE"]:
                    def_cnt = 10
                elif s_type == "CALCULATION":
                    def_cnt = 4
                elif s_type == "APPLICATION":
                    def_cnt = 4
                else:
                    def_cnt = 4
                final_questions = [str(i) for i in range(1, def_cnt + 1)]

        sec["questions"] = final_questions

    return exam_title, sections

# ==========================================
# 4. 空白答案卡卷（答題卷）Docx 表格繪製
# ==========================================

def generate_blank_answer_sheet_docx(pdf_bytes):
    exam_title, sections = extract_pdf_structure(pdf_bytes)

    doc = Document()

    # 設定邊距（上下左右皆為 1.8 公分，約 0.7 英吋）
    sections_page = doc.sections[0]
    sections_page.top_margin = Inches(0.7)
    sections_page.bottom_margin = Inches(0.7)
    sections_page.left_margin = Inches(0.7)
    sections_page.right_margin = Inches(0.7)

    # 1. 建立頂部標頭表格
    add_header_table(doc, exam_title)

    # 2. 逐一產生各題型的答案卡框格
    for sec in sections:
        # 大題標題段落
        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_before = Pt(8)
        p_title.paragraph_format.space_after = Pt(3)
        p_title.paragraph_format.line_spacing = 1.0
        run = p_title.add_run(sec["title"])
        run.bold = True
        run.font.name = '微軟正黑體'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
        run.font.size = Pt(11.5)

        q_list = sec["questions"]
        s_type = sec["type"]

        if s_type in ["CHOICE", "TRUE_FALSE", "FILL_BLANK"]:
            # --- 選擇題 / 是非題 / 填填看 / 比大小填入符號：網格表格 ---
            col_max = 10
            total_q = len(q_list)
            for chunk_start in range(0, total_q, col_max):
                chunk = q_list[chunk_start:chunk_start + col_max]
                num_cols = len(chunk)
                
                tbl = doc.add_table(rows=2, cols=num_cols)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.autofit = False

                tbl.rows[0].height = Pt(22)
                tbl.rows[1].height = Pt(36) # 預留充足作答高度

                cell_w = Inches(7.2 / max(num_cols, 1))

                for idx, q_num in enumerate(chunk):
                    # 第 1 列：題號（必為 1., 2., 3. 正確順序）
                    c_num = tbl.cell(0, idx)
                    c_num.width = cell_w
                    format_cell_text(c_num, f"{q_num}.", bold=True, size_pt=10, align=WD_ALIGN_PARAGRAPH.CENTER)
                    set_cell_background(c_num, "F2F2F2")
                    apply_cell_borders(c_num, color="000000", sz="6", val="single")

                    # 第 2 列：空白作答格
                    c_ans = tbl.cell(1, idx)
                    c_ans.width = cell_w
                    c_ans.text = ""
                    c_ans.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    apply_cell_borders(c_ans, color="000000", sz="6", val="single")

                # 強制加上表格全黑邊框
                apply_table_borders(tbl, color="000000", sz="6", val="single")

                p_space = doc.add_paragraph()
                p_space.paragraph_format.space_before = Pt(2)
                p_space.paragraph_format.space_after = Pt(2)

        elif s_type == "CALCULATION":
            # --- 直式算算看 / 計算題：2 欄或 3 欄大型計算網格 ---
            total_q = len(q_list)
            # 若為 2 題或 4 題，採用 2 欄對稱排版（如考卷原格式 2x2 格）
            if total_q in [2, 4]:
                col_cnt = 2
            else:
                col_cnt = 3 if total_q >= 3 else total_q
            col_cnt = max(col_cnt, 1)

            for chunk_start in range(0, total_q, col_cnt):
                chunk = q_list[chunk_start:chunk_start + col_cnt]
                num_cols = len(chunk)

                tbl = doc.add_table(rows=1, cols=num_cols)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.autofit = False
                tbl.rows[0].height = Pt(130) # 直式算式垂直高度

                cell_w = Inches(7.2 / num_cols)

                for idx, q_num in enumerate(chunk):
                    cell = tbl.cell(0, idx)
                    cell.width = cell_w
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_before = Pt(4)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.left_indent = Inches(0.08)
                    run_num = p.add_run(f"({q_num})")
                    run_num.bold = True
                    run_num.font.name = '微軟正黑體'
                    run_num._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
                    run_num.font.size = Pt(11)

                    apply_cell_borders(cell, color="000000", sz="6", val="single")

                apply_table_borders(tbl, color="000000", sz="6", val="single")

                p_space = doc.add_paragraph()
                p_space.paragraph_format.space_before = Pt(3)
                p_space.paragraph_format.space_after = Pt(3)

        elif s_type == "APPLICATION":
            # --- 應用題 / 問答題 / 把做法和答案記下來：單欄大型作答框 ---
            for q_num in q_list:
                tbl = doc.add_table(rows=1, cols=1)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.autofit = False
                tbl.rows[0].height = Pt(105) # 寬敞的解題作答空間

                cell = tbl.cell(0, 0)
                cell.width = Inches(7.2)
                p_q = cell.paragraphs[0]
                p_q.paragraph_format.space_before = Pt(4)
                p_q.paragraph_format.left_indent = Inches(0.08)
                run_q = p_q.add_run(f"({q_num}) ")
                run_q.bold = True
                run_q.font.name = '微軟正黑體'
                run_q._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
                run_q.font.size = Pt(11)

                # 底部加入「答：」
                p_ans = cell.add_paragraph()
                p_ans.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_ans.paragraph_format.space_before = Pt(60)
                run_ans = p_ans.add_run("答：____________________    ")
                run_ans.font.name = '微軟正黑體'
                run_ans._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
                run_ans.font.size = Pt(10.5)

                apply_cell_borders(cell, color="000000", sz="6", val="single")
                apply_table_borders(tbl, color="000000", sz="6", val="single")

                p_space = doc.add_paragraph()
                p_space.paragraph_format.space_before = Pt(3)
                p_space.paragraph_format.space_after = Pt(3)
        else:
            # 其他題型：標準寬度作答框
            col_max = 5
            total_q = len(q_list)
            for chunk_start in range(0, total_q, col_max):
                chunk = q_list[chunk_start:chunk_start + col_max]
                num_cols = len(chunk)
                tbl = doc.add_table(rows=2, cols=num_cols)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.autofit = False
                tbl.rows[0].height = Pt(22)
                tbl.rows[1].height = Pt(38)

                cell_w = Inches(7.2 / num_cols)

                for idx, q_num in enumerate(chunk):
                    c_num = tbl.cell(0, idx)
                    c_num.width = cell_w
                    format_cell_text(c_num, f"{q_num}.", bold=True, size_pt=10)
                    set_cell_background(c_num, "F2F2F2")
                    apply_cell_borders(c_num, color="000000", sz="6", val="single")

                    c_ans = tbl.cell(1, idx)
                    c_ans.width = cell_w
                    c_ans.text = ""
                    apply_cell_borders(c_ans, color="000000", sz="6", val="single")

                apply_table_borders(tbl, color="000000", sz="6", val="single")

                p_space = doc.add_paragraph()
                p_space.paragraph_format.space_before = Pt(2)
                p_space.paragraph_format.space_after = Pt(2)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    return doc_io.getvalue()

# ==========================================
# 5. 主介面 Streamlit
# ==========================================

def main():
    st.set_page_config(
        page_title="考卷轉換與空白答案卡卷生成器",
        page_icon="📝",
        layout="centered"
    )

    st.markdown("""
        <style>
        .main-header {
            font-size: 2.1rem;
            font-weight: 700;
            color: #1E3A8A;
            text-align: center;
            margin-bottom: 0.5rem;
        }
        .sub-header {
            font-size: 1.05rem;
            color: #4B5563;
            text-align: center;
            margin-bottom: 1.8rem;
        }
        .stButton>button {
            width: 100%;
            height: 3.2rem;
            font-size: 1.2rem;
            font-weight: bold;
        }
        </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="main-header">📝 考卷轉換與空白答案卡卷生成系統</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">上傳「題目既答案卷」，一鍵消除答案轉為純題目卷，或自動排版生成正式「空白答案卡卷（答題卷）」</div>', unsafe_allow_html=True)

    # 1. 檔案上傳區
    uploaded_file = st.file_uploader(
        "請上傳考卷 PDF 檔（支援 A4、A3、B4、直式或橫式）：",
        type=["pdf"],
        help="支援包含紅色解答的題目卷 PDF"
    )

    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        base_name = os.path.splitext(uploaded_file.name)[0]

        st.markdown("---")
        st.subheader("⚙️ 匯出選項設定")

        # 匯出模式單選 (一次打包下載，免除分開下載)
        export_mode = st.radio(
            "請選擇匯出模式：",
            (
                "📦 同時匯出兩者（純題目卷 + 空白答題卷，打包為 ZIP 壓縮檔）",
                "📄 僅匯出純題目卷（PDF）",
                "📝 僅匯出空白答案卡卷（Docx）"
            ),
            index=0
        )

        st.caption("💡 提示：選擇「同時匯出兩者」會將題目卷與答案卷自動打包成單一 ZIP 壓縮檔，一次點擊即可全部下載！")

        # 執行轉換按鈕
        if st.button("🚀 開始處理並產生檔案", type="primary"):
            with st.spinner("正在分析試卷結構、處理圖像與排版框格中，請稍候..."):
                try:
                    if "同時匯出兩者" in export_mode:
                        # 處理純題目卷 PDF
                        clean_pdf_bytes = remove_red_from_pdf(file_bytes)
                        # 處理空白答案卡卷 Docx
                        answer_docx_bytes = generate_blank_answer_sheet_docx(file_bytes)

                        # 打包成單一 ZIP 檔案
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            zip_file.writestr(f"{base_name}_純題目卷.pdf", clean_pdf_bytes)
                            zip_file.writestr(f"{base_name}_空白答題卷.docx", answer_docx_bytes)
                        
                        zip_data = zip_buffer.getvalue()

                        st.success("🎉 轉換成功！純題目卷與空白答案卡卷已自動打包完成。")
                        st.download_button(
                            label="📥 一鍵下載完整檔案 (ZIP 壓縮包)",
                            data=zip_data,
                            file_name=f"{base_name}_考卷與答題卷.zip",
                            mime="application/zip",
                            type="primary"
                        )

                    elif "僅匯出純題目卷" in export_mode:
                        clean_pdf_bytes = remove_red_from_pdf(file_bytes)
                        st.success("🎉 純題目卷生成成功！")
                        st.download_button(
                            label="📥 下載純題目卷 (PDF)",
                            data=clean_pdf_bytes,
                            file_name=f"{base_name}_純題目卷.pdf",
                            mime="application/pdf",
                            type="primary"
                        )

                    elif "僅匯出空白答案卡卷" in export_mode:
                        answer_docx_bytes = generate_blank_answer_sheet_docx(file_bytes)
                        st.success("🎉 空白答題卷生成成功！")
                        st.download_button(
                            label="📥 下載空白答案卡卷 (Docx)",
                            data=answer_docx_bytes,
                            file_name=f"{base_name}_空白答題卷.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )

                except Exception as e:
                    st.error(f"❌ 處理過程中發生錯誤：{str(e)}")
                    st.exception(e)

if __name__ == "__main__":
    main()
