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

# --- 輔助函式：Docx 樣式與表格微調 ---

def set_cell_background(cell, fill_hex):
    """設定表格儲存格背景色 (例如 'F2F2F2')"""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def format_cell_text(cell, text, bold=False, size_pt=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, font_name='微軟正黑體'):
    """統一格式化儲存格內的文字"""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size_pt)
    run.bold = bold
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def add_header_table(doc, exam_title):
    """
    建立標準台灣國小/國中試卷的頂部標頭表格：
    包含試卷名稱、班級、年班、座號、姓名、得分格
    """
    table = doc.add_table(rows=2, cols=8)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # 第 1 列：試卷名稱 (合併 8 欄)
    header_cell = table.cell(0, 0)
    for c in range(1, 8):
        header_cell.merge(table.cell(0, c))
    
    format_cell_text(header_cell, exam_title, bold=True, size_pt=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    table.rows[0].height = Pt(32)

    # 第 2 列：學生資訊欄位
    # 欄位配比: 班級 | ___年___班 | 座號 | ___號 | 姓名 | ________ | 得分 | ____
    headers = [
        ("班級", True, "F4F4F4", 1.0),
        ("     年     班", False, "FFFFFF", 1.8),
        ("座號", True, "F4F4F4", 1.0),
        ("     號", False, "FFFFFF", 1.2),
        ("姓名", True, "F4F4F4", 1.0),
        ("           ", False, "FFFFFF", 2.0),
        ("得分", True, "F4F4F4", 1.0),
        (" ", False, "FFFFFF", 1.4)
    ]

    for col_idx, (text, is_bold, bg_color, width_in) in enumerate(headers):
        cell = table.cell(1, col_idx)
        cell.width = Inches(width_in)
        format_cell_text(cell, text, bold=is_bold, size_pt=11, align=WD_ALIGN_PARAGRAPH.CENTER)
        if bg_color != "FFFFFF":
            set_cell_background(cell, bg_color)

    table.rows[1].height = Pt(26)
    
    # 加上段落間距
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)

# --- 題目卷：消除紅色答案 (純題目卷) ---

def remove_red_from_pdf(pdf_bytes):
    """
    透過精確 HSV 濾色消除紅字答案，並保留原始頁面格式 (A4/A3/B4/直式/橫式)。
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    processed_images = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        # 以 200 DPI 高畫質轉為影像
        pix = page.get_pixmap(dpi=200)

        # 轉換為 numpy 陣列
        img_np = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)

        if pix.n == 4:
            img_np = cv2.cvtColor(img_np, cv2.COLOR_RGBA2RGB)
        elif pix.n == 1:
            img_np = cv2.cvtColor(img_np, cv2.COLOR_GRAY2RGB)

        img_cv = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
        hsv = cv2.cvtColor(img_cv, cv2.COLOR_BGR2HSV)

        # 紅色在 HSV 色彩空間的兩個區間 (色相 0~10 與 170~180)
        lower_red1 = np.array([0, 45, 45])
        upper_red1 = np.array([12, 255, 255])
        lower_red2 = np.array([168, 45, 45])
        upper_red2 = np.array([180, 255, 255])

        mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
        mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
        mask = cv2.bitwise_or(mask1, mask2)

        # 輕微膨脹遮罩以徹底消除紅字邊緣粉紅色反鋸齒
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
        mask_dilated = cv2.dilate(mask, kernel, iterations=1)

        # 將紅字區域直接覆蓋成純白色 (255, 255, 255)
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

# --- 答題卷分析與產生 (空白答案卡卷) ---

def extract_pdf_structure(pdf_bytes):
    """
    智慧解析考卷結構：
    - 自動判別雙欄排版 (左欄優先再右欄)
    - 識別標題、各階段大題、題型與題數
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    lines_in_order = []

    for page in doc:
        rect = page.rect
        page_width = rect.width
        page_height = rect.height
        blocks = page.get_text("blocks")
        # 篩選有效文字塊
        text_blocks = [b for b in blocks if b[6] == 0 and b[4].strip()]

        # 頂部標題區閥值
        header_thresh = page_height * 0.15
        header_blocks = [b for b in text_blocks if b[1] < header_thresh]
        body_blocks = [b for b in text_blocks if b[1] >= header_thresh]

        # 判斷是否為雙欄
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

    # 擷取試卷總標題
    exam_title = "國小 評量試卷"
    for line in lines_in_order[:10]:
        if any(k in line for k in ["國小", "國中", "學年度", "試卷", "評量", "數學", "語文", "自然"]):
            exam_title = line
            break

    # 解析大題與題目
    section_regex = re.compile(r'^[（\(]?[一二三四五六七八九十壹貳參肆伍陸柒捌玖拾]+[）\)]?[、.．\s]+(.*)')
    score_regex = re.compile(r'每[小題題格份]+[：:\s]*(\d+)\s*分.*?共\s*(\d+)\s*分')
    alt_score_regex = re.compile(r'共\s*(\d+)\s*分.*?每[小題題格份]+[：:\s]*(\d+)\s*分')

    sections = []
    current_sec = None

    for line in lines_in_order:
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
            elif any(k in title for k in ["填填看", "填入", "填寫", "填充", "填空"]):
                sec_type = "FILL_BLANK"
            elif any(k in title for k in ["算算看", "計算", "直式", "做做看"]):
                sec_type = "CALCULATION"
            elif any(k in title for k in ["做法", "應用", "問答", "過程", "解決", "寫出"]):
                sec_type = "APPLICATION"

            # 推估題數 (透過「每題X分，共Y分」)
            expected_count = 0
            sc_m = score_regex.search(title)
            if sc_m:
                each_pt = int(sc_m.group(1))
                total_pt = int(sc_m.group(2))
                if each_pt > 0:
                    expected_count = total_pt // each_pt
            else:
                alt_m = alt_score_regex.search(title)
                if alt_m:
                    total_pt = int(alt_m.group(1))
                    each_pt = int(alt_m.group(2))
                    if each_pt > 0:
                        expected_count = total_pt // each_pt

            current_sec = {
                "title": title,
                "type": sec_type,
                "expected_count": expected_count,
                "raw_lines": [],
                "questions": []
            }
        else:
            if current_sec:
                current_sec["raw_lines"].append(line)

    if current_sec:
        sections.append(current_sec)

    # 進一步從 raw_lines 抽取題目標號或小題
    for sec in sections:
        detected_q = []
        for line in sec["raw_lines"]:
            # 題號模式：如 "1.", "1、", "( ) 1.", "（ ）1.", "1. ( )", "3-a", "7-1"
            q_num_match = re.search(r'(?:[\(（]\s*[\)）\dOoxX✓✔\s]*[\)）]\s*)?(\d+(?:[-_][a-zA-Z0-9]+)?)[\.．、\s]', line)
            if q_num_match:
                q_id = q_num_match.group(1)
                if q_id not in detected_q:
                    detected_q.append(q_id)
            elif re.match(r'^[a-d][\.．\s]', line):
                sub_id = line[0]
                if sub_id not in detected_q:
                    detected_q.append(sub_id)

        # 綜合判斷：如果抓到的題號太少，但有預期題數，則用預期題數補齊
        if len(detected_q) < sec["expected_count"] and sec["expected_count"] > 0:
            if len(detected_q) == 0:
                detected_q = [str(i) for i in range(1, sec["expected_count"] + 1)]
            else:
                # 補充至預期題數
                last_num = int(detected_q[-1]) if detected_q[-1].isdigit() else len(detected_q)
                while len(detected_q) < sec["expected_count"]:
                    last_num += 1
                    detected_q.append(str(last_num))
        elif len(detected_q) == 0:
            # 預設題數
            default_cnt = 10 if sec["type"] in ["CHOICE", "TRUE_FALSE", "FILL_BLANK"] else (3 if sec["type"] == "CALCULATION" else 2)
            detected_q = [str(i) for i in range(1, default_cnt + 1)]

        sec["questions"] = detected_q

    return exam_title, sections

def generate_blank_answer_sheet_docx(pdf_bytes):
    """
    產生與使用者提供的範例 100% 吻合的標準「空白答案卡卷（答題卷）」Docx 檔
    """
    exam_title, sections = extract_pdf_structure(pdf_bytes)

    doc = Document()

    # 設定版面邊界為標準 2 公分
    sections_page = doc.sections[0]
    sections_page.top_margin = Inches(0.7)
    sections_page.bottom_margin = Inches(0.7)
    sections_page.left_margin = Inches(0.7)
    sections_page.right_margin = Inches(0.7)

    # 1. 頂部標頭表格 (Header Table)
    add_header_table(doc, exam_title)

    # 2. 逐一產生各題型的答案卡表格
    for sec in sections:
        # 大題標題
        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_before = Pt(8)
        p_title.paragraph_format.space_after = Pt(4)
        run = p_title.add_run(sec["title"])
        run.bold = True
        run.font.name = '微軟正黑體'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
        run.font.size = Pt(11.5)

        q_list = sec["questions"]
        s_type = sec["type"]

        if s_type in ["CHOICE", "TRUE_FALSE", "FILL_BLANK"]:
            # --- 選擇題 / 是非題 / 填填看：10 欄網格表格 ---
            col_max = 10
            total_q = len(q_list)
            # 以 10 題為一組分排
            for chunk_start in range(0, total_q, col_max):
                chunk = q_list[chunk_start:chunk_start + col_max]
                num_cols = len(chunk)
                
                # 建立 2 列 (第 1 列題號，第 2 列空白填寫格)
                tbl = doc.add_table(rows=2, cols=num_cols)
                tbl.style = 'Table Grid'
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.autofit = False

                tbl.rows[0].height = Pt(22)
                tbl.rows[1].height = Pt(36)

                for idx, q_num in enumerate(chunk):
                    # 題號列
                    c_num = tbl.cell(0, idx)
                    format_cell_text(c_num, f"{q_num}.", bold=True, size_pt=10, align=WD_ALIGN_PARAGRAPH.CENTER)
                    set_cell_background(c_num, "F2F2F2")

                    # 作答列
                    c_ans = tbl.cell(1, idx)
                    c_ans.text = ""
                    c_ans.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                # 表格下方微距
                p_space = doc.add_paragraph()
                p_space.paragraph_format.space_before = Pt(2)
                p_space.paragraph_format.space_after = Pt(2)

        elif s_type == "CALCULATION":
            # --- 算算看 / 計算題：3 欄大型計算網格 ---
            col_cnt = 3
            total_q = len(q_list)
            for chunk_start in range(0, total_q, col_cnt):
                chunk = q_list[chunk_start:chunk_start + col_cnt]
                num_cols = len(chunk)

                tbl = doc.add_table(rows=1, cols=num_cols)
                tbl.style = 'Table Grid'
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.autofit = False
                tbl.rows[0].height = Pt(130) # 足夠寫直式的垂直高度

                for idx, q_num in enumerate(chunk):
                    cell = tbl.cell(0, idx)
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_before = Pt(4)
                    run_num = p.add_run(f"{q_num}.")
                    run_num.bold = True
                    run_num.font.name = '微軟正黑體'
                    run_num._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
                    run_num.font.size = Pt(11)

                p_space = doc.add_paragraph()
                p_space.paragraph_format.space_before = Pt(3)
                p_space.paragraph_format.space_after = Pt(3)

        elif s_type == "APPLICATION":
            # --- 應用題 / 問答題 / 過程題：單欄大作答框 ---
            for q_num in q_list:
                tbl = doc.add_table(rows=1, cols=1)
                tbl.style = 'Table Grid'
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.autofit = False
                tbl.rows[0].height = Pt(110) # 寬敞的解題與作答框

                cell = tbl.cell(0, 0)
                p_q = cell.paragraphs[0]
                p_q.paragraph_format.space_before = Pt(4)
                run_q = p_q.add_run(f"{q_num}. ")
                run_q.bold = True
                run_q.font.name = '微軟正黑體'
                run_q._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
                run_q.font.size = Pt(11)

                # 底部加入「答：」
                p_ans = cell.add_paragraph()
                p_ans.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_ans.paragraph_format.space_before = Pt(65)
                run_ans = p_ans.add_run("答：____________________")
                run_ans.font.name = '微軟正黑體'
                run_ans._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
                run_ans.font.size = Pt(10.5)

                p_space = doc.add_paragraph()
                p_space.paragraph_format.space_before = Pt(3)
                p_space.paragraph_format.space_after = Pt(3)
        else:
            # 其他綜合型：5 欄答題格
            col_max = 5
            total_q = len(q_list)
            for chunk_start in range(0, total_q, col_max):
                chunk = q_list[chunk_start:chunk_start + col_max]
                tbl = doc.add_table(rows=2, cols=len(chunk))
                tbl.style = 'Table Grid'
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.rows[0].height = Pt(22)
                tbl.rows[1].height = Pt(36)

                for idx, q_num in enumerate(chunk):
                    c_num = tbl.cell(0, idx)
                    format_cell_text(c_num, f"{q_num}.", bold=True, size_pt=10)
                    set_cell_background(c_num, "F2F2F2")

                    c_ans = tbl.cell(1, idx)
                    c_ans.text = ""

                p_space = doc.add_paragraph()
                p_space.paragraph_format.space_before = Pt(2)
                p_space.paragraph_format.space_after = Pt(2)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    return doc_io.getvalue()

# --- 主介面 Streamlit ---

def main():
    st.set_page_config(
        page_title="考卷轉換與空白答案卷生成器",
        page_icon="📝",
        layout="centered"
    )

    st.markdown("""
        <style>
        .main-header {
            font-size: 2.2rem;
            font-weight: 700;
            color: #1E3A8A;
            text-align: center;
            margin-bottom: 0.5rem;
        }
        .sub-header {
            font-size: 1.05rem;
            color: #4B5563;
            text-align: center;
            margin-bottom: 2rem;
        }
        .stButton>button {
            width: 100%;
            height: 3rem;
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

        # 匯出模式單選 (解決分開下載的困擾)
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
            with st.spinner("正在分析試卷結構、處理圖像與排版中，請稍候..."):
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
