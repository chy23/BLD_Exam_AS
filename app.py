import streamlit as st
import fitz  # PyMuPDF
import cv2
import numpy as np
from PIL import Image
import io
from docx import Document
from docx.shared import Pt
import re

def remove_red_from_image(pdf_bytes):
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
        
        # Red hue wraps around 0 and 180
        lower_red1 = np.array([0, 50, 50])
        upper_red1 = np.array([10, 255, 255])
        lower_red2 = np.array([170, 50, 50])
        upper_red2 = np.array([180, 255, 255])
        
        mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
        mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
        mask = mask1 + mask2
        
        # Replace red pixels with white
        img_cv[mask > 0] = [255, 255, 255]
        
        result_image = cv2.cvtColor(img_cv, cv2.COLOR_BGR2RGB)
        processed_images.append(Image.fromarray(result_image))
        
    pdf_bytes_io = io.BytesIO()
    if processed_images:
        processed_images[0].save(
            pdf_bytes_io, 
            format='PDF', 
            save_all=True, 
            append_images=processed_images[1:]
        )
    return pdf_bytes_io.getvalue()

def generate_answer_sheet(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    docx_doc = Document()
    
    full_text = ""
    for page in doc:
        full_text += page.get_text()
        
    lines = full_text.split('\n')
    
    header_line = ""
    for line in lines:
        if "學年度" in line or "試卷" in line or "國小" in line or "國中" in line:
            header_line = line
            break
    if not header_line and lines:
        header_line = lines[0]
        
    header_para = docx_doc.add_paragraph()
    run = header_para.add_run(header_line)
    run.font.size = Pt(14)
    run.bold = True
    header_para.alignment = 1 # Center
    
    info_para = docx_doc.add_paragraph("班級：_______ 年_______班 座號：_______ 姓名：_____________")
    info_para.alignment = 1
    
    section_pattern = re.compile(r'^([一二三四五六七八九十]、.*)')
    question_pattern = re.compile(r'^(\d+)[\.．]\s*(.*)')
    
    current_section = ""
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        sec_match = section_pattern.match(line)
        if sec_match:
            current_section = sec_match.group(1)
            p = docx_doc.add_paragraph()
            run = p.add_run(current_section)
            run.bold = True
            run.font.size = Pt(12)
            continue
            
        q_match = question_pattern.match(line)
        if q_match:
            q_num = q_match.group(1)
            
            if "算算看" in current_section or "計算" in current_section or "應用題" in current_section or "問答題" in current_section or "過程" in current_section:
                docx_doc.add_paragraph(f"{q_num}.")
                docx_doc.add_paragraph("\n\n") 
            elif "選擇題" in current_section or "是非題" in current_section:
                docx_doc.add_paragraph(f"{q_num}. (   )")
            elif "填填看" in current_section or "填空題" in current_section:
                docx_doc.add_paragraph(f"{q_num}. ________________")
            elif "連連看" in current_section:
                 docx_doc.add_paragraph(f"{q_num}. (請在題目卷上作答)")
            else:
                docx_doc.add_paragraph(f"{q_num}. ________________")

    docx_bytes_io = io.BytesIO()
    docx_doc.save(docx_bytes_io)
    return docx_bytes_io.getvalue()

st.set_page_config(page_title="考卷轉換工具", page_icon="📄")

st.title("考卷轉換工具")
st.write("上傳包含答案的考卷（紅色字體），可自動消除答案並匯出「純題目卷」，或自動產生「空白答題卷」。支援 A4, A3, B4 等各種格式。")

uploaded_file = st.file_uploader("請上傳考卷 PDF 檔", type=["pdf"])

if uploaded_file is not None:
    pdf_bytes = uploaded_file.read()
    
    st.write("### 選擇匯出選項")
    export_blank = st.checkbox("匯出純題目卷 (PDF)", value=True, help="將考卷中的紅色字體（答案）消除，保留原始題目格式。")
    export_answer_sheet = st.checkbox("匯出空白答案卡卷 (Docx)", value=True, help="自動分析題目結構，產生對應的空白答題卷。")
    
    if st.button("開始轉換", type="primary"):
        if not export_blank and not export_answer_sheet:
            st.warning("請至少選擇一種匯出格式。")
        else:
            with st.spinner("處理中，請稍候..."):
                if export_blank:
                    blank_pdf_bytes = remove_red_from_image(pdf_bytes)
                    st.download_button(
                        label="⬇️ 下載純題目卷 (PDF)",
                        data=blank_pdf_bytes,
                        file_name=uploaded_file.name.replace(".pdf", "_純題目卷.pdf"),
                        mime="application/pdf"
                    )
                
                if export_answer_sheet:
                    answer_sheet_bytes = generate_answer_sheet(pdf_bytes)
                    st.download_button(
                        label="⬇️ 下載空白答案卡卷 (Docx)",
                        data=answer_sheet_bytes,
                        file_name=uploaded_file.name.replace(".pdf", "_空白答題卷.docx"),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
