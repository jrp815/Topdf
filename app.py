"""
文件合并转PDF工具 - Streamlit Web应用
支持：Word(doc/docx)、PDF、Excel(xls/xlsx)、图片(jpg/png/bmp)

作者: QClaw
部署: Streamlit Cloud
"""

import streamlit as st
import os
import tempfile
import io
from datetime import datetime
from pathlib import Path

# PDF处理
from pypdf import PdfWriter, PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Image as RLImage, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 文档处理
from docx import Document
import openpyxl
from PIL import Image

# 页面配置
st.set_page_config(
    page_title="文件合并转PDF",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 注册中文字体（Streamlit Cloud环境）
def setup_chinese_font():
    """注册CJK字体，适配不同平台"""
    import platform
    system = platform.system()
    
    font_paths = []
    if system == 'Windows':
        windir = os.environ.get('WINDIR', 'C:\\Windows')
        font_paths = [
            os.path.join(windir, 'Fonts', 'msyh.ttc'),
            os.path.join(windir, 'Fonts', 'simhei.ttf'),
            os.path.join(windir, 'Fonts', 'simsun.ttc'),
        ]
    elif system == 'Darwin':  # macOS
        font_paths = [
            '/System/Library/Fonts/STHeiti Light.ttc',
            '/System/Library/Fonts/Supplemental/Songti.ttc',
        ]
    else:  # Linux (Streamlit Cloud)
        font_paths = [
            '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
            '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
            '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
            '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        ]
    
    cn_font = 'Helvetica'  # 默认fallback
    for fp in font_paths:
        if os.path.exists(fp):
            try:
                font_name = os.path.splitext(os.path.basename(fp))[0]
                if fp.endswith('.ttc'):
                    pdfmetrics.registerFont(TTFont(font_name, fp, subfontIndex=0))
                else:
                    pdfmetrics.registerFont(TTFont(font_name, fp))
                cn_font = font_name
                break
            except Exception as e:
                continue
    
    # 创建样式
    styles = getSampleStyleSheet()
    for style in styles.byName.values():
        if isinstance(style, ParagraphStyle):
            style.fontName = cn_font
    
    return cn_font, styles

@st.cache_resource
def get_font_setup():
    return setup_chinese_font()

# 转换函数
def convert_docx_to_pdf(docx_path, pdf_path, cn_font, styles):
    """将docx转换为PDF"""
    doc = Document(docx_path)
    
    pdf_doc = SimpleDocTemplate(pdf_path, pagesize=A4,
                                 leftMargin=20*mm, rightMargin=20*mm,
                                 topMargin=20*mm, bottomMargin=20*mm)
    
    story = []
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            story.append(Spacer(1, 6))
            continue
        
        # 根据段落样式选择PDF样式
        style_name = 'Normal'
        if para.style.name.startswith('Heading'):
            style_name = 'Heading1'
        elif 'Title' in para.style.name:
            style_name = 'Title'
        
        try:
            p = Paragraph(text, styles[style_name])
        except:
            p = Paragraph(text, styles['Normal'])
        story.append(p)
    
    # 处理表格
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                row_data.append(Paragraph(cell_text, styles['Normal']))
            table_data.append(row_data)
        
        if table_data:
            # 计算列宽
            num_cols = len(table_data[0])
            col_width = (A4[0] - 40*mm) / num_cols
            
            t = Table(table_data, colWidths=[col_width]*num_cols)
            t.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E0E0E0')),
                ('FONTNAME', (0, 0), (-1, -1), cn_font),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            story.append(t)
            story.append(Spacer(1, 12))
    
    pdf_doc.build(story)

def convert_xlsx_to_pdf(xlsx_path, pdf_path, cn_font, styles):
    """将xlsx转换为PDF"""
    wb = openpyxl.load_workbook(xlsx_path)
    
    pdf_doc = SimpleDocTemplate(pdf_path, pagesize=A4,
                                 leftMargin=15*mm, rightMargin=15*mm,
                                 topMargin=15*mm, bottomMargin=15*mm)
    
    story = []
    
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        
        # 添加工作表标题
        story.append(Paragraph(f"工作表: {sheet_name}", styles['Heading2']))
        story.append(Spacer(1, 10))
        
        # 获取表格数据
        table_data = []
        max_rows = min(sheet.max_row, 100)  # 限制行数
        max_cols = min(sheet.max_column, 20)  # 限制列数
        
        for row_idx in range(1, max_rows + 1):
            row_data = []
            for col_idx in range(1, max_cols + 1):
                cell = sheet.cell(row=row_idx, column=col_idx)
                value = str(cell.value) if cell.value is not None else ''
                row_data.append(Paragraph(value, styles['Normal']))
            table_data.append(row_data)
        
        if table_data:
            col_width = (A4[0] - 30*mm) / max_cols
            t = Table(table_data, colWidths=[col_width]*max_cols)
            t.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E0E0E0')),
                ('FONTNAME', (0, 0), (-1, -1), cn_font),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
            ]))
            story.append(t)
        
        story.append(PageBreak())
    
    pdf_doc.build(story)

def convert_image_to_pdf(img_path, pdf_path):
    """将图片转换为PDF（A4页面）"""
    img = Image.open(img_path)
    if img.mode in ('RGBA', 'P'):
        img = img.convert('RGB')
    
    page_w, page_h = A4
    margin = 15 * mm
    avail_w = page_w - 2 * margin
    avail_h = page_h - 2 * margin
    
    img_w, img_h = img.size
    ratio = min(avail_w / img_w, avail_h / img_h)
    display_w = img_w * ratio
    display_h = img_h * ratio
    
    pdf_doc = SimpleDocTemplate(pdf_path, pagesize=A4,
                                 leftMargin=margin, rightMargin=margin,
                                 topMargin=margin, bottomMargin=margin)
    story = [RLImage(img_path, width=display_w, height=display_h)]
    pdf_doc.build(story)

def merge_pdfs(pdf_paths, output_path):
    """合并多个PDF"""
    writer = PdfWriter()
    for pdf_path in pdf_paths:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            writer.add_page(page)
    
    with open(output_path, 'wb') as f:
        writer.write(f)

# Streamlit应用
def main():
    st.title("📄 文件合并转PDF工具")
    st.markdown("---")
    
    # 侧边栏说明
    with st.sidebar:
        st.header("📖 使用说明")
        st.markdown("""
        **支持的文件格式：**
        - Word文档：`.docx`
        - Excel表格：`.xlsx`
        - PDF文档：`.pdf`
        - 图片文件：`.jpg`, `.png`, `.bmp`
        
        **使用步骤：**
        1. 上传需要合并的文件
        2. 调整文件顺序（拖拽或重新上传）
        3. 点击"合并生成PDF"
        4. 下载合并后的PDF文件
        
        **注意事项：**
        - 图片会自动适配A4页面
        - Excel表格最多显示100行、20列
        - `.doc`格式建议先转换为`.docx`
        """)
        
        st.markdown("---")
        st.markdown("💡 提示：文件顺序按上传顺序排列")
    
    # 文件上传区域
    st.header("1️⃣ 上传文件")
    uploaded_files = st.file_uploader(
        "选择要合并的文件",
        accept_multiple_files=True,
        type=['docx', 'xlsx', 'pdf', 'jpg', 'jpeg', 'png', 'bmp'],
        help="支持多文件上传，可混合不同格式"
    )
    
    if uploaded_files:
        st.success(f"已上传 {len(uploaded_files)} 个文件")
        
        # 显示文件列表
        st.header("2️⃣ 文件列表")
        
        # 创建可排序的文件列表
        file_order = []
        for i, file in enumerate(uploaded_files):
            col1, col2, col3 = st.columns([1, 4, 2])
            with col1:
                st.write(f"**{i+1}**")
            with col2:
                st.write(f"{file.name}")
            with col3:
                ext = Path(file.name).suffix.lower()
                type_icon = {
                    '.docx': '📝 Word',
                    '.xlsx': '📊 Excel',
                    '.pdf': '📄 PDF',
                    '.jpg': '🖼️ 图片',
                    '.jpeg': '🖼️ 图片',
                    '.png': '🖼️ 图片',
                    '.bmp': '🖼️ 图片'
                }.get(ext, '📁 文件')
                st.write(type_icon)
            file_order.append(i)
        
        # 合并按钮
        st.header("3️⃣ 生成PDF")
        
        if st.button("🔗 合并生成PDF", type="primary", use_container_width=True):
            with st.spinner("正在处理文件..."):
                try:
                    cn_font, styles = get_font_setup()
                    
                    # 创建临时目录
                    temp_dir = tempfile.mkdtemp()
                    pdf_files = []
                    
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    for i, file in enumerate(uploaded_files):
                        status_text.text(f"处理: {file.name}")
                        progress_bar.progress((i + 1) / len(uploaded_files))
                        
                        # 保存上传文件到临时目录
                        file_path = os.path.join(temp_dir, file.name)
                        with open(file_path, 'wb') as f:
                            f.write(file.getvalue())
                        
                        # 转换为PDF
                        ext = Path(file.name).suffix.lower()
                        pdf_path = os.path.join(temp_dir, f"{i:03d}.pdf")
                        
                        if ext == '.pdf':
                            # 直接复制
                            with open(file_path, 'rb') as src:
                                with open(pdf_path, 'wb') as dst:
                                    dst.write(src.read())
                            pdf_files.append(pdf_path)
                        
                        elif ext == '.docx':
                            convert_docx_to_pdf(file_path, pdf_path, cn_font, styles)
                            pdf_files.append(pdf_path)
                        
                        elif ext == '.xlsx':
                            convert_xlsx_to_pdf(file_path, pdf_path, cn_font, styles)
                            pdf_files.append(pdf_path)
                        
                        elif ext in ['.jpg', '.jpeg', '.png', '.bmp']:
                            convert_image_to_pdf(file_path, pdf_path)
                            pdf_files.append(pdf_path)
                    
                    # 合并PDF
                    status_text.text("正在合并PDF...")
                    merged_path = os.path.join(temp_dir, "merged.pdf")
                    merge_pdfs(pdf_files, merged_path)
                    
                    # 读取合并后的PDF
                    with open(merged_path, 'rb') as f:
                        pdf_bytes = f.read()
                    
                    # 清理临时文件
                    import shutil
                    shutil.rmtree(temp_dir, ignore_errors=True)
                    
                    progress_bar.progress(100)
                    status_text.text("完成！")
                    
                    # 显示下载按钮
                    st.success(f"✅ 合并完成！共 {len(uploaded_files)} 个文件，{len(PdfReader(io.BytesIO(pdf_bytes)).pages)} 页")
                    
                    # 生成文件名
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    download_filename = f"合并文档_{timestamp}.pdf"
                    
                    st.download_button(
                        label="📥 下载合并后的PDF",
                        data=pdf_bytes,
                        file_name=download_filename,
                        mime="application/pdf",
                        use_container_width=True
                    )
                    
                except Exception as e:
                    st.error(f"处理失败: {str(e)}")
                    st.exception(e)
    
    else:
        st.info("👆 请先上传需要合并的文件")

if __name__ == "__main__":
    main()