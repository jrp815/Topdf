"""
文件合并转PDF工具 - Streamlit Web应用
支持：Word(doc/docx)、PDF、Excel(xls/xlsx)、图片(jpg/png/bmp)

作者: QClaw
部署: Streamlit Cloud
版本: v1.3.0

转换引擎（自动选择）：
1. LibreOffice headless  → 完美保留格式（需系统安装）
2. ReportLab + python-docx/openpyxl → 纯Python fallback（保留大部分格式）
3. 图片 → Pillow 嵌入A4
"""

import streamlit as st
import os
import tempfile
import io
import subprocess
import platform
import shutil
from datetime import datetime
from pathlib import Path

from pypdf import PdfWriter, PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.platypus import (SimpleDocTemplate, Image as RLImage,
                                 Paragraph, Spacer, Table, TableStyle, PageBreak)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY

from docx import Document
from docx.shared import Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openpyxl
from PIL import Image


# ============================================================
# 版本信息
# ============================================================
__version__ = "v1.3.0"
__version_date = "2026-05-31"


# ============================================================
# 中文字体
# ============================================================
FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts')
EMBEDDED_FONT = os.path.join(FONT_DIR, 'ChineseFont.ttf')


def setup_chinese_font():
    """注册中文字体，返回 (font_name, styles)"""
    cn_font = 'ChineseFont'

    if os.path.exists(EMBEDDED_FONT) and os.path.getsize(EMBEDDED_FONT) > 100000:
        pdfmetrics.registerFont(TTFont(cn_font, EMBEDDED_FONT))
    else:
        # Fallback 系统字体
        system = platform.system()
        cands = []
        if system == 'Windows':
            base = os.environ.get('WINDIR', 'C:\\Windows')
            cands = [os.path.join(base, 'Fonts', f) for f in ['msyh.ttc', 'simhei.ttf', 'simsun.ttc']]
        elif system == 'Darwin':
            cands = ['/System/Library/Fonts/STHeiti Light.ttc']
        else:
            cands = [
                '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
                '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
            ]
        for fp in cands:
            if os.path.exists(fp):
                try:
                    kw = {'subfontIndex': 0} if fp.endswith('.ttc') else {}
                    pdfmetrics.registerFont(TTFont(cn_font, fp, **kw))
                    break
                except Exception:
                    continue

    styles = getSampleStyleSheet()

    style_map = {
        'Title':     {'fontSize': 22, 'spaceAfter': 18, 'alignment': TA_CENTER},
        'Heading1':  {'fontSize': 16, 'spaceBefore': 14, 'spaceAfter': 8},
        'Heading2':  {'fontSize': 13, 'spaceBefore': 10, 'spaceAfter': 6},
        'Heading3':  {'fontSize': 11, 'spaceBefore': 8,  'spaceAfter': 4},
        'Normal':    {'fontSize': 10.5, 'leading': 18, 'alignment': TA_JUSTIFY,
                      'firstLineIndent': Pt(21),  # 中文首行缩进2字符
                      },
        'BodyText':  {'fontSize': 10.5, 'leading': 18, 'alignment': TA_JUSTIFY,
                      'firstLineIndent': Pt(21),
                      },
    }

    for name, props in style_map.items():
        if name in styles.byName:
            s = styles.byName[name]
        else:
            s = ParagraphStyle(name, parent=styles['Normal'])
            styles.add(s)
        s.fontName = cn_font
        for k, v in props.items():
            setattr(s, k, v)

    return cn_font, styles


@st.cache_resource
def get_font_setup():
    return setup_chinese_font()


# ============================================================
# LibreOffice 引擎（可选增强）
# ============================================================
def check_libreoffice():
    try:
        r = subprocess.run(['libreoffice', '--version'], capture_output=True, text=True, timeout=10)
        return bool(r.stdout.strip()), r.stdout.strip()
    except Exception:
        return False, None


def convert_with_libreoffice(src_path, output_dir):
    """用 LibreOffice headless 转换 → PDF（完美保留格式）"""
    try:
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, src_path],
            capture_output=True, text=True, timeout=120,
            env={**os.environ, 'HOME': os.environ.get('HOME', '/tmp')}
        )
        base = Path(src_path).stem
        pdf_path = os.path.join(output_dir, f"{base}.pdf")
        if os.path.exists(pdf_path):
            return pdf_path
        for f in os.listdir(output_dir):
            if f.endswith('.pdf'):
                fp = os.path.join(output_dir, f)
                if os.path.getmtime(fp) > os.path.gettime(src_path) - 1:
                    return fp
        return None
    except Exception:
        return None


# ============================================================
# 颜色 / 对齐 辅助函数
# ============================================================
def _rl_color(c):
    if c is None:
        return colors.black
    if isinstance(c, RGBColor):
        return colors.Color(c.red / 255, c.green / 255, c.blue / 255)
    if isinstance(c, str) and c.startswith('#'):
        return colors.HexColor(c)
    return colors.black


def _rl_align(align):
    m = {WD_ALIGN_PARAGRAPH.LEFT: TA_LEFT, WD_ALIGN_PARAGRAPH.CENTER: TA_CENTER,
         WD_ALIGN_PARAGRAPH.RIGHT: TA_RIGHT, WD_ALIGN_PARAGRAPH.JUSTIFY: TA_JUSTIFY}
    return m.get(align, TA_JUSTIFY)


def _emu_to_pt(emu_val):
    """将 EMU 单位转换为 Point (1pt = 12700 EMU)"""
    if emu_val is None:
        return None
    try:
        return emu_val.pt if hasattr(emu_val, 'pt') else float(emu_val) / 12700
    except (AttributeError, TypeError):
        return None


# ============================================================
# Word → PDF（增强版：保留格式细节）
# ============================================================
def convert_docx_to_pdf(docx_path, pdf_path, cn_font, styles):
    """python-docx + ReportLab，尽可能保留原始格式"""
    doc = Document(docx_path)

    pdf_doc = SimpleDocTemplate(
        pdf_path, pagesize=A4,
        leftMargin=25*mm, rightMargin=25*mm,
        topMargin=25*mm, bottomMargin=25*mm
    )
    story = []

    # 按文档 body 元素顺序处理（保持段落和表格的原始顺序）
    body = doc.element.body

    for elem in body:
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

        # ---- 段落 ----
        if tag == 'p':
            from docx.text.paragraph import Paragraph as DPara
            para = DPara(elem, doc)

            # 收集所有 run 的文本和格式 → 构建带格式的 XML 片段
            # ReportLab Paragraph 支持 <b>、<i>、<font>、<color> 等 XML 标签
            parts = []
            has_formatting = False
            for run in para.runs:
                text = run.text or ''
                if not text.strip() and not text:
                    continue
                
                # 获取 run 格式
                bold = run.bold
                italic = run.italic
                underline = run.underline
                font_size = None
                font_name = None
                color_val = None
                
                try:
                    rpr = run._element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
                    if rpr is not None:
                        # 字体大小
                        sz = rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
                        if sz is not None and sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'):
                            font_size = int(sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')) / 2  # half-points to points
                        
                        # 字体名称
                        rFonts = rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
                        if rFonts is not None:
                            font_name = rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii') or \
                                       rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
                        
                        # 颜色
                        color_elem = rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
                        if color_elem is not None and color_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006main}val'):
                            color_val = color_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                except Exception:
                    pass
                
                # 如果有特殊格式，构建 XML 标签
                formatted = text
                if bold or italic or underline or font_size or color_val or font_name:
                    has_formatting = True
                    
                    tags_open = []
                    tags_close = []
                    
                    if font_name:
                        tags_open.append(f'<font face="{cn_font}" name="{cn_font}">')
                        tags_close.insert(0, '</font>')
                    
                    if font_size:
                        tags_open.append(f'<font size="{font_size}">')
                        tags_close.insert(0, '</font>')
                    
                    if color_val:
                        tags_open.append(f'<font color="#{color_val}">')
                        tags_close.insert(0, '</font>')
                    
                    if bold:
                        tags_open.append('<b>')
                        tags_close.insert(0, '</b>')
                    
                    if italic:
                        tags_open.append('<i>')
                        tags_close.insert(0, '</i>')
                    
                    if underline:
                        tags_open.append('<u>')
                        tags_close.insert(0, '</u>')
                    
                    formatted = ''.join(tags_open) + text + ''.join(tags_close)
                
                parts.append(formatted)
            
            full_text = ''.join(parts).strip()
            if not full_text:
                story.append(Spacer(1, 6))
                continue

            # 判断段落样式级别
            sn = para.style.name
            if 'Heading1' in sn or ('Heading' in sn and '1' in sn):
                base_sn = 'Heading1'
            elif 'Heading2' in sn or ('Heading' in sn and '2' in sn):
                base_sn = 'Heading2'
            elif 'Heading3' in sn or ('Heading' in sn and '3' in sn):
                base_sn = 'Heading3'
            elif 'Title' in sn:
                base_sn = 'Title'
            else:
                base_sn = 'Normal'

            pf = para.paragraph_format
            
            # 安全获取数值，确保不为 None
            def safe_pt(val, default=0):
                """安全地将值转换为 Pt，如果为 None 则返回默认值"""
                if val is None:
                    return Pt(default)
                try:
                    return Pt(val.pt) if hasattr(val, 'pt') else Pt(float(val))
                except (AttributeError, TypeError):
                    return Pt(default)

            ps = ParagraphStyle(
                'auto_para',
                parent=styles[base_sn],
                fontSize=getattr(styles[base_sn], 'fontSize', None) or 10.5,
                leading=getattr(styles[base_sn], 'leading', None) or (getattr(styles[base_sn], 'fontSize', None) or 10.5) * 1.7,
                alignment=_rl_align(para.alignment),
                spaceBefore=safe_pt(pf.space_before, 0),
                spaceAfter=safe_pt(pf.space_after, 0),
                leftIndent=safe_pt(pf.left_indent, 0),
                rightIndent=safe_pt(pf.right_indent, 0),
                firstLineIndent=safe_pt(pf.first_line_indent, 0),  # Normal样式已有首行缩进
            )

            story.append(Paragraph(full_text, ps))

        # ---- 表格 ----
        elif tag == 'tbl':
            from docx.table import Table as DTbl
            table = DTbl(elem, doc)
            table_data = []
            ncols = len(table.columns)

            for row_idx, row in enumerate(table.rows):
                row_data = []
                for cell in row.cells:
                    # 处理单元格内的段落，保留基本格式
                    cell_parts = []
                    for p in cell.paragraphs:
                        if p.text.strip():
                            # 尝试保留单元格内文本的粗体等格式
                            run_texts = []
                            for run in p.runs:
                                txt = run.text or ''
                                if run.bold:
                                    txt = f'<b>{txt}</b>'
                                run_texts.append(txt)
                            cell_parts.append(''.join(run_texts))
                    cell_text = '<br/>'.join(cell_parts) if cell_parts else ' '
                    row_data.append(Paragraph(cell_text, styles['Normal']))
                table_data.append(row_data)

            if table_data and ncols > 0:
                col_w = (A4[0] - 50*mm) / ncols
                t = Table(table_data, colWidths=[col_w] * ncols)

                cmds = [
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('FONTNAME', (0, 0), (-1, -1), cn_font),
                    ('FONTSIZE', (0, 0), (-1, -1), 9),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('TOPPADDING', (0, 0), (-1, -1), 5),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                    ('LEFTPADDING', (0, 0), (-1, -1), 4),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ]
                if len(table_data) > 1:
                    cmds.append(('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E8E8E8')))

                t.setStyle(TableStyle(cmds))
                story.append(t)
                story.append(Spacer(1, 12))

    pdf_doc.build(story)


# ============================================================
# Excel → PDF（增强版：保留表格结构）
# ============================================================
def convert_xlsx_to_pdf(xlsx_path, pdf_path, cn_font, styles):
    """openpyxl + ReportLab，保留 Excel 表格结构和数据"""
    wb = openpyxl.load_workbook(xlsx_path)

    pdf_doc = SimpleDocTemplate(
        pdf_path, pagesize=A4,
        leftMargin=20*mm, rightMargin=20*mm,
        topMargin=20*mm, bottomMargin=20*mm
    )
    story = []

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        story.append(Paragraph(f"工作表: {sheet_name}", styles['Heading2']))
        story.append(Spacer(1, 8))

        max_r = min(sheet.max_row, 200)
        max_c = min(sheet.max_column, 26)

        table_data = []
        for r in range(1, max_r + 1):
            row_data = []
            for c in range(1, max_c + 1):
                val = sheet.cell(row=r, column=c).value
                txt = str(val) if val is not None else ''
                row_data.append(Paragraph(txt, styles['Normal']))
            table_data.append(row_data)

        if table_data:
            col_w = (A4[0] - 40*mm) / max_c
            t = Table(table_data, colWidths=[col_w] * max_c)

            cmds = [
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('FONTNAME', (0, 0), (-1, -1), cn_font),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]
            if len(table_data) > 1:
                cmds.append(('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E8E8E8')))

            t.setStyle(TableStyle(cmds))
            story.append(t)

        story.append(PageBreak())

    pdf_doc.build(story)


# ============================================================
# 图片 → PDF
# ============================================================
def convert_image_to_pdf(img_path, pdf_path):
    """图片嵌入 A4 页面"""
    img = Image.open(img_path)
    if img.mode in ('RGBA', 'P'):
        img = img.convert('RGB')

    margin = 15 * mm
    avail_w = A4[0] - 2 * margin
    avail_h = A4[1] - 2 * margin

    img_w, img_h = img.size
    ratio = min(avail_w / img_w, avail_h / img_h)

    pdf_doc = SimpleDocTemplate(pdf_path, pagesize=A4,
                                leftMargin=margin, rightMargin=margin,
                                topMargin=margin, bottomMargin=margin)
    story = [RLImage(img_path, width=img_w * ratio, height=img_h * ratio)]
    pdf_doc.build(story)


# ============================================================
# 合并 PDF
# ============================================================
def merge_pdfs(pdf_paths, output_path):
    writer = PdfWriter()
    for p in pdf_paths:
        reader = PdfReader(p)
        for page in reader.pages:
            writer.add_page(page)
    with open(output_path, 'wb') as f:
        writer.write(f)


# ============================================================
# Streamlit 主界面
# ============================================================
st.set_page_config(page_title="📄 文件合并转PDF", page_icon="📄", layout="wide")


def main():
    st.title("📄 文件合并转PDF工具")
    
    # 版本显示
    version_col1, version_col2, version_col3 = st.columns([1, 2, 1])
    with version_col2:
        st.markdown(f"<div style='text-align:center; padding:4px 12px; background:#f0f2f6; border-radius:8px; font-size:14px;'>"
                   f"🏷️ 版本 <b>{__version__}</b> · {__version_date} · "
                   f"<a href='https://github.com/jrp815/Topdf' target='_blank'>GitHub</a>"
                   f"</div>", unsafe_allow_html=True)
    
    st.markdown("保持源文件格式 · 支持 Word/Excel/PDF/图片")

    # 侧边栏
    with st.sidebar:
        st.header("📖 使用说明")
        st.markdown(f"""**版本：{__version__}**

**支持格式：**
- 📝 Word：`.doc` `.docx`
- 📊 Excel：`.xls` `.xlsx`
- 📄 PDF：`.pdf`
- 🖼️ 图片：`.jpg` `.png` `.bmp`

**步骤：**
1. 上传文件
2. 点击「合并生成PDF」
3. 下载结果

**格式保留：**
- ✅ 有 LibreOffice 时：**完美保留**
- ⚠️ 纯 Python 时：保留文字、表格、标题层级、对齐方式、**粗体/斜体/下划线**
""")
        lo_ok, lo_ver = check_libreoffice()
        if lo_ok:
            st.success(f'✅ LibreOffice\n`{lo_ver.split(chr(10))[0]}`')
        else:
            st.info('ℹ️ 使用纯 Python 引擎')

    # 上传
    st.header("1️⃣ 上传文件")
    uploaded_files = st.file_uploader(
        "选择要合并的文件",
        accept_multiple_files=True,
        type=['doc', 'docx', '.xls', '.xlsx', 'pdf', 'jpg', 'jpeg', 'png', 'bmp'],
    )

    if not uploaded_files:
        st.info("👆 请先上传需要合并的文件")
        return

    st.success(f"已上传 **{len(uploaded_files)}** 个文件")

    # 文件列表
    st.header("2️⃣ 文件列表")
    icons = {
        '.doc': '📝', '.docx': '📝', '.xls': '📊', '.xlsx': '📊',
        '.pdf': '📄', '.jpg': '🖼️', '.jpeg': '🖼️', '.png': '🖼️', '.bmp': '🖼️'
    }
    types = {
        '.doc': 'Word', '.docx': 'Word', '.xls': 'Excel', '.xlsx': 'Excel',
        '.pdf': 'PDF', '.jpg': '图片', '.jpeg': '图片', '.png': '图片', '.bmp': '图片'
    }

    for i, f in enumerate(uploaded_files):
        ext = Path(f.name).suffix.lower()
        c = st.columns(5)
        c[0].write(f"`{i+1}`")
        c[1].write(f.name)
        c[2].write(f"{icons.get(ext,'📁')} {types.get(ext,'其他')}")
        c[3].write(f"{f.size/1024:.1f} KB")

    # 合并按钮
    st.header("3️⃣ 生成PDF")
    if st.button("🔗 合并生成PDF", type="primary", use_container_width=True):
        with st.spinner("正在处理..."):
            try:
                cn_font, styles = get_font_setup()

                temp_dir = tempfile.mkdtemp()
                conv_dir = os.path.join(temp_dir, "conv")
                os.makedirs(conv_dir, exist_ok=True)

                pdf_files = []
                logs = []
                pb = st.progress(0)
                status = st.empty()

                lo_available, _ = check_libreoffice()

                for i, f in enumerate(uploaded_files):
                    status.text(f"📂 ({i+1}/{len(uploaded_files)}) {f.name}")
                    pb.progress((i + 0.5) / len(uploaded_files))

                    # 保存上传文件
                    fpath = os.path.join(temp_dir, f.name)
                    with open(fpath, 'wb') as fh:
                        fh.write(f.getvalue())

                    ext = Path(f.name).suffix.lower()
                    out_pdf = os.path.join(conv_dir, f"{i:03d}.pdf")

                    if ext == '.pdf':
                        shutil.copy(fpath, out_pdf)
                        pdf_files.append(out_pdf)
                        logs.append(f"✅ {f.name} → 直接使用")

                    elif ext in ('.doc', '.docx', '.xls', '.xlsx'):
                        # 优先用 LibreOffice
                        converted = False
                        if lo_available:
                            result = convert_with_libreoffice(fpath, conv_dir)
                            if result:
                                shutil.move(result, out_pdf)
                                pdf_files.append(out_pdf)
                                logs.append(f"✅ {f.name} → LibreOffice（完美格式）")
                                converted = True

                        if not converted:
                            # Fallback: 纯 Python
                            if ext in ('.doc', '.docx'):
                                convert_docx_to_pdf(fpath, out_pdf, cn_font, styles)
                            else:
                                convert_xlsx_to_pdf(fpath, out_pdf, cn_font, styles)
                            pdf_files.append(out_pdf)
                            logs.append(f"⚠️ {f.name} → Python引擎（基本格式）")

                    elif ext in ('.jpg', '.jpeg', '.png', '.bmp'):
                        convert_image_to_pdf(fpath, out_pdf)
                        pdf_files.append(out_pdf)
                        logs.append(f"✅ {f.name} → 图片嵌入")

                    else:
                        logs.append(f"⏭️ {f.name} → 不支持")

                    pb.progress((i + 1) / len(uploaded_files))

                # 日志
                st.markdown("\n".join(logs))

                if not pdf_files:
                    st.error("没有成功转换的文件")
                    return

                # 合并
                status.text("📋 合并中...")
                merged = os.path.join(temp_dir, "merged.pdf")
                merge_pdfs(pdf_files, merged)

                with open(merged, 'rb') as fh:
                    pdf_bytes = fh.read()

                shutil.rmtree(temp_dir, ignore_errors=True)

                pb.progress(100)
                status.text("✅ 完成！")

                n_pages = len(PdfReader(io.BytesIO(pdf_bytes)).pages)
                st.success(f"🎉 合并完成！{len(pdf_files)} 个文件 → {n_pages} 页")

                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label="📥 下载 PDF",
                    data=pdf_bytes,
                    file_name=f"合并文档_{ts}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ 失败: {e}")
                st.exception(e)


if __name__ == "__main__":
    main()
