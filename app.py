"""
文件合并转PDF工具 - Streamlit Web应用
支持：Word(doc/docx)、PDF、Excel(xls/xlsx)、图片(jpg/png/bmp)

作者: QClaw
部署: Streamlit Cloud

转换引擎：
- Linux/macOS: LibreOffice headless (保持原始格式)
- Windows:     COM 自动化 (Word/Excel) + LibreOffice fallback
- 图片:        Pillow (嵌入A4)
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

# PDF处理
from pypdf import PdfWriter, PdfReader

# 图片处理
from PIL import Image


# ============================================================
# LibreOffice 检测与安装 (Streamlit Cloud / Linux)
# ============================================================
def check_libreoffice():
    """检查 LibreOffice 是否可用"""
    try:
        result = subprocess.run(
            ['libreoffice', '--version'],
            capture_output=True, text=True, timeout=10
        )
        return True, result.stdout.strip()
    except FileNotFoundError:
        return False, None
    except Exception as e:
        return False, str(e)


def install_libreoffice():
    """在 Linux 上安装 LibreOffice"""
    if platform.system() != 'Linux':
        return False
    
    st.warning('⏳ 正在安装 LibreOffice（首次约需1-2分钟）...')
    
    # 更新包列表并安装
    subprocess.run(['apt-get', 'update', '-qq'], capture_output=True, timeout=120)
    result = subprocess.run(
        ['apt-get', 'install', '-y', '-qq',
         'libreoffice', 'libreoffice-writer'],
        capture_output=True, timeout=300
    )
    
    ok, ver = check_libreoffice()
    if ok:
        st.success(f'✅ LibreOffice 安装成功: {ver}')
        return True
    else:
        st.error('❌ LibreOffice 安装失败')
        return False


def ensure_libreoffice():
    """确保 LibreOffice 可用"""
    ok, ver = check_libreoffice()
    if ok:
        return True
    
    # 尝试安装
    return install_libreoffice()


# ============================================================
# 文件转换函数
# ============================================================
def convert_with_libreoffice(src_path, output_dir):
    """
    用 LibreOffice headless 将文件转换为 PDF。
    支持格式: doc, docx, xls, xlsx, ppt, pptx, odt, etc.
    返回生成的 PDF 路径，失败返回 None。
    """
    try:
        result = subprocess.run(
            [
                'libreoffice', '--headless', '--convert-to', 'pdf',
                '--outdir', output_dir,
                src_path
            ],
            capture_output=True,
            text=True,
            timeout=120,
            env={**os.environ, 'HOME': os.environ.get('HOME', '/tmp')}
        )
        
        # LibreOffice 输出的 PDF 文件名与源文件相同（扩展名改为 .pdf）
        base_name = Path(src_path).stem
        pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
        
        if os.path.exists(pdf_path):
            return pdf_path
        
        # 有时 LibreOffice 会用原始文件名（如果含特殊字符会变化）
        # 尝试查找目录下新生成的 PDF
        for f in os.listdir(output_dir):
            if f.endswith('.pdf'):
                fp = os.path.join(output_dir, f)
                # 检查是否是最近创建的（10秒内）
                if os.path.getmtime(fp) > os.path.gettime(src_path):
                    return fp
        
        return None
        
    except subprocess.TimeoutExpired:
        st.error(f'⏰ 转换超时: {Path(src_path).name}')
        return None
    except Exception as e:
        st.error(f'❌ LibreOffice 转换失败 ({Path(src_path).name}): {e}')
        return None


def convert_image_to_pdf(img_path, pdf_path):
    """将图片转换为 PDF（居中显示在 A4 页面上）"""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Image as RLImage

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
    """合并多个 PDF 为一个"""
    writer = PdfWriter()
    for pdf_path in pdf_paths:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            writer.add_page(page)

    with open(output_path, 'wb') as f:
        writer.write(f)


# ============================================================
# Streamlit 应用主界面
# ============================================================

st.set_page_config(
    page_title="📄 文件合并转PDF",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)


def main():
    st.title("📄 文件合并转PDF工具")
    st.markdown("保持原始格式 · 支持 Word/Excel/PDF/图片")

    # 侧边栏
    with st.sidebar:
        st.header("📖 使用说明")
        st.markdown("""
**支持的文件格式：**
- 📝 Word：`.doc` `.docx`
- 📊 Excel：`.xls` `.xlsx`
- 📄 PDF：`.pdf`
- 🖼️ 图片：`.jpg` `.png` `.bmp`

**使用步骤：**
1. 上传需要合并的文件
2. 调整顺序（按上传顺序排列）
3. 点击 **"合并生成PDF"**
4. 下载结果

**✨ 格式保留说明：**
- Word/Excel 通过 **LibreOffice** 转换
- **完整保留**字体、表格样式、图片、页眉页脚等
- 图片自动适配 A4 页面
        """)
        
        st.markdown("---")
        # 显示 LibreOffice 状态
        lo_ok, lo_ver = check_libreoffice()
        if lo_ok:
            st.success(f'✅ LibreOffice 就绪\n`{lo_ver.split(chr(10))[0]}`')
        else:
            st.warning('⚠️ LibreOffice 未安装\n将在首次转换时自动安装')

    # 文件上传
    st.header("1️⃣ 上传文件")
    uploaded_files = st.file_uploader(
        "选择要合并的文件（支持多文件、混合格式）",
        accept_multiple_files=True,
        type=['doc', 'docx', 'xls', 'xlsx', 'pdf', 'jpg', 'jpeg', 'png', 'bmp'],
        help="可同时上传 Word、Excel、PDF、图片"
    )

    if not uploaded_files:
        st.info("👆 请先上传需要合并的文件")
        return

    st.success(f"已上传 **{len(uploaded_files)}** 个文件")

    # 文件列表预览
    st.header("2️⃣ 文件列表")
    cols = st.columns(5)
    cols[0].markdown("**序号**")
    cols[1].markdown("**文件名**")
    cols[2].markdown("**类型**")
    cols[3].markdown("**大小**")
    cols[4].markdown("**状态**")

    type_icons = {
        '.doc': '📝 Word', '.docx': '📝 Word',
        '.xls': '📊 Excel', '.xlsx': '📊 Excel',
        '.pdf': '📄 PDF',
        '.jpg': '🖼️ 图片', '.jpeg': '🖼️ 图片', '.png': '🖼️ 图片', '.bmp': '🖼️ 图片'
    }

    for i, file in enumerate(uploaded_files):
        ext = Path(file.name).suffix.lower()
        c = st.columns(5)
        c[0].write(f"`{i+1}`")
        c[1].write(file.name)
        c[2].write(type_icons.get(ext, '📁 其他'))
        c[3].write(f"{file.size / 1024:.1f} KB")
        c[4].write("✅ 待处理")

    # 合并按钮
    st.header("3️⃣ 生成PDF")

    if st.button("🔗 合并生成PDF", type="primary", use_container_width=True):
        with st.spinner("正在处理..."):
            try:
                # 确保 LibreOffice 可用
                if not ensure_libreoffice():
                    st.error("无法初始化转换引擎，请刷新重试")
                    return

                # 创建临时工作目录
                temp_dir = tempfile.mkdtemp()
                convert_dir = os.path.join(temp_dir, "convert")
                os.makedirs(convert_dir, exist_ok=True)

                pdf_files = []
                progress_bar = st.progress(0)
                status_text = st.empty()
                log_area = st.empty()

                logs = []

                for i, file in enumerate(uploaded_files):
                    status_text.text(f"📂 处理中 ({i+1}/{len(uploaded_files)}): {file.name}")
                    progress_bar.progress((i + 0.5) / len(uploaded_files))

                    # 保存上传文件
                    file_path = os.path.join(temp_dir, file.name)
                    with open(file_path, 'wb') as f:
                        f.write(file.getvalue())

                    ext = Path(file.name).suffix.lower()

                    if ext == '.pdf':
                        # PDF 直接复制
                        pdf_path = os.path.join(convert_dir, f"{i:03d}.pdf")
                        shutil.copy(file_path, pdf_path)
                        pdf_files.append(pdf_path)
                        logs.append(f"✅ {file.name} → 直接使用")

                    elif ext in ('.doc', '.docx', '.xls', '.xlsx'):
                        # 用 LibreOffice 转换（保持格式！）
                        pdf_path = convert_with_libreoffice(file_path, convert_dir)
                        if pdf_path:
                            # 重命名以保序
                            ordered_path = os.path.join(convert_dir, f"{i:03d}.pdf")
                            shutil.move(pdf_path, ordered_path)
                            pdf_files.append(ordered_path)
                            logs.append(f"✅ {file.name} → LibreOffice 转换（保留格式）")
                        else:
                            logs.append(f"❌ {file.name} → 转换失败")

                    elif ext in ('.jpg', '.jpeg', '.png', '.bmp'):
                        # 图片转 PDF
                        pdf_path = os.path.join(convert_dir, f"{i:03d}.pdf")
                        convert_image_to_pdf(file_path, pdf_path)
                        pdf_files.append(pdf_path)
                        logs.append(f"✅ {file.name} → 图片嵌入 A4")

                    else:
                        logs.append(f"⏭️ {file.name} → 不支持的格式，已跳过")

                    progress_bar.progress((i + 1) / len(uploaded_files))

                # 显示处理日志
                log_area.markdown("\n".join(logs))

                if not pdf_files:
                    st.error("没有成功转换的文件")
                    return

                # 合并所有 PDF
                status_text.text("📋 正在合并 PDF...")
                merged_path = os.path.join(temp_dir, "merged.pdf")
                merge_pdfs(pdf_files, merged_path)

                # 读取结果
                with open(merged_path, 'rb') as f:
                    pdf_bytes = f.read()

                # 清理临时文件
                shutil.rmtree(temp_dir, ignore_errors=True)

                progress_bar.progress(100)
                status_text.text("✅ 完成！")

                page_count = len(PdfReader(io.BytesIO(pdf_bytes)).pages)
                st.success(f"🎉 合并完成！{len(pdf_files)} 个文件 → {page_count} 页")

                # 下载按钮
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label="📥 下载合并后的 PDF",
                    data=pdf_bytes,
                    file_name=f"合并文档_{timestamp}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ 处理失败: {e}")
                st.exception(e)


if __name__ == "__main__":
    main()
