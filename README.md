# 文件合并转PDF工具

一个基于 Streamlit 的 Web 应用，可以将多种格式的文件合并成一个 PDF 文档。

## 功能特点

- ✅ 支持多种文件格式：
  - Word 文档（`.docx`）
  - Excel 表格（`.xlsx`）
  - PDF 文档（`.pdf`）
  - 图片文件（`.jpg`, `.png`, `.bmp`）
- ✅ 自动适配 A4 页面尺寸
- ✅ 支持中文内容
- ✅ 在线使用，无需安装

## 部署方式

### Streamlit Cloud 部署（推荐）

1. 将此仓库上传到 GitHub
2. 登录 [Streamlit Cloud](https://streamlit.io/cloud)
3. 点击 "New app"，选择你的 GitHub 仓库
4. 主文件路径设置为 `app.py`
5. 点击 "Deploy"

### 本地运行

```bash
# 安装依赖
pip install -r requirements.txt

# 运行应用
streamlit run app.py
```

## 使用说明

1. 上传需要合并的文件（支持多文件上传）
2. 文件会按上传顺序排列
3. 点击 "合并生成PDF" 按钮
4. 下载生成的 PDF 文件

## 注意事项

- 图片会自动缩放适配 A4 页面
- Excel 表格最多显示 100 行、20 列
- `.doc` 格式建议先转换为 `.docx`

## 技术栈

- Streamlit - Web 界面
- pypdf - PDF 合并
- reportlab - PDF 生成
- python-docx - Word 文档解析
- openpyxl - Excel 文档解析
- Pillow - 图片处理

## License

MIT License