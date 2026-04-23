# 📄 发票 OCR 识别工具

**版本：** v1.0.0  
**作者：** OpenClaw  
**日期：** 2026-04-23

自动识别 PDF 凭证中的发票信息，生成 Excel 汇总表。

---

## ✨ 功能特点

- ✅ 自动识别发票号码、日期、金额
- ✅ 批量处理 PDF 文件
- ✅ 生成 Excel 汇总表
- ✅ 支持命令行和图形界面
- ✅ 100% 本地处理，保护隐私

---

## 📦 安装说明

### 方法 1：使用安装包（推荐）

1. 下载 `invoice-ocr-windows.zip`
2. 解压到任意目录
3. 安装 Tesseract OCR（含中文包）
4. 运行程序

### 方法 2：从源码运行

```bash
# 安装依赖
pip install -r requirements.txt

# 安装 Tesseract OCR
# Windows: https://github.com/UB-Mannheim/tesseract/releases
# Linux: sudo apt-get install tesseract-ocr tesseract-ocr-chi-sim

# 运行命令行版
python invoice_ocr.py 凭证.pdf

# 运行 GUI 版
python invoice_ocr_gui.py
```

---

## 💻 使用方法

### 命令行版

```bash
# 单个文件
invoice-ocr.exe 凭证.pdf

# 批量处理
invoice-ocr.exe *.pdf

# 指定输出目录
invoice-ocr.exe 凭证.pdf D:\output
```

### GUI 版

```bash
# 双击运行
invoice-ocr-gui.exe
```

界面操作：
1. 点击"➕ 添加 PDF"选择文件
2. 选择输出目录
3. 点击"🚀 开始识别"
4. 查看结果

---

## 📊 输出示例

```
识别结果:
  总发票数：72 张
  发票号码：71 张 (98.6%)
  开票日期：29 张 (40.3%)
  金   额：47 张 (65.3%)
  总金额：¥623,363.11
```

输出文件：
- `XXX_invoices.json` - 原始数据
- `XXX_发票信息汇总表.xlsx` - Excel 表格

---

## 🔧 依赖

- Python 3.8+
- Tesseract OCR 5.x（含中文语言包）
- PyMuPDF
- Pillow
- openpyxl

---

## 📝 许可证

MIT License

---

## 📞 技术支持

如有问题，请联系系统管理员。
