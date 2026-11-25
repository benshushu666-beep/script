# -*- coding: utf-8 -*-
"""生成虚构测试样本（不含任何真实个人信息）。"""
import shutil
from pathlib import Path

import pymupdf
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Pt
from openpyxl import Workbook

BASE = Path(__file__).parent / "data"

ID_W = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
ID_C = "10X98765432"
USCC_A = "0123456789ABCDEFGHJKLMNPQRTUWXY"
USCC_W = [1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28]


def id_check(b17):
    return ID_C[sum(int(c) * w for c, w in zip(b17, ID_W)) % 11]


def uscc_check(b17):
    t = sum(USCC_A.index(c) * w for c, w in zip(b17, USCC_W))
    c = 31 - t % 31
    if c == 31:
        c = 0
    return USCC_A[c]


ID = "33010219900307123" + id_check("33010219900307123")
ID2 = "33010219880615234" + id_check("33010219880615234")
USCC_BODY = "91330106MA2ABCDEFX"[:17]
USCC = USCC_BODY + uscc_check(USCC_BODY)

print("测试身份证:", ID, "校验:", ID[-1])
print("测试统一社会信用代码:", USCC, "长度:", len(USCC))

if BASE.exists():
    shutil.rmtree(BASE)
BASE.mkdir(parents=True)

# ---------------- docx ----------------
d1 = BASE / "01_诉讼文书"
d1.mkdir()
doc = Document()
sec = doc.sections[0]
sec.header.paragraphs[0].text = "浙江某某律师事务所  承办人：王强  电话13812345678"
sec.footer.paragraphs[0].text = "地址：杭州市西湖区文三路100号3幢201室"

doc.add_heading("民事起诉状", level=1)
doc.add_paragraph(
    "原告：张伟，男，1985年3月7日出生，汉族，住杭州市西湖区文三路100号3幢201室，"
    "身份证号码" + ID + "，联系电话13812345678。"
)
doc.add_paragraph(
    "被告：杭州测试科技有限公司，住所地：杭州市西湖区文三路100号3幢201室，"
    "统一社会信用代码" + USCC + "，法定代表人：李娜。"
)
doc.add_paragraph(
    "委托代理人：王强，浙江测试律师事务所律师，手机13812345678，"
    "邮箱zhangwei@example.com，固定电话0571-88886666。"
)
doc.add_paragraph(
    "其他联系方式：QQ：123456789，微信号：zhangwei_1985，护照号：E12345678。"
)

table = doc.add_table(rows=3, cols=3)
data = [
    ("当事人", "姓名", "身份证号"),
    ("原告", "张伟", ID),
    ("被告", "李娜", ID2),
]
for i, row in enumerate(data):
    for j, val in enumerate(row):
        table.cell(i, j).text = val

doc.save(d1 / "起诉状.docx")

# ---------------- xlsx ----------------
d2 = BASE / "02_证据"
d2.mkdir()
wb = Workbook()
ws = wb.active
ws.title = "当事人信息"
rows = [
    ["姓名", "身份证号", "手机号", "住址", "工作单位"],
    ["张伟", ID, "13812345678", "杭州市西湖区文三路100号3幢201室", "杭州测试科技有限公司"],
    ["李娜", ID2, "13900001111", "杭州市拱墅区莫干山路200号", "杭州示例贸易有限公司"],
]
for r in rows:
    ws.append(r)

ws2 = wb.create_sheet("联系方式")
ws2["A1"] = "邮箱"
ws2["B1"] = "zhangwei@example.com"
ws2["A2"] = "QQ"
ws2["B2"] = "123456789"
ws2["A3"] = "微信号"
ws2["B3"] = "zhangwei_1985"
ws2["A4"] = "护照号"
ws2["B4"] = "E12345678"
ws2["A5"] = "出生日期"
ws2["B5"] = "1985年3月7日"
wb.save(d2 / "证据清单.xlsx")

# ---------------- 图片 ----------------
d3 = BASE / "03_扫描件"
d3.mkdir()
font_path = "C:/Windows/Fonts/msyh.ttc"
font = ImageFont.truetype(font_path, 28)
img = Image.new("RGB", (1400, 700), "white")
dr = ImageDraw.Draw(img)
lines = [
    "身份证复印件",
    "姓名：张伟    性别：男    出生：1985年3月7日",
    "住址：杭州市西湖区文三路100号3幢201室",
    "公民身份号码：" + ID,
    "联系电话：13812345678",
    "工作单位：杭州测试科技有限公司",
    "邮箱：zhangwei@example.com    QQ：123456789",
]
y = 40
for ln in lines:
    dr.text((40, y), ln, fill="black", font=font)
    y += 70
img.save(d3 / "身份证扫描件.png")

# 扫描版 PDF（纯图片）
img.save(d3 / "身份证扫描件.pdf", "PDF", resolution=150)

# ---------------- 文字版 PDF ----------------
d4 = BASE / "04_PDF"
d4.mkdir()
pdf = pymupdf.open()
page = pdf.new_page(width=595, height=842)
fontfile = "C:/Windows/Fonts/msyh.ttc"
page.insert_font(fontname="msyh", fontfile=fontfile)
text = (
    "借款合同\n\n"
    "甲方（借款人）：张伟\n"
    "身份证号：" + ID + "\n"
    "住址：杭州市西湖区文三路100号3幢201室\n"
    "电话：13812345678\n\n"
    "乙方（出借人）：杭州测试科技有限公司\n"
    "统一社会信用代码：" + USCC + "\n"
    "法定代表人：李娜\n"
    "联系邮箱：zhangwei@example.com\n"
    "固定电话：0571-88886666\n"
)
page.insert_textbox(pymupdf.Rect(50, 50, 545, 780), text,
                    fontname="msyh", fontsize=12, lineheight=1.6)
pdf.save(d4 / "借款合同.pdf")
pdf.close()

# ---------------- 旧格式（Word/Excel COM 转换） ----------------
try:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(str(d1 / "起诉状.docx"))
    d.SaveAs2(str(d1 / "起诉状_旧格式.doc"), FileFormat=0)  # wdFormatDocument
    d.Close(False)
    word.Quit()

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb2 = excel.Workbooks.Open(str(d2 / "证据清单.xlsx"))
    wb2.SaveAs(str(d2 / "证据清单_旧格式.xls"), FileFormat=56)  # xlExcel8
    wb2.Close(False)
    excel.Quit()
    pythoncom.CoUninitialize()
    print("旧格式样本已生成")
except Exception as e:
    print("旧格式样本生成失败（可忽略）:", e)

print("测试样本目录:", BASE)
for p in sorted(BASE.rglob("*")):
    if p.is_file():
        print("  ", p.relative_to(BASE), p.stat().st_size)
