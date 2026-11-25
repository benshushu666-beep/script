#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
材料脱敏工具（redact.py）

功能：
  - 识别并脱敏：身份证号、手机号、固定电话、出生日期、邮箱、QQ、微信、
    护照号、银行卡号（可选）、统一社会信用代码、姓名、企业名称、地址。
  - 企业名称、银行/法院/公安等机构名称默认保留；
    需分别加 --redact-companies / --redact-institutions 才脱敏。
  - 支持格式：.docx / .xlsx / .doc / .xls / .pdf / 图片（png jpg bmp tif webp）
  - 保留原始目录结构，绝不修改源文件。
  - 生成脱敏报告与（可选的）映射表。

重要说明：
  1. 姓名、地址、企业名称的识别为启发式规则，不能保证 100% 覆盖，
     输出后必须人工复核。
  2. 扫描件/图片依赖 OCR（RapidOCR），OCR 本身存在漏识和错识。
  3. 公章、印章、手写体不做强制识别，报告中列为需人工处理项。
  4. PDF 采用“真删除”式修订（redaction），不是覆盖黑块；姓名跨行时按行分别遮盖。
"""

from __future__ import annotations

import argparse
import io
import json
import os
import re
import shutil
import sys
import threading
import time
import traceback
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

try:
    from lxml import etree
except ImportError:  # pragma: no cover
    print("缺少依赖 lxml（python-docx 会自带），请先安装：pip install lxml", file=sys.stderr)
    raise

# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_P = "{%s}p" % W_NS
W_T = "{%s}t" % W_NS
W_COMMENT = "{%s}comment" % W_NS
W_TBL = "{%s}tbl" % W_NS
W_TR = "{%s}tr" % W_NS
W_TC = "{%s}tc" % W_NS
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# 表格中“孤立姓名”识别：仅当单元格本身是姓名，且可被表头或角色列佐证时启用
ROLE_WORDS = {
    "原告", "被告", "第三人", "申请人", "被申请人", "上诉人", "被上诉人",
    "法定代表人", "负责人", "联系人", "委托代理人", "代理人", "当事人",
    "证人", "借款人", "出借人", "甲方", "乙方", "丙方", "承租人", "出租人",
    "买受人", "出卖人", "收款人",
}
NAME_HEADERS = {
    "姓名", "名称", "当事人", "法定代表人", "负责人", "联系人", "代理人",
    "原告", "被告", "申请人", "借款人", "出借人", "承租人", "出租人",
}
ISOLATED_STOPWORDS = {
    "公司", "单位", "工厂", "学校", "医院", "法院", "公安", "银行", "支行", "分行", "集团",
    "金额", "单价", "合计", "小计", "总计", "数量", "名称", "姓名", "性别", "民族", "出生",
    "日期", "时间", "地址", "电话", "手机", "邮箱", "账号", "编号", "序号", "备注", "用途",
    "说明", "证据", "材料", "附件", "原告", "被告", "第三人", "申请人", "被申请人", "代理人",
    "当事人", "法人", "负责人", "联系人", "开户", "户名", "杭州", "浙江", "上海", "北京",
    "南京", "广州", "深圳", "中国", "美国", "日本", "韩国", "男", "女", "汉族", "已婚",
    "未婚", "高中", "大学", "本科", "硕士", "博士", "经理", "董事", "监事", "股东", "员工",
    "客户", "高级", "中级", "初级",
}

# 姓名误报黑名单：这些词以姓氏字开头，但显然不是人名
NAME_STOPWORDS = {
    "申请", "支付", "要求", "认为", "提出", "主张", "陈述", "辩称", "答辩",
    "起诉", "上诉", "履行", "承担", "偿还", "归还", "返还", "借款", "贷款",
    "出借", "转账", "交付", "签订", "约定", "同意", "拒绝", "承认", "否认",
    "表示", "说明", "证明", "提供", "提交", "出具", "收到", "催要", "提醒",
    "答复", "承诺", "违反", "损害", "维护", "提起", "诉讼", "法院", "判决",
    "裁定", "调解", "执行", "保全", "冻结", "查封", "拍卖", "抵押", "担保",
    "保证", "利息", "本金", "违约", "损失", "费用", "律师", "代理", "委托",
    "授权", "证据", "材料", "事实", "理由", "请求", "判令", "依法", "根据",
    "依据", "由于", "因为", "所以", "但是", "并且", "或者", "如果", "虽然",
    "然而", "此外", "同时", "已经", "正在", "将要", "可以", "应当", "必须",
    "需要", "希望", "愿意", "决定", "选择", "使用", "利用", "进行", "实施",
    "完成", "开始", "结束", "继续", "停止", "上述", "以下", "其中", "期间",
    "之后", "之前", "当时", "现在", "目前", "一直", "多次", "共计", "合计",
    "以及", "通过", "关于", "对于", "为了", "除了", "包括", "例如", "比如",
    "即使", "无论", "只要", "只有", "而是", "就是", "都是", "仍是", "还是",
    "全国", "人民", "商业", "银行", "公司", "单位", "个人", "双方", "各方",
    "对方", "本人", "当事人", "原告", "被告", "第三人", "申请人", "被申请人",
}

CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"

CATEGORIES = [
    "id_card", "uscc", "bankcard", "phone", "landline", "birth",
    "email", "qq", "wechat", "passport", "person", "company", "address",
]
DEFAULT_CATEGORIES = [c for c in CATEGORIES if c != "bankcard"]
NUMERIC_CATEGORIES = {"id_card", "uscc", "bankcard", "phone", "landline"}

# 公共机构、金融机构名称默认不脱敏（与“企业名称”区分）
INSTITUTION_KEYWORDS = (
    "银行", "信用社", "法院", "人民检察院", "检察院", "公安", "派出所",
    "人民政府", "管理局", "税务局", "民政局", "司法局", "财政局", "教育局",
    "市场监督", "人力资源和社会保障", "卫生健康", "医疗保障", "自然资源",
    "生态环境", "住房和城乡建设", "交通运输", "水利局", "农业农村",
    "商务局", "文化和旅游", "应急管理", "退役军人", "审计局", "统计局",
    "气象局", "不动产登记", "仲裁委员会", "公证处", "街道办事处",
    "居民委员会", "村民委员会", "看守所", "监狱", "海关", "消防救援",
)


def is_institution(name: str) -> bool:
    return any(k in name for k in INSTITUTION_KEYWORDS)

CAT_LABEL = {
    "id_card": "身份证号", "uscc": "统一社会信用代码", "bankcard": "银行卡号",
    "phone": "手机号", "landline": "固定电话", "birth": "出生日期",
    "email": "邮箱", "qq": "QQ号", "wechat": "微信号",
    "passport": "护照号", "person": "姓名", "company": "企业名称", "address": "地址",
}

PRIORITY = {
    "id_card": 100, "uscc": 96, "bankcard": 92, "passport": 86,
    "phone": 82, "landline": 78, "email": 74, "qq": 70, "wechat": 69,
    "birth": 64, "company": 60, "address": 55, "person": 44,
}

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
TEXT_EXTS = {".docx", ".xlsx", ".pdf", ".doc", ".xls"}

MASK_MARKERS = ("*", "×", "某")

# 百家姓（常用姓氏，用于姓名高置信校验）
SURNAMES = set(
    "赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜戚谢邹喻柏水窦章"
    "云苏潘葛奚范彭郎鲁韦昌马苗凤花方俞任袁柳酆鲍史唐费廉岑薛雷贺倪汤滕殷罗毕郝邬安常"
    "乐于时傅皮卞齐康伍余元卜顾孟平黄和穆萧尹姚邵湛汪祁毛禹狄米贝明臧计伏成戴谈宋茅庞"
    "熊纪舒屈项祝董梁杜阮蓝闵席季麻强贾路娄危江童颜郭梅盛林刁钟徐邱骆高夏蔡田樊胡凌霍"
    "虞万支柯昝管卢莫经房裘缪干解应宗丁宣贲邓郁单杭洪包诸左石崔吉钮龚程嵇邢滑裴陆荣翁"
    "荀羊於惠甄曲家封芮羿储靳汲邴糜松井段富巫乌焦巴弓牧隗山谷车侯宓蓬全郗班仰秋仲伊宫"
    "宁仇栾暴甘钭厉戎祖武符刘景詹束龙叶幸司韶郜黎蓟薄印宿白怀蒲邰从鄂索咸籍赖卓蔺屠蒙"
    "池乔阴胥能苍双闻莘党翟谭贡劳逄姬申扶堵冉宰郦雍却璩桑桂濮牛寿通边扈燕冀郏浦尚农温"
    "别庄晏柴瞿阎充慕连茹习宦艾鱼容向古易慎戈廖庾终暨居衡步都耿满弘匡国文寇广禄阙东欧"
    "殳沃利蔚越夔隆师巩厍聂晁勾敖融冷訾辛阚那简饶空曾毋沙乜养鞠须丰巢关蒯相查后荆红游"
    "竺权逯盖益桓公"
)
COMPOUND_SURNAMES = [
    "欧阳", "司马", "上官", "诸葛", "东方", "皇甫", "尉迟", "公孙",
    "慕容", "长孙", "宇文", "司徒", "司空", "夏侯", "令狐", "端木", "南宫",
]

# 代号池（刻意不使用真实姓氏，避免泄露）
SURNAME_POOL = list(
    "赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜"
    "戚谢邹喻柏窦章云苏潘葛范彭郎鲁韦马苗方俞任袁柳鲍史唐薛雷贺倪汤"
    "罗毕郝邬安常傅卞齐康伍余顾孟黄穆萧尹姚邵汪祁毛禹狄贝明臧成戴谈"
    "宋庞熊纪舒屈项祝董梁杜阮蓝闵席季强贾路娄江童颜郭梅盛林钟徐邱骆"
    "高夏蔡田樊胡凌霍万支柯管卢莫房裘解应宗丁宣邓郁单杭洪包诸左石崔"
    "吉龚程邢裴陆荣翁荀羊惠甄曲封芮储靳段富巫乌焦巴牧山谷车侯全班仰"
    "秋仲伊宫宁仇甘厉祖武符刘景詹龙叶幸司黎薄印宿白怀蒲鄂索赖卓屠池"
)
company_pool = list("甲乙丙丁戊己庚辛壬癸")


# --------------------------------------------------------------------------
# 工具函数
# --------------------------------------------------------------------------

def now_stamp() -> str:
    return datetime.now().strftime("%Y-%m-%d_%H%M")


def contains_marker(s: str) -> bool:
    return any(m in s for m in MASK_MARKERS)


def is_cjk(ch: str) -> bool:
    return "\u4e00" <= ch <= "\u9fa5"


# --------------------------------------------------------------------------
# 号码校验
# --------------------------------------------------------------------------

ID_WEIGHTS = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
ID_CHECK = "10X98765432"


def valid_id(s: str) -> bool:
    if len(s) != 18:
        return False
    body = s[:17]
    if not body.isdigit():
        return False
    total = sum(int(c) * w for c, w in zip(body, ID_WEIGHTS))
    return ID_CHECK[total % 11] == s[17].upper()


USCC_ALPHABET = "0123456789ABCDEFGHJKLMNPQRTUWXY"
USCC_WEIGHTS = [1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28]


def valid_uscc(s: str) -> bool:
    if len(s) != 18:
        return False
    try:
        vals = [USCC_ALPHABET.index(c) for c in s]
    except ValueError:
        return False
    total = sum(v * w for v, w in zip(vals[:17], USCC_WEIGHTS))
    check = 31 - total % 31
    if check == 31:
        check = 0
    return USCC_ALPHABET[check] == s[17]


def valid_luhn(s: str) -> bool:
    digits = [int(c) for c in s if c.isdigit()]
    if not digits:
        return False
    total = 0
    for i, d in enumerate(reversed(digits)):
        if i % 2 == 1:
            d *= 2
            if d > 9:
                d -= 9
        total += d
    return total % 10 == 0


# --------------------------------------------------------------------------
# 正则规则
# --------------------------------------------------------------------------

RE_ID = re.compile(
    r"(?<![0-9A-Za-z])[1-9]\d{5}(?:18|19|20)\d{2}"
    r"(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx](?![0-9A-Za-z])"
)
RE_USCC = re.compile(r"(?<![0-9A-Za-z])[0-9A-HJ-NPQRTUWXY]{18}(?![0-9A-Za-z])")
RE_PHONE = re.compile(r"(?<![0-9])1[3-9]\d{9}(?![0-9])")
RE_LANDLINE_AREA = re.compile(
    r"(?<![0-9])(0\d{2,3})[-—–]?(\d{7,8})(?:[-—–](\d{1,6}))?(?![0-9])"
)
RE_LANDLINE_CTX = re.compile(
    r"(?:电话|传真|座机|Tel|TEL|Fax)\s*(?:号码)?\s*[：:]?\s*"
    r"(\d{7,8}(?:[-—–]\d{1,6})?)(?![0-9])"
)
RE_BIRTH = re.compile(
    r"(?:出生(?:日期|年月|时间)?|生日|生于|DOB)\s*(?:为|是|：|:)?\s*"
    r"((?:19|20)\d{2}\s*[年\-/.]\s*(?:0?[1-9]|1[0-2])"
    r"(?:\s*[月\-/.]\s*(?:0?[1-9]|[12]\d|3[01])\s*日?)?)"
)
# “1985年3月7日出生”这类日期在前、提示词在后的语序
RE_BIRTH_POST = re.compile(
    r"((?:19|20)\d{2}\s*[年\-/.]\s*(?:0?[1-9]|1[0-2])\s*[月\-/.]\s*"
    r"(?:0?[1-9]|[12]\d|3[01])\s*日?)\s*(?:出生|生)"
)
RE_EMAIL = re.compile(
    r"(?<![A-Za-z0-9._%+\-])[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}(?![A-Za-z0-9.\-])"
)
RE_QQ = re.compile(r"(?:QQ|Qq|qq)\s*(?:号码|号)?\s*[：:]?\s*(\d{5,12})(?![0-9])")
RE_WECHAT = re.compile(
    r"(?:微信号|微信|WeChat|wechat|Weixin|weixin|VX|vx)\s*(?:号码|号)?\s*[：:]?\s*"
    r"([A-Za-z][\-_A-Za-z0-9]{5,19})(?![A-Za-z0-9])"
)
RE_PASSPORT = re.compile(
    r"(?:护照(?:号码|号)?|Passport(?:\s*No\.?)?)\s*[：:]?\s*"
    r"([A-Za-z]{1,2}\d{7,8})(?![A-Za-z0-9])", re.I
)
RE_BANKCARD = re.compile(
    r"(?:银行卡号|卡号|账号|账户)\s*[：:]?\s*(\d{16,19})(?![0-9])"
)

NAME_CONTEXT = (
    "原告|被告|第三人|上诉人|被上诉人|申请人|被申请人|再审申请人|被执行人|"
    "申请执行人|法定代表人|法定代理人|负责人|委托代理人|诉讼代理人|代理人|"
    "当事人|证人|联系人|经办人|收款人|借款人|出借人|承租人|出租人|买受人|出卖人|"
    "甲方|乙方|丙方|姓名|曾用名|户名|权利人|义务人|承办人|承办律师|主审法官|"
    "审判员|书记员|记录人|送达人|签收人|代收人|委托人|受托人|"
    "授权委托人|被委托人|委托方|受托方"
)
_COMPOUND_ALT = "|".join(COMPOUND_SURNAMES)
RE_NAME = re.compile(
    "(?:" + NAME_CONTEXT + ")\s*(?:（[^）]{0,20}）)?\s*"
    "(?:为|是|系|与|和|：|:)?\s*"
    "((?:" + _COMPOUND_ALT + r"|[一-龥])"
    r"(?:\s*[一-龥]){1,2}?)"
    "(?=(?:\s*(?:[，。、；：？！（）()《》【】,;:!?]|$|\d|诉|向|与|和|因|称|认为|"
    "辩称|请求|申请|提交|出具|签订|于|在|未|已|的|等|及|被|名下|处|中|先生|"
    "女士|同志|律师|法官|性别|年龄|民族|籍贯|住址|出生|身份证|联系电话|"
    "手机号|电话|邮箱|民|借|合|纠|案|侵|离|继|劳|买|卖|租|房|车|款|货|股|"
    "债|赔|还|欠|婚|抚养|赡养|结|经|到|拒|无|不|同|负|担|享|履|承|应|签|男|女|围))|[\r\n](?=(?:联系|住址|地址|电话|手机|身份|出生|民族|性别|工作|职务|单位|户籍|文化|政治|案由|诉讼|事实|本院|审判|书记|此致|附件|证据|日期|原告|被告|第三人|委托|法定|负责|代理|邮编|传真|邮箱|微信|QQ|户名|开户|账号|籍贯|年龄|职业)))"
)
RE_NAME_TITLE = re.compile(
    r"([一-龥]{2,4}?)\s*(?:先生|女士|同志|律师)(?!事务所)")
# 判决书落款常把“审判员/书记员”逐字换行，姓名在下一行
RE_SIGN_ROLE = re.compile(
    r"(?:审\s*判\s*员|审\s*判\s*长|书\s*记\s*员|人\s*民\s*陪\s*审\s*员|"
    r"法\s*官\s*助\s*理)\s*[:：]?\s*([一-龥]{2,4}?)"
    r"(?=\s|$|[〇一二三四五六七八九十]|[，。、；：])")
# 仅 --aggressive-names 启用：识别“张伟与李娜”“张伟向李娜”这类无角色词的并列姓名
RE_NAME_PAIR = re.compile(
    r"([一-龥]{2,3}?)(?:与|和|向|诉)([一-龥]{2,3}?)"
    r"(?=[，。、；：？！\s（）()《》【】,;:!?]|$|于|在|签|借|约|因|系|为|称|未|已|被|"
    r"将|就|均|都|曾|是|的|等|及|出|收|付|还|欠|转|汇|结|离|抚|赡|劳|买|卖|租|房|车|款|货|股|债|赔|婚)"
)

ADDR_CONTEXT = (
    "住址|地址|住所地|住所|现住址|现住|户籍地|户籍地址|联系地址|送达地址|"
    "通讯地址|注册地址|经营场所|营业场所|所在地|住"
)
RE_ADDR_CTX = re.compile(
    "(?:" + ADDR_CONTEXT + r")\s*(?:为|是|：|:)?\s*([^\n，。；;！？]{4,60})"
)
RE_ADDR_STRUCT = re.compile(
    r"[\u4e00-\u9fa5]{2,8}(?:省|自治区|特别行政区)"
    r"[\u4e00-\u9fa5]{2,10}(?:市|自治州|地区|盟)"
    r"[\u4e00-\u9fa5]{2,10}(?:区|县|旗)"
    r"[\u4e00-\u9fa5A-Za-z0-9\-]{0,30}?"
    r"(?:路|街|道|巷|弄|胡同|村|镇|乡)"
    r"[^\s，。；;：:]{0,40}"
)
RE_ADDR_STRUCT2 = re.compile(
    r"[\u4e00-\u9fa5]{2,8}市[\u4e00-\u9fa5]{2,10}(?:区|县)"
    r"[^\s，。；;：:]{2,40}?(?:号|室|栋|幢|单元|楼)"
    r"[^\s，。；;：:]{0,20}"
)
STRUCT_LEAD = re.compile(r"^(?:住址|地址|住所|现住|户籍地|位于|坐落于|住|在|于|址|地)[：:\s]*")

ORG_SUFFIXES = [
    "股份有限公司", "有限责任公司", "集团有限公司", "有限公司", "合伙企业",
    "律师事务所", "会计师事务所", "农民专业合作社", "信用合作社",
    "支行", "分行", "银行", "集团", "公司", "厂", "店", "商行", "经营部",
]
_ORG_SUFFIX_SORTED = sorted(ORG_SUFFIXES, key=len, reverse=True)
# 向左扩展时遇到这些字符立即停止（含 ASCII 标点与数字），避免跨句吞并
ORG_BOUNDARY_CHARS = set("，。、；：:;!?！？（）()《》【】[]{}<>\"'“”‘’,.")
ORG_BAD_CONTENT = set("元年的了被将让向由与为是在按据至起共计及等0123456789")
# “字号”位置不允许出现这些泛称
ORG_BAD_CORE = {
    "全国", "人民", "商业", "政策", "中央", "地方", "该", "本", "各", "上述",
    "相关", "有关", "开户", "贷款", "放贷", "存款", "收款", "付款", "转出",
    "转入", "外资", "民营", "国有", "股份制", "同业",
    "多家", "两家", "一家", "这家", "那家", "贷款", "放贷", "存款", "借款",
    "还款", "放款",
}
# “银行间”“银行业”“银行贷款”等不是机构名
ORG_BANK_NEXT = set("间业卡账流转存利系机网柜")
ORG_BANK_NEXT2 = {"流水", "转账", "利率", "系统", "机构", "网点", "柜台",
                  "账户", "业务"}
ORG_LEAD_STRIP = re.compile(
    r"^(?:原告|被告|第三人|上诉人|被上诉人|申请人|被申请人|再审申请人|被执行人|"
    r"申请执行人|法定代表人|法定代理人|负责人|委托代理人|诉讼代理人|代理人|"
    r"当事人|甲方|乙方|丙方|其中|以及|并且|但是|因为|所以|由于|根据|依据|"
    r"按照|通过|关于|对于|为了|包括|例如|比如|同时|此外|目前|现在|已经|"
    r"正在|将要|可以|应当|必须|需要|双方|各方|对方|本人|"
    r"系|为|是|由|向|在|经|及|、|，)+"
)


def org_code_type(name: str) -> str:
    """代号保留机构类型，避免把银行写成“公司”。"""
    if name.endswith(("银行", "支行", "分行")):
        return "银行"
    if name.endswith("律师事务所"):
        return "律师事务所"
    if name.endswith("会计师事务所"):
        return "会计师事务所"
    if name.endswith(("农民专业合作社", "信用合作社")):
        return "合作社"
    if name.endswith("集团"):
        return "集团"
    if name.endswith("合伙企业"):
        return "合伙企业"
    return "公司"


# --------------------------------------------------------------------------
# 检测器
# --------------------------------------------------------------------------

@dataclass
class Hit:
    start: int
    end: int
    category: str
    text: str
    confidence: float = 1.0


class Detector:
    def __init__(self, categories=None, aggressive_names=False,
                 redact_companies=False, redact_institutions=False):
        self.categories = set(categories or DEFAULT_CATEGORIES)
        self.aggressive_names = aggressive_names
        # 企业名称与机构名称默认都保留，需显式开启才脱敏
        self.redact_companies = redact_companies
        self.redact_institutions = redact_institutions

    # -- 各规则 ---------------------------------------------------------

    def _id(self, text):
        out = []
        for m in RE_ID.finditer(text):
            s = m.group(0)
            if valid_id(s):
                out.append(Hit(m.start(), m.end(), "id_card", s))
        return out

    def _uscc(self, text):
        if "uscc" not in self.categories:
            return []
        out = []
        for m in RE_USCC.finditer(text):
            s = m.group(0)
            if valid_uscc(s):
                out.append(Hit(m.start(), m.end(), "uscc", s))
        return out

    def _phone(self, text):
        out = []
        for m in RE_PHONE.finditer(text):
            out.append(Hit(m.start(), m.end(), "phone", m.group(0)))
        return out

    def _landline(self, text):
        out = []
        for m in RE_LANDLINE_AREA.finditer(text):
            out.append(Hit(m.start(), m.end(), "landline", m.group(0)))
        for m in RE_LANDLINE_CTX.finditer(text):
            out.append(Hit(m.start(1), m.end(1), "landline", m.group(1)))
        return out

    def _birth(self, text):
        out = []
        for m in RE_BIRTH.finditer(text):
            out.append(Hit(m.start(1), m.end(1), "birth", m.group(1)))
        for m in RE_BIRTH_POST.finditer(text):
            out.append(Hit(m.start(1), m.end(1), "birth", m.group(1)))
        return out

    def _email(self, text):
        return [Hit(m.start(), m.end(), "email", m.group(0))
                for m in RE_EMAIL.finditer(text)]

    def _qq(self, text):
        return [Hit(m.start(1), m.end(1), "qq", m.group(1))
                for m in RE_QQ.finditer(text)]

    def _wechat(self, text):
        return [Hit(m.start(1), m.end(1), "wechat", m.group(1))
                for m in RE_WECHAT.finditer(text)]

    def _passport(self, text):
        return [Hit(m.start(1), m.end(1), "passport", m.group(1))
                for m in RE_PASSPORT.finditer(text)]

    def _bankcard(self, text):
        if "bankcard" not in self.categories:
            return []
        out = []
        for m in RE_BANKCARD.finditer(text):
            s = m.group(1)
            if valid_luhn(s):
                out.append(Hit(m.start(1), m.end(1), "bankcard", s))
        return out

    def _person(self, text):
        out = []
        for m in RE_NAME.finditer(text):
            cand = m.group(1)
            if not self._valid_person(cand):
                continue
            out.append(Hit(m.start(1), m.end(1), "person", cand, 0.9))
        for m in RE_NAME_TITLE.finditer(text):
            cand = m.group(1)
            if not self._valid_person(cand):
                continue
            out.append(Hit(m.start(1), m.end(1), "person", cand, 0.8))
        for m in RE_SIGN_ROLE.finditer(text):
            cand = m.group(1)
            if not self._valid_person(cand):
                continue
            out.append(Hit(m.start(1), m.end(1), "person", cand, 0.9))
        if self.aggressive_names:
            for m in RE_NAME_PAIR.finditer(text):
                for g in (1, 2):
                    cand = m.group(g)
                    if self.isolated_person(cand) and not contains_marker(cand):
                        out.append(Hit(m.start(g), m.end(g), "person", cand, 0.5))
        return out

    @staticmethod
    def isolated_person(s: str):
        """整段/整格仅为姓名时的低置信识别，需外部提供表头或角色列佐证。"""
        s = re.sub(r"\s+", "", s or "")
        if not (2 <= len(s) <= 4):
            return None
        if not all(is_cjk(c) for c in s):
            return None
        if (contains_marker(s) or s in ISOLATED_STOPWORDS
                or s in NAME_STOPWORDS):
            return None
        if s[:2] in COMPOUND_SURNAMES:
            return Hit(0, len(s), "person", s, 0.6)
        if s[0] in SURNAMES:
            return Hit(0, len(s), "person", s, 0.6)
        return None

    def _valid_person(self, s: str) -> bool:
        s = re.sub(r"\s+", "", s)
        if not (2 <= len(s) <= 4):
            return False
        if not all(is_cjk(c) for c in s):
            return False
        if contains_marker(s):
            return False
        # “申请贷款”“被告支付”“从原告”这类词以姓氏字开头，但不是人名
        if s in NAME_STOPWORDS or s[:2] in NAME_STOPWORDS:
            return False
        if s in ROLE_WORDS or s[1:] in ROLE_WORDS or s[:2] in ROLE_WORDS:
            return False
        # “师事务所”是“律师事务所”的截断，不是人名
        if any(k in s for k in ("事务所", "公司", "银行", "法院", "公安")):
            return False
        if s[:2] in COMPOUND_SURNAMES:
            return True
        return s[0] in SURNAMES

    def _company(self, text):
        out = []
        n = len(text)
        i = 0
        while i < n:
            matched = None
            for suf in _ORG_SUFFIX_SORTED:
                if not text.startswith(suf, i):
                    continue
                # “银行间”“银行业”“银行贷款”等不是机构名
                if suf in ("银行", "支行", "分行"):
                    tail = text[i + len(suf):i + len(suf) + 2]
                    if (tail[:1] in ORG_BANK_NEXT) or (tail in ORG_BANK_NEXT2):
                        continue
                matched = suf
                break
            if not matched:
                i += 1
                continue
            j = i + len(matched)
            k = i
            while k > 0 and i - k < 24:
                ch = text[k - 1]
                if ch in ORG_BOUNDARY_CHARS or ch.isdigit():
                    break
                k -= 1
            for cand in self._clean_org(text[k:j], len(matched)):
                if contains_marker(cand) or not self._valid_org(cand, len(matched)):
                    continue
                start = text.find(cand, k, j)
                if start >= 0:
                    out.append(Hit(start, start + len(cand), "company", cand, 0.85))
                    break
            i = j
        return out

    @staticmethod
    def _clean_org(cand: str, suffix_len: int) -> list:
        raw = cand
        for sep in ("号", "室"):
            idx = raw.rfind(sep)
            if idx > 0 and len(raw) - idx - len(sep) >= suffix_len + 2:
                raw = raw[idx + len(sep):]
                break
        # 相邻企业名（如“甲公司与乙公司”）只取当前这一段
        last_sep = None
        for m in re.finditer(r"(?:有限公司|有限责任公司|股份有限公司|公司|银行|"
                             r"支行|分行|事务所|集团)(?:与|和|及|、)", raw):
            last_sep = m.end()
        if last_sep:
            raw = raw[last_sep:]
        stripped = ORG_LEAD_STRIP.sub("", raw)
        # “与南京银行”这类连接词单独处理，避免把“和平有限公司”的“和”一起剥掉
        alt = re.sub(r"^[与和]", "", stripped)
        out = []
        # 依次尝试：去上下文前缀 → 再去连接词 → 原文
        for c in (stripped, alt, raw):
            if (len(c) >= suffix_len + 2 and c[0] not in "本该某各贵我上下"
                    and any(is_cjk(ch) for ch in c) and c not in out):
                out.append(c)
        return out

    def _valid_org(self, cand: str, suffix_len: int) -> bool:
        core = cand[:-suffix_len]
        if len(core) < 2:
            return False
        if any(ch in ORG_BAD_CONTENT or ch.isdigit() for ch in cand):
            return False
        if core in ORG_BAD_CORE or core[:2] in ORG_BAD_CORE:
            return False
        # 机构与企业分别由两个开关控制，默认都保留
        if is_institution(cand):
            return self.redact_institutions
        return self.redact_companies

    def _address(self, text):
        out = []
        for m in RE_ADDR_CTX.finditer(text):
            s = m.group(1).strip()
            s = self._trim_addr(s)
            if self._valid_addr(s):
                start = m.start(1) + m.group(1).find(s)
                out.append(Hit(start, start + len(s), "address", s, 0.85))
        for rx in (RE_ADDR_STRUCT, RE_ADDR_STRUCT2):
            for m in rx.finditer(text):
                raw = m.group(0)
                s = STRUCT_LEAD.sub("", raw, count=1)
                start = m.start() + (len(raw) - len(s))
                if self._valid_addr(s) and not contains_marker(s):
                    out.append(Hit(start, start + len(s), "address", s, 0.9))
        return out

    @staticmethod
    def _trim_addr(s: str) -> str:
        m = re.search(r"(?:有限公司|有限责任公司|股份有限公司|合伙企业|律师事务所|"
                      r"会计师事务所|银行|支行|分行|公司)", s)
        if m:
            idx = m.start()
            cut = -1
            for sep in ("号", "室"):
                cut = max(cut, s.rfind(sep, 0, idx))
            if cut > 0:
                s = s[:cut + 1]
            else:
                s = s[:idx]
        return s.strip(" ：:，,、")

    @staticmethod
    def _valid_addr(s: str) -> bool:
        if len(s) < 4:
            return False
        if not any(is_cjk(c) for c in s):
            return False
        has_region = any(k in s for k in ("省", "市", "区", "县", "镇", "乡", "村"))
        has_detail = any(k in s for k in ("路", "街", "道", "巷", "号", "室", "栋",
                                          "幢", "单元", "楼", "村", "镇", "乡"))
        return has_region and has_detail

    # -- 汇总 -----------------------------------------------------------

    def detect(self, text: str) -> list:
        if not text:
            return []
        c = self.categories
        rules = []
        if "id_card" in c:
            rules.append(self._id)
        if "bankcard" in c:
            rules.append(self._bankcard)
        if "phone" in c:
            rules.append(self._phone)
        if "landline" in c:
            rules.append(self._landline)
        if "birth" in c:
            rules.append(self._birth)
        if "email" in c:
            rules.append(self._email)
        if "qq" in c:
            rules.append(self._qq)
        if "wechat" in c:
            rules.append(self._wechat)
        if "passport" in c:
            rules.append(self._passport)
        if "address" in c:
            rules.append(self._address)
        if "company" in c and (self.redact_companies or self.redact_institutions):
            rules.append(self._company)
        if "person" in c:
            rules.append(self._person)
        rules.append(self._uscc)

        hits = []
        for fn in rules:
            try:
                hits.extend(fn(text))
            except Exception:
                continue
        return self._resolve(hits)

    @staticmethod
    def _resolve(hits):
        hits.sort(key=lambda h: (-PRIORITY.get(h.category, 0),
                                -(h.end - h.start), h.start))
        accepted = []
        for h in hits:
            if any(not (h.end <= a.start or h.start >= a.end) for a in accepted):
                continue
            accepted.append(h)
        accepted.sort(key=lambda h: h.start)
        return accepted


# --------------------------------------------------------------------------
# 映射与替换
# --------------------------------------------------------------------------

class Mapping:
    def __init__(self, mode="code"):
        self.mode = mode
        self._map = {}
        self.entries = []
        self._person_idx = 0
        self._company_idx = 0
        self._person_used = set()
        self._company_used = set()
        self._real_surnames = set()

    def replacement(self, category: str, original: str) -> str:
        # PDF 提取常在姓名/企业名中间插入空格（如“王小 明”），统一按去空格后的值建档
        norm = (re.sub(r"\s+", "", original)
                if category in ("person", "company") else original)
        key = (category, norm)
        if key in self._map:
            return self._map[key]
        if self.mode == "code":
            if category == "person":
                rep = self._next_person(norm)
            elif category == "company":
                rep = self._next_company(norm)
            else:
                rep = self._code_mask(category, original)
        elif self.mode == "star":
            rep = "***"
        else:
            rep = self._keep_mask(category, original)
        self._map[key] = rep
        self.entries.append({"category": category, "label": CAT_LABEL.get(category, category),
                             "original": norm, "replacement": rep})
        return rep

    def _next_person(self, original: str):
        orig_surname = original[:2] if original[:2] in COMPOUND_SURNAMES else original[:1]
        self._real_surnames.add(orig_surname)
        while self._person_idx < len(SURNAME_POOL):
            s = SURNAME_POOL[self._person_idx]
            self._person_idx += 1
            # 代号姓氏不能与本批任何真实姓氏相同，否则等于没脱敏
            if s in self._person_used or s in self._real_surnames:
                continue
            self._person_used.add(s)
            return s + "某"
        return "某甲"

    def _next_company(self, original: str):
        typ = org_code_type(original)
        while self._company_idx < len(company_pool):
            s = company_pool[self._company_idx]
            self._company_idx += 1
            if s not in self._company_used:
                self._company_used.add(s)
                return s + typ
        return "第%d%s" % (self._company_idx + 1, typ)

    @staticmethod
    def _code_mask(category, s):
        if category == "id_card" and len(s) == 18:
            return s[:6] + "*" * 8 + s[-4:]
        if category == "phone" and len(s) == 11:
            return s[:3] + "****" + s[-4:]
        if category == "landline":
            m = re.match(r"(0\d{2,3})", s)
            digits = re.sub(r"\D", "", s)
            if m:
                return m.group(1) + "-" + "*" * max(7, len(digits) - len(m.group(1)))
            return "*" * len(digits)
        if category == "birth":
            return Mapping._mask_birth(s)
        if category == "email":
            local, _, domain = s.partition("@")
            tld = domain.rsplit(".", 1)[-1] if "." in domain else "com"
            return (local[0] if local else "*") + "***@***." + tld
        if category == "qq":
            return "*" * 6
        if category == "wechat":
            return "*" * 6
        if category == "passport":
            return (s[0] + "*" * (len(s) - 1)) if s else "***"
        if category == "bankcard":
            return "*" * max(0, len(s) - 4) + s[-4:]
        if category == "uscc" and len(s) == 18:
            return s[:2] + "*" * 12 + s[-4:]
        if category == "address":
            m = re.match(
                r"^([一-龥]{2,8}?(?:省|自治区|特别行政区))?"
                r"([一-龥]{2,10}?(?:市|自治州|地区|盟))?"
                r"([一-龥]{2,10}?(?:区|县|旗))?", s)
            prov, city, county = (m.groups() if m else (None, None, None))
            if city:
                # 有市级时保留到市，区县一并遮蔽
                prefix = (prov or "") + city
            elif county:
                # 只写到县的（如“浙江省新昌县”），保留到县，便于判断管辖
                prefix = (prov or "") + county
            else:
                prefix = prov or ""
            return prefix + "〔地址已脱敏〕"
        return "***"

    @staticmethod
    def _mask_birth(s):
        y = re.match(r"\s*((?:19|20)\d{2})", s)
        year = y.group(1) if y else "****"
        if "年" in s:
            return year + "年**月" + ("**日" if "日" in s else "")
        if any(sep in s for sep in "-/."):
            sep = next(x for x in "-/." if x in s)
            if s.count(sep) >= 2:
                return year + sep + "**" + sep + "**"
            return year + sep + "**"
        return year + "****"

    @staticmethod
    def _keep_mask(category, s):
        if not s:
            return "***"
        if category == "company":
            if len(s) <= 4:
                return s[0] + "***"
            return s[:2] + "***" + s[-2:]
        if category == "address":
            return s[:6] + "***" if len(s) > 6 else s[0] + "***"
        if len(s) == 1:
            return "*"
        if len(s) == 2:
            return s[0] + "*"
        return s[0] + "*" * (len(s) - 2) + s[-1]


def redact_text(text: str, detector: Detector, mapping: Mapping):
    hits = detector.detect(text)
    if not hits:
        return text, []
    out = []
    last = 0
    applied = []
    for h in hits:
        rep = mapping.replacement(h.category, h.text)
        out.append(text[last:h.start])
        out.append(rep)
        last = h.end
        applied.append((h, rep))
    out.append(text[last:])
    return "".join(out), applied


# --------------------------------------------------------------------------
# OOXML（docx）处理
# --------------------------------------------------------------------------

DOCX_TEXT_PART = re.compile(
    r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$"
)


def _nearest_p(el):
    p = el.getparent()
    while p is not None and p.tag != W_P:
        p = p.getparent()
    return p


def _nearest_tag(el, tag):
    p = el.getparent()
    while p is not None and p.tag != tag:
        p = p.getparent()
    return p


def _cell_text(tc):
    return "".join(t.text or "" for t in tc.iter(W_T))


def _preserve_space(t):
    txt = t.text
    if txt and (txt[:1].isspace() or txt[-1:].isspace()):
        t.set(XML_SPACE, "preserve")


def _apply_to_ts(ts, hits, reps):
    spans = []
    pos = 0
    for t in ts:
        s = t.text or ""
        spans.append((pos, pos + len(s), t))
        pos += len(s)

    for h in sorted(hits, key=lambda x: -x.start):
        rep = reps[(h.start, h.end)]
        affected = [(a, b, t) for a, b, t in spans if b > h.start and a < h.end]
        if not affected:
            continue
        a0, b0, t0 = affected[0]
        ls = max(h.start, a0) - a0
        le = min(h.end, b0) - a0
        cur = t0.text or ""
        t0.text = cur[:ls] + rep + cur[le:]
        _preserve_space(t0)
        for a, b, t in affected[1:]:
            cur = t.text or ""
            ls2 = max(h.start, a) - a
            le2 = min(h.end, b) - a
            t.text = cur[:ls2] + cur[le2:]
            _preserve_space(t)


def _replace_cell_text(tc, txt: str, rep: str):
    changed = 0
    for p in tc.iter(W_P):
        ts = [t for t in p.iter(W_T) if _nearest_p(t) is p]
        if not ts:
            continue
        text = "".join(t.text or "" for t in ts)
        if text.strip() != txt:
            continue
        start = text.find(txt)
        if start < 0:
            continue
        _apply_to_ts(ts, [Hit(start, start + len(txt), "person", txt)],
                     {(start, start + len(txt)): rep})
        changed += 1
    return changed


def _redact_table_names(root, detector: Detector, mapping: Mapping, stats):
    """表格中孤立姓名：仅在表头列或角色列佐证时才替换，避免误伤。"""
    count = 0
    for tbl in root.iter(W_TBL):
        trs = tbl.findall(W_TR)
        if not trs:
            continue
        header_cells = trs[0].findall(W_TC)
        name_cols = {i for i, tc in enumerate(header_cells)
                     if _cell_text(tc).strip() in NAME_HEADERS}
        for ri, tr in enumerate(trs):
            cells = tr.findall(W_TC)
            targets = set()
            if ri > 0:
                for idx in name_cols:
                    if idx < len(cells):
                        targets.add(idx)
            for idx, tc in enumerate(cells):
                if _cell_text(tc).strip() in ROLE_WORDS and idx + 1 < len(cells):
                    targets.add(idx + 1)
            for idx in targets:
                tc = cells[idx]
                txt = _cell_text(tc).strip()
                if not detector.isolated_person(txt):
                    continue
                rep = mapping.replacement("person", txt)
                if _replace_cell_text(tc, txt, rep):
                    stats["person"] += 1
                    count += 1
    return count


def redact_ooxml_part(data: bytes, detector: Detector, mapping: Mapping, stats):
    try:
        parser = etree.XMLParser(recover=True, huge_tree=True, remove_blank_text=False)
        root = etree.fromstring(data, parser)
    except Exception:
        return data, 0

    count = 0
    for p in root.iter(W_P):
        ts = [t for t in p.iter(W_T) if _nearest_p(t) is p]
        if not ts:
            continue
        text = "".join(t.text or "" for t in ts)
        if not text.strip():
            continue
        hits = detector.detect(text)
        if not hits:
            continue
        reps = {}
        valid_hits = []
        for h in hits:
            reps[(h.start, h.end)] = mapping.replacement(h.category, h.text)
            valid_hits.append(h)
            stats[h.category] += 1
        _apply_to_ts(ts, valid_hits, reps)
        count += len(valid_hits)

    # 批注作者
    for c in root.iter(W_COMMENT):
        if c.get("author"):
            c.set("author", "***")
            count += 1
        if c.get("initials"):
            c.set("initials", "***")

    # 表格中的孤立姓名
    count += _redact_table_names(root, detector, mapping, stats)

    if count:
        out = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                             standalone=True)
        return out, count
    return data, 0


def redact_doc_props(data: bytes, stats):
    try:
        parser = etree.XMLParser(recover=True, huge_tree=True)
        root = etree.fromstring(data, parser)
    except Exception:
        return data, 0
    count = 0
    targets = [
        "{%s}creator" % DC_NS, "{%s}title" % DC_NS, "{%s}subject" % DC_NS,
        "{%s}description" % DC_NS,
        "{%s}lastModifiedBy" % CORE_NS, "{%s}keywords" % CORE_NS,
        "{%s}category" % CORE_NS,
        "{%s}Company" % APP_NS, "{%s}Manager" % APP_NS,
    ]
    for tag in targets:
        for el in root.iter(tag):
            if el.text and el.text.strip():
                el.text = ""
                count += 1
    if count:
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                              standalone=True), count
    return data, 0


def write_zip(zin: zipfile.ZipFile, dst: Path, replacements: dict):
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replacements.get(item.filename)
            if data is None:
                data = zin.read(item.filename)
            new_item = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            new_item.compress_type = item.compress_type
            new_item.external_attr = item.external_attr
            new_item.internal_attr = item.internal_attr
            new_item.create_system = item.create_system
            zout.writestr(new_item, data)


def process_docx(src: Path, dst: Path, ctx):
    stats = defaultdict(int)
    warnings = []
    with zipfile.ZipFile(src) as zin:
        repl = {}
        for name in zin.namelist():
            if DOCX_TEXT_PART.match(name):
                new, n = redact_ooxml_part(zin.read(name), ctx.detector, ctx.mapping, stats)
                if n:
                    repl[name] = new
            elif name in ("docProps/core.xml", "docProps/app.xml"):
                new, n = redact_doc_props(zin.read(name), stats)
                if n:
                    repl[name] = new
                    stats["__props__"] += n
        if not ctx.dry_run:
            write_zip(zin, dst, repl)
    return stats, warnings


# --------------------------------------------------------------------------
# xlsx 处理
# --------------------------------------------------------------------------

RISKY_XLSX_PARTS = ("xl/charts/", "xl/pivotCache/", "xl/pivotTables/",
                    "xl/vbaProject.bin", "xl/drawings/")


def _num_to_str(v):
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


CELL_LABELS = {
    "身份证号", "身份证号码", "身份证", "手机号", "手机号码", "手机",
    "联系电话", "电话", "固定电话", "座机", "传真", "出生日期", "出生年月",
    "生日", "邮箱", "电子邮箱", "电子邮件", "QQ", "QQ号", "微信号", "微信",
    "护照号", "护照号码", "护照", "银行卡号", "卡号", "账号", "账户",
    "住址", "地址", "住所", "住所地", "姓名", "名称", "统一社会信用代码",
    "信用代码", "纳税人识别号",
}


def _apply_labeled(value: str, label: str, ctx, stats):
    combined = label + "：" + value
    offset = len(label) + 1
    hits = [h for h in ctx.detector.detect(combined) if h.end > offset]
    if not hits:
        return value
    out = value
    for h in sorted(hits, key=lambda x: -x.start):
        s = max(h.start, offset) - offset
        e = min(h.end, offset + len(value)) - offset
        if s >= e:
            continue
        rep = ctx.mapping.replacement(h.category, h.text)
        out = out[:s] + rep + out[e:]
        stats[h.category] += 1
    return out


def _xlsx_labeled_cells(ws, ctx, stats):
    """处理“标签在左/上、值在右/下”的表格结构，补足缺少上下文的号码类字段。"""
    maxr = ws.max_row or 0
    maxc = ws.max_column or 0
    if not maxr or not maxc:
        return

    for r in range(1, maxr + 1):
        for c in range(1, maxc):
            lab = ws.cell(r, c).value
            if not isinstance(lab, str) or lab.strip() not in CELL_LABELS:
                continue
            cell = ws.cell(r, c + 1)
            v = cell.value
            is_num = False
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                v = _num_to_str(v)
                is_num = True
            elif not isinstance(v, str):
                continue
            if not v:
                continue
            new = _apply_labeled(v, lab.strip(), ctx, stats)
            if new != v:
                cell.value = new
                if is_num:
                    cell.number_format = "@"

    for c in range(1, maxc + 1):
        for r in range(1, min(maxr, 5) + 1):
            lab = ws.cell(r, c).value
            if not isinstance(lab, str) or lab.strip() not in CELL_LABELS:
                continue
            for rr in range(r + 1, maxr + 1):
                cell = ws.cell(rr, c)
                v = cell.value
                is_num = False
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    v = _num_to_str(v)
                    is_num = True
                elif not isinstance(v, str):
                    continue
                if not v:
                    continue
                new = _apply_labeled(v, lab.strip(), ctx, stats)
                if new != v:
                    cell.value = new
                    if is_num:
                        cell.number_format = "@"


def _xlsx_isolated_names(ws, ctx, stats):
    maxr = ws.max_row or 0
    maxc = ws.max_column or 0
    if maxr == 0 or maxc == 0:
        return
    name_cols = {}
    for r in range(1, min(maxr, 5) + 1):
        for c in range(1, maxc + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() in NAME_HEADERS:
                name_cols[c] = r
    for c, hr in name_cols.items():
        for r in range(hr + 1, maxr + 1):
            cell = ws.cell(r, c)
            if not isinstance(cell.value, str):
                continue
            if ctx.detector.isolated_person(cell.value):
                cell.value = ctx.mapping.replacement("person", cell.value)
                stats["person"] += 1
    for r in range(1, maxr + 1):
        for c in range(1, maxc):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() in ROLE_WORDS:
                cell = ws.cell(r, c + 1)
                if isinstance(cell.value, str) and ctx.detector.isolated_person(cell.value):
                    cell.value = ctx.mapping.replacement("person", cell.value)
                    stats["person"] += 1


def process_xlsx(src: Path, dst: Path, ctx):
    import openpyxl

    stats = defaultdict(int)
    warnings = []

    with zipfile.ZipFile(src) as z:
        names = z.namelist()
    risky = [n for n in names if any(p in n for p in RISKY_XLSX_PARTS)]
    if risky:
        warnings.append("工作簿含图表/透视表/宏/图形对象，openpyxl 可能无法完整保留，请人工复核")

    wb = openpyxl.load_workbook(src)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v:
                    new, applied = redact_text(v, ctx.detector, ctx.mapping)
                    if applied:
                        cell.value = new
                        for h, _ in applied:
                            stats[h.category] += 1
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    s = _num_to_str(v)
                    hits = ctx.detector.detect(s)
                    if (hits and hits[0].start == 0 and hits[0].end == len(s)
                            and hits[0].category in NUMERIC_CATEGORIES):
                        h = hits[0]
                        cell.value = ctx.mapping.replacement(h.category, h.text)
                        cell.number_format = "@"
                        stats[h.category] += 1
                if cell.comment:
                    ctext = cell.comment.text or ""
                    new, applied = redact_text(ctext, ctx.detector, ctx.mapping)
                    if applied:
                        cell.comment.text = new
                        for h, _ in applied:
                            stats[h.category] += 1
                    if cell.comment.author:
                        cell.comment.author = "***"
        # 页眉页脚
        for hf in (ws.oddHeader, ws.oddFooter, ws.evenHeader,
                   ws.evenFooter, ws.firstHeader, ws.firstFooter):
            for attr in ("left", "center", "right"):
                part = getattr(hf, attr, None)
                if part is not None and getattr(part, "text", None):
                    new, applied = redact_text(part.text, ctx.detector, ctx.mapping)
                    if applied:
                        part.text = new
                        for h, _ in applied:
                            stats[h.category] += 1
        _xlsx_labeled_cells(ws, ctx, stats)
        _xlsx_isolated_names(ws, ctx, stats)

    if not ctx.dry_run:
        wb.save(dst)
    return stats, warnings


# --------------------------------------------------------------------------
# Office 旧格式转换
# --------------------------------------------------------------------------

_com_lock = threading.Lock()
_com_failures = 0


def convert_old_format(src: Path, dst: Path, kind: str, timeout=90):
    """用本机 Word/Excel 将 .doc/.xls 转为 .docx/.xlsx。返回 (ok, msg)。"""
    global _com_failures
    if _com_failures >= 2:
        return False, "Office 转换连续失败，已对本批次禁用自动转换"

    result = {"ok": False, "msg": ""}

    def worker():
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            if kind == "word":
                app = win32com.client.DispatchEx("Word.Application")
                app.Visible = False
                app.DisplayAlerts = 0
                try:
                    app.AutomationSecurity = 3  # 禁用宏
                except Exception:
                    pass
                doc = app.Documents.Open(str(src), ReadOnly=True,
                                         AddToRecentFiles=False)
                try:
                    doc.SaveAs2(str(dst), FileFormat=16)  # wdFormatDocumentDefault
                finally:
                    doc.Close(False)
                app.Quit()
            else:
                app = win32com.client.DispatchEx("Excel.Application")
                app.Visible = False
                app.DisplayAlerts = False
                wb = app.Workbooks.Open(str(src), ReadOnly=True, UpdateLinks=0)
                try:
                    wb.SaveAs(str(dst), FileFormat=51)  # xlOpenXMLWorkbook
                finally:
                    wb.Close(False)
                app.Quit()
            result["ok"] = True
        except Exception as e:
            result["msg"] = "%s: %s" % (type(e).__name__, e)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    t = threading.Thread(target=worker, daemon=True)
    t.start()
    t.join(timeout)
    if t.is_alive():
        _com_failures += 1
        return False, "转换超时（%ds），已跳过" % timeout
    if not result["ok"]:
        _com_failures += 1
        return False, "转换失败：" + (result["msg"] or "未知错误")
    _com_failures = 0
    return True, ""


# --------------------------------------------------------------------------
# OCR 与图片
# --------------------------------------------------------------------------

_ocr_engine = None
_ocr_error = None


def get_ocr():
    global _ocr_engine, _ocr_error
    if _ocr_engine is not None:
        return _ocr_engine
    if _ocr_error is not None:
        raise RuntimeError(_ocr_error)
    try:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
        return _ocr_engine
    except Exception as e:
        _ocr_error = ("OCR 引擎不可用（%s）。请安装：pip install rapidocr-onnxruntime"
                      % e)
        raise RuntimeError(_ocr_error)


def ocr_lines(img):
    """返回 [(text, (x0,y0,x1,y1), score), ...]"""
    import numpy as np
    engine = get_ocr()
    arr = np.array(img.convert("RGB"))[:, :, ::-1].copy()  # RGB -> BGR
    out = engine(arr)
    if isinstance(out, tuple) and len(out) == 2:
        res = out[0]
    else:
        res = out
    lines = []
    for item in (res or []):
        try:
            box, text, score = item[0], item[1], item[2]
            xs = [float(p[0]) for p in box]
            ys = [float(p[1]) for p in box]
            lines.append((str(text), (min(xs), min(ys), max(xs), max(ys)), float(score)))
        except Exception:
            continue
    return lines


def redact_pil(img, ctx, stats, warnings):
    lines = ocr_lines(img)
    if not lines:
        warnings.append("未识别到文字")
        return img, 0

    texts = []
    ranges = []
    pos = 0
    for text, box, score in lines:
        ranges.append((pos, pos + len(text), box))
        texts.append(text)
        pos += len(text) + 1
    full = "\n".join(texts)
    hits = ctx.detector.detect(full)
    if not hits:
        return img, 0

    from PIL import ImageDraw
    draw = ImageDraw.Draw(img)
    drawn = set()
    for h in hits:
        touched = [box for s, e, box in ranges if e > h.start and s < h.end]
        if not touched:
            continue
        for box in touched:
            if box in drawn:
                continue
            drawn.add(box)
            x0, y0, x1, y1 = box
            draw.rectangle([x0 - 2, y0 - 2, x1 + 2, y1 + 2], fill=(0, 0, 0))
        stats[h.category] += 1

    low = [t for t, b, s in lines if s < 0.5]
    if low:
        warnings.append("OCR 置信度低于 0.5 的文本行 %d 条，请人工复核" % len(low))
    warnings.append("OCR 按整行遮盖，可能连带遮盖同行其他内容")
    return img, len(hits)


def process_image(src: Path, dst: Path, ctx):
    from PIL import Image

    stats = defaultdict(int)
    warnings = []
    if ctx.no_ocr:
        raise RuntimeError("图片脱敏需要 OCR，但已通过 --no-ocr 禁用")
    img = Image.open(src).convert("RGB")
    img, n = redact_pil(img, ctx, stats, warnings)
    if not ctx.dry_run:
        ext = dst.suffix.lower()
        if ext in (".jpg", ".jpeg"):
            img.save(dst, quality=95, optimize=True)
        else:
            img.save(dst)
    return stats, warnings


# --------------------------------------------------------------------------
# PDF 处理
# --------------------------------------------------------------------------

def _pdf_page_textmap(page):
    """逐字符建立“文本偏移 -> bbox”映射，保证只抹掉命中字符本身。"""
    try:
        d = page.get_text("rawdict")
    except Exception:
        d = page.get_text("dict")
    parts = []
    entries = []
    pos = 0
    line_id = 0
    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                chars = span.get("chars")
                if chars:
                    for ch in chars:
                        c = ch.get("c", "")
                        if not c:
                            continue
                        entries.append((pos, pos + len(c), ch.get("bbox"), line_id))
                        parts.append(c)
                        pos += len(c)
                else:
                    t = span.get("text", "")
                    if not t:
                        continue
                    entries.append((pos, pos + len(t), span.get("bbox"), line_id))
                    parts.append(t)
                    pos += len(t)
            line_id += 1
            parts.append("\n")
            pos += 1
    return "".join(parts), entries


def _redact_pdf_page(page, ctx, stats, warnings, fitz):
    # 先清除非链接批注（批注内容可能含 PII），必须在添加修订注释之前做，
    # 否则会把刚添加的 Redact 注释一并删掉，导致原文未被删除。
    try:
        for annot in list(page.annots() or []):
            try:
                if annot.type[0] != fitz.PDF_ANNOT_LINK:
                    page.delete_annot(annot)
                    stats["__annots__"] += 1
            except Exception:
                pass
    except Exception:
        pass

    # 表单域
    try:
        for w in page.widgets() or []:
            if getattr(w, "field_value", None) not in (None, ""):
                try:
                    w.field_value = ""
                    w.update()
                    stats["__forms__"] += 1
                except Exception:
                    pass
    except Exception:
        pass

    text, entries = _pdf_page_textmap(page)
    if text.strip():
        hits = ctx.detector.detect(text)
        for h in hits:
            # 按行分组：姓名若被 PDF 拆到两行，分别遮盖，避免跨行大黑块
            rects = {}
            for s, e, bbox, ln in entries:
                if e > h.start and s < h.end and bbox:
                    r = fitz.Rect(bbox)
                    rects[ln] = r if ln not in rects else (rects[ln] | r)
            if not rects:
                continue
            for r in rects.values():
                r = fitz.Rect(r.x0 - 1, r.y0 - 1, r.x1 + 1, r.y1 + 1)
                page.add_redact_annot(r, fill=(0, 0, 0))
            stats[h.category] += 1

    page.apply_redactions()


def _import_fitz():
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    return fitz


def _page_image_coverage(page):
    """图片覆盖页面面积的比例，用于识别“扫描件 + OCR 文字层”。"""
    try:
        page_area = abs(page.rect)
        if page_area <= 0:
            return 0.0
        covered = 0.0
        for img in page.get_images(full=True):
            xref = img[0]
            for r in page.get_image_rects(xref):
                covered += abs(r)
        return min(covered / page_area, 1.0)
    except Exception:
        return 0.0


def process_pdf(src: Path, dst: Path, ctx):
    fitz = _import_fitz()
    from PIL import Image

    stats = defaultdict(int)
    warnings = ["PDF 采用修订式删除，命中处以黑块遮盖，不回填代号文本；"
                "如需保留可读代号，请使用 docx 原件处理"]

    doc = fitz.open(src)
    if doc.needs_pass:
        doc.close()
        raise RuntimeError("PDF 已加密，请先去除密码后再处理")

    if doc.embfile_names():
        warnings.append("PDF 含内嵌附件 %d 个，已剔除以免携带原始材料"
                        % len(doc.embfile_names()))

    out = fitz.open()
    try:
        for pno in range(doc.page_count):
            page = doc[pno]
            txt = page.get_text("text")
            has_img = bool(page.get_images())
            coverage = _page_image_coverage(page)
            # 覆盖过半即视为图片型页面；带 OCR 文字层的扫描件也走这条分支，
            # 否则只删文字层、底图仍可见。
            is_scan = coverage >= 0.5 or (len(txt.strip()) < 20 and has_img)
            if is_scan:
                if ctx.no_ocr:
                    raise RuntimeError("第 %d 页为扫描页，需要 OCR，但已禁用" % (pno + 1))
                pix = page.get_pixmap(dpi=ctx.dpi)
                img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
                img, n = redact_pil(img, ctx, stats, warnings)
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                newpage = out.new_page(width=page.rect.width, height=page.rect.height)
                newpage.insert_image(newpage.rect, stream=buf.getvalue())
                if txt.strip():
                    warnings.append("第 %d 页为图片型页面（含文字层），"
                                    "已整体栅格化重建，原文字层已丢弃" % (pno + 1))
                else:
                    warnings.append("第 %d 页为扫描页，已栅格化重建（失去文字层）" % (pno + 1))
            else:
                if has_img:
                    warnings.append("第 %d 页含图片，图片内文字未做 OCR" % (pno + 1))
                if not txt.strip():
                    warnings.append("第 %d 页无可提取文字且无图片，"
                                    "可能为矢量图形或手写，未处理" % (pno + 1))
                _redact_pdf_page(page, ctx, stats, warnings, fitz)
                out.insert_pdf(doc, from_page=pno, to_page=pno)

        if not ctx.keep_embeds:
            pass  # insert_pdf 不复制附件

        out.set_metadata({
            "title": "", "author": "", "subject": "", "keywords": "",
            "creator": "", "producer": "redact.py", "creationDate": "", "modDate": "",
        })
        try:
            out.del_xml_metadata()
        except Exception:
            pass

        if not ctx.dry_run:
            out.save(dst, garbage=4, deflate=True)
    finally:
        out.close()
        doc.close()

    return stats, warnings


# --------------------------------------------------------------------------
# 文本抽取（用于残余复扫）
# --------------------------------------------------------------------------

def extract_text(path) -> str:
    path = Path(path)
    ext = path.suffix.lower()
    try:
        if ext == ".docx":
            with zipfile.ZipFile(path) as z:
                chunks = []
                for name in z.namelist():
                    if DOCX_TEXT_PART.match(name):
                        root = etree.fromstring(z.read(name),
                                                etree.XMLParser(recover=True, huge_tree=True))
                        chunks.extend(t.text or "" for t in root.iter(W_T))
                return "\n".join(chunks)
        if ext == ".xlsx":
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            chunks = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    for v in row:
                        if isinstance(v, str):
                            chunks.append(v)
            wb.close()
            return "\n".join(chunks)
        if ext == ".pdf":
            fitz = _import_fitz()
            doc = fitz.open(path)
            text = "\n".join(doc[i].get_text("text") for i in range(doc.page_count))
            doc.close()
            return text
    except Exception:
        return ""
    return ""


# --------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------

@dataclass
class Ctx:
    detector: Detector
    mapping: Mapping
    mode: str
    dpi: int
    dry_run: bool
    no_ocr: bool
    keep_embeds: bool
    overwrite: bool


@dataclass
class FileResult:
    rel: str
    status: str
    counts: dict = field(default_factory=dict)
    warnings: list = field(default_factory=list)
    error: str = ""
    output: str = ""


def _dst_path(outdir: Path, rel: Path, new_suffix: str = None, overwrite=False):
    rel = Path(rel)
    if new_suffix:
        rel = rel.with_suffix(new_suffix)
    dst = outdir / rel
    if dst.exists() and not overwrite:
        stem, suf = dst.stem, dst.suffix
        i = 1
        while True:
            cand = dst.with_name("%s_%d%s" % (stem, i, suf))
            if not cand.exists():
                return cand
            i += 1
    return dst


def process_one(src: Path, rel: Path, outdir: Path, ctx: Ctx) -> FileResult:
    rel_s = str(rel)
    ext = src.suffix.lower()
    res = FileResult(rel=rel_s, status="成功")

    try:
        if ext in (".doc", ".xls"):
            conv_ext = ".docx" if ext == ".doc" else ".xlsx"
            tmpdir = Path(os.environ.get("TEMP", ".")) / ("redact_conv_%d" % os.getpid())
            tmpdir.mkdir(parents=True, exist_ok=True)
            conv = tmpdir / (src.stem + conv_ext)
            ok, msg = convert_old_format(src, conv, "word" if ext == ".doc" else "excel")
            if not ok:
                res.status = "失败"
                res.error = msg
                return res
            dst = _dst_path(outdir, rel, conv_ext, ctx.overwrite)
            dst.parent.mkdir(parents=True, exist_ok=True)
            stats, warns = (process_docx if conv_ext == ".docx" else process_xlsx)(conv, dst, ctx)
            res.warnings.append("旧格式已转换为 %s，请复核版式" % conv_ext)
            try:
                conv.unlink()
            except Exception:
                pass
        elif ext == ".docx":
            dst = _dst_path(outdir, rel, None, ctx.overwrite)
            dst.parent.mkdir(parents=True, exist_ok=True)
            stats, warns = process_docx(src, dst, ctx)
        elif ext == ".xlsx":
            dst = _dst_path(outdir, rel, None, ctx.overwrite)
            dst.parent.mkdir(parents=True, exist_ok=True)
            stats, warns = process_xlsx(src, dst, ctx)
        elif ext == ".pdf":
            dst = _dst_path(outdir, rel, None, ctx.overwrite)
            dst.parent.mkdir(parents=True, exist_ok=True)
            stats, warns = process_pdf(src, dst, ctx)
        elif ext in IMAGE_EXTS:
            dst = _dst_path(outdir, rel, None, ctx.overwrite)
            dst.parent.mkdir(parents=True, exist_ok=True)
            stats, warns = process_image(src, dst, ctx)
        else:
            res.status = "跳过"
            res.error = "不支持的文件类型"
            return res

        res.counts = {k: v for k, v in stats.items() if not k.startswith("__")}
        res.warnings.extend(warns)
        res.output = str(dst)
        if ctx.dry_run:
            res.status = "预检"
            res.output = ""
        total = sum(res.counts.values())
        res.error = "命中 %d 处" % total
    except Exception as e:
        res.status = "失败"
        res.error = "%s: %s" % (type(e).__name__, e)
        res.warnings.append(traceback.format_exc(limit=3))
    return res


def collect_files(root: Path, outdir: Path):
    if root.is_file():
        return [root]
    files = []
    for p in sorted(root.rglob("*")):
        if p.is_dir():
            continue
        try:
            if outdir in p.parents or p == outdir:
                continue
            # 跳过以往的输出目录，避免把脱敏件再脱敏一遍
            if any(part.startswith("redacted_") for part in p.relative_to(root).parts[:-1]):
                continue
        except Exception:
            pass
        if p.name.startswith("~$"):
            continue
        files.append(p)
    return files


def write_report(outdir: Path, src_roots, results, elapsed, mode_label, total_counts):
    ts = now_stamp()
    path = outdir / ("脱敏报告_%s.md" % ts)
    if isinstance(src_roots, (str, Path)):
        src_roots = [src_roots]
    lines = []
    lines.append("# 材料脱敏报告\n")
    lines.append("- 生成时间：%s" % datetime.now().strftime("%Y-%m-%d %H:%M"))
    lines.append("- 输入：%s" % "；".join("`%s`" % p for p in src_roots))
    lines.append("- 输出：`%s`" % outdir)
    lines.append("- 脱敏模式：%s" % mode_label)
    lines.append("- 耗时：%.1f 秒\n" % elapsed)

    ok = sum(1 for r in results if r.status in ("成功", "预检"))
    skip = sum(1 for r in results if r.status == "跳过")
    fail = sum(1 for r in results if r.status == "失败")
    lines.append("## 一、总体情况\n")
    lines.append("| 项目 | 数量 |")
    lines.append("|---|---|")
    lines.append("| 文件总数 | %d |" % len(results))
    lines.append("| 成功 | %d |" % ok)
    lines.append("| 跳过 | %d |" % skip)
    lines.append("| 失败 | %d |\n" % fail)

    lines.append("## 二、脱敏命中统计\n")
    lines.append("| 类别 | 命中数 |")
    lines.append("|---|---|")
    for cat in CATEGORIES:
        if total_counts.get(cat):
            lines.append("| %s | %d |" % (CAT_LABEL[cat], total_counts[cat]))
    if not any(total_counts.get(c) for c in CATEGORIES):
        lines.append("| （无命中） | 0 |")
    lines.append("")

    lines.append("## 三、逐文件明细\n")
    lines.append("| 文件 | 状态 | 命中 | 说明 |")
    lines.append("|---|---|---|---|")
    for r in results:
        total = sum(r.counts.values())
        note = r.error.replace("|", "\\|")
        lines.append("| `%s` | %s | %d | %s |" % (r.rel, r.status, total, note))
    lines.append("")

    warned = [r for r in results if r.warnings]
    if warned:
        lines.append("## 四、警告与需人工复核项\n")
        for r in warned:
            lines.append("### `%s`\n" % r.rel)
            for w in r.warnings:
                if w.strip():
                    lines.append("- %s" % w.strip().replace("\n", " / "))
            lines.append("")

    lines.append("## 五、需要人工处理的固定项\n")
    lines.append("- 公章、印章中的单位名称：图形化文字，OCR 不保证识别。")
    lines.append("- 手写体、签名、指印：不做识别。")
    lines.append("- 图片内嵌于 Word/Excel/PDF 正文中的文字：未做 OCR。")
    lines.append("- PDF 中以矢量曲线绘制、既无文字层也无图片的内容：无法识别。")
    lines.append("- 压缩包（zip/rar/7z）未处理。")
    lines.append("- 文件名本身未脱敏（仅保留原文件名，建议自行改名）。")
    lines.append("")

    lines.append("## 六、残余复扫\n")
    residual = defaultdict(int)
    for r in results:
        if r.status != "成功" or not r.output:
            continue
        p = Path(r.output)
        if p.suffix.lower() not in (".docx", ".xlsx", ".pdf"):
            continue
        text = extract_text(p)
        if not text:
            continue
        for h in _residual_detector.detect(text):
            residual[h.category] += 1
    if residual:
        lines.append("以下类别在输出文件中仍能检出，请人工核对（可能是未覆盖写法或规则漏检）：\n")
        lines.append("| 类别 | 剩余命中 |")
        lines.append("|---|---|")
        for cat, n in residual.items():
            lines.append("| %s | %d |" % (CAT_LABEL.get(cat, cat), n))
        lines.append("")
    else:
        lines.append("文本可提取的格式未检出残余命中（不含图片/扫描页的再次 OCR）。\n")

    lines.append("## 七、使用提醒\n")
    lines.append("1. 本报告与输出文件均为机器初筛结果，对外使用前必须逐份人工复核。")
    lines.append("2. 映射表含真实信息，属敏感文件，请与脱敏件分开存放，不要外传。")
    lines.append("3. 若需回溯原始信息，仅使用本地映射表，不要通过联网工具查询。")
    lines.append("4. 规则基于上下文与结构特征，姓名、地址、企业名称存在漏检可能。")
    lines.append("")

    path.write_text("\n".join(lines), encoding="utf-8")
    return path


_residual_detector = Detector(DEFAULT_CATEGORIES)


def write_mapping(outdir: Path, mapping: Mapping, src_root: Path):
    path = outdir / ("映射表_%s.json" % now_stamp())
    payload = {
        "_warning": "本文件包含真实个人信息，属敏感材料，请本地保管，禁止外传。",
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "input": str(src_root),
        "mode": mapping.mode,
        "entries": mapping.entries,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def parse_args(argv=None):
    p = argparse.ArgumentParser(
        description="材料脱敏工具：支持 docx/xlsx/doc/xls/pdf/图片，保留目录结构。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：\n"
               "  python redact.py D:\\案件\\材料 -o D:\\输出\n"
               "  python redact.py D:\\案件\\起诉状.docx --dry-run\n"
               "  python redact.py 起诉状.docx 证据清单.xlsx -o D:\\输出\n"
               "  python redact.py D:\\材料 --mode star --categories id_card,phone,person\n",
    )
    p.add_argument("inputs", nargs="*",
                   help="待处理文件或目录，可一次传多个（支持拖拽多个文件）")
    p.add_argument("-o", "--output", help="输出目录（默认：输入同级 redacted_时间戳）")
    p.add_argument("--mode", choices=["code", "star", "keep"], default="code",
                   help="脱敏呈现：code 代号一致化（默认）/ star 全遮蔽 / keep 保留首尾")
    p.add_argument("--categories", default="all",
                   help="启用的类别，逗号分隔；默认除银行卡、企业名称外全部，"
                        "企业名称需再加 --redact-companies")
    p.add_argument("--dry-run", action="store_true", help="只扫描出报告，不写文件")
    p.add_argument("--dpi", type=int, default=300, help="扫描件/图片 OCR 渲染 DPI，默认 300")
    p.add_argument("--no-ocr", action="store_true", help="禁用 OCR（图片与扫描页将失败）")
    p.add_argument("--keep-embeds", action="store_true", help="保留 PDF 内嵌附件（默认剔除）")
    p.add_argument("--aggressive-names", action="store_true", help="更激进的姓名识别（误报率上升）")
    p.add_argument("--overwrite", action="store_true", help="允许覆盖同名输出文件")
    p.add_argument("--no-mapping", action="store_true", help="不生成映射表")
    p.add_argument("--redact-companies", action="store_true",
                   help="脱敏公司、律所、会计师事务所、合伙企业、工厂、商行等企业名称（默认保留）")
    p.add_argument("--redact-institutions", action="store_true",
                   help="脱敏银行、法院、公安等机构名称（默认保留）")
    p.add_argument("--list-categories", action="store_true", help="列出全部类别后退出")
    return p.parse_args(argv)


def main(argv=None):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    args = parse_args(argv)
    if args.list_categories:
        for c in CATEGORIES:
            print("%-10s %s" % (c, CAT_LABEL[c]))
        return 0

    if not args.inputs:
        print("错误：缺少输入路径。使用 -h 查看用法。", file=sys.stderr)
        return 2

    inputs = []
    for x in args.inputs:
        p = Path(x).expanduser().resolve()
        if not p.exists():
            print("输入不存在：%s" % p, file=sys.stderr)
            return 2
        inputs.append(p)

    if args.categories.strip().lower() == "all":
        cats = list(DEFAULT_CATEGORIES)
    else:
        cats = [c.strip() for c in args.categories.split(",") if c.strip()]
        unknown = [c for c in cats if c not in CATEGORIES]
        if unknown:
            print("未知类别：%s" % ", ".join(unknown), file=sys.stderr)
            print("可用类别：%s" % ", ".join(CATEGORIES), file=sys.stderr)
            return 2

    if args.output:
        outdir = Path(args.output).expanduser().resolve()
    else:
        # 单个目录输出在其内部；单个文件或多个输入输出在第一个输入的旁边
        first = inputs[0]
        base = first if first.is_dir() else first.parent
        outdir = base / ("redacted_%s" % now_stamp())
    outdir.mkdir(parents=True, exist_ok=True)

    # 显式写了 --categories company 的，视为要求脱敏企业名称
    redact_companies = args.redact_companies or (
        args.categories.strip().lower() != "all" and "company" in cats)
    detector = Detector(cats, aggressive_names=args.aggressive_names,
                        redact_companies=redact_companies,
                        redact_institutions=args.redact_institutions)
    mapping = Mapping(args.mode)
    ctx = Ctx(detector=detector, mapping=mapping, mode=args.mode, dpi=args.dpi,
              dry_run=args.dry_run, no_ocr=args.no_ocr,
              keep_embeds=args.keep_embeds, overwrite=args.overwrite)

    pairs = []
    seen = set()
    multi = len(inputs) > 1
    for root in inputs:
        for f in collect_files(root, outdir):
            key = str(f.resolve()).lower()
            if key in seen:
                continue
            seen.add(key)
            if root.is_dir():
                rel = f.relative_to(root)
                if multi:
                    rel = Path(root.name) / rel
            else:
                rel = Path(f.name)
            pairs.append((f, rel))

    print("输入：%s" % "；".join(str(p) for p in inputs))
    print("输出：%s" % outdir)
    print("待处理：%d 个文件\n" % len(pairs))

    t0 = time.time()
    results = []
    total_counts = defaultdict(int)

    for idx, (f, rel) in enumerate(pairs, 1):
        res = process_one(f, rel, outdir, ctx)
        results.append(res)
        for k, v in res.counts.items():
            total_counts[k] += v
        flag = {"成功": "OK ", "预检": "DRY", "跳过": "SKIP", "失败": "FAIL"}.get(res.status, "    ")
        print("[%d/%d] %s %s  (%s)" % (idx, len(pairs), flag, rel, res.error))

    elapsed = time.time() - t0
    mode_label = {"code": "代号一致化", "star": "全遮蔽", "keep": "保留首尾"}[args.mode]

    # 残余复扫需要 detector
    globals()["_residual_detector"] = Detector(
        cats, aggressive_names=args.aggressive_names,
        redact_companies=redact_companies,
        redact_institutions=args.redact_institutions)

    report = write_report(outdir, inputs, results, elapsed, mode_label, total_counts)
    print("\n报告：%s" % report)

    if mapping.entries and not args.dry_run and not args.no_mapping:
        mpath = write_mapping(outdir, mapping, "；".join(str(p) for p in inputs))
        print("映射表（敏感）：%s" % mpath)

    ok = sum(1 for r in results if r.status in ("成功", "预检"))
    fail = sum(1 for r in results if r.status == "失败")
    print("\n完成：成功 %d，失败 %d，耗时 %.1f 秒" % (ok, fail, elapsed))
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
