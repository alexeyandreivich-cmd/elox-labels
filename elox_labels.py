# -*- coding: utf-8 -*-
"""
Генератор этикеток (Data Sheet) систем соединительных АО "ЭЛОКС-ПРОМ".
Справочник типов и кодов документов — из файла 4.pdf (колонки: тип, ЭТ, DS, код RU, код EN).
Макет: альбомный лист, две колонки, логотип в шапке, код документа внизу в рамке (отключаемо).
Логотип вшивается в exe через --add-data "logo.png;." (fallback — векторная копия).
Сборка: pyinstaller --onefile --noconsole --add-data "logo.png;." --name EloxLabels elox_labels.py
"""

import os
import re
import sys
import json
import threading
import traceback
import datetime
import subprocess

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.platypus import Paragraph, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader

APP_TITLE = "ЭЛОКС-ПРОМ — генератор этикеток / Data Sheets"
START_ROW_DEFAULT = 3
COL_NAME, COL_QTY, COL_KKS, COL_UNIT = 4, 6, 7, 9
NB = "\u00a0"  # неразрывный пробел для линий заполнения

# =============================================================================
#  СПРАВОЧНИК ИЗ 4.pdf
#  тип : (ЭТ рус., DS англ., код документа RU (.R), код документа EN (.E))
# =============================================================================
TYPES_TABLE = [
    ("ЕА-М1250",     "ЕР.1.001.00-00 ЭТ", "ЕР.1.001.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0001.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0001.E"),
    ("ЕА-М500",      "ЕР.1.002.00-00 ЭТ", "ЕР.1.002.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0002.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0002.E"),
    ("ЕА-М350",      "ЕР.1.003.00-00 ЭТ", "ЕР.1.003.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0003.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0003.E"),
    ("ЕА-М250",      "ЕР.1.004.00-00 ЭТ", "ЕР.1.004.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0004.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0004.E"),
    ("ЕА-А02",       "ЕР.1.005.00-00 ЭТ", "ЕР.1.005.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0005.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0005.E"),
    ("ЕА-А01",       "ЕР.1.006.00-00 ЭТ", "ЕР.1.006.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0006.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0006.E"),
    ("ЕА-А1",        "ЕР.1.007.00-00 ЭТ", "ЕР.1.007.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0007.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0007.E"),
    ("ЕА-А2",        "ЕР.1.008.00-00 ЭТ", "ЕР.1.008.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0008.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0008.E"),
    ("ЕА-А4",        "ЕР.1.009.00-00 ЭТ", "ЕР.1.009.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0009.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0009.E"),
    ("ЕА-А6",        "ЕР.1.010.00-00 ЭТ", "ЕР.1.010.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0010.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0010.E"),
    ("ЕА-А8",        "ЕР.1.011.00-00 ЭТ", "ЕР.1.011.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0011.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0011.E"),
    ("ЕА-А10",       "ЕР.1.012.00-00 ЭТ", "ЕР.1.012.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0012.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0012.E"),
    ("ЕА-А12",       "ЕР.1.013.00-00 ЭТ", "ЕР.1.013.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0013.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0013.E"),
    ("ЕА-А14",       "ЕР.1.014.00-00 ЭТ", "ЕР.1.014.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0014.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0014.E"),
    ("ЕА-А16",       "ЕР.1.015.00-00 ЭТ", "ЕР.1.015.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0015.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0015.E"),
    ("ЕА-А16хк",     "ЕР.1.015.01-00 ЭТ", "ЕР.1.015.01-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0016.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0016.E"),
    ("ЕА-А16ха",     "ЕР.1.015.02-00 ЭТ", "ЕР.1.015.02-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0017.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0017.E"),
    ("ЕА-А18",       "ЕР.1.016.00-00 ЭТ", "ЕР.1.016.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0018.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0018.E"),
    ("ЕА-А18хк",     "ЕР.1.016.01-00 ЭТ", "ЕР.1.016.01-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0019.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0019.E"),
    ("ЕА-А18ха",     "ЕР.1.016.02-00 ЭТ", "ЕР.1.016.02-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0020.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0020.E"),
    ("ЕА-А18ТК-13",  "ЕР.1.016.03-00 ЭТ", "ЕР.1.016.03-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0021.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0021.E"),
    ("ЕА-А18ТК",     "ЕР.1.016.04-00 ЭТ", "ЕР.1.016.04-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0022.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0022.E"),
    ("ЕА-А18хкТК",   "ЕР.1.016.05-00 ЭТ", "ЕР.1.016.05-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0023.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0023.E"),
    ("ЕА-А18хаТК",   "ЕР.1.016.06-00 ЭТ", "ЕР.1.016.06-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0024.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0024.E"),
    ("ЕА-COAX50",    "ЕР.1.017.00-00 ЭТ", "ЕР.1.017.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0025.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0025.E"),
    ("ЕА-COAX75",    "ЕР.1.018.01-00 ЭТ", "ЕР.1.018.01-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0026.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0026.E"),
    ("ЕА-COAX3x50",  "ЕР.1.018.00-00 ЭТ", "ЕР.1.018.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0027.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0027.E"),
    ("ЕА-COAX3x75",  "ЕР.1.018.01-00 ЭТ", "ЕР.1.018.01-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0028.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0028.E"),
    ("ЕА-TRIAX",     "ЕР.1.019.00-00 ЭТ", "ЕР.1.019.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0029.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0029.E"),
    ("ЕА-OPTIC",     "ЕР.1.020.00-00 ЭТ", "ЕР.1.020.00-00 DS", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0030.R", "ED.S.A940.1.0UJA&&.JML&&&.092.ZB.0030.E"),
]

# Приведение регистра и замена кириллических букв на латинические "двойники".
# Цифры и состав обозначения НЕ меняются: ЕА-А02 и ЕА-А2 остаются разными типами.
_CYR2LAT = str.maketrans("ЕАМСОРТХКВНУІЅЈЁ", "EAMCOPTXKBHUISJE")


def norm_type(s: str) -> str:
    """Сравнение типов без учёта написания: ЕА-М350 (кир.) == EA-M350 (лат.)."""
    s = (s or "").strip().upper().translate(_CYR2LAT)
    s = s.replace("–", "-").replace("—", "-").replace("‑", "-")
    s = re.sub(r"\s+", "", s)
    return s


TYPES = {norm_type(t): (t, ru, en, dr, de) for t, ru, en, dr, de in TYPES_TABLE}

# =============================================================================
#  ТЕКСТЫ ФОРМ (1ru.pdf / 1en.pdf)
#  Внимание: шаблоны содержат разметку <u>...</u> — их нельзя прогонять через esc()
# =============================================================================
TEXTS = {
    "ru": {
        "head_left": "Атомное производство",
        "head_right": 'АО "ЭЛОКС-ПРОМ"',
        "title": "СИСТЕМА СОЕДИНИТЕЛЬНАЯ ТИПА {type}",
        "sub1": "ЭТИКЕТКА",
        "sub2": "{label}",
        "kks_label": "Код KKS:",
        "s1": "1  Основные технические данные",
        "p11": "1.1 Основные параметры системы соединительной приведены в ТУ 6981-002-33680530-02.",
        "s12_pre": "1.2 Класс безопасности",
        "s12_post": "по НП-001-97 (ПНАЭ Г-01-011-97)",
        "s13_pre": "1.3 Категория сейсмостойкости",
        "s13_post": "по НП-031-01",
        "p14": "1.4 Драгоценные металлы отсутствуют. Цветные металлы содержатся в незначительном количестве, учету и дальнейшей переработке не подлежат.",
        "s2": "2  Данные об изготовлении",
        "p21": "2.1 Дата изготовления:",
        "p22": "2.2 Изготовитель: АО «ЭЛОКС-ПРОМ».",
        "p23": "2.3 Адрес изготовителя: 307220, Российская Федерация, Курская обл., Курчатовский район, п. Иванино, ул. Ленина, 1В.",
        "p24": "2.4 Лицензия регистрационный номер ЦО-12-101-13757 от 15.02.2024 на изготовление оборудования для ядерных установок. Срок действия до 15.02.2029.",
        "s3": "3  Состав партии приведен в таблице 1",
        "tab_caption": "Таблица 1",
        "tab_head1": "Обозначение системы соединительной типа {type}, ТУ 6981-002-33680530-02",
        "tab_head2": "Количество",
        "s4": "4  Гарантии изготовителя (Поставщика)",
        "p41": "4.1 Срок службы системы соединительной типа {type} – {life}. Режим работы – непрерывный.",
        "p42": "4.2 Средний срок сохраняемости системы соединительной типа {type} до ввода в эксплуатацию – {shelf}.",
        "p43": "4.3 Гарантийный срок эксплуатации – {warranty} со дня ввода в эксплуатацию.",
        "s5": "5  Свидетельство о приемке",
        "p51": "5.1 Партия систем соединительных типа {type} изготовлена в соответствии с требованиями действующей технической документацией и отвечает требованиям ТУ 6981-002-33680530-02.",
        "p52": "5.2 План качества № <u>{plan}</u>.",
        "p53": "5.3 Системы соединительные признаны годными для эксплуатации.",
        "p54": "5.4 Обозначение документа, по которому производится поставка: <u>Договор № {contract}</u>",
        "sign_head": "Руководитель СОК",
        "cap_pos": "должность",
        "cap_sign": "подпись",
        "cap_name": "расшифровка подписи",
        "seal": "МП",
        "cap_date": "год, месяц, число",
        "s6": "6  Свидетельство об упаковывании",
        "p61": '6.1 Партия систем соединительных типа {type} упакована в АО "ЭЛОКС-ПРОМ" согласно требованиям действующей технической документации.',
        "years": ("год", "года", "лет"),
        "months": ("месяц", "месяца", "месяцев"),
    },
    "en": {
        "head_left": "Atomic Industry",
        "head_right": "ELOX-PROM JSC",
        "title": "CONNECTING SYSTEM TYPE {type}",
        "sub1": "DATA SHEET",
        "sub2": "{label}",
        "kks_label": "KKS Code:",
        "s1": "1  Main Technical Data",
        "p11": "1.1 Main parameters of the connecting system are given in Specifications TU 6981-002-33680530-02.",
        "s12_pre": "1.2 Safety class",
        "s12_post": "as per NP-001-97 (PNAE G-01-011-97)",
        "s13_pre": "1.3 Seismic category",
        "s13_post": "as per NP-031-01",
        "p14": "1.4 Contain no precious metals. Insignificant amount of non-ferrous metals; no subject to registration or further processing.",
        "s2": "2.  Manufacturing Data",
        "p21": "2.1. Date of manufacture:",
        "p22": "2.2. Manufacturer: ELOX-PROM JSC.",
        "p23": "2.3. Address: 1V, Lenina street, Ivanino village, Kurchatov district, Kursk Oblast, the Russian Federation, 307220.",
        "p24": "2.4. License for the manufacture of the equipment for nuclear power facilities № ЦО-12-101-13757 от 15.02.2024. Valid through 15.02.2029.",
        "s3": "3.  Composition of the batch is given in Table 1",
        "tab_caption": "Table 1",
        "tab_head1": "Designation of connecting system type {type}, Specifications TU 6981-002-33680530-02",
        "tab_head2": "Q-TY",
        "s4": "4.  Manufacturer’s (Supplier’s) Warranty",
        "p41": "4.1. Operational lifetime of connecting system type {type} – {life}. Operation mode – continuous.",
        "p42": "4.2. Average shelf life of connecting system type {type} before putting into operation – {shelf}.",
        "p43": "4.3. Warranty service life – {warranty} from the date of putting into operation.",
        "s5": "5.  Certificate of Acceptance",
        "p51": "5.1. The batch of connecting systems type {type} is manufactured in accordance with the effective technical documentation and meet the requirements of Specifications TU 6981-002-33680530-02.",
        "p52": "5.2. Quality Plan № <u>{plan}</u>.",
        "p53": "5.3. Connecting systems are classified as fit for operation and maintenance.",
        "p54": "5.4. Index of the document referenced to when supplying: <u>Contract № {contract}</u>.",
        "sign_head": "Head of QAS",
        "cap_pos": "position",
        "cap_sign": "signature",
        "cap_name": "full name",
        "seal": "SEAL HERE",
        "cap_date": "year, month, date",
        "s6": "6.  Packing Certificate",
        "p61": "6.1. The batch of connecting systems type {type} is packed at ELOX-PROM JSC according to the requirements set forth in the effective technical documentation.",
        "years": ("year", "years", "years"),
        "months": ("month", "months", "months"),
    },
}

# =============================================================================
#  ЧТЕНИЕ ТАБЛИЦЫ
# =============================================================================
def _cell_to_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _read_xls(path):
    import xlrd
    wb = xlrd.open_workbook(path, on_demand=True)
    sh = None
    for n in wb.sheet_names():
        if n.strip().lower() == "rptsheet":
            sh = wb.sheet_by_name(n)
            break
    if sh is None:
        sh = wb.sheet_by_index(0)
    rows = []
    for r in range(sh.nrows):
        row = []
        for c in range(sh.ncols):
            cell = sh.cell(r, c)
            v = cell.value
            if cell.ctype == xlrd.XL_CELL_NUMBER and float(v).is_integer():
                v = int(v)
            row.append(_cell_to_str(v))
        rows.append(row)
    return rows


def _read_xlsx(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = None
    for n in wb.sheetnames:
        if n.strip().lower() == "rptsheet":
            ws = wb[n]
            break
    if ws is None:
        ws = wb[wb.sheetnames[0]]
    return [[_cell_to_str(v) for v in row] for row in ws.iter_rows(values_only=True)]


def _read_html_xls(path):
    from html.parser import HTMLParser

    class P(HTMLParser):
        def __init__(self):
            super().__init__()
            self.rows, self.row, self.cell, self.inside = [], None, None, False

        def handle_starttag(self, tag, attrs):
            if tag == "tr":
                self.row = []
            elif tag in ("td", "th"):
                self.cell, self.inside = [], True

        def handle_endtag(self, tag):
            if tag in ("td", "th") and self.inside:
                self.row.append("".join(self.cell).strip())
                self.inside = False
            elif tag == "tr" and self.row is not None:
                self.rows.append(self.row)
                self.row = None

        def handle_data(self, data):
            if self.inside:
                self.cell.append(data)

    raw = open(path, "rb").read()
    text = None
    for enc in ("utf-8", "cp1251", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    p = P()
    p.feed(text)
    return p.rows


def _read_xml_xls(path):
    import xml.etree.ElementTree as ET
    ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
    root = ET.parse(path).getroot()
    sheets = root.findall("ss:Worksheet", ns)
    chosen = None
    for w in sheets:
        name = w.get("{urn:schemas-microsoft-com:office:spreadsheet}Name", "")
        if name.strip().lower() == "rptsheet":
            chosen = w
            break
    if chosen is None:
        chosen = sheets[0]
    rows = []
    for r in chosen.findall("ss:Table/ss:Row", ns):
        row, idx = [], 0
        for c in r.findall("ss:Cell", ns):
            i = c.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
            if i:
                i = int(i) - 1
                while idx < i:
                    row.append("")
                    idx += 1
            d = c.find("ss:Data", ns)
            row.append("".join(d.itertext()).strip() if d is not None else "")
            idx += 1
        rows.append(row)
    return rows


def load_table(path):
    with open(path, "rb") as f:
        head = f.read(4096)
    low = head.lower()
    errors = []
    if head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        order = [_read_xls]
    elif head[:2] == b"PK":
        order = [_read_xlsx]
    else:
        head_txt = low.decode("latin-1", "ignore")
        if "<table" in head_txt or "<html" in head_txt:
            order = [_read_html_xls, _read_xml_xls, _read_xls]
        elif "spreadsheet" in head_txt and "<?xml" in head_txt:
            order = [_read_xml_xls, _read_html_xls, _read_xls]
        else:
            order = [_read_xls, _read_xlsx, _read_html_xls, _read_xml_xls]
    for fn in order:
        try:
            rows = fn(path)
            if rows:
                return rows
        except Exception as e:
            errors.append(f"{fn.__name__}: {e}")
    raise RuntimeError("Не удалось прочитать файл.\n" + "\n".join(errors))


NAME_RE = re.compile(
    r"^\s*(Комплект\s+систем\s+соединительных|Система\s+соединительная)\s+(.+?)\s*$",
    re.IGNORECASE | re.DOTALL,
)


def parse_qty(s):
    s = (s or "").strip().replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except ValueError:
        return None


def fmt_qty(q):
    return str(int(q)) if float(q).is_integer() else str(q).replace(".", ",")


def unit_ru(u):
    u = (u or "").strip().lower().rstrip(".")
    return {"компл": "компл.", "шт": "шт."}.get(u, (u + ".") if u else "")


def unit_en(u, q):
    u = (u or "").strip().lower().rstrip(".")
    if u == "компл":
        return "set" if q == 1 else "sets"
    if u == "шт":
        return "pc" if q == 1 else "pcs"
    return u


def parse_rows(rows, start_row):
    records, warnings, skipped = [], [], 0
    for i in range(start_row - 1, len(rows)):
        row = rows[i]

        def get(idx):
            return row[idx].strip() if idx < len(row) and row[idx] is not None else ""

        name = get(COL_NAME)
        if not name:
            skipped += 1
            continue
        m = NAME_RE.match(name)
        if not m:
            skipped += 1
            continue
        tail = re.sub(r"\s+", " ", m.group(2)).strip()
        typ = tail.split(" ")[0]
        qty = parse_qty(get(COL_QTY))
        if qty is None or qty <= 0:
            warnings.append(f"Строка {i+1}: некорректное количество «{get(COL_QTY)}» — пропущена.")
            skipped += 1
            continue
        kks = re.split(r"[/\s]", get(COL_KKS))[0].strip()
        if not kks:
            warnings.append(f"Строка {i+1}: пустой KKS (столбец H).")
        ref = TYPES.get(norm_type(typ))
        if ref is None:
            warnings.append(f"Строка {i+1}: тип «{typ}» не найден в справочнике (4.pdf).")
            ref = (typ, "", "", "", "")
        u = get(COL_UNIT)
        records.append({
            "row": i + 1, "type": ref[0], "label_ru": ref[1], "label_en": ref[2],
            "doc_ru": ref[3], "doc_en": ref[4],
            "tail": tail, "kks": kks, "qty": qty,
            "qty_str": fmt_qty(qty), "unit_ru": unit_ru(u), "unit_en": unit_en(u, qty),
        })
    return records, warnings, skipped


# =============================================================================
#  ШРИФТЫ И ЛОГОТИП
# =============================================================================
FONT, FONT_B = "Helvetica", "Helvetica-Bold"


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource(name):
    base = getattr(sys, "_MEIPASS", app_dir())
    return os.path.join(base, name)


def register_fonts():
    global FONT, FONT_B
    win = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    regular = [resource("fonts/DejaVuSans.ttf"), os.path.join(app_dir(), "DejaVuSans.ttf"),
               os.path.join(win, "arial.ttf"), os.path.join(win, "calibri.ttf"),
               os.path.join(win, "segoeui.ttf"), os.path.join(win, "times.ttf")]
    bold = [resource("fonts/DejaVuSans-Bold.ttf"), os.path.join(app_dir(), "DejaVuSans-Bold.ttf"),
            os.path.join(win, "arialbd.ttf"), os.path.join(win, "calibrib.ttf"),
            os.path.join(win, "segoeuib.ttf"), os.path.join(win, "timesbd.ttf")]
    ok = False
    for p in regular:
        if os.path.isfile(p):
            try:
                pdfmetrics.registerFont(TTFont("LBL", p))
                FONT = "LBL"
                ok = True
                break
            except Exception:
                pass
    for p in bold:
        if os.path.isfile(p):
            try:
                pdfmetrics.registerFont(TTFont("LBL-B", p))
                FONT_B = "LBL-B"
                break
            except Exception:
                pass
    if FONT_B == "Helvetica-Bold" and ok:
        FONT_B = FONT
    if not ok:
        raise RuntimeError("Не найден шрифт с кириллицей (arial.ttf / DejaVuSans.ttf).")


LOGO_PATH = None


def find_logo():
    """Логотип ищется внутри exe (вшит через --add-data), затем рядом с exe."""
    global LOGO_PATH
    for p in (resource("logo.png"), os.path.join(app_dir(), "logo.png")):
        if os.path.isfile(p):
            LOGO_PATH = p
            return


def draw_logo(c, cx, cy, h):
    """Логотип: файл logo.png (вшит в exe), иначе векторная копия значка."""
    if LOGO_PATH:
        try:
            ir = ImageReader(LOGO_PATH)
            iw, ih = ir.getSize()
            w = h * iw / float(ih)
            c.drawImage(ir, cx - w / 2.0, cy - h / 2.0, w, h, mask="auto")
            return
        except Exception:
            pass
    w = h * 1.55
    beige = colors.Color(0.72, 0.70, 0.66)
    gray = colors.Color(0.62, 0.60, 0.57)
    red = colors.Color(0.72, 0.02, 0.10)
    c.setFillColor(beige)
    c.circle(cx - w * 0.16, cy, h * 0.52, stroke=0, fill=1)
    c.setStrokeColor(gray)
    c.setLineWidth(0.7)
    c.line(cx - w * 0.55, cy + h * 0.42, cx + w * 0.55, cy + h * 0.02)
    c.line(cx - w * 0.45, cy + h * 0.10, cx + w * 0.55, cy - h * 0.28)

    def box(x0, x1, y0, y1):
        p = c.beginPath()
        p.moveTo(x0 + h * 0.10, y1)
        p.lineTo(x1 + h * 0.10, y1)
        p.lineTo(x1, y0)
        p.lineTo(x0, y0)
        p.close()
        c.setFillColor(red)
        c.drawPath(p, fill=1, stroke=0)

    b1y0, b1y1 = cy + h * 0.06, cy + h * 0.40
    b2y0, b2y1 = cy - h * 0.40, cy - h * 0.06
    box(cx - w * 0.18, cx + w * 0.30, b1y0, b1y1)
    box(cx - w * 0.50, cx + w * 0.02, b2y0, b2y1)
    c.setFillColor(colors.white)
    c.setFont(FONT_B, h * 0.24)
    c.drawCentredString(cx + w * 0.06 + h * 0.05, (b1y0 + b1y1) / 2 - h * 0.08, "ELOX")
    c.drawCentredString(cx - w * 0.24 + h * 0.05, (b2y0 + b2y1) / 2 - h * 0.08, "ЭЛОКС")
    c.setFillColor(gray)
    c.setFont(FONT_B, h * 0.15)
    c.drawString(cx + w * 0.30 + h * 0.16, (b1y0 + b1y1) / 2 - h * 0.05, "PROM")
    c.drawString(cx + w * 0.02 + h * 0.16, (b2y0 + b2y1) / 2 - h * 0.05, "ПРОМ")


# =============================================================================
#  МАКЕТ ЛИСТА (альбомный, две колонки, код внизу в рамке)
# =============================================================================
PAGE_W, PAGE_H = landscape(A4)
LEFT_M, RIGHT_M, TOP_M = 14 * mm, 12 * mm, 12 * mm
GUTTER = 10 * mm
COL_W = (PAGE_W - LEFT_M - RIGHT_M - GUTTER) / 2.0
BOX_H, BOX_Y = 7 * mm, 7 * mm
AVAIL_TOP = PAGE_H - TOP_M
AVAIL_BOT = BOX_Y + BOX_H + 4 * mm
AVAIL = AVAIL_TOP - AVAIL_BOT


def plural_ru(n, forms):
    try:
        n = int(float(str(n).replace(",", ".")))
    except ValueError:
        return forms[2]
    n = abs(n) % 100
    n1 = n % 10
    if 10 < n < 20:
        return forms[2]
    if n1 == 1:
        return forms[0]
    if 2 <= n1 <= 4:
        return forms[1]
    return forms[2]


def plural_en(n, forms):
    try:
        v = float(str(n).replace(",", "."))
    except ValueError:
        return forms[1]
    return forms[0] if v == 1 else forms[1]


def esc(s):
    """Экранирование ТОЛЬКО динамических значений (не шаблонов с разметкой <u>)."""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


class Flow:
    """Поток текста в одной колонке."""

    def __init__(self, c, x, w, y, fs, dry):
        self.c, self.x, self.w, self.y, self.fs, self.dry = c, x, w, y, fs, dry

    def gap(self, d):
        self.y -= d

    def para(self, text, size=None, bold=False, align=TA_JUSTIFY, before=0, after=2, lead=None):
        size = size or self.fs
        st = ParagraphStyle("x", fontName=FONT_B if bold else FONT, fontSize=size,
                            leading=lead or size * 1.25, alignment=align, textColor=colors.black)
        p = Paragraph(text, st)
        _, h = p.wrap(self.w, 10000)
        self.y -= before
        if not self.dry:
            p.drawOn(self.c, self.x, self.y - h)
        self.y -= h + after

    def heading(self, text):
        self.para(text, size=self.fs + 1, bold=True, align=TA_LEFT, before=5, after=4)

    def uline(self, x, w, caption=None):
        if not self.dry:
            c = self.c
            c.setLineWidth(0.5)
            c.line(x, self.y, x + w, self.y)
            if caption:
                c.setFont(FONT, 6.5)
                c.drawCentredString(x + w / 2.0, self.y - 7, caption)


def render_label(c, rec, man, lang, fs=10.0, dry=False):
    T = TEXTS[lang]
    typ = rec["type"]
    label = rec["label_ru"] if lang == "ru" else rec["label_en"]
    if lang == "ru":
        life = f"{man['life']} {plural_ru(man['life'], T['years'])}"
        shelf = f"{man['shelf']} {plural_ru(man['shelf'], T['months'])}"
        warr = f"{man['warranty']} {plural_ru(man['warranty'], T['months'])}"
        contract = man["contract"]
    else:
        life = f"{man['life']} {plural_en(man['life'], T['years'])}"
        shelf = f"{man['shelf']} {plural_en(man['shelf'], T['months'])}"
        warr = f"{man['warranty']} {plural_en(man['warranty'], T['months'])}"
        contract = re.sub(r"\bот\b", "of", man["contract"], flags=re.IGNORECASE)
    contract = re.sub(r"^\s*№\s*", "", contract)
    # значения экранируются ЗДЕСЬ, один раз; шаблоны с <u> после format() не экранируются
    ctx = dict(type=esc(typ), label=esc(label), plan=esc(man["plan"]), contract=esc(contract),
               life=esc(life), shelf=esc(shelf), warranty=esc(warr))
    date_ul = f"<u>{NB*3}{esc(man['date'])}{NB*3}</u>" if man["date"] else f"<u>{NB*24}</u>"

    xL, xR = LEFT_M, LEFT_M + COL_W + GUTTER
    y0 = AVAIL_TOP
    L = Flow(c, xL, COL_W, y0, fs, dry)
    R = Flow(c, xR, COL_W, y0, fs, dry)

    # ---------- шапка левой колонки: текст | логотип | текст ----------
    logo_h = 15 * mm
    logo_cy = y0 - logo_h / 2.0 - 1 * mm
    if not dry:
        c.setFont(FONT, fs)
        c.drawString(xL, logo_cy - fs * 0.35, T["head_left"])
        c.drawRightString(xL + COL_W, logo_cy - fs * 0.35, T["head_right"])
        draw_logo(c, xL + COL_W / 2.0, logo_cy, logo_h)
    L.y = logo_cy - logo_h / 2.0 - 4 * mm

    # ---------- заголовки ----------
    L.para(T["title"].format(**ctx), size=fs + 3, bold=True, align=TA_CENTER, before=2, after=6)
    L.para(T["sub1"], size=fs + 2, bold=True, align=TA_CENTER, after=1)
    L.para(T["sub2"].format(**ctx), size=fs + 2, bold=True, align=TA_CENTER, after=6)
    L.para(f'{T["kks_label"]} <u>{NB*4}{esc(rec["kks"])}{NB*4}</u>',
           size=fs, align=TA_CENTER, after=6)

    # ---------- 1 ----------
    L.heading(T["s1"])
    L.para(T["p11"])
    L.para(f'{T["s12_pre"]} <u>{NB*5}{esc(man["safety"])}{NB*5}</u> {T["s12_post"]}',
           align=TA_LEFT)
    L.para(f'{T["s13_pre"]} <u>{NB*4}{esc(man["seismic"])}{NB*4}</u> {T["s13_post"]}',
           align=TA_LEFT)
    L.para(T["p14"])

    # ---------- 2 ----------
    L.heading(T["s2"])
    L.para(f'{T["p21"]} {date_ul}', align=TA_LEFT)
    L.para(T["p22"], align=TA_LEFT)
    L.para(T["p23"])
    L.para(T["p24"])

    # ---------- 3 + таблица ----------
    L.heading(T["s3"])
    L.para(T["tab_caption"], align=TA_LEFT, after=3)
    hst = ParagraphStyle("h", fontName=FONT_B, fontSize=fs - 0.5, leading=fs * 1.2, alignment=TA_CENTER)
    lft = ParagraphStyle("l", fontName=FONT, fontSize=fs, leading=fs * 1.2, alignment=TA_LEFT)
    ctr = ParagraphStyle("c", fontName=FONT, fontSize=fs, leading=fs * 1.2, alignment=TA_CENTER)
    qty_text = f"{rec['qty_str']} {rec['unit_ru'] if lang == 'ru' else rec['unit_en']}".strip()
    data = [
        [Paragraph(T["tab_head1"].format(**ctx), hst), Paragraph(T["tab_head2"], hst)],
        [Paragraph(esc(rec["tail"]), lft), Paragraph(esc(qty_text), ctr)],
        ["", ""], ["", ""],
    ]
    tb = Table(data, colWidths=[COL_W * 0.72, COL_W * 0.28], rowHeights=[None, None, 6 * mm, 6 * mm])
    tb.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    _, th = tb.wrap(COL_W, 10000)
    if not dry:
        tb.drawOn(c, xL, L.y - th)
    L.y -= th + 2

    # ---------- правая колонка: 4 ----------
    R.heading(T["s4"])
    R.para(T["p41"].format(**ctx))
    R.para(T["p42"].format(**ctx))
    R.para(T["p43"].format(**ctx))

    # ---------- 5 ----------
    R.heading(T["s5"])
    R.para(T["p51"].format(**ctx))
    R.para(T["p52"].format(**ctx))          # без esc(): внутри шаблона разметка <u>
    R.para(T["p53"])
    R.para(T["p54"].format(**ctx))          # без esc(): внутри шаблона разметка <u>
    R.gap(4 * mm)
    R.para(T["sign_head"], align=TA_LEFT, after=0)
    R.gap(12 * mm)                          # увеличенный отступ над "подпись"/"расшифровка подписи"
    R.uline(xR + 2 * mm, 40 * mm, T["cap_sign"])
    R.uline(xR + 48 * mm, 62 * mm, T["cap_name"])
    R.gap(12 * mm)                          # увеличенный отступ над "год, месяц, число"
    if not dry:
        c.setFont(FONT, fs)
        c.drawString(xR + 2 * mm, R.y, T["seal"])
    R.uline(xR + 25 * mm, 45 * mm, T["cap_date"])
    R.gap(11 * mm)

    # ---------- 6 ----------
    R.heading(T["s6"])
    R.para(T["p61"].format(**ctx))
    R.gap(14 * mm)                          # увеличенный отступ над "должность/подпись/расшифровка"
    R.uline(xR + 2 * mm, 35 * mm, T["cap_pos"])
    R.uline(xR + 43 * mm, 32 * mm, T["cap_sign"])
    R.uline(xR + 81 * mm, 45 * mm, T["cap_name"])
    R.gap(12 * mm)                          # увеличенный отступ над "год, месяц, число"
    R.uline(xR + 81 * mm, 45 * mm, T["cap_date"])
    R.gap(8 * mm)

    # ---------- код документа внизу в рамке (отключаемо галочкой) ----------
    code = ""
    if man.get("show_doc", True):
        code = rec["doc_ru"] if lang == "ru" else rec["doc_en"]
    if code:
        tw = pdfmetrics.stringWidth(code, "Times-Roman", fs)
        bw = tw + 14 * mm
        bx = (PAGE_W - bw) / 2.0
        if not dry:
            c.setStrokeColor(colors.black)
            c.setLineWidth(0.8)
            c.rect(bx, BOX_Y, bw, BOX_H)
            c.setFont("Times-Roman", fs)
            c.drawCentredString(PAGE_W / 2.0, BOX_Y + (BOX_H - fs) / 2.0 + 1, code)
    return y0 - min(L.y, R.y)


def make_pdf(path, records, man, lang, progress=None):
    c = canvas.Canvas(path, pagesize=landscape(A4))
    c.setTitle("Этикетки / Data Sheets " + ("RU" if lang == "ru" else "EN"))
    c.setAuthor("ELOX-PROM JSC")
    for i, rec in enumerate(records, 1):
        size = 10.0
        for fs in (10.0, 9.5, 9.0, 8.5, 8.0, 7.5, 7.0):
            size = fs
            if render_label(c, rec, man, lang, fs=fs, dry=True) <= AVAIL:
                break
        render_label(c, rec, man, lang, fs=size, dry=False)
        c.showPage()
        if progress:
            progress(i)
    c.save()


# =============================================================================
#  GUI
# =============================================================================
FIELDS = [
    ("safety",   "Класс безопасности (mhgbd)", "2", 12),
    ("seismic",  "Категория сейсмостойкости (hkn)", "I", 12),
    ("life",     "Срок службы, лет (jyxcgh)", "60", 12),
    ("shelf",    "Средний срок сохраняемости, мес. (laqbe)", "60", 12),
    ("warranty", "Гарантийный срок, мес. (mssky)", "24", 12),
    ("date",     "Дата изготовления (необязательно)", "", 20),
    ("plan",     "План качества № (ED-QP-…)", "ED-QP-ELOX-PROM-08-2026", 46),
    ("contract", "Договор № … от …", "ЭП-100/24 от 30.05.2024", 46),
]
SETTINGS_FILE = "elox_labels_settings.json"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("900x700")
        self.minsize(860, 660)
        self.records, self.xls_path = [], None
        self.vars = {}
        self._build()
        self._load_settings()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build(self):
        pad = dict(padx=6, pady=3)
        f_file = ttk.LabelFrame(self, text="1. Файл с данными (.xls)")
        f_file.pack(fill="x", padx=10, pady=(10, 4))
        row = ttk.Frame(f_file)
        row.pack(fill="x", **pad)
        ttk.Button(row, text="Выбрать файл…", command=self.choose_file).pack(side="left")
        self.lbl_file = ttk.Label(row, text="файл не выбран", foreground="#666")
        self.lbl_file.pack(side="left", padx=8)
        row2 = ttk.Frame(f_file)
        row2.pack(fill="x", **pad)
        ttk.Label(row2, text="Начальная строка данных:").pack(side="left")
        self.start_row = tk.IntVar(value=START_ROW_DEFAULT)
        ttk.Spinbox(row2, from_=1, to=1000, width=6, textvariable=self.start_row,
                    command=self.reload_rows).pack(side="left", padx=6)
        self.lbl_count = ttk.Label(row2, text="листов: 0", foreground="#0a6")
        self.lbl_count.pack(side="left", padx=12)

        f_man = ttk.LabelFrame(self, text="2. Поля, заполняемые вручную (одинаковы для всех листов)")
        f_man.pack(fill="x", padx=10, pady=4)
        for i, (key, cap, default, width) in enumerate(FIELDS):
            r, col = divmod(i, 2)
            box = ttk.Frame(f_man)
            box.grid(row=r, column=col, sticky="ew", **pad)
            f_man.columnconfigure(col, weight=1)
            ttk.Label(box, text=cap).pack(anchor="w")
            var = tk.StringVar(value=default)
            ttk.Entry(box, textvariable=var, width=width).pack(fill="x")
            self.vars[key] = var

        f_opt = ttk.Frame(self)
        f_opt.pack(fill="x", padx=10, pady=4)
        self.separate = tk.BooleanVar(value=False)
        ttk.Checkbutton(f_opt, text="Сохранять каждый лист отдельным PDF-файлом",
                        variable=self.separate).pack(side="left")
        self.show_doc = tk.BooleanVar(value=True)
        ttk.Checkbutton(f_opt, text="Отображать код документа внизу листа (рамка)",
                        variable=self.show_doc).pack(side="left", padx=(24, 0))

        f_btn = ttk.Frame(self)
        f_btn.pack(fill="x", padx=10, pady=6)
        self.btn_ru = ttk.Button(f_btn, text="Русские версии", command=lambda: self.generate("ru"))
        self.btn_ru.pack(side="left", expand=True, fill="x", padx=(0, 5), ipady=8)
        self.btn_en = ttk.Button(f_btn, text="Английские версии", command=lambda: self.generate("en"))
        self.btn_en.pack(side="left", expand=True, fill="x", padx=(5, 0), ipady=8)

        self.pb = ttk.Progressbar(self, mode="determinate")
        self.pb.pack(fill="x", padx=10, pady=(4, 2))

        f_log = ttk.LabelFrame(self, text="Журнал")
        f_log.pack(fill="both", expand=True, padx=10, pady=(2, 10))
        self.log = tk.Text(f_log, height=10, wrap="word")
        self.log.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(f_log, command=self.log.yview)
        sb.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=sb.set)

    def say(self, msg):
        self.log.insert("end", msg + "\n")
        self.log.see("end")

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="Выберите файл с данными",
            filetypes=[("Excel", "*.xls *.xlsx"), ("Все файлы", "*.*")])
        if not path:
            return
        self.xls_path = path
        self.lbl_file.config(text=os.path.basename(path), foreground="#000")
        self.reload_rows()

    def reload_rows(self):
        if not self.xls_path:
            return
        try:
            rows = load_table(self.xls_path)
            self.records, warns, skipped = parse_rows(rows, self.start_row.get())
        except Exception as e:
            messagebox.showerror("Ошибка чтения", str(e))
            self.say("ОШИБКА: " + str(e))
            return
        self.lbl_count.config(text=f"листов: {len(self.records)}")
        self.log.delete("1.0", "end")
        self.say(f"Загружено: {self.xls_path}")
        self.say(f"Строк-листов: {len(self.records)}; пропущено строк: {skipped}")
        for w in warns[:40]:
            self.say("⚠ " + w)
        if len(warns) > 40:
            self.say(f"… и ещё {len(warns)-40} предупреждений.")
        if not self.records:
            self.say("Нет подходящих строк. Проверьте номер начальной строки и столбцы E/G/H/J.")

    def manual(self):
        m = {k: v.get().strip() for k, v in self.vars.items()}
        for key in ("safety", "seismic", "life", "shelf", "warranty", "plan", "contract"):
            if not m[key]:
                raise ValueError(f"Не заполнено поле: {key}")
        m["show_doc"] = bool(self.show_doc.get())
        return m

    def generate(self, lang):
        if not self.records:
            messagebox.showwarning("Нет данных", "Сначала загрузите .xls файл.")
            return
        try:
            man = self.manual()
        except ValueError as e:
            messagebox.showwarning("Проверьте поля", str(e))
            return
        stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        suffix = "RU" if lang == "ru" else "EN"
        targets = []
        if self.separate.get():
            folder = filedialog.askdirectory(title="Папка для PDF-файлов")
            if not folder:
                return
            for rec in self.records:
                safe_tail = re.sub(r'[\\/:*?"<>|]', "_", rec["tail"])
                fname = f"{rec['kks']}_{safe_tail}.pdf"
                targets.append((os.path.join(folder, fname), [rec]))
        else:
            path = filedialog.asksaveasfilename(
                title="Сохранить PDF", defaultextension=".pdf",
                initialfile=f"Этикетки_{suffix}_{stamp}.pdf",
                filetypes=[("PDF", "*.pdf")])
            if not path:
                return
            targets.append((path, self.records))

        self._save_settings()
        self.btn_ru.state(["disabled"])
        self.btn_en.state(["disabled"])
        total = sum(len(rr) for _, rr in targets)
        self.pb.configure(value=0, maximum=total)
        self.say(f"\nГенерация ({suffix}): листов {total}, файлов {len(targets)}…")

        def work():
            done = 0
            try:
                for path, recs in targets:
                    def prog(i):
                        self.after(0, lambda v=done + i: self.pb.configure(value=v))
                    make_pdf(path, recs, man, lang, progress=prog)
                    done += len(recs)
                self.after(0, lambda: self._done(targets))
            except Exception as e:
                tb = traceback.format_exc()
                self.after(0, lambda: self._fail(str(e), tb))

        threading.Thread(target=work, daemon=True).start()

    def _done(self, targets):
        self.btn_ru.state(["!disabled"])
        self.btn_en.state(["!disabled"])
        self.pb.configure(value=self.pb["maximum"])
        for p, _ in targets[:10]:
            self.say("✔ " + p)
        if len(targets) > 10:
            self.say(f"… всего файлов: {len(targets)}")
        first = targets[0][0]
        if messagebox.askyesno("Готово", f"Создано файлов: {len(targets)}.\nОткрыть результат?"):
            try:
                os.startfile(first)  # noqa
            except Exception:
                subprocess.Popen(["explorer", os.path.dirname(first)])

    def _fail(self, msg, tb):
        self.btn_ru.state(["!disabled"])
        self.btn_en.state(["!disabled"])
        self.say("ОШИБКА: " + tb)
        messagebox.showerror("Ошибка", msg)

    def _settings_path(self):
        p = os.path.join(app_dir(), SETTINGS_FILE)
        try:
            with open(p, "a"):
                pass
            return p
        except Exception:
            return os.path.join(os.environ.get("TEMP", app_dir()), SETTINGS_FILE)

    def _save_settings(self):
        data = {k: v.get() for k, v in self.vars.items()}
        data["start_row"] = self.start_row.get()
        data["separate"] = self.separate.get()
        data["show_doc"] = self.show_doc.get()
        try:
            with open(self._settings_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_settings(self):
        try:
            with open(self._settings_path(), encoding="utf-8") as f:
                data = json.load(f)
            for k, v in self.vars.items():
                if k in data:
                    v.set(data[k])
            self.start_row.set(int(data.get("start_row", START_ROW_DEFAULT)))
            self.separate.set(bool(data.get("separate", False)))
            self.show_doc.set(bool(data.get("show_doc", True)))
        except Exception:
            pass

    def _on_close(self):
        self._save_settings()
        self.destroy()


def main():
    try:
        register_fonts()
        find_logo()
    except Exception as e:
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror("Шрифт", str(e))
        return
    App().mainloop()


if __name__ == "__main__":
    main()
