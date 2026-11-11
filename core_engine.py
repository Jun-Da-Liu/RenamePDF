#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
CORE ENGINE - PHẦN MỀM TỰ ĐỘNG NHẬN DIỆN OCR & ĐỔI TÊN HÀNG LOẠT FILE PDF
Hỗ trợ nạp động danh mục Excel (Upload / Chỉnh sửa trực tiếp), bộ nhớ đệm SQLite,
Dry-Run có callback tiến trình cho GUI, Đổi tên 2 pha an toàn và Rollback.
"""

import sys
import os
import re
import time
import json
import shutil
import sqlite3
import zipfile
import ctypes
import pathlib
import unicodedata
import concurrent.futures
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
from collections import defaultdict, Counter

import pymupdf
from rapidfuzz import fuzz

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

APP_DIR = pathlib.Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else pathlib.Path(__file__).resolve().parent
DATA_DIR = APP_DIR / 'data'
CONFIG_PATH = APP_DIR / 'config.json'
DEFAULT_EXCEL_PATH = DATA_DIR / 'Danh_muc_tai_lieu_khong_dau.xlsx'

_BUNDLED_TESS_DIR = APP_DIR / 'tesseract'
_ENV_TESS_DIR = os.environ.get('TESSERACT_DIR', '').strip()
_DEFAULT_TESS_DIR = _BUNDLED_TESS_DIR if _BUNDLED_TESS_DIR.is_dir() else pathlib.Path(_ENV_TESS_DIR) if _ENV_TESS_DIR else None
_DEFAULT_ROOT_DIR = pathlib.Path.home() / 'Documents'
_DEFAULT_REPORT_DIR = pathlib.Path.home() / 'Documents' / 'PhanMemDoiTenPDF_Reports'

DEFAULT_CONFIG = {
    'root_dir': str(_DEFAULT_ROOT_DIR),
    'excel_catalog_path': str(DEFAULT_EXCEL_PATH),
    'report_dir': str(_DEFAULT_REPORT_DIR),
    'tess_dir': str(_DEFAULT_TESS_DIR) if _DEFAULT_TESS_DIR else '',
    'tessdata_dir': str(_DEFAULT_TESS_DIR / 'tessdata') if _DEFAULT_TESS_DIR else '',
    'workers': 10,
    'allow_outside_excel': True,
}

EXCLUDED_FOLDER_NAMES = {
    '_BAO_CAO_DOI_TEN',
    'PhanMemDoiTenPDF',
    '__pycache__',
    '.git',
    '.vscode',
}

os.environ['OMP_THREAD_LIMIT'] = '1'


# ==============================================================================
# 0. QUẢN LÝ CẤU HÌNH (CONFIG)
# ==============================================================================

def load_config() -> dict:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    cfg = dict(DEFAULT_CONFIG)
    if CONFIG_PATH.exists():
        try:
            loaded = json.loads(CONFIG_PATH.read_text(encoding='utf-8'))
            if isinstance(loaded, dict):
                cfg.update(loaded)
        except Exception:
            pass
    # Đảm bảo file Excel mặc định tồn tại
    excel_p = pathlib.Path(cfg.get('excel_catalog_path', str(DEFAULT_EXCEL_PATH)))
    if not excel_p.exists():
        fallback_root = pathlib.Path(cfg['root_dir']) / 'Danh_muc_tai_lieu_khong_dau.xlsx'
        if DEFAULT_EXCEL_PATH.exists():
            cfg['excel_catalog_path'] = str(DEFAULT_EXCEL_PATH)
        elif fallback_root.exists():
            shutil.copy2(fallback_root, DEFAULT_EXCEL_PATH)
            cfg['excel_catalog_path'] = str(DEFAULT_EXCEL_PATH)
    return cfg


def save_config(cfg: dict):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')


# ==============================================================================
# 1. CHUẨN HÓA TIẾNG VIỆT & TIỆN ÍCH XỬ LÝ VĂN BẢN
# ==============================================================================

def remove_accents(s: str) -> str:
    if not s:
        return ""
    s = s.replace('đ', 'd').replace('Đ', 'D')
    s = unicodedata.normalize('NFD', s)
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn')


def norm_text(s: str) -> str:
    s = remove_accents(s).lower()
    s = re.sub(r'[^a-z0-9\s]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def sanitize_windows_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '-', name)
    name = re.sub(r'\s+', ' ', name).strip()
    name = name.rstrip('. ')
    if len(name) > 180:
        name = name[:180].rstrip('. ')
    return name


def natural_sort_key(path_obj: pathlib.Path):
    name = path_obj.name.lower()
    return [int(text) if text.isdigit() else text for text in re.split(r'(\d+)', name)]


def get_dept_and_person(rel_dir: pathlib.Path):
    parts = rel_dir.parts
    if not parts or parts == ('.',):
        return "(Thư mục gốc)", "(Thư mục gốc)"
    dept = parts[0]
    if len(parts) == 1:
        person = parts[0]
    elif parts[-1].lower() in ('tài liệu khác', 'hồ sơ khác', 'tai lieu khac', 'ho so khac') and len(parts) >= 3:
        person = f"{parts[-2]} ({parts[-1]})"
    elif len(parts) >= 3 and parts[-1] == parts[-2]:
        person = parts[-1]
    else:
        person = " - ".join(parts[1:])
    return dept, person


# ==============================================================================
# 2. ĐỌC & GHI FILE EXCEL (.XLSX) CHUẨN OPENXML (LINH HOẠT MỌI ĐỊNH DẠNG CỘT)
# ==============================================================================

def _col_ref_to_index(cell_ref: str) -> int:
    """Chuyển đổi địa chỉ ô Excel (ví dụ 'A1', 'BC23') thành chỉ số cột 0-based."""
    m = re.match(r'^([A-Za-z]+)', cell_ref or '')
    if not m:
        return 0
    col_str = m.group(1).upper()
    idx = 0
    for ch in col_str:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def read_raw_xlsx_rows(xlsx_path: pathlib.Path) -> list:
    """Đọc toàn bộ các dòng từ sheet đầu tiên của file .xlsx (hỗ trợ cả sharedStrings và inlineStr)."""
    xlsx_path = pathlib.Path(xlsx_path)
    if not xlsx_path.exists():
        raise FileNotFoundError(f"Không tìm thấy file Excel: {xlsx_path}")

    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        namelist = set(z.namelist())
        sst = []
        if 'xl/sharedStrings.xml' in namelist:
            sst_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in sst_root.findall(ns + 'si'):
                texts = [t.text or '' for t in si.iter(ns + 't')]
                sst.append(''.join(texts))

        # Tìm sheet đầu tiên
        sheet_name = 'xl/worksheets/sheet1.xml'
        if sheet_name not in namelist:
            sheets = sorted(n for n in namelist if n.startswith('xl/worksheets/sheet') and n.endswith('.xml'))
            if not sheets:
                raise ValueError("File Excel không chứa worksheet hợp lệ.")
            sheet_name = sheets[0]

        sheet = ET.fromstring(z.read(sheet_name))
        rows = []
        for r in sheet.iter(ns + 'row'):
            row_dict = {}
            max_col = -1
            next_col = 0
            for c in r.findall(ns + 'c'):
                ref = c.get('r')
                c_idx = _col_ref_to_index(ref) if ref else next_col
                next_col = c_idx + 1
                if c_idx > max_col:
                    max_col = c_idx

                cell_type = c.get('t')
                val = ''
                if cell_type == 's':
                    v = c.find(ns + 'v')
                    if v is not None and v.text is not None:
                        s_idx = int(v.text)
                        val = sst[s_idx] if 0 <= s_idx < len(sst) else ''
                elif cell_type == 'inlineStr':
                    is_el = c.find(ns + 'is')
                    if is_el is not None:
                        val = ''.join(t.text or '' for t in is_el.iter(ns + 't'))
                else:
                    v = c.find(ns + 'v')
                    if v is not None and v.text is not None:
                        val = v.text
                row_dict[c_idx] = val.strip()

            if max_col >= 0:
                row_vals = [row_dict.get(i, '') for i in range(max_col + 1)]
                if any(x for x in row_vals):
                    rows.append(row_vals)
        return rows


def load_excel_catalog(xlsx_path: pathlib.Path) -> dict:
    """
    Đọc danh mục đặt tên từ file Excel (.xlsx).
    Hỗ trợ linh hoạt:
      - Cấu trúc 2 cột chuẩn: [Tên đặt File không dấu, Tên tài liệu có dấu, (Từ khóa tùy chọn)]
      - Cấu trúc có cột STT ở đầu: [STT, Tên đặt File không dấu, Tên tài liệu có dấu, (Từ khóa)]
      - Cấu trúc 1 cột: [Tên tài liệu] -> tự động sinh tên không dấu an toàn
    """
    rows = read_raw_xlsx_rows(xlsx_path)
    if not rows:
        return {}

    # Kiểm tra dòng đầu tiên có phải là tiêu đề (header) không
    first_row_norm = ' '.join(norm_text(c) for c in rows[0])
    start_idx = 1
    if any(h in first_row_norm for h in ['ten dat', 'ten tai lieu', 'ten file', 'stt', 'danh muc', 'ma tai lieu', 'khong dau']):
        start_idx = 1
    else:
        # Nếu dòng 1 bắt đầu bằng "01." hoặc số 1 thì có thể file không có dòng tiêu đề
        r0_c0 = rows[0][0].strip() if rows[0] else ''
        if re.match(r'^0?1(\.|\s|$)', r0_c0):
            start_idx = 0

    data_rows = rows[start_idx:]
    if not data_rows:
        return {}

    # Phát hiện xem cột 0 có phải là cột STT thuần túy (1, 2, 3...) không
    col0_is_stt = False
    if len(data_rows[0]) >= 2:
        sample = data_rows[:min(10, len(data_rows))]
        if all(re.match(r'^\d{1,4}$', r[0].strip()) and len(r[1].strip()) > 3 for r in sample if len(r) >= 2 and r[0].strip()):
            col0_is_stt = True

    catalog = {}
    for row_num, r in enumerate(data_rows, start=1):
        if col0_is_stt:
            stt_str = r[0].strip()
            col_safe = r[1].strip() if len(r) > 1 else ''
            col_orig = r[2].strip() if len(r) > 2 else ''
            col_kw = r[3].strip() if len(r) > 3 else ''
        else:
            stt_str = ''
            col_safe = r[0].strip() if len(r) > 0 else ''
            col_orig = r[1].strip() if len(r) > 1 else ''
            col_kw = r[2].strip() if len(r) > 2 else ''

        if not col_safe and not col_orig:
            continue

        if not col_safe and col_orig:
            col_safe = remove_accents(col_orig)
        if not col_orig:
            col_orig = col_safe

        safe_name = sanitize_windows_filename(col_safe)
        orig_name = col_orig

        # Xác định ID của mục (ưu tiên số ở đầu tên file như "04. ..." hoặc cột STT, nếu trùng thì dùng row_num)
        item_id = row_num
        m_num = re.match(r'^(\d{1,4})\.\s*(.+)$', safe_name)
        if m_num:
            parsed_id = int(m_num.group(1))
            clean_title = m_num.group(2).strip()
            if parsed_id not in catalog:
                item_id = parsed_id
        else:
            clean_title = safe_name
            if stt_str.isdigit():
                parsed_id = int(stt_str)
                if parsed_id not in catalog:
                    item_id = parsed_id

        while item_id in catalog:
            item_id += 1

        kw_list = [norm_text(k) for k in re.split(r'[,;|]', col_kw) if norm_text(k)]

        catalog[item_id] = {
            'id': item_id,
            'safe_name': safe_name,
            'orig_name': orig_name,
            'keywords': kw_list,
            'keywords_raw': col_kw,
            'norm_title': norm_text(clean_title),
            'compact_title': norm_text(clean_title).replace(' ', '')
        }

    return catalog


def write_xlsx(filepath: pathlib.Path, sheets: list):
    """Ghi danh sách các sheet [(sheet_name, headers, rows)] ra file .xlsx."""
    filepath = pathlib.Path(filepath)
    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for i in range(len(sheets)):
        content_types.append(f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    content_types.append('</Types>')

    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )

    wb_xml = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets>'
    ]
    wb_rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    ]
    for i, (sname, _, _) in enumerate(sheets):
        wb_xml.append(f'<sheet name="{escape(sname)}" sheetId="{i+1}" r:id="rId{i+1}"/>')
        wb_rels.append(f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>')
    wb_xml.append('</sheets></workbook>')
    wb_rels.append('</Relationships>')

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Times New Roman"/></font>'
        '<font><b/><sz val="11"/><name val="Times New Roman"/></font>'
        '</fonts>'
        '<fills count="3">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/></patternFill></fill>'
        '</fills>'
        '<borders count="2">'
        '<border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border>'
        '<left style="thin"><color auto="1"/></left>'
        '<right style="thin"><color auto="1"/></right>'
        '<top style="thin"><color auto="1"/></top>'
        '<bottom style="thin"><color auto="1"/></bottom>'
        '<diagonal/>'
        '</border>'
        '</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>'
        '</cellXfs>'
        '</styleSheet>'
    )

    def col_letter(n):
        res = ""
        while n >= 0:
            res = chr(65 + (n % 26)) + res
            n = n // 26 - 1
        return res

    def clean_xml_str(val):
        s = str(val) if val is not None else ""
        s = "".join(ch for ch in s if ch in ('\t', '\n', '\r') or (0x20 <= ord(ch) <= 0xD7FF) or (0xE000 <= ord(ch) <= 0xFFFD))
        return escape(s)

    filepath.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', ''.join(content_types))
        zf.writestr('_rels/.rels', rels)
        zf.writestr('xl/workbook.xml', ''.join(wb_xml))
        zf.writestr('xl/_rels/workbook.xml.rels', ''.join(wb_rels))
        zf.writestr('xl/styles.xml', styles_xml)

        for idx, (sname, headers, rows) in enumerate(sheets):
            lines = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
                '<sheetData>'
            ]
            h_cells = []
            for c_idx, h in enumerate(headers):
                ref = f"{col_letter(c_idx)}1"
                h_cells.append(f'<c r="{ref}" t="inlineStr" s="1"><is><t>{clean_xml_str(h)}</t></is></c>')
            lines.append(f'<row r="1">{"".join(h_cells)}</row>')

            for r_idx, row in enumerate(rows, start=2):
                r_cells = []
                for c_idx, val in enumerate(row):
                    ref = f"{col_letter(c_idx)}{r_idx}"
                    if isinstance(val, (int, float)) and not isinstance(val, bool):
                        r_cells.append(f'<c r="{ref}" s="2"><v>{val}</v></c>')
                    else:
                        r_cells.append(f'<c r="{ref}" t="inlineStr" s="2"><is><t>{clean_xml_str(val)}</t></is></c>')
                lines.append(f'<row r="{r_idx}">{"".join(r_cells)}</row>')

            lines.append('</sheetData></worksheet>')
            zf.writestr(f'xl/worksheets/sheet{idx+1}.xml', ''.join(lines))


def save_catalog_to_xlsx(catalog_dict: dict, filepath: pathlib.Path):
    """Lưu từ điển catalog hiện tại ra file Excel chuẩn."""
    headers = [
        "Tên đặt File/Folder an toàn (Bỏ dấu '/')",
        "Tên tài liệu (Có dấu gốc)",
        "Từ khóa nhận diện bổ sung (Tùy chọn, cách nhau dấu phẩy)"
    ]
    rows = []
    for cid in sorted(catalog_dict.keys()):
        item = catalog_dict[cid]
        rows.append([
            item.get('safe_name', ''),
            item.get('orig_name', ''),
            item.get('keywords_raw', '')
        ])
    write_xlsx(filepath, [('Danh_muc_tai_lieu', headers, rows)])


# ==============================================================================
# 3. NHẬN DẠNG VÀ PHÂN LOẠI TÀI LIỆU (OCR + DYNAMIC CATALOG + FUZZY MATCHING)
# ==============================================================================

# Biến toàn cục trong từng tiến trình (cập nhật khi nạp file Excel mới)
CATALOG = {}


def reload_global_catalog(excel_path: pathlib.Path = None) -> dict:
    global CATALOG
    if excel_path is None:
        cfg = load_config()
        excel_path = pathlib.Path(cfg['excel_catalog_path'])
    excel_path = pathlib.Path(excel_path)
    if excel_path.exists():
        CATALOG = load_excel_catalog(excel_path)
    return CATALOG


def parse_existing_stem(orig_stem: str, catalog: dict = None):
    if catalog is None:
        catalog = CATALOG
    s = orig_stem.strip()
    if re.match(r'^\d+(_\d+)?$', s):
        return None, None

    m = re.match(r'^(\d{1,3})[\.,\s_-]*(.+)$', s)
    if m:
        num = int(m.group(1))
        rest = m.group(2)
        rest_clean = re.sub(r'(\s*nam\s*[\d\.]+|\d{4}|\.\d+|\s*\(\d+\))+$', '', rest, flags=re.I).strip()
        n_rest = norm_text(rest_clean).replace(' ', '')
        if 'phieudangvien' in n_rest and 'bosung' not in n_rest and 3 in catalog:
            return 3, None
        if ('doithedang' in n_rest or 'doilaithedang' in n_rest) and 73 in catalog:
            return 73, None
        if 'caplaithedang' in n_rest and 72 in catalog:
            return 72, None
        if 'tanghuyhieudang' in n_rest and 'truy' not in n_rest and num == 103 and 102 in catalog:
            return 102, None
        if ('nhanthucvedang' in n_rest or 'nhanthucvaodang' in n_rest or 'nhanthucdang' in n_rest) and 38 in catalog:
            return 38, None
        if ('bantukiemdiemhang' in n_rest or 'bankiemdiemcanhan' in n_rest or 'bankiemdiemnam' in n_rest or 'bankiemrdiemnam' in n_rest) and 61 in catalog:
            return 61, None
        if 'nghi' in n_rest and 'congnhandangvienchinhthuc' in n_rest:
            if 'chibo' in n_rest and 50 in catalog:
                return 50, None
            if 'danguy' in n_rest and 52 in catalog:
                return 52, None
        if 'quyetdinhcongnhandangvienchinhthuc' in n_rest and 7 in catalog:
            return 7, None
        if 'quyetdinhketnapdangvien' in n_rest and 5 in catalog:
            return 5, None
        if ('bannhanxet' in n_rest or 'cuanguoigiupdo' in n_rest or 'phanconggiupdo' in n_rest) and 48 in catalog:
            return 48, None
        if 'nghiquyetketnapdangviencuadanguy' in n_rest and 45 in catalog:
            return 45, None
        if num in catalog:
            return num, None

    n_s = norm_text(s)
    n_comp = n_s.replace(' ', '')
    if 'bienbangapdangvientruockhithihanhkiluat' in n_comp or 'bienbangapdangvientruockhithihanhkyluat' in n_comp:
        return None, "Bien ban gap dang vien truoc khi thi hanh ky luat"
    if 'bienbancongbovagiaoquyetdinhkiluat' in n_comp or 'bienbancongbovagiaoquyetdinhkyluat' in n_comp:
        return None, "Bien ban cong bo va giao quyet dinh ky luat"
    if 'phieubieuquyet' in n_comp:
        return None, "Phieu bieu quyet"
    if 'giaitrinhbanthan' in n_comp or 'bangiaitrinh' in n_comp:
        return None, "Ban giai trinh"
    if ('lylichdangvien' in n_comp or 'soyeulylich' in n_comp or 'soyeull' in n_comp or n_comp == 'lylich') and 2 in catalog:
        return 2, None
    if ('lilichcuanguoixinvaodang' in n_comp or 'lylichcuanguoixinvaodang' in n_comp or 'bankhailylich' in n_comp) and 1 in catalog:
        return 1, None
    if 'phieudangvien' in n_comp and 'bosung' not in n_comp and 3 in catalog:
        return 3, None
    if 'phieubosung' in n_comp and 4 in catalog:
        return 4, None
    if ('ketnapdang' in n_comp or 'ketnapdv' in n_comp) and ('quyetdinh' in n_comp or 'qd' in n_comp or 'quyetnghi' in n_comp) and 5 in catalog:
        return 5, None
    if ('congnhandangvien' in n_comp or 'congnhandv' in n_comp or 'dangvienchinhthuc' in n_comp or 'dvchinhthuc' in n_comp or 'chuyendangchinhthuc' in n_comp or 'dubisangchinhthuc' in n_comp or 'dubithanhdangvienchinhthuc' in n_comp) and ('quyetdinh' in n_comp or 'qd' in n_comp or 'quyetnghi' in n_comp) and 7 in catalog:
        return 7, None
    if ('phatthedang' in n_comp or 'phatthedv' in n_comp) and 19 in catalog:
        return 19, None
    if ('doithedang' in n_comp or 'doilaithedang' in n_comp) and 73 in catalog:
        return 73, None
    if ('llct' in n_comp or 'lyluanchinhtri' in n_comp or 'chungnhancaplu' in n_comp or 'chungnhanclct' in n_comp) and 70 in catalog:
        return 70, None
    if ('camtinhdang' in n_comp or 'nhanthucvedang' in n_comp) and 38 in catalog:
        return 38, None
    if any(k in n_s for k in ['bang ', 'bang tn', 'bang dh', 'bang cd', 'bang cap', 'bang cu nhan', 'chung chi', 'chung nhan', 'tieng mong', 'tieng tay', 'boi duong', 'qpan', 'kt qpan', 'hoc vien', 'cu nhan', 'dai tao', 'tn dh', 'hvct', 'sqlq1']):
        if 'nhanthuc' not in n_comp and 'dangvienmoi' not in n_comp and 'camtinhdang' not in n_comp and 'bangtotnghieptrunghoc' not in n_comp and 86 in catalog:
            return 86, None
    if any(k in n_s for k in ['dieu dong', 'dieudong', 'bo nhiem', 'bonhiem', 'thang quan ham', 'thangquanham', 'phong quan ham', 'phongquanham', 'nang bac luong', 'nangbacluong', 'nang luong', 'nangluong', 'chuyen che do', 'chuyenchedo', 'xep luong', 'xepluong', 'xep lung', 'chuyen lung', 'phien qh', 'thieu uy', 'trung uy', 'thuong uy', 'thung uy', 'dai uy', 'thieu ta', 'thieus ta', 'trung ta', 'thuong ta', 'thung ta', 'chuan uy', 'tlkh', 'tltc', 'tlhc', 'tlkt', 'tltuh', 'tlcab', 'tlct', 'tlbtm', 'tltm', 'tltoc', 'tltn', 'tldv', 'tldqtv', 'tlql', 'tlqk', 'tlpb', 'tlhh', 'tlbvdv', 'tltcldt', 'tltlldtl', 'cnct', 'cnhckt', 'ctv', 'ctvp', 'pdt', 'dt doi', 'dt dd', 'dt d ', 'dp d', 'bt ', 'ct ', 'pct ', 'bp c', 'b3 c2', 'gd nlt', 'pgd nlt', 'ptbtmkh', 'tbtc', 'tbtmkh', 'tbtuh', 've qk', 've 338', 've d338', 've f3', 'f3 ve', 'tu sqct', 'tu ls', 'tang cuong', 'kien toan', 'kientoan', 'chi uy vien', 'cuv ', 'duv ', 'btddu', 'pbtcb', 'pbtdu', 'bch dang bo', 'dang uy vien', 'ubktdu', 'uv bch', 'doi trung', 'doi truong', 'pho dt', 'gv khoa', 'ctp loc binh', 'sq hoat dong', 'truong qd', 'qk5', 'tuyendung']) and 87 in catalog:
        return 87, None
    if 'mucluctailieutronghosodangvien' in n_comp:
        return None, "Muc luc tai lieu trong ho so dang vien"
    if 'trichbienbansinhhoatchidoan' in n_comp or 'bienbansinhhoatchidoan' in n_comp or 'bienbanhoinghichidoan' in n_comp or 'bienbanhopchidoan' in n_comp or 'sinhhoatchidoan' in n_comp or 'sinhhoatbchdoan' in n_comp or 'sinhhoatbchchidoan' in n_comp or 'sinhhoatlienchidoan' in n_comp or 'sinhoatbchdoan' in n_comp or 'sinhhoatbanchaphanhdoan' in n_comp or 'sinhhoatbanchaphanhchidoan' in n_comp or 'sinhhoatdoancoso' in n_comp:
        return None, "Trich bien ban sinh hoat chi doan"
    if 'bienbansinhhoatchibo' in n_comp or 'bienbanhopchibo' in n_comp or 'bienbanhoinghichibo' in n_comp or 'trichbbhoinghichibo' in n_comp:
        return None, "Trich bien ban sinh hoat chi bo"
    if 'bienbansinhhoattodang' in n_comp or 'bienbantodang' in n_comp or 'sinhhoattodang' in n_comp:
        return None, "Bien ban sinh hoat to dang"
    if 'giaykhaisinh' in n_comp and 75 in catalog:
        return 75, None
    if 'bangtotnghieptrunghoc' in n_comp:
        return None, "Bang tot nghiep trung hoc"
    if ('giaybaodam' in n_comp or 'giayambao' in n_comp or 'giaydambao' in n_comp) and 40 in catalog:
        return 40, None
    if 'bantukiemdiem' in n_comp or 'bankiemdiem' in n_comp or 'tunhanxetdangvien' in n_comp or 'tunnhanxetdangvien' in n_comp or 'bantunhanxetdangvien' in n_comp:
        if ('dubi' in n_comp or 'chinhthuc' in n_comp) and 47 in catalog:
            return 47, None
        if ('chuyensinhhoat' in n_comp or 'gioithieusinhhoat' in n_comp) and 62 in catalog:
            return 62, None
        if 61 in catalog:
            return 61, None
    if 'bannhanxet' in n_comp or 'nhanetketqua' in n_comp:
        if ('dubi' in n_comp or 'phandau' in n_comp) and 48 in catalog:
            return 48, None
    if 'ykiennhanxet' in n_comp or 'ykiencuatochuc' in n_comp or 'tonghopykien' in n_comp:
        if 'dubi' in n_comp and 49 in catalog:
            return 49, None
        if 'vaodang' in n_comp and 42 in catalog:
            return 42, None
    if 'nghi' in n_comp and 'ketnap' in n_comp:
        if 'danguy' in n_comp and 45 in catalog:
            return 45, None
        if 43 in catalog:
            return 43, None
    if 'nghi' in n_comp and ('congnhan' in n_comp or 'chinhthuc' in n_comp):
        if 'danguy' in n_comp and 52 in catalog:
            return 52, None
        if 50 in catalog:
            return 50, None
    if 'phieudanhgia' in n_comp:
        return None, "Phieu danh gia xep loai dang vien"
    if 'banxacminhlylich' in n_comp:
        return None, "Ban xac minh ly lich"
    if 'bancamket' in n_comp:
        return None, "Ban cam ket tu duong ren luyen phan dau"
    return None, None


def classify_text(ocr_p1: str, orig_stem: str = "", ocr_extra: str = "", catalog: dict = None, allow_outside: bool = True):
    """
    Đối chiếu văn bản OCR với danh mục tài liệu trong Excel (hỗ trợ cả danh mục chuẩn 103 mục
    và bất kỳ danh mục Excel tùy chỉnh nào do người dùng tải lên).
    Trả về: (status, excel_id, excel_name, proposed_base, confidence, reason)
    """
    if catalog is None:
        catalog = CATALOG
    if not catalog:
        catalog = reload_global_catalog()

    n_p1 = norm_text(ocr_p1)
    n_ex = norm_text(ocr_extra)
    n_full = (n_p1 + " " + n_ex).strip()
    hint_id, hint_out = parse_existing_stem(orig_stem, catalog)

    def match_excel(cid, conf, reason=""):
        if cid in catalog:
            item = catalog[cid]
            return ('EXCEL', cid, item['safe_name'], item['safe_name'], conf, reason)
        # Nếu người dùng đã tải lên file Excel mới thay đổi STT, tìm theo tiêu đề tương đương
        n_reason = norm_text(reason)
        best_cid = None
        best_sc = 0.0
        for k_id, it in catalog.items():
            sc = fuzz.token_set_ratio(it['norm_title'], n_reason)
            if sc > best_sc:
                best_sc = sc
                best_cid = k_id
        if best_cid is not None and best_sc >= 82.0:
            item = catalog[best_cid]
            return ('EXCEL', best_cid, item['safe_name'], item['safe_name'], conf, reason)
        if allow_outside and reason:
            clean_r = re.sub(r'^mau\s*[\w\-]+\s*/\s*', '', reason, flags=re.I).strip()
            clean_t = sanitize_windows_filename(remove_accents(clean_r))
            return ('OUTSIDE', None, '', clean_t, conf, f"{reason} (Ngoài Excel hiện tại)")
        return ('CHECK', None, '', '', 0.0, f"Không có mã {cid} trong danh mục Excel hiện tại")

    def match_outside(title_no_accent, conf, reason="Ngoài danh mục Excel"):
        # Kiểm tra xem mục "Ngoài Excel" này có vừa được người dùng thêm vào file Excel mới không!
        n_target = norm_text(title_no_accent)
        for cid, item in catalog.items():
            if item['norm_title'] == n_target or (len(item['norm_title']) >= 12 and item['norm_title'] in n_target):
                return ('EXCEL', cid, item['safe_name'], item['safe_name'], conf, f"Khớp mục mới trong Excel ({item['safe_name']})")
        if not allow_outside:
            return ('CHECK', None, '', '', conf, f"Ngoài danh mục Excel: {reason}")
        clean_t = sanitize_windows_filename(remove_accents(title_no_accent))
        return ('OUTSIDE', None, '', clean_t, conf, reason)

    def need_check(cand_id, conf, reason="Không xác định chắc chắn"):
        item = catalog.get(cand_id) if cand_id is not None else None
        ename = item['safe_name'] if item else ''
        return ('CHECK', cand_id, ename, ename, conf, reason)

    # 0. Ưu tiên kiểm tra từ khóa tùy chỉnh (cột thứ 3 trong Excel nếu người dùng thêm)
    if n_full:
        for cid, item in catalog.items():
            for kw in item.get('keywords', []):
                if len(kw) >= 4 and kw in n_full:
                    return ('EXCEL', cid, item['safe_name'], item['safe_name'], 98.0, f"Khớp từ khóa Excel: '{kw}'")

    # 1. Mục lục tài liệu trong hồ sơ đảng viên (Mẫu 12-HSĐV)
    if re.search(r'\b(mau\s*12\s*hs\s*dv|mau\s*12\s*hsdv|muc\s*luc\s*tai\s*(lieu\s*|u\s*h5\s*)?trong\s*ho\s*so\s*dang\s*vien|danh\s*muc\s*ho\s*so\s*dang\s*vien|minliin\s*tai\s*he\s*twodorn\s*ho\s*so)\b', n_full):
        return match_outside("Muc luc tai lieu trong ho so dang vien", 96.0, "Mẫu 12-HSĐV / Mục lục hồ sơ đảng viên")

    # 2. 04. Phiếu bổ sung hồ sơ đảng viên (Mẫu 3-HSĐV / Mẫu số 06-HD)
    if re.search(r'\b(mau\s*3\s*hs\s*dv|mau\s*3\s*hsdv|mau\s*so\s*0?6\s*hd|phieu\s*bo\s*sung|[tu]h?ieu\s*bo\s*sung|u+\s*bo\s*sung|sung\s*ho\s*so|bo\s*sung\s*ho\s*so\s*(can\s*bo\s*)?dang\s*vien|ho\s*so\s*can\s*bo\s*dang\s*vien\s*so\s*tdv|chi\s*gh[ij]\s*n?hung\s*muc\s*co\s*noi\s*dung|ghi\s*ro\s*nhung\s*noi\s*dung\s*thay\s*doi\s*so\s*voi\s*nam\s*truoc|nhung\s*thay\s*doi\s*so\s*voi|moi\s*thay\s*doi\s*ve\s*nghe\s*nghiep\s*don\s*vi\s*cong\s*tac)\b', n_full):
        return match_excel(4, 98.0, "Mẫu 3-HSĐV / Phiếu bổ sung hồ sơ đảng viên")
    if re.search(r'\b(so\s*ly\s*lich.*so\s*tdv|dang\s*bo\s*quan\s*doi.*so\s*tdv)\b', n_full) and not re.search(r'\b(mau\s*[12]\s*hs|phieu\s*dang\s*vien|ly\s*lich\s*dang\s*vien|giay\s*gioi\s*thieu)\b', n_full):
        return match_excel(4, 96.0, "Mẫu 3-HSĐV / Phiếu bổ sung hồ sơ đảng viên")

    # 3. 03. Phiếu đảng viên (Mẫu 2-HSĐV)
    if re.search(r'\b(giay\s*chung\s*nhan\s*.*phieu\s*dang\s*vien)\b', n_full):
        return match_outside("Giay chung nhan chua co phieu dang vien", 95.0, "Giấy chứng nhận chưa có phiếu đảng viên (Ngoài Excel)")
    if re.search(r'\b(mau\s*2\s*hs\s*[db]v|mau\s*2\s*hs[db]v|so\s*ly\s*lich.*so\s*m?[ae]\s*dang\s*vien.*phieu)\b', n_full) or (re.search(r'\bphieu\s*dang\s*vien\b', n_full) and not re.search(r'\b(bo\s*sung|nhan\s*xet|danh\s*gia|xep\s*loai|bao\s*dang\s*vien|cu|chua\s*co|xin\s*y\s*kien|giay\s*chung\s*nhan)\b', n_full)):
        if re.search(r'\b(so\s*ly\s*lich|so\s*the\s*dang\s*vien|so\s*m[ae]\s*dang\s*vien|ho\s*va\s*ten\s*khai\s*sinh|mau\s*2)\b', n_full):
            return match_excel(3, 97.0, "Mẫu 2-HSĐV / Phiếu đảng viên")

    # Bản thẩm tra xác minh lý lịch & Phiếu chứng nhận thẩm tra lý lịch (94)
    if re.search(r'\b(phieu\s*chung\s*nhan\s*(tham\s*tra\s*)?l[yi]?\s*lich\s*cua\s*nguoi\s*(xin\s*)?vao\s*dang)\b', n_full):
        return match_excel(94, 96.0, "Phiếu chứng nhận thẩm tra lý lịch của người vào Đảng")
    if re.search(r'\b(ban\s*tuyen\s*sinh\s*quan\s*su|ban\s*tham\s*tra\s*xac\s*minh\s*ly\s*lich|ban\s*xac\s*[mn]inh\s*l[yi]\s*(l[ij]ch|h6n)|ban\s*bao\s*ve\s*so\s*yeu\s*ly\s*lich|dung\s*cho\s*thi\s*sinh\s*du\s*tuyen)\b', n_full):
        return match_outside("Ban tham tra xac minh ly lich", 95.0, "Bản thẩm tra xác minh lý lịch (Ngoài Excel)")

    # 4. 01. Lý lịch của người vào Đảng (Mẫu 2-KNĐ / M2-PTĐ)
    if re.search(r'\b(mau\s*2\s*kn\s*d|mau\s*2\s*knd|mau\s*2\s*kn|m2\s*p\s*t\s*d|m2\s*ptd|ly\s*lich\s*[cg]ua\s*nguoi\s*(xin\s*)?vao\s*dang|[cg]ua\s*nguoi\s*xin\s*vao\s*dang.*ho\s*(va\s*)?[ts]en\s*(dang\s*dung|khai\s*sinh))\b', n_full):
        if not re.search(r'\b(don\s*x[iy]n\s*vao\s*[dq]?b?ang|y\s*kien\s*nhan\s*xet|tham\s*tra|giay\s*gioi\s*thieu|nghi\s*quyet|quyet\s*dinh|nhan\s*xet)\b', n_full[:220]):
            return match_excel(1, 97.0, "Mẫu 2-KNĐ / Lý lịch của người vào Đảng")
    if re.search(r'\b(ban\s*[rk]hai\s*ly\s*lich)\b', n_full) and (re.search(r'\b(nguoi\s*xet\s*ket\s*nap|xin\s*vao\s*dang)\b', n_full) or hint_id == 1):
        return match_excel(1, 96.0, "Bản khai lý lịch của người vào Đảng")

    # 5. 02. Lý lịch đảng viên (Mẫu 1-HSĐV / Sơ yếu lý lịch / Sơ lược lý lịch / Bản khai lý lịch)
    if re.search(r'\b(mau\s*1\s*hs\s*dv|mau\s*1\s*hsdv|mau\s*1\s*115dv|so\s*yeu\s*ly\s*lich\s*dang\s*(vien|vin)|so\s*yeu\s*ly\s*1\s*1\s*hang\s*vien)\b', n_full) or (re.search(r'\bly\s*lich\s*([r\s]*[dtvx]?ang\s*vien)\b', n_full) and not re.search(r'\b(xin\s*vao\s*dang|nguoi\s*vao\s*dang|muc\s*luc|danh\s*muc|bo\s*sung|tham\s*tra|bao\s*cao|mat\s*the)\b', n_full)):
        return match_excel(2, 96.0, "Mẫu 1-HSĐV / Lý lịch đảng viên")

    if re.search(r'\b(so\s*luoc\s*ly\s*lich|so\s*yeu\s*ly\s*lich|yeu\s*ly\s*lich\s*tu\s*thuat|ban\s*[rk]hai\s*ly\s*(lich|uch)|ban\s*tom\s*tat\s*ly\s*lich)\b', n_full):
        if re.search(r'\b(ho\s*(va\s*)?ten\s*(dang\s*dung|khai\s*sinh)|tu\s*thuat|nghe\s*nghiep|noi\s*dang\s*ky\s*ho\s*khau|ban\s*tom\s*tat\s*ly\s*lich)\b', n_full):
            return match_excel(2, 94.0, "Sơ yếu / Sơ lược lý lịch đảng viên")

    if re.search(r'\b(bo\s*sung\s*ly\s*lich\s*gia\s*dinh)\b', n_full):
        return match_outside("Bo sung ly lich gia dinh", 94.0, "Bổ sung lý lịch gia đình (Ngoài Excel)")

    # 6. 37. Đơn xin vào Đảng (Mẫu 1-KNĐ)
    if re.search(r'\b(mau\s*1\s*[kx]nd|don\s*x[iy]n\s*vao\s*[dq]?b?ang|don\s*vn\s*vao\s*dang)\b', n_full):
        if not re.search(r'\b(xet\s*don\s*xin\s*vao\s*dang|quyet\s*dinh|nghi\s*quyet|ket\s*nap\s*dang\s*vien|ke\s*nap\s*dang\s*yien|mau\s*9|miu\s*9|qd\s*du|qd\s*tv|ban\s*nhan\s*xet|dang\s*vien\s*du\s*bi|bao\s*cao\s*cua\s*bi\s*thu)\b', n_full):
            return match_excel(37, 97.0, "Mẫu 1-KNĐ / Đơn xin vào Đảng")

    # 7. 38 vs 46: Giấy chứng nhận học lớp nhận thức về Đảng (38) / lớp đảng viên mới (46)
    if re.search(r'\b(lop\s*(boi\s*duong\s*)?(nhan\s*thuc\s*ve\s*dang|ket\s*nap\s*dang|cam\s*tinh\s*dang|doi\s*tuong\s*(phat\s*t[rh]ien\s*)?dang)|boi\s*duong\s*(kien\s*thuc\s*|nhan\s*thuc\s*|doi\s*tuong\s*phat\s*t[rh]ien\s*)?(ve\s*)?dang|h[oe]c\s*l[o6]p\s*nhan\s*thuc\s*v[e]?d[a]ng|lop\s*hoc\s*nhan\s*thuc\s*ve\s*dang|hoc\s*(lop\s*)?doi\s*tuong\s*dang|(giay\s*chung\s*nhan|cong\s*nhan)\s*doi\s*tuong\s*dang|mau\s*cn\s*ntvd)\b', n_full):
        return match_excel(38, 96.0, "Giấy chứng nhận học lớp nhận thức về Đảng")

    if re.search(r'\b(lop\s*(boi\s*duong\s*)?dang\s*vien\s*moi|boi\s*duong\s*ly\s*luan\s*chinh\s*tri\s*danh\s*cho\s*dang\s*vien\s*moi|h[oaeg][ce]?\s*(l[o6e]p\s*|giang\s*ke\s*)?d[a]ng\s*(vien\s*|vam\s*)?m[o0g][il]|mau\s*(16\s*|cn\s*)?dvm|m[ra]c?o?ndvm[i]?|khcn\s*dvk|da\s*hoc\s*qua\s*lop\s*dang\s*vien\s*moi)\b', n_full):
        return match_excel(46, 96.0, "Giấy chứng nhận học lớp đảng viên mới")

    # 8. 39. Giấy giới thiệu người ưu tú vào Đảng (Mẫu 3-KNĐ / M3-PTĐ)
    if re.search(r'\b(mau\s*3\s*knd|m3\s*pt[db]|giay\s*(dam\s*bao\s*|bao\s*dam\s*)?gioi\s*thieu\s*[qt]uan\s*[ceg0]?hung\s*uu\s*t[u0]\s*vao\s*[db]ang|gioi\s*thieu\s*[qt]uan\s*[ceg0]?hung\s*uu\s*t[u0]\s*vao\s*[db]ang|giay\s*gioi\s*thieu\s*nguoi\s*(uu\s*tu\s*)?vao\s*dang|gioi\s*thieu\s*nguoi\s*uu\s*tu\s*vao\s*dang|giay\s*bao\s*dam\s*nguoi\s*(uu\s*tu\s*|xin\s*)?vao\s*dang|doi\s*chieu\s*tieu\s*chuan\s*va\s*dieu\s*kien\s*ket\s*nap\s*dang\s*vien\s*toi\s*gioi\s*thieu)\b', n_full):
        if not re.search(r'\b(doan\s*vien\s*uu\s*tu|ban\s*chap\s*hanh\s*doan|doan\s*tncs)\b', n_full[:200]):
            return match_excel(39, 96.0, "Mẫu 3-KNĐ / Giấy giới thiệu người vào Đảng")

    # 9. 40. Nghị quyết / Giấy bảo đảm giới thiệu đoàn viên ưu tú vào Đảng (Mẫu 4-KNĐ)
    if re.search(r'\b(mau\s*4b?\s*knd|nghi\s*quyet\s*gioi\s*(thieu|t\s*pheu)\s*d?o?an\s*vien\s*uu\s*(tu|10)|nghi\s*quyet\s*gioi\s*thieu\s*de\s*nghi\s*xet\s*bao\s*(dam|in)|giay\s*(bao\s*dam|dam\s*bao)\s*gioi\s*(thieu|hieu|t\s*thieu)\s*doan\s*vien\s*uu\s*tu|gioi\s*(thieu|hieu)\s*[d2]?o?an\s*vien\s*(uu|ru|u)\s*tu\s*(vao\s*dang|de\s*nghi\s*xet\s*ket\s*nap\s*dang)|xet\s*gioi\s*thieu\s*dang\s*vien\s*du\s*bi\s*chuyen)\b', n_full):
        if re.search(r'\b(trich\s*bien\s*(ban|bon)\s*sinh\s*hoat|bien\s*ban\s*sinh\s*hoat\s*chi\s*doan|lien\s*ban\s*sinh\s*hoat\s*chi\s*doan)\b', n_full):
            return match_outside("Trich bien ban sinh hoat chi doan", 94.0, "Trích biên bản sinh hoạt chi đoàn (Ngoài Excel)")
        return match_excel(40, 96.0, "Mẫu 4-KNĐ / Nghị quyết giới thiệu đoàn viên ưu tú vào Đảng")

    # 10. 41. Nghị quyết giới thiệu đoàn viên công đoàn vào Đảng (Mẫu 4b-KNĐ)
    if re.search(r'\b(doan\s*vien\s*cong\s*doan\s*vao\s*dang|ban\s*chap\s*hanh\s*cong\s*doan.*vao\s*dang|hoi\s*phu\s*nu\s*co\s*so.*gt\s*pn)\b', n_full):
        return match_excel(41, 95.0, "Nghị quyết giới thiệu đoàn viên công đoàn/hội viên vào Đảng")

    # 11. 42 vs 49: Tổng hợp ý kiến nhận xét của tổ chức đoàn thể / nơi cư trú (Mẫu 5-KNĐ / Mẫu 12-KNĐ)
    if re.search(r'\b(y\s*kien\s*nhan\s*x[ei]t\s*cua\s*(cac\s*)?(to\s*chuc\s*)?(doan\s*the|chi\s*uy|cap\s*uy|bch\s*chi\s*doan|chi\s*doan)|phieu\s*xin\s*y\s*kien\s*(chi\s*uy\s*|cap\s*uy\s*)?noi\s*cu\s*tru|tong\s*hop\s*y\s*[kr]ien\s*nhan\s*x[ei]t|ban\s*chap\s*hanh\s*doan\s*the\s*noi\s*cong\s*tac\s*va\s*chi\s*uy\s*noi\s*cu\s*tru|cua\s*[qt]uan\s*[cg]hung\s*(noi\s*cu\s*tru\s*|noi\s*sinh\s*hoat\s*)?doi\s*voi\s*(nguoi\s*xin\s*vao\s*dang|dang\s*vien\s*du\s*bi)|nhan\s*xet\s*cua\s*(chi\s*doan|doan\s*co\s*so|to\s*dang)\s*doi\s*voi\s*dang\s*vien\s*du\s*bi|mau\s*12\s*2?knd|mau\s*5[abc]?\s*knd)\b', n_full):
        if not re.search(r'\b(nghi\s*quyet)\b', n_full[:180]):
            if re.search(r'\b(dang\s*vien\s*du\s*(bi|5)|tro\s*thanh\s*dang\s*vien\s*chinh\s*thuc|cong\s*nhan\s*dang\s*vien\s*chinh\s*thuc|mau\s*12\s*2?knd)\b', n_full):
                return match_excel(49, 95.0, "Tổng hợp ý kiến nhận xét đối với đảng viên dự bị")
            if re.search(r'\b(nguoi\s*xin\s*vao\s*dang|nguoi\s*vao\s*dang|quan\s*chung|noi\s*cu\s*tru|chi\s*doan|mau\s*5[abc]?\s*knd)\b', n_full):
                return match_excel(42, 95.0, "Tổng hợp ý kiến nhận xét đối với người vào Đảng")

    # 12. 47. Bản tự kiểm điểm của đảng viên dự bị (Mẫu 10-KNĐ)
    if re.search(r'\b(mau\s*[1i][0o]\s*knd|ban\s*tu\s*[kr]iem\s*(diem\s*)?(dang\s*vien\s*)?(de\s*nghi\s*(chuyen\s*dang\s*vien\s*du\s*bi|[cdg0]?o?n[gh]\s*nhan\s*dang\s*vien\s*chinh\s*thu)|cua\s*dang\s*vien\s*du\s*bi)|de\s*nghi\s*chuyen\s*dang\s*vien\s*du\s*bi\s*thanh\s*dang\s*vien\s*chinh\s*thuc|de\s*hb?h?\s*(cong|dong)\s*nhan\s*dang\s*vien\s*chinh\s*thuc|kiem\s*diem\s*dang\s*vien\s*trong\s*thoi\s*gian\s*du\s*bi)\b', n_full):
        if not re.search(r'\b(nghi\s*quyet|quyet\s*dinh)\b', n_full[:180]):
            return match_excel(47, 96.0, "Mẫu 10-KNĐ / Bản tự kiểm điểm của đảng viên dự bị")

    # 13. 48. Bản nhận xét đảng viên dự bị (Mẫu 11-KNĐ)
    if re.search(r'\b(mau\s*11[ab]?\s*knd|mau\s*[il1][il1][ab]?\s*knd|mau\s*hia\s*knd|mau\s*11k\s*2|ban\s*nhan\s*xet\s*(ket\s*qua\s*phan\s*dau\s*(cua\s*)?)?dang\s*vien\s*du\s*bi|[kh]et\s*[q(]ua\s*phan\s*dau\s*[cg]ua\s*dang\s*vien\s*du\s*[bd]i|nhan\s*xet\s*ket\s*qua\s*phan\s*dau\s*cua\s*dang\s*vien\s*du\s*bi|phan\s*cong\s*giup\s*do\s*dang\s*vien\s*du\s*bi)\b', n_full):
        return match_excel(48, 96.0, "Mẫu 11-KNĐ / Bản nhận xét đảng viên dự bị")

    # 14. 44 & 51: Báo cáo thẩm định / đề nghị của đảng ủy bộ phận
    if re.search(r'\b(bao\s*cao\s*(ve\s*viec\s*)?(tham\s*dinh\s*(nghi\s*quyet|quyet\s*dinh)|de\s*nghi\s*(ket\s*nap\s*dang\s*vien|cong\s*nhan\s*dang\s*vien\s*chinh\s*thuc))|mau\s*14\s*kn[dp].*bao\s*cao|mau\s*h4?\s*knd.*bao\s*cao|mau\s*7\s*knd.*bao\s*cao)\b', n_full):
        if re.search(r'\b(cong\s*nhan\s*dang\s*vien\s*chinh\s*thuc|mau\s*14|mau\s*h4?\s*knd)\b', n_full):
            return match_excel(51, 95.0, "Báo cáo đề nghị công nhận đảng viên chính thức của đảng ủy bộ phận")
        return match_excel(44, 95.0, "Báo cáo đề nghị kết nạp đảng viên của đảng ủy bộ phận")

    # 15. 43 vs 45: Nghị quyết đề nghị / chuẩn y kết nạp đảng viên
    if re.search(r'\b(nghi\s*quyet\s*(chi\s*bo\s*)?(xet\s*)?([db]?e\s*nghi\s*|ve\s*viec\s*|[cgdb]?huan\s*y\s*)?(xet\s*)?ket\s*[nhr]ap\s*[db]ang\s*vien|trich\s*nghi\s*quyet\s*(chi\s*bo\s*|dang\s*uy\s*)?.*ket\s*nap\s*dang\s*vien|bien\s*ban\s*(sinh\s*hoat\s*|hop\s*)?chi\s*uy.*ket\s*nap\s*(quan\s*chung|dang\s*vien)|mau\s*6\s*knd|mau\s*7\s*knd.*nghi\s*quyet|mau\s*80?\s*knd|mau\s*8\s*kd|m6\s*pt[db])\b', n_full):
        if not (re.search(r'\b(quyet\s*dinh|qd\s*du|qd\s*tv)\b', n_full[:200]) and not re.search(r'\b(nghi\s*quyet|nq)\b', n_full[:200])):
            if re.search(r'\b(mau\s*7\s*knd|mau\s*80?\s*knd|mau\s*8\s*kd|nq\s*du|no\s*du|ng\s*du|ban\s*chap\s*hanh\s*dang\s*bo|dang\s*uy\s*tieu\s*doan|dang\s*uy\s*trung\s*doan|dang\s*uy\s*co\s*so|trich\s*nghi\s*quyet\s*dang\s*uy|[cgdb]?huan\s*y\s*ket\s*nap\s*dang\s*vien|tong\s*so\s*uy\s*vien\s*ban\s*chap\s*hanh|tong\s*so\s*dang\s*uy\s*vien)\b', n_full):
                return match_excel(45, 95.0, "Mẫu 8-KNĐ / Nghị quyết kết nạp đảng viên của đảng ủy")
            if re.search(r'\b(mau\s*6\s*knd|m6\s*pt[db]|nq\s*cb|nghi\s*quyet\s*chi\s*bo|chi\s*bo\s*.*hop\s*de\s*xet|tien\s*hanh\s*hop\s*hoi|tong\s*so\s*dang\s*vien\s*cua\s*chi\s*bo)\b', n_full):
                return match_excel(43, 95.0, "Mẫu 6-KNĐ / Nghị quyết đề nghị kết nạp đảng viên của chi bộ")
            if hint_id in (43, 45):
                return match_excel(hint_id, 93.0, "Nghị quyết đề nghị kết nạp đảng viên")
            return match_excel(43, 90.0, "Nghị quyết đề nghị kết nạp đảng viên của chi bộ")

    # 16. 50 vs 52: Nghị quyết công nhận đảng viên chính thức
    if re.search(r'\b(nghi\s*quyet\s*(chi\s*bo\s*)?(xet\s*)?([dbu]?e\s*nghi\s*|ve\s*viec\s*|[cgdb]?huan\s*y\s*)?(chuyen\s*dang\s*vien\s*du\s*bi\s*thanh\s*|[cgdb0]ong\s*nhan\s*)[db]ang\s*vien\s*[cg]hinh\s*thuc|xet\s*chuyen\s*dang\s*vien\s*du\s*bi\s*thanh\s*dang\s*vien\s*chinh\s*thuc|mau\s*13\s*knd|mau\s*14\s*kn[do].*nghi\s*quyet|mau\s*[1i][5s]\s*k?e?nd|mau\s*16\s*kn[d2].*nghi\s*quyet|mau\s*20\s*knd)\b', n_full):
        if not (re.search(r'\b(quyet\s*dinh|qd\s*du|qd\s*tv|mau\s*15a)\b', n_full[:200]) and not re.search(r'\b(nghi\s*quyet|nq)\b', n_full[:200])):
            if re.search(r'\b(mau\s*[1i][5s]\s*k?e?nd|mau\s*16\s*kn[d2]|nq\s*du|nqfdu|ng\s*du|ban\s*chap\s*hanh\s*dang\s*bo|dang\s*uy\s*.*hop\s*de\s*xet|[cgdb]?huan\s*y\s*[cg0]ong\s*nhan\s*dang\s*vien|tong\s*so\s*uy\s*vien\s*ban\s*chap\s*hanh|tong\s*so\s*dang\s*uy\s*vien)\b', n_full):
                return match_excel(52, 95.0, "Nghị quyết công nhận đảng viên chính thức của đảng ủy")
            if re.search(r'\b(mau\s*13\s*knd|mau\s*14\s*kn[do]|mau\s*20\s*knd|nq\s*cb|nghi\s*quyet\s*chi\s*bo|chi\s*bo\s*.*hop\s*de\s*xet|tien\s*hanh\s*hop\s*hoi|tong\s*so\s*dang\s*vien\s*cua\s*chi\s*bo)\b', n_full):
                return match_excel(50, 95.0, "Nghị quyết công nhận đảng viên chính thức của chi bộ")
            if hint_id in (50, 52):
                return match_excel(hint_id, 93.0, "Nghị quyết công nhận đảng viên chính thức")
            return match_excel(50, 90.0, "Nghị quyết công nhận đảng viên chính thức của chi bộ")

    # 16b. TỜ KHAI ĐỀ NGHỊ TẶNG HUY HIỆU ĐẢNG (102/103) & CÁC BIÊN BẢN/PHIẾU KỶ LUẬT NGOÀI EXCEL
    if re.search(r'\b(to\s*khai\s*de\s*nghi\s*(truy\s*)?tang\s*huy\s*hieu\s*dang)\b', n_full):
        if re.search(r'\btruy\s*tang\b', n_full):
            return match_excel(103, 96.0, "Tờ khai đề nghị truy tặng Huy hiệu đảng")
        return match_excel(102, 96.0, "Tờ khai đề nghị tặng Huy hiệu đảng")

    if re.search(r'\b(bien\s*ban\s*gap\s*dang\s*vien\s*truoc\s*khi\s*thi\s*hanh\s*ky\s*luat)\b', n_full):
        return match_outside("Bien ban gap dang vien truoc khi thi hanh ky luat", 95.0, "Biên bản gặp đảng viên trước khi thi hành kỷ luật (Ngoài Excel)")
    if re.search(r'\b(bien\s*ban\s*cong\s*bo\s*va\s*giao\s*quyet\s*dinh\s*ky\s*luat)\b', n_full):
        return match_outside("Bien ban cong bo va giao quyet dinh ky luat", 95.0, "Biên bản công bố và giao quyết định kỷ luật (Ngoài Excel)")
    if re.search(r'\b(phieu\s*bieu\s*quyet)\b', n_full) or hint_out == "Phieu bieu quyet":
        return match_outside("Phieu bieu quyet", 95.0, "Phiếu biểu quyết kỷ luật (Ngoài Excel)")
    if re.search(r'\b(bao\s*cao\s*cua\s*bi\s*thu\s*chi\s*bo)\b', n_full):
        return match_outside("Bao cao cua bi thu chi bo", 94.0, "Báo cáo của bí thư chi bộ (Ngoài Excel)")

    # 17. CÁC LOẠI QUYẾT ĐỊNH & QUYẾT NGHỊ
    if re.search(r'\b(can\s*cu\s*(vao\s*)?dieu\s*le\s*dang\s*cong\s*san\s*viet\s*nam.*xet\s*de\s*nghi\s*cua\s*chi\s*bo|m[73]\s*ptd|m10\s*ptd|mau\s*17[ab]?\s*knd|mau\s*15a\s*knd|mau\s*16[ab]?\s*knd|mau\s*9[ab]?\s*knd|miu\s*9\s*knd)\b', n_full):
        if re.search(r'\b(chuyen\s*dang\s*vien\s*du\s*bi\s*thanh\s*dang\s*vien\s*chinh\s*thuc|cong\s*nhan\s*dang\s*vien\s*chinh\s*thuc|m10\s*ptd|m30\s*ptd|mau\s*17[ab]?\s*knd|mau\s*15a\s*knd|mau\s*16[ab]?\s*knd)\b', n_full):
            return match_excel(7, 96.0, "Quyết định/Quyết nghị công nhận đảng viên chính thức")
        if re.search(r'\b(ve\s*viec\s*ket\s*nap\s*dang\s*vien|ket\s*nap\s*(quan\s*chung|dang\s*vien)|m7\s*ptd|mau\s*9[ab]?\s*knd|miu\s*9\s*knd)\b', n_full):
            return match_excel(5, 96.0, "Quyết định/Quyết nghị kết nạp đảng viên")

    if re.search(r'\b(quyet\s*dinh|quyet\s*din|quyet\s*nghi|qd\s*du|qd\s*tv|ban\s*thuong\s*vu\s*quyet\s*dinh|thuong\s*vu\s*dang\s*uy|chinh\s*uy\s*.*quyet\s*dinh|tu\s*lenh\s*.*quyet\s*dinh|bo\s*truong\s*bo\s*quoc\s*phong|doan\s*truong\s*.*quyet\s*dinh|su\s*doan\s*truong\s*.*quyet\s*dinh|quyet\s*dinh\s*cua\s*su\s*doan|cuc\s*truong\s*cuc\s*hau\s*can)\b', n_full):
        if re.search(r'\b(mau\s*9[ab]?\s*knd|mu\s*9h?\s*knd|mau\s*8\s*knd|quyet\s*dinh\s*(cua\s*dang\s*uy\s*)?(v\s*v\s*|ve\s*viec\s*)?k[e]?\s*nap\s*dang\s*[vy]ien|k[e]?\s*nap\s*dang\s*[vy]ien|xet\s*don\s*xin\s*vao\s*dang|de\s*nghi\s*ket\s*nap\s*quan\s*chung|ve\s*viec\s*ket\s*nap\s*dang\s*vien\s*moi)\b', n_full):
            if not re.search(r'\b(nghi\s*quyet|cong\s*nhan\s*dang\s*vien\s*chinh\s*thuc)\b', n_full[:150]):
                return match_excel(5, 97.0, "Quyết định kết nạp đảng viên")
        if re.search(r'\b(mau\s*15a\s*knd|mau\s*16[ab]?\s*knd|mau\s*17[ab]?\s*knd|quyet\s*dinh\s*cong\s*nhan\s*dang\s*vien\s*chinh\s*thuc|cong\s*nhan\s*dang\s*vien\s*(chinh\s*thuc|du\s*bi)|chuyen\s*dang\s*vien\s*du\s*bi\s*thanh\s*dang\s*vien\s*chinh\s*thuc)\b', n_full):
            return match_excel(7, 97.0, "Quyết định công nhận đảng viên chính thức")
        if re.search(r'\b(phat\s*the\s*dang\s*vien|cap\s*the\s*dang\s*vien)\b', n_full):
            return match_excel(19, 97.0, "Quyết định phát thẻ Đảng viên")
        if re.search(r'\b(doi\s*lai\s*the\s*dang\s*vien|doi\s*the\s*dang\s*vien|mau\s*3\s*dtdv)\b', n_full):
            return match_excel(73, 96.0, "Quyết định đổi lại thẻ đảng viên")
        if re.search(r'\b(cap\s*lai\s*the\s*dang\s*vien\s*bi\s*mat)\b', n_full):
            return match_excel(72, 96.0, "Quyết định cấp lại thẻ đảng viên bị mất")
        if re.search(r'\b(tang\s*huy\s*hieu\s*dang|tang\s*huy\s*hieu\s*30\s*nam|tang\s*huy\s*hieu\s*40\s*nam)\b', n_full):
            return match_excel(20, 96.0, "Quyết định tặng Huy hiệu Đảng")
        if re.search(r'\b(thi\s*hanh\s*ky\s*luat|ky\s*luat\s*dang|ky\s*luat\s*hanh\s*chinh)\b', n_full) and not re.search(r'\b(quan\s*triet\s*thuc\s*hien\s*nghi\s*quyet)\b', n_full):
            if hint_id == 36 or re.search(r'\b(ky\s*luat\s*hanh\s*chinh|cong\s*hoa\s*xa\s*hoi\s*chu\s*nghia|doc\s*lap\s*tu\s*do\s*hanh\s*phuc)\b', n_full[:220]):
                return match_excel(36, 95.0, "Quyết định kỷ luật hành chính")
            return match_excel(22, 95.0, "Quyết định kỷ luật Đảng")
        if re.search(r'\b(quyet\s*dinh\s*ve\s*viec\s*khen\s*thuong|tang\s*danh\s*hieu|tang\s*bang\s*khen|tang\s*giay\s*khen)\b', n_full):
            return match_excel(23, 94.0, "Quyết định khen thưởng")
        if re.search(r'\b(nghi\s*huu|huong\s*che\s*do\s*huu\s*tri)\b', n_full):
            return match_excel(88, 95.0, "Quyết định nghỉ hưu")
        if re.search(r'\b(phuc\s*vien|chuyen\s*nganh)\b', n_full):
            return match_excel(90, 95.0, "Quyết định phục viên/chuyển ngành")
        if re.search(r'\b(xuat\s*ngu)\b', n_full):
            return match_excel(91, 95.0, "Quyết định xuất ngũ")
        if re.search(r'\b(dieu\s*dong|bo\s*nhiem|thang\s*quan\s*ham|phong\s*quan\s*ham|nang\s*bac\s*luong|nang\s*luong|chuyen\s*che\s*do|xep\s*luong|giao\s*nhiem\s*vu|kien\s*toan\s*(bi\s*thu|chi\s*uy|dang\s*uy|cap\s*uy)|chuan\s*y\s*ket\s*qua\s*bau|chi\s*dinh\s*(ban\s*chap\s*hanh|chu\s*nhiem)|luat\s*ve\s*si\s*quan|luat\s*si\s*quan\s*quan\s*doi|quan\s*nhan\s*chuyen\s*nghiep|cong\s*nhan\s*vien\s*quoc\s*phong|tuyen\s*dung|thay\s*doi\s*phien\s*hieu)\b', n_full):
            return match_excel(87, 96.0, "Các quyết định điều động bổ nhiệm")

    # 18. 61 vs 62: BẢN TỰ KIỂM ĐIỂM / BẢN KIỂM ĐIỂM CÁ NHÂN
    if re.search(r'\b(b[a]n\s*g?\s*(tu\s*)?([krlt]i[e]?m|riem|1u\s*om)\s*(d[i]?[e]?m\s*)?(ca\s*nhan|dang\s*vien|cua\s*doan\s*vien)?|tu\s*[kr]iem\s*diem\s*dang\s*vien|ban\s*diem\s*dang\s*vien|tu\s*nhan\s*xet\s*dang\s*vien|ban\s*nhan\s*xet\s*dang\s*vien|nhan\s*xet\s*can\s*bo\s*dang\s*vien|nhan\s*xet\s*dang\s*vien\s*(thuc\s*hien|cua\s*cap\s*uy|tu\s*thang)|ban\s*dong\s*gop\s*y\s*kien\s*dang\s*vien|xac\s*nhan\s*da\s*kiem\s*diem\s*tu\s*phe\s*binh)\b', n_full):
        if hint_id == 87 and not re.search(r'\b(ban\s*(tu\s*)?kiem\s*diem)\b', n_full):
            return match_excel(87, 94.0, f"Đối chiếu kết hợp tên file gốc ({orig_stem})")
        if re.search(r'\b(de\s*nghi\s*(chuyen\s*dang\s*vien\s*du\s*bi|[cdg0]?o?n[gh]\s*nhan\s*dang\s*vien\s*chinh\s*thu)|cua\s*dang\s*vien\s*du\s*bi|trong\s*thoi\s*gian\s*du\s*bi)\b', n_full):
            return match_excel(47, 95.0, "Bản tự kiểm điểm của đảng viên dự bị")
        if re.search(r'\b(khi\s*chuyen\s*(sinh\s*hoat\s*dang|shd)|dung\s*(khi\s*|cho\s*)?chuyen\s*(sinh\s*hoat\s*dang|shd|sh|nh|sim)|dung\s*khi\s*ch[ij]\s*shd|tai\s*thoi\s*diem\s*chuyen\s*sinh\s*hoat\s*dang)\b', n_full):
            if hint_id == 61:
                return match_excel(61, 95.0, "Bản tự kiểm điểm đảng viên (3 năm gần nhất) và khi chuyển sinh hoạt đảng")
            return match_excel(62, 95.0, "Bản tự kiểm điểm tại thời điểm chuyển sinh hoạt đảng")
        return match_excel(61, 96.0, "Bản tự kiểm điểm đảng viên hàng năm")

    # 19. 54 vs 55 vs 59: GIẤY GIỚI THIỆU SINH HOẠT ĐẢNG
    if re.search(r'\b(giay\s*gioi\s*thieu\s*sinh\s*hoat\s*dang|giay\s*gioi\s*thieu.*sinh\s*(hoat|homt|hot)\s*dang|gioi\s*thieu\s*.*s[id]n\s*hoat\s*dang|sinh\s*hoat\s*dang\s*noi\s*bo|sinh\s*hoat\s*dang\s*tam\s*thoi|dung\s*trong\s*noi\s*bo\s*dang\s*bo\s*co\s*so|da\s*(dong|đong)\s*dang\s*phi\s*(het\s*)?thang|kinh\s*chuyen.*kinh\s*chuyen|de\s*nghi\s*nhan\s*va\s*g\s*t\s*s|gt\s*shd|gtshd|gt5hp|gt54d|gtsnd|6tsiip|g\s*t\s*s\s*h\s*d|g\s*l\s*5\s*h\s*d|ten\s*dang\s*bo\s*chi\s*bo\s*co\s*so\s*chuyen\s*den)\b', n_full):
        if re.search(r'\b(sinh\s*hoat\s*dang\s*noi\s*bo|dung\s*trong\s*noi\s*bo\s*dang\s*bo)\b', n_full) or hint_id == 59:
            return match_excel(59, 96.0, "Giấy giới thiệu sinh hoạt đảng nội bộ")
        if re.search(r'\b(sinh\s*hoat\s*dang\s*tam\s*thoi|hoat\s*dang\s*tam\s*thoi)\b', n_full) or hint_id == 55:
            return match_excel(55, 96.0, "Giấy giới thiệu sinh hoạt đảng tạm thời")
        return match_excel(54, 95.0, "Giấy giới thiệu sinh hoạt đảng chính thức")

    # 20. 60 / 100 / 101: PHIẾU BÁO ĐẢNG VIÊN
    if re.search(r'\b(phieu\s*bao\s*dang\s*vien|mau\s*11\s*shd.*phieu\s*bao|dang\s*vien\s*chuyen\s*sinh\s*hoat\s*dang\s*chinh\s*thuc)\b', n_full):
        if re.search(r'\b(cong\s*nhan\s*chinh\s*thuc)\b', n_full):
            return match_excel(100, 95.0, "Phiếu báo đảng viên được công nhận chính thức")
        if re.search(r'\b(ra\s*khoi\s*dang)\b', n_full):
            return match_excel(101, 95.0, "Phiếu báo đảng viên ra khỏi đảng")
        return match_excel(60, 95.0, "Phiếu báo đảng viên chuyển sinh hoạt đảng chính thức")

    # Các tài liệu ngoài danh mục Excel có tiêu đề rõ ràng ở trang đầu
    if re.search(r'\b(phieu\s*theo\s*d[o0]i\s*diem\s*hoc\s*vien)\b', n_full):
        return match_outside("Phieu theo doi diem hoc vien", 95.0, "Phiếu theo dõi điểm học viên (Ngoài Excel)")
    if re.search(r'\b(ban\s*thu\s*hoach|bai\s*thu\s*hoach)\b', n_full):
        return match_outside("Ban thu hoach", 95.0, "Bản thu hoạch (Ngoài Excel)")
    if re.search(r'\b(bien\s*ban\s*ban\s*giao\s*ho\s*so\s*dang\s*vien)\b', n_full):
        return match_outside("Bien ban ban giao ho so dang vien va the dang vien", 95.0, "Biên bản bàn giao hồ sơ đảng viên (Ngoài Excel)")
    if re.search(r'\b(phieu\s*danh\s*gia\s*xep\s*loai\s*(chat\s*luong\s*)?(chi\s*bo\s*va\s*)?dang\s*vien|eu\s*danh\s*gia\s*xep\s*loai\s*dang\s*vien|phieu\s*tu\s*cham\s*diem\s*danh\s*gia)\b', n_full):
        return match_outside("Phieu danh gia xep loai dang vien", 94.0, "Phiếu đánh giá xếp loại đảng viên (Ngoài Excel)")
    if re.search(r'\b(phieu\s*nhan\s*xet\s*dang\s*vien|phieu\s*lay\s*y\s*kien\s*nhan\s*xet)\b', n_full):
        return match_outside("Phieu nhan xet dang vien", 94.0, "Phiếu nhận xét đảng viên (Ngoài Excel)")
    if re.search(r'\b(ban\s*dang\s*ky\s*phan\s*dau|ban\s*cam\s*ket)\b', n_full):
        return match_outside("Ban cam ket", 94.0, "Bản cam kết / đăng ký phấn đấu (Ngoài Excel)")
    if re.search(r'\b(don\s*xin\s*(dang\s*k[yi]\s*)?ket\s*hon|don\s*xin\s*(xay\s*dung|lap)\s*gia\s*dinh|don\s*xin\s*lay\s*vo)\b', n_full):
        return match_outside("Don xin ket hon", 95.0, "Đơn xin kết hôn (Ngoài Excel)")
    if re.search(r'\b(bao\s*cao\s*(v\s*v\s*|ve\s*viec\s*)?mat\s*(the\s*dang\s*vien|ho\s*so))\b', n_full):
        return match_outside("Bao cao ve viec mat the dang vien", 94.0, "Báo cáo mất thẻ đảng viên (Ngoài Excel)")
    if re.search(r'\b(ban\s*tuong\s*trinh)\b', n_full):
        return match_outside("Ban tuong trinh", 94.0, "Bản tường trình (Ngoài Excel)")
    if re.search(r'\b(ket\s*luan\s*kiem\s*tra\s*dang\s*vien)\b', n_full):
        return match_outside("Ket luan kiem tra dang vien", 94.0, "Kết luận kiểm tra đảng viên (Ngoài Excel)")
    if re.search(r'\b(bien\s*ban\s*kiem\s*phieu)\b', n_full):
        return match_outside("Bien ban kiem phieu bieu quyet", 94.0, "Biên bản kiểm phiếu biểu quyết (Ngoài Excel)")
    if re.search(r'\b(hoi\s*dong\s*quan\s*nhan|bb\s*h[dn][oq]n|bien\s*ban\s*hoi\s*nghi.*quan\s*nhan)\b', n_full):
        return match_outside("Bien ban hoi nghi quan nhan", 94.0, "Biên bản hội nghị quân nhân (Ngoài Excel)")
    if re.search(r'\b(to\s*k[hu]a[it]\s*(lam\s*)?the\s*dang\s*vien)\b', n_full):
        return match_outside("To khai lam the dang vien", 94.0, "Tờ khai làm thẻ đảng viên (Ngoài Excel)")
    if re.search(r'\b(giay\s*chung\s*nhan\s*chua\s*lam\s*the\s*dang)\b', n_full):
        return match_outside("Giay chung nhan chua lam the Dang", 94.0, "Giấy chứng nhận chưa làm thẻ Đảng (Ngoài Excel)")
    if re.search(r'\b(giay\s*xac\s*nhan\s*dan\s*toc)\b', n_full):
        return match_outside("Giay xac nhan dan toc", 94.0, "Giấy xác nhận dân tộc (Ngoài Excel)")
    if re.search(r'\b(ho\s*so\s*ket\s*nap\s*dang\s*vien)\b', n_full) and len(n_full) < 160:
        return match_outside("Ho so ket nap dang vien", 93.0, "Bìa hồ sơ kết nạp đảng viên (Ngoài Excel)")
    if re.search(r'\b(ho\s*so\s*cong\s*nhan\s*dang\s*vien\s*chinh\s*thuc)\b', n_full) and len(n_full) < 160:
        return match_outside("Ho so cong nhan dang vien chinh thuc", 93.0, "Bìa hồ sơ công nhận đảng viên chính thức (Ngoài Excel)")
    if re.search(r'\b(trich\s*bien\s*(ban|bon)|bien\s*ban\s*(sinh\s*hoat|hoi\s*nghi|hop|tham\s*gia\s*dong\s*gop)|sien\s*ban\s*inh\s*woat|bien\s*ban\s*ginh\s*hot)\b', n_full):
        if re.search(r'\b(chi\s*doan|doan\s*thanh\s*nien|doan\s*vien\s*uu\s*tu|lien\s*chi\s*doan)\b', n_full):
            return match_outside("Trich bien ban sinh hoat chi doan", 94.0, "Trích biên bản sinh hoạt chi đoàn (Ngoài Excel)")
        if re.search(r'\b(to\s*dang)\b', n_full):
            return match_outside("Bien ban sinh hoat to dang", 94.0, "Biên bản sinh hoạt tổ đảng (Ngoài Excel)")
        if re.search(r'\b(chi\s*bo|hoi\s*nghi\s*chi\s*bo|phe\s*binh)\b', n_full):
            return match_outside("Trich bien ban sinh hoat chi bo", 94.0, "Trích biên bản sinh hoạt chi bộ (Ngoài Excel)")

    # 22. 70 vs 86: VĂN BẰNG, CHỨNG CHỈ, BẰNG TỐT NGHIỆP
    if re.search(r'\b(bang\s*tot\s*nghiep\s*trung\s*hoc\s*pho\s*thong|bang\s*tot\s*nghiep\s*trung\s*hoc\s*so\s*cap|hoc\s*ba)\b', n_full):
        return match_outside("Bang tot nghiep trung hoc pho thong", 94.0, "Bằng tốt nghiệp phổ thông (Ngoài Excel)")
    if re.search(r'\b(ly\s*luan\s*chinh\s*tri|trung\s*cap\s*ly\s*luan|cao\s*cap\s*ly\s*luan|so\s*cap\s*ly\s*luan|boi\s*duong\s*chinh\s*tri)\b', n_full) and re.search(r'\b(bang\s*tot\s*nghiep|chung\s*chi|giay\s*chung\s*nhan)\b', n_full):
        return match_excel(70, 95.0, "Bằng/chứng chỉ lý luận chính trị")
    if re.search(r'\b(bang\s*tot\s*nghiep\s*(dai\s*hoc|cao\s*dang|trung\s*cap)|the\s*degree\s*of\s*bachelor|chung\s*chi|giay\s*chung\s*nhan\s*(tot\s*nghiep|boi\s*duong|dao\s*tao|kien\s*thuc\s*quoc\s*phong|ket\s*qua\s*hoc\s*tap)|chung\s*nhan\s*ket\s*qua\s*hoc\s*tap|phieu\s*diem|bang\s*diem)\b', n_full):
        return match_excel(86, 95.0, "Các văn bằng, chứng chỉ chuyên môn nghiệp vụ")

    # 23. 75. Giấy khai sinh
    if re.search(r'\b(giay\s*khai\s*sinh|ban\s*sao\s*giay\s*khai\s*sinh)\b', n_full):
        return match_excel(75, 95.0, "Giấy khai sinh")

    # 24. 92-99: Công văn / Giấy giới thiệu thẩm tra lý lịch
    if re.search(r'\b(di\s*tham\s*tra\s*ly\s*lich|xac\s*minh\s*ly\s*lich\s*cua\s*nguoi\s*vao\s*dang)\b', n_full):
        if re.search(r'\bgiay\s*gioi\s*thieu\b', n_full):
            return match_excel(96, 94.0, "Giấy giới thiệu cử đảng viên đi thẩm tra lý lịch")
        if re.search(r'\bcong\s*van\b', n_full):
            return match_excel(97, 94.0, "Công văn đề nghị thẩm tra lý lịch của người vào Đảng")

    # 25. Kết hợp gợi ý từ tên file gốc
    if hint_id is not None:
        return match_excel(hint_id, 91.0, f"Đối chiếu kết hợp tên file gốc ({orig_stem})")
    if hint_out is not None:
        return match_outside(hint_out, 91.0, f"Ngoài danh mục Excel ({orig_stem})")

    # 26. Trang lẻ tách rời của Lý lịch đảng viên hoặc Bản kiểm điểm
    if re.search(r'\b(lich\s*su\s*ban\s*than|nhung\s*cong\s*tac\s*va\s*chuc\s*vu\s*da\s*qua|nhung\s*lop\s*dao\s*tao\s*boi\s*duong\s*da\s*qua|hoan\s*canh\s*gia\s*dinh|quan\s*he\s*gia\s*dinh|cam\s*doan\s*ky\s*ten|chung\s*nhan\s*cua\s*cap\s*uy|ket\s*luan\s*cua\s*dia\s*phuong|nhung\s*diem\s*can\s*chu\s*y)\b', n_full):
        return need_check(2, 70.0, "Trang lẻ tách rời của Lý lịch đảng viên (Cần kiểm tra xem có bị tách trang)")

    if re.search(r'\b(han\s*che\s*khuyet\s*diem\s*va\s*nguyen\s*nhan|ket\s*qua\s*khac\s*phuc\s*nhung\s*han\s*che|giai\s*trinh\s*nhung\s*van\s*de\s*duoc\s*goi\s*y\s*kiem\s*diem|xac\s*dinh\s*trach\s*nhiem\s*cua\s*ca\s*nhan|lam\s*ro\s*trach\s*nhiem\s*cua\s*ca\s*nhan|phuong\s*huong\s*bien\s*phap\s*khac\s*phuc\s*nhung\s*han\s*che|phuong\s*huong\s*bien\s*phap\s*khac\s*phuc\s*han\s*che|tu\s*nhan\s*muc\s*xep\s*loai\s*chat\s*luong|chi\s*bo\s*de\s*xuat\s*xep\s*loai\s*muc\s*chat\s*luong|danh\s*gia\s*xep\s*loai\s*chat\s*luong\s*dang\s*vien|y\s*kien\s*nhan\s*xet\s*va\s*(ket\s*qua\s*)?bieu\s*quyet\s*cua\s*chi\s*bo|ket\s*qua\s*bieu\s*quyet\s*hoac\s*bo\s*phieu\s*cua\s*chi\s*bo)\b', n_full):
        return need_check(61, 70.0, "Trang lẻ tách rời của Bản kiểm điểm đảng viên (Cần kiểm tra xem có bị tách trang)")

    # 27. Fuzzy Matching với toàn bộ các mục trong danh mục Excel hiện tại
    if len(n_full) >= 10:
        best_id = None
        best_score = 0.0
        second_score = 0.0
        head = n_full[:400]
        for cid, item in catalog.items():
            nt = item['norm_title']
            if not nt:
                continue
            if len(nt) >= 10 and nt in head:
                return ('EXCEL', cid, item['safe_name'], item['safe_name'], 95.0, "Khớp trực tiếp tiêu đề trong Excel")
            if len(nt) < 24 and nt not in head:
                continue
            s1 = fuzz.token_set_ratio(nt, head)
            s2 = fuzz.partial_ratio(nt, head)
            score = 0.6 * s1 + 0.4 * s2
            if score > best_score:
                second_score = best_score
                best_score = score
                best_id = cid
            elif score > second_score:
                second_score = score
        if best_score >= 86.0 and (best_score - second_score) >= 6.0:
            item = catalog[best_id]
            return ('EXCEL', best_id, item['safe_name'], item['safe_name'], round(best_score, 1), "Fuzzy matching tiêu đề OCR với Excel")
        return need_check(best_id, round(best_score, 1), f"Điểm đối chiếu thấp hoặc nhiều kết quả tương đồng ({best_score:.1f} vs {second_score:.1f})")

    return need_check(None, 0.0, "Trang trắng hoặc chữ viết tay mờ không nhận dạng được")


# ==============================================================================
# 4. WORKER ĐA TIẾN TRÌNH QUÉT OCR CỤC BỘ QUA LIBTESSERACT C-API
# ==============================================================================

_api = None
_lib = None
_dll_directory_handle = None
_worker_root_dir = None
_worker_allow_outside = True


def init_ocr_worker(excel_path_str: str, root_dir_str: str, tess_dir_str: str, tessdata_dir_str: str, allow_outside: bool = True):
    global _api, _lib, _dll_directory_handle, _worker_root_dir, _worker_allow_outside, CATALOG
    os.environ['OMP_THREAD_LIMIT'] = '1'
    _worker_root_dir = pathlib.Path(root_dir_str)
    _worker_allow_outside = allow_outside
    CATALOG = load_excel_catalog(pathlib.Path(excel_path_str))

    _dll_directory_handle = os.add_dll_directory(tess_dir_str)
    _lib = ctypes.CDLL(os.path.join(tess_dir_str, 'libtesseract-5.dll'))
    _lib.TessBaseAPICreate.restype = ctypes.c_void_p
    _lib.TessBaseAPIInit3.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_char_p]
    _lib.TessBaseAPIInit3.restype = ctypes.c_int
    _lib.TessBaseAPISetPageSegMode.argtypes = [ctypes.c_void_p, ctypes.c_int]
    _lib.TessBaseAPISetVariable.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_char_p]
    _lib.TessBaseAPISetImage.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int]
    _lib.TessBaseAPIGetUTF8Text.argtypes = [ctypes.c_void_p]
    _lib.TessBaseAPIGetUTF8Text.restype = ctypes.c_void_p
    _lib.TessDeleteText.argtypes = [ctypes.c_void_p]

    _api = _lib.TessBaseAPICreate()
    _lib.TessBaseAPIInit3(_api, tessdata_dir_str.encode('utf-8'), b'vie')
    _lib.TessBaseAPISetPageSegMode(_api, 6)
    _lib.TessBaseAPISetVariable(_api, b'omp_thread_limit', b'1')


def _ocr_pixmap(pix):
    global _api, _lib
    _lib.TessBaseAPISetImage(_api, pix.samples, pix.width, pix.height, 1, pix.stride)
    ptr = _lib.TessBaseAPIGetUTF8Text(_api)
    if ptr:
        txt = ctypes.string_at(ptr).decode('utf-8', errors='replace')
        _lib.TessDeleteText(ptr)
        return " | ".join(line.strip() for line in txt.splitlines() if line.strip())
    return ""


def process_single_pdf(rel_path_str: str):
    full_path = _worker_root_dir / rel_path_str
    orig_stem = full_path.stem
    try:
        doc = pymupdf.open(full_path)
        num_pages = len(doc)
        if num_pages == 0:
            return (rel_path_str, "", "", "CHECK", None, "", "", 0.0, "File PDF rỗng (0 trang)")

        page0 = doc[0]
        r = page0.rect

        # Bước 1: OCR nhanh 40% đầu trang 1
        clip_top = pymupdf.Rect(r.width * 0.04, r.height * 0.02, r.width * 0.96, r.height * 0.40)
        pix_top = page0.get_pixmap(dpi=85, clip=clip_top, colorspace=pymupdf.csGRAY)
        ocr_p1 = _ocr_pixmap(pix_top)

        st, cid, ename, prop, conf, rsn = classify_text(ocr_p1, orig_stem, "", CATALOG, _worker_allow_outside)
        if st in ('EXCEL', 'OUTSIDE') and conf >= 90.0:
            doc.close()
            return (rel_path_str, ocr_p1, "", st, cid, ename, prop, conf, rsn)

        # Bước 2: Nếu chưa rõ, đọc phần còn lại của trang 1 hoặc xoay trang nếu bị ngang
        extra_texts = []
        if len(norm_text(ocr_p1)) < 35 or st == 'CHECK':
            clip_full = pymupdf.Rect(r.width * 0.04, r.height * 0.02, r.width * 0.96, r.height * 0.85)
            pix_full = page0.get_pixmap(dpi=85, clip=clip_full, colorspace=pymupdf.csGRAY)
            ocr_p1_full = _ocr_pixmap(pix_full)
            if len(ocr_p1_full) > len(ocr_p1):
                ocr_p1 = ocr_p1_full

            st, cid, ename, prop, conf, rsn = classify_text(ocr_p1, orig_stem, "", CATALOG, _worker_allow_outside)
            if st in ('EXCEL', 'OUTSIDE') and conf >= 90.0:
                doc.close()
                return (rel_path_str, ocr_p1, "", st, cid, ename, prop, conf, rsn)

            orig_rot = page0.rotation
            if orig_rot in (90, 270):
                page0.set_rotation(0)
                r0 = page0.rect
                pix_rot0 = page0.get_pixmap(dpi=85, clip=pymupdf.Rect(r0.width * 0.04, r0.height * 0.02, r0.width * 0.96, r0.height * 0.55), colorspace=pymupdf.csGRAY)
                txt_rot0 = _ocr_pixmap(pix_rot0)
                if txt_rot0:
                    extra_texts.append(txt_rot0)
                    st, cid, ename, prop, conf, rsn = classify_text(ocr_p1, orig_stem, " | ".join(extra_texts), CATALOG, _worker_allow_outside)
                    if st in ('EXCEL', 'OUTSIDE') and conf >= 90.0:
                        doc.close()
                        return (rel_path_str, ocr_p1, " | ".join(extra_texts), st, cid, ename, prop, conf, rsn)

        # Bước 3: Chỉ đọc thêm trang 2 khi thực sự cần thiết
        if num_pages >= 2 and st == 'CHECK':
            page1 = doc[1]
            r1 = page1.rect
            pix_p2 = page1.get_pixmap(dpi=85, clip=pymupdf.Rect(r1.width * 0.04, r1.height * 0.02, r1.width * 0.96, r1.height * 0.55), colorspace=pymupdf.csGRAY)
            txt_p2 = _ocr_pixmap(pix_p2)
            if txt_p2:
                extra_texts.append(f"[Trang 2] {txt_p2}")
                st, cid, ename, prop, conf, rsn = classify_text(ocr_p1, orig_stem, " | ".join(extra_texts), CATALOG, _worker_allow_outside)

        doc.close()
        return (rel_path_str, ocr_p1, " | ".join(extra_texts), st, cid, ename, prop, conf, rsn)

    except Exception as e:
        return (rel_path_str, f"ERROR: {e}", "", "CHECK", None, "", "", 0.0, f"Lỗi đọc file PDF: {e}")


# ==============================================================================
# 5. QUẢN LÝ BỘ NHỚ ĐỆM (SQLITE CACHE) & LẬP BÁO CÁO DRY-RUN
# ==============================================================================

def init_cache_db(db_path: pathlib.Path):
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS pdf_ocr_cache (
            rel_path TEXT PRIMARY KEY,
            file_size INTEGER,
            mtime_ns INTEGER,
            ocr_p1 TEXT,
            ocr_extra TEXT
        )
    """)
    conn.commit()
    return conn


def run_dry_run_engine(
    root_dir: pathlib.Path,
    excel_path: pathlib.Path,
    report_dir: pathlib.Path,
    tess_dir: str = DEFAULT_CONFIG['tess_dir'],
    tessdata_dir: str = DEFAULT_CONFIG['tessdata_dir'],
    workers: int = 10,
    allow_outside: bool = True,
    progress_callback=None,
    stop_event=None
) -> dict:
    """
    Thực hiện quét OCR (có cache) và đối chiếu với danh mục Excel `excel_path`.
    Gọi `progress_callback(stage, current, total, message)` để cập nhật GUI.
    """
    t_start = time.time()
    root_dir = pathlib.Path(root_dir)
    excel_path = pathlib.Path(excel_path)
    report_dir = pathlib.Path(report_dir)
    report_dir.mkdir(parents=True, exist_ok=True)
    cache_db_path = report_dir / 'ocr_cache.sqlite'

    catalog = reload_global_catalog(excel_path)
    if not catalog:
        raise ValueError(f"Danh mục Excel rỗng hoặc không hợp lệ: {excel_path}")

    if progress_callback:
        progress_callback("scan", 0, 100, f"Đang quét danh sách file PDF trong {root_dir} ...")

    all_pdfs = []
    for p in root_dir.rglob('*.pdf'):
        if stop_event and stop_event.is_set():
            return {'cancelled': True}
        rel = p.relative_to(root_dir)
        if any(part in EXCLUDED_FOLDER_NAMES for part in rel.parts):
            continue
        if report_dir in p.parents:
            continue
        all_pdfs.append(p)

    total_pdfs = len(all_pdfs)
    if total_pdfs == 0:
        return {
            'cancelled': False,
            'total_pdfs': 0,
            'excel_matched': 0,
            'outside_excel': 0,
            'need_check': 0,
            'duplicates': 0,
            'records': [],
            'dept_stats': {},
            'elapsed': 0.0
        }

    conn = init_cache_db(cache_db_path)
    cached_ocr = {}
    cur = conn.cursor()
    cur.execute("SELECT rel_path, file_size, mtime_ns, ocr_p1, ocr_extra FROM pdf_ocr_cache")
    for rel_s, fsize, mtime_ns, p1, p_ex in cur.fetchall():
        cached_ocr[rel_s] = (fsize, mtime_ns, p1, p_ex)

    to_ocr = []
    results_map = {}

    for idx, p in enumerate(all_pdfs, start=1):
        if stop_event and stop_event.is_set():
            conn.close()
            return {'cancelled': True}
        rel_s = str(p.relative_to(root_dir))
        stat = p.stat()
        if rel_s in cached_ocr:
            c_size, c_mtime, ocr_p1, ocr_extra = cached_ocr[rel_s]
            if c_size == stat.st_size and c_mtime == stat.st_mtime_ns:
                st, cid, ename, prop, conf, rsn = classify_text(ocr_p1, p.stem, ocr_extra, catalog, allow_outside)
                results_map[rel_s] = (rel_s, ocr_p1, ocr_extra, st, cid, ename, prop, conf, rsn)
                if idx % 400 == 0 and progress_callback:
                    progress_callback("cache", idx, total_pdfs, f"Đang đối chiếu từ bộ nhớ đệm: {idx:,}/{total_pdfs:,} file...")
                continue
        to_ocr.append((rel_s, stat.st_size, stat.st_mtime_ns))

    if to_ocr:
        meta_lookup = {item[0]: (item[1], item[2]) for item in to_ocr}
        rel_list = [item[0] for item in to_ocr]
        batch_insert = []
        done_count = 0
        ocr_t0 = time.time()

        if progress_callback:
            progress_callback("ocr", 0, len(to_ocr), f"Bắt đầu nhận diện OCR {len(to_ocr):,} file mới ({len(results_map):,} file đã có cache)...")

        with concurrent.futures.ProcessPoolExecutor(
            max_workers=workers,
            initializer=init_ocr_worker,
            initargs=(str(excel_path), str(root_dir), str(tess_dir), str(tessdata_dir), allow_outside)
        ) as executor:
            for res in executor.map(process_single_pdf, rel_list, chunksize=8):
                if stop_event and stop_event.is_set():
                    executor.shutdown(wait=False, cancel_futures=True)
                    conn.close()
                    return {'cancelled': True}

                rel_s, ocr_p1, ocr_extra, st, cid, ename, prop, conf, rsn = res
                results_map[rel_s] = res
                fsize, mtime_ns = meta_lookup[rel_s]
                batch_insert.append((rel_s, fsize, mtime_ns, ocr_p1, ocr_extra))
                done_count += 1

                if done_count % 20 == 0 or done_count == len(to_ocr):
                    elapsed = time.time() - ocr_t0
                    rate = done_count / elapsed if elapsed > 0 else 0
                    eta = (len(to_ocr) - done_count) / rate if rate > 0 else 0
                    if progress_callback:
                        progress_callback(
                            "ocr",
                            done_count,
                            len(to_ocr),
                            f"Đang OCR: {done_count:,}/{len(to_ocr):,} ({done_count*100/len(to_ocr):.1f}%) | Tốc độ: {rate:.1f} file/s | Còn khoảng: {eta/60:.1f} phút"
                        )

                if len(batch_insert) >= 200:
                    conn.executemany(
                        "INSERT OR REPLACE INTO pdf_ocr_cache (rel_path, file_size, mtime_ns, ocr_p1, ocr_extra) VALUES (?, ?, ?, ?, ?)",
                        batch_insert
                    )
                    conn.commit()
                    batch_insert.clear()

        if batch_insert:
            conn.executemany(
                "INSERT OR REPLACE INTO pdf_ocr_cache (rel_path, file_size, mtime_ns, ocr_p1, ocr_extra) VALUES (?, ?, ?, ?, ?)",
                batch_insert
            )
            conn.commit()
            batch_insert.clear()

    conn.close()

    if progress_callback:
        progress_callback("report", 95, 100, "Đang xử lý đánh số trùng lặp và xuất các báo cáo Excel...")

    by_folder = defaultdict(list)
    for p in all_pdfs:
        by_folder[p.parent].append(p)

    all_records = []
    check_records = []
    ui_records = []

    stat_total = total_pdfs
    stat_excel_matched = 0
    stat_outside_excel = 0
    stat_need_check = 0
    stat_duplicates = 0

    dept_stats = defaultdict(lambda: {'total': 0, 'excel': 0, 'outside': 0, 'check': 0, 'dup': 0})

    for folder_path in sorted(by_folder.keys()):
        pdf_list = sorted(by_folder[folder_path], key=natural_sort_key)
        rel_dir = folder_path.relative_to(root_dir)
        dept, person = get_dept_and_person(rel_dir)

        base_groups = defaultdict(list)
        folder_items = []
        for p in pdf_list:
            rel_s = str(p.relative_to(root_dir))
            rel_s, ocr_p1, ocr_extra, st, cid, ename, prop, conf, rsn = results_map[rel_s]
            folder_items.append((p, rel_s, ocr_p1, ocr_extra, st, cid, ename, prop, conf, rsn))
            if st in ('EXCEL', 'OUTSIDE') and prop:
                base_groups[prop].append(p)

        base_counters = defaultdict(int)
        for p, rel_s, ocr_p1, ocr_extra, st, cid, ename, prop, conf, rsn in folder_items:
            dept_stats[dept]['total'] += 1
            ocr_summary = ocr_p1 if not ocr_extra else f"{ocr_p1} | {ocr_extra}"
            ocr_summary_short = ocr_summary[:500]

            if st in ('EXCEL', 'OUTSIDE') and prop:
                total_same = len(base_groups[prop])
                if total_same > 1:
                    base_counters[prop] += 1
                    idx_num = base_counters[prop]
                    final_new_name = f"{prop}.{idx_num}.pdf"
                    is_dup = "Có"
                    stat_duplicates += 1
                    dept_stats[dept]['dup'] += 1
                else:
                    final_new_name = f"{prop}.pdf"
                    is_dup = "Không"

                if st == 'EXCEL':
                    stat_excel_matched += 1
                    dept_stats[dept]['excel'] += 1
                    match_label = ename
                else:
                    stat_outside_excel += 1
                    dept_stats[dept]['outside'] += 1
                    match_label = "(Ngoài danh mục Excel - Đặt theo tiêu đề PDF)"

                all_records.append([
                    dept, person, str(rel_dir), p.name, ocr_summary_short,
                    match_label, conf, final_new_name, is_dup, rsn
                ])
            else:
                stat_need_check += 1
                dept_stats[dept]['check'] += 1
                final_new_name = "(CẦN KIỂM TRA - CHƯA ĐỔI TÊN)"
                match_label = ename if ename else "(Chưa xác định)"
                is_dup = "Không"
                all_records.append([
                    dept, person, str(rel_dir), p.name, ocr_summary_short,
                    match_label, conf, final_new_name, is_dup, rsn
                ])
                check_records.append([
                    dept, person, str(rel_dir), p.name, ocr_summary_short,
                    match_label, conf, rsn
                ])

            ui_records.append({
                'rel_path': rel_s,
                'full_path': str(p),
                'dept': dept,
                'person': person,
                'rel_dir': str(rel_dir),
                'old_name': p.name,
                'status': st,
                'excel_id': cid,
                'match_label': match_label,
                'confidence': conf,
                'new_name': final_new_name,
                'is_dup': is_dup,
                'reason': rsn,
                'ocr_text': ocr_summary_short
            })

    # Ghi 3 file báo cáo Excel + file JSON kế hoạch
    save_reports_and_plan(
        report_dir, ui_records, all_records, check_records,
        stat_total, stat_excel_matched, stat_outside_excel, stat_duplicates, stat_need_check,
        dept_stats
    )

    elapsed = time.time() - t_start
    if progress_callback:
        progress_callback("done", 100, 100, f"Hoàn tất trong {elapsed:.1f}s! Nhận diện được {stat_excel_matched + stat_outside_excel:,}/{stat_total:,} file.")

    return {
        'cancelled': False,
        'total_pdfs': stat_total,
        'excel_matched': stat_excel_matched,
        'outside_excel': stat_outside_excel,
        'need_check': stat_need_check,
        'duplicates': stat_duplicates,
        'records': ui_records,
        'dept_stats': dict(dept_stats),
        'elapsed': elapsed,
        'report_dir': str(report_dir)
    }


def save_reports_and_plan(
    report_dir: pathlib.Path,
    ui_records: list,
    all_records: list = None,
    check_records: list = None,
    stat_total: int = 0,
    stat_excel_matched: int = 0,
    stat_outside_excel: int = 0,
    stat_duplicates: int = 0,
    stat_need_check: int = 0,
    dept_stats: dict = None
):
    report_dir = pathlib.Path(report_dir)
    report_dir.mkdir(parents=True, exist_ok=True)

    if all_records is None:
        all_records = []
        check_records = []
        stat_total = len(ui_records)
        stat_excel_matched = 0
        stat_outside_excel = 0
        stat_duplicates = 0
        stat_need_check = 0
        dept_stats = defaultdict(lambda: {'total': 0, 'excel': 0, 'outside': 0, 'check': 0, 'dup': 0})
        for r in ui_records:
            dept = r['dept']
            dept_stats[dept]['total'] += 1
            if r.get('is_dup') == 'Có':
                stat_duplicates += 1
                dept_stats[dept]['dup'] += 1
            if r['status'] == 'EXCEL' and r['new_name'] != "(CẦN KIỂM TRA - CHƯA ĐỔI TÊN)":
                stat_excel_matched += 1
                dept_stats[dept]['excel'] += 1
            elif r['status'] == 'OUTSIDE' and r['new_name'] != "(CẦN KIỂM TRA - CHƯA ĐỔI TÊN)":
                stat_outside_excel += 1
                dept_stats[dept]['outside'] += 1
            else:
                stat_need_check += 1
                dept_stats[dept]['check'] += 1
                check_records.append([
                    r['dept'], r['person'], r['rel_dir'], r['old_name'],
                    r['ocr_text'], r['match_label'], r['confidence'], r['reason']
                ])
            all_records.append([
                r['dept'], r['person'], r['rel_dir'], r['old_name'],
                r['ocr_text'], r['match_label'], r['confidence'],
                r['new_name'], r['is_dup'], r['reason']
            ])

    du_kien_path = report_dir / 'Danh_sach_du_kien_doi_ten.xlsx'
    can_kiem_tra_path = report_dir / 'Danh_sach_can_kiem_tra.xlsx'
    thong_ke_path = report_dir / 'Thong_ke.xlsx'
    plan_json_path = report_dir / 'rename_plan.json'
    ui_cache_json_path = report_dir / 'ui_records_cache.json'

    headers_du_kien = [
        'Phòng ban', 'Người', 'Thư mục tương đối', 'Tên file gốc',
        'Nội dung OCR nhận dạng', 'Tên Excel được đối chiếu',
        'Điểm tin cậy (%)', 'Tên file dự kiến', 'Trùng lặp trong thư mục', 'Ghi chú / Lý do'
    ]
    write_xlsx(du_kien_path, [('Danh_sach_du_kien', headers_du_kien, all_records)])

    headers_kiem_tra = [
        'Phòng ban', 'Người', 'Thư mục tương đối', 'Tên file gốc',
        'Nội dung OCR nhận dạng', 'Gợi ý gần nhất (nếu có)',
        'Điểm tin cậy (%)', 'Lý do cần kiểm tra'
    ]
    write_xlsx(can_kiem_tra_path, [('Danh_sach_can_kiem_tra', headers_kiem_tra, check_records)])

    summary_rows = [
        ['Tổng số file PDF đã quét', stat_total],
        ['Số file nhận dạng được (Tổng cộng)', stat_excel_matched + stat_outside_excel],
        ['  - Thuộc danh mục chuẩn Excel', stat_excel_matched],
        ['  - Không thuộc danh sách Excel (Đặt theo tiêu đề trang đầu PDF)', stat_outside_excel],
        ['Số file trùng lặp cùng loại trong thư mục cá nhân (được đánh số .1, .2...)', stat_duplicates],
        ['Số file cần kiểm tra (Chưa đổi tên)', stat_need_check],
    ]
    dept_rows = []
    for d_name, d_val in sorted(dept_stats.items()):
        dept_rows.append([
            d_name,
            d_val['total'],
            d_val['excel'] + d_val['outside'],
            d_val['excel'],
            d_val['outside'],
            d_val['dup'],
            d_val['check']
        ])
    write_xlsx(thong_ke_path, [
        ('Thong_ke_chung', ['Chỉ tiêu thống kê', 'Số lượng file PDF'], summary_rows),
        ('Theo_phong_ban', ['Phòng ban', 'Tổng số PDF', 'Tổng nhận dạng được', 'Trong danh sách Excel', 'Ngoài danh sách Excel', 'Số file trùng lặp (.1, .2...)', 'Số cần kiểm tra'], dept_rows)
    ])

    plan_data = []
    for r in ui_records:
        if r['new_name'] and r['new_name'] != "(CẦN KIỂM TRA - CHƯA ĐỔI TÊN)":
            plan_data.append({
                'rel_dir': r['rel_dir'],
                'old_name': r['old_name'],
                'new_name': r['new_name'],
                'confidence': r['confidence'],
                'reason': r['reason']
            })
    plan_json_path.write_text(json.dumps(plan_data, ensure_ascii=False, indent=2), encoding='utf-8')
    ui_cache_json_path.write_text(json.dumps(ui_records, ensure_ascii=False, indent=2), encoding='utf-8')


def load_saved_ui_records(report_dir: pathlib.Path, root_dir: pathlib.Path = None) -> list:
    """Đọc lại kết quả Dry-Run gần nhất để hiển thị ngay khi mở phần mềm."""
    report_dir = pathlib.Path(report_dir)
    ui_cache_json_path = report_dir / 'ui_records_cache.json'
    if ui_cache_json_path.exists():
        try:
            return json.loads(ui_cache_json_path.read_text(encoding='utf-8'))
        except Exception:
            pass

    # Fallback: đọc từ Danh_sach_du_kien_doi_ten.xlsx nếu đã có từ bản CLI trước đó
    du_kien_xlsx = report_dir / 'Danh_sach_du_kien_doi_ten.xlsx'
    if du_kien_xlsx.exists():
        try:
            rows = read_raw_xlsx_rows(du_kien_xlsx)
            if len(rows) > 1:
                records = []
                r_dir = pathlib.Path(root_dir) if root_dir else _DEFAULT_ROOT_DIR
                for r in rows[1:]:
                    if len(r) < 10:
                        r = r + [''] * (10 - len(r))
                    dept, person, rel_dir_s, old_name, ocr_sum, match_lbl, conf_s, new_name, is_dup, rsn = r[:10]
                    try:
                        conf = float(conf_s)
                    except Exception:
                        conf = 0.0
                    if new_name == "(CẦN KIỂM TRA - CHƯA ĐỔI TÊN)":
                        st = 'CHECK'
                    elif "Ngoài danh mục Excel" in match_lbl:
                        st = 'OUTSIDE'
                    else:
                        st = 'EXCEL'
                    full_p = r_dir / rel_dir_s / old_name
                    records.append({
                        'rel_path': str(pathlib.Path(rel_dir_s) / old_name),
                        'full_path': str(full_p),
                        'dept': dept,
                        'person': person,
                        'rel_dir': rel_dir_s,
                        'old_name': old_name,
                        'status': st,
                        'excel_id': None,
                        'match_label': match_lbl,
                        'confidence': conf,
                        'new_name': new_name,
                        'is_dup': is_dup,
                        'reason': rsn,
                        'ocr_text': ocr_sum
                    })
                return records
        except Exception:
            pass
    return []


# ==============================================================================
# 6. GIAI ĐOẠN THỰC THI ĐỔI TÊN AN TOÀN & HOÀN TÁC (ROLLBACK)
# ==============================================================================

def execute_rename_engine(root_dir: pathlib.Path, report_dir: pathlib.Path) -> dict:
    root_dir = pathlib.Path(root_dir)
    report_dir = pathlib.Path(report_dir)
    plan_json_path = report_dir / 'rename_plan.json'
    if not plan_json_path.exists():
        raise FileNotFoundError("Không tìm thấy file kế hoạch rename_plan.json! Hãy chạy 'Quét & Dự kiến đổi tên (Dry-Run)' trước.")

    plan_data = json.loads(plan_json_path.read_text(encoding='utf-8'))
    log_entries = []
    skipped_entries = []

    by_dir = defaultdict(list)
    for item in plan_data:
        if item['old_name'] != item['new_name'] and item['new_name'] != "(CẦN KIỂM TRA - CHƯA ĐỔI TÊN)":
            by_dir[item['rel_dir']].append(item)

    for rel_dir_s, items in sorted(by_dir.items()):
        folder = root_dir / rel_dir_s
        if not folder.exists():
            continue

        planned_targets = Counter(it['new_name'].lower() for it in items)
        existing_files = {f.name.lower(): f.name for f in folder.iterdir() if f.is_file()}
        old_set_lower = {it['old_name'].lower() for it in items}

        # Bước 1: Chuyển sang tên tạm
        temp_moves = []
        for idx, it in enumerate(items):
            src = folder / it['old_name']
            dst = folder / it['new_name']
            if not src.exists():
                skipped_entries.append({**it, 'error': 'File nguồn không tồn tại'})
                continue
            if planned_targets[it['new_name'].lower()] > 1:
                skipped_entries.append({**it, 'error': 'Trùng tên đích trong kế hoạch'})
                continue
            if it['new_name'].lower() in existing_files and it['new_name'].lower() not in old_set_lower:
                skipped_entries.append({**it, 'error': 'File đích đã tồn tại trên đĩa'})
                continue

            tmp_name = f"__tmp_sohoa_{idx}_{it['old_name']}"
            tmp_path = folder / tmp_name
            try:
                src.rename(tmp_path)
                temp_moves.append((tmp_path, dst, src, it))
            except Exception as e:
                skipped_entries.append({**it, 'error': str(e)})

        # Bước 2: Chuyển từ tên tạm sang tên chuẩn cuối cùng
        for tmp_path, dst, orig_src, it in temp_moves:
            if dst.exists():
                tmp_path.rename(orig_src)
                skipped_entries.append({**it, 'error': 'Phát hiện xung đột ở bước 2, đã giữ nguyên tên gốc'})
            else:
                try:
                    tmp_path.rename(dst)
                    log_entries.append({
                        'rel_dir': rel_dir_s,
                        'old_name': it['old_name'],
                        'new_name': it['new_name'],
                        'old_full_path': str(orig_src),
                        'new_full_path': str(dst),
                        'timestamp': time.strftime('%Y-%m-%d %H:%M:%S')
                    })
                except Exception as e:
                    tmp_path.rename(orig_src)
                    skipped_entries.append({**it, 'error': str(e)})

    log_json_path = report_dir / 'nhat_ky_doi_ten.json'
    already_named_count = sum(1 for it in plan_data if it['old_name'] == it['new_name'])
    log_payload = {
        'renamed_count': len(log_entries),
        'already_named_count': already_named_count,
        'skipped_count': len(skipped_entries),
        'renamed': log_entries,
        'skipped': skipped_entries
    }
    log_json_path.write_text(json.dumps(log_payload, ensure_ascii=False, indent=2), encoding='utf-8')

    log_xlsx_path = report_dir / 'Nhat_ky_doi_ten.xlsx'
    log_rows = [
        [e['rel_dir'], e['old_name'], e['new_name'], e['old_full_path'], e['new_full_path'], e['timestamp']]
        for e in log_entries
    ]
    write_xlsx(log_xlsx_path, [(
        'Nhat_ky_doi_ten',
        ['Thư mục tương đối', 'Tên file cũ', 'Tên file mới', 'Đường dẫn cũ đầy đủ', 'Đường dẫn mới đầy đủ', 'Thời gian đổi tên'],
        log_rows
    )])

    return log_payload


def rollback_rename_engine(report_dir: pathlib.Path) -> dict:
    report_dir = pathlib.Path(report_dir)
    log_file = report_dir / 'nhat_ky_doi_ten.json'
    if not log_file.exists():
        raise FileNotFoundError(f"Không tìm thấy file nhật ký đổi tên tại: {log_file}")

    data = json.loads(log_file.read_text(encoding='utf-8'))
    entries = data.get('renamed', [])
    if not entries:
        return {'total': 0, 'restored': 0}

    temp_moves = []
    for idx, item in enumerate(reversed(entries)):
        new_p = pathlib.Path(item['new_full_path'])
        old_p = pathlib.Path(item['old_full_path'])
        if new_p.exists():
            tmp_p = new_p.parent / f"__rb_tmp_{idx}_{new_p.name}"
            try:
                new_p.rename(tmp_p)
                temp_moves.append((tmp_p, old_p))
            except Exception:
                pass

    restored = 0
    for tmp_p, old_p in temp_moves:
        if not old_p.exists():
            try:
                tmp_p.rename(old_p)
                restored += 1
            except Exception:
                pass
    return {'total': len(entries), 'restored': restored}
