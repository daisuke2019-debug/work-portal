# -*- coding: utf-8 -*-
"""デスクトップの「◯月マスタミス」フォルダ（10月マスタミス、11月マスタミス…）にある
*_resultMMDD.xlsx を全自動検出して portal_json_data.json を再生成する。
各店舗データには month（例 2026-11）を付け、ポータルで月ごとに切り替えられるようにする。"""
import openpyxl, json, os, re, sys, datetime, ctypes, ctypes.wintypes
sys.stdout.reconfigure(encoding='utf-8')

DESKTOP = r'C:\Users\marum\OneDrive\デスクトップ'
# 「10月マスタミス」「11月マスタミス」のような名前のフォルダを全部読む（12月以降もフォルダを作るだけでよい）
SRC_DIR_PATTERN = re.compile(r'^\d{1,2}月マスタミス$')
# 一時ファイルと集計JSONは data/ に置く（Gitには入らない）
WORK_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'data')

def safe_copy(src, dst):
    k = ctypes.windll.kernel32
    k.CreateFileW.restype = ctypes.wintypes.HANDLE
    h = k.CreateFileW(src, 0x80000000, 0x7, None, 3, 0x80, None)
    if h in (-1, 0xFFFFFFFFFFFFFFFF, None):
        raise PermissionError(src)
    o = k.CreateFileW(dst, 0x40000000, 0, None, 2, 0x80, None)
    buf = ctypes.create_string_buffer(65536)
    r = ctypes.wintypes.DWORD(); w = ctypes.wintypes.DWORD()
    while k.ReadFile(h, buf, 65536, ctypes.byref(r), None) and r.value:
        k.WriteFile(o, buf, r.value, ctypes.byref(w), None)
    k.CloseHandle(h); k.CloseHandle(o)

def parse_file(path, date_str):
    tmp = os.path.join(WORK_DIR, 'tmp_scan.xlsx')
    safe_copy(path, tmp)
    wb = openpyxl.load_workbook(tmp, data_only=True)
    sh = wb['店舗マスタ'] if '店舗マスタ' in wb.sheetnames else wb[wb.sheetnames[0]]
    headers = [str(sh.cell(row=1, column=c).value or f'Col{c}').strip() for c in range(1, sh.max_column + 1)]
    # 11/5: 列は見出しの名前で探す（11月から先頭に「店舗コード」列が増え、位置がずれたため）。
    # 見出しが無い古い形式は、これまでどおりの位置（1列目=店舗・4列目=最終更新スタッフ・5列目）を使う。
    def col_of(name, default):
        return headers.index(name) + 1 if name in headers else default
    c_shop = col_of('店舗', 1)
    c_staff = col_of('最終更新スタッフ', 4)
    c_sv = col_of('シフト稼働日', 5)
    c_code = col_of('店舗コード', None)
    shop_cols = {c_shop} | ({c_code} if c_code else set())  # 店舗名・店舗コードの赤＝「A列赤（マスタミス）」扱い
    shops = []
    for r in range(2, sh.max_row + 1):
        v = sh.cell(row=r, column=c_shop).value
        if not v or str(v).strip() == '店舗':
            continue
        staff = str(sh.cell(row=r, column=c_staff).value or '未設定').strip()
        sv = str(sh.cell(row=r, column=c_sv).value or '').strip()
        is_a_red, non_a = False, []
        for c in range(1, sh.max_column + 1):
            f = sh.cell(row=r, column=c).fill
            if f and f.fill_type and f.fill_type != 'none':
                fg = str(getattr(f.fgColor, 'rgb', '') or getattr(f.start_color, 'rgb', '') or '').upper()
                if 'FF0000' in fg or fg == 'RED':
                    if c in shop_cols:
                        is_a_red = True
                    else:
                        non_a.append(headers[c - 1])
        shops.append({'name': str(v).strip(), 'staff': staff, 'sv': sv, 'date': date_str,
                      'is_a_red': is_a_red, 'non_a_count': len(non_a), 'non_a_items': non_a})
    wb.close()
    os.remove(tmp)
    return shops

def month_key(fn, mm):
    """ファイル名の作成日時(shop_masters_YYYYMM...)から年を決め、データの月(mm)と合わせて 'YYYY-MM' を返す"""
    m = re.search(r'shop_masters_(\d{4})(\d{2})', fn)
    if m:
        year, file_mm = int(m.group(1)), int(m.group(2))
        if int(mm) > file_mm:  # 例: 1月に作られた12月分のファイル → 前の年
            year -= 1
    else:
        year = datetime.date.today().year
    return f'{year}-{mm}'

files = {}  # 'YYYY-MM-DD' -> (ファイルの場所, 'MM/DD', 'YYYY-MM')
for folder in sorted(os.listdir(DESKTOP)):
    src = os.path.join(DESKTOP, folder)
    if not (SRC_DIR_PATTERN.match(folder) and os.path.isdir(src)):
        continue
    for fn in os.listdir(src):
        m = re.search(r'_result(\d{2})(\d{2})\.xlsx$', fn, re.I)
        if m:
            mk = month_key(fn, m.group(1))
            files[f'{mk}-{m.group(2)}'] = (os.path.join(src, fn), f'{m.group(1)}/{m.group(2)}', mk)

staff_data, col_ranking = {}, {}
for key in sorted(files):
    path, d, mk = files[key]
    shops = parse_file(path, d)
    print(f'{d}: {len(shops)}店舗 ({os.path.basename(os.path.dirname(path))}/{os.path.basename(path)})')
    for it in shops:
        it['month'] = mk
        s = staff_data.setdefault(it['staff'], {'a_count': 0, 'non_a_red_total': 0, 'shops': []})
        s['a_count'] += it['is_a_red']
        s['non_a_red_total'] += it['non_a_count']
        for col in it['non_a_items']:
            col_ranking[col] = col_ranking.get(col, 0) + 1
        s['shops'].append(it)

for info in staff_data.values():
    t = len(info['shops'])
    info['miss_rate'] = round(info['a_count'] / t * 100, 1) if t else 0.0

with open(os.path.join(WORK_DIR, 'portal_json_data.json'), 'w', encoding='utf-8') as f:
    json.dump({'staff_data': staff_data, 'col_ranking': col_ranking,
               'available_dates': [files[k][1] for k in sorted(files)],
               'months': sorted({v[2] for v in files.values()})}, f, ensure_ascii=False, indent=2, default=str)
print('合計:', sum(len(v['shops']) for v in staff_data.values()), '店舗')
