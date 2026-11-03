"""SRP現場ナビ: ドコモ代理店早見表を月次Excelから更新する。

使い方:
  python scripts/srp_update.py "<Excelのパス>"            # 更新（ファイルを書き換える。pushはしない）
  python scripts/srp_update.py "<Excelのパス>" --dry-run  # 差分レポートだけ表示
  python scripts/srp_update.py "<Excelのパス>" --init     # 現在のページ内容=このExcel として基準データだけ保存

仕様:
  - Excelの「DCM」を含むシートだけを使う（KDDIシートは使わない）
  - 「担当」列（個人名）は読み込まない・保存しない・公開しない
  - 前月から備考が変わっていない行は、ページの既存HTML（手で整えた表記）をそのまま使い、開始日だけ更新
  - 新規行・備考が変わった行は自動整形し、レポートで【要確認】として出す
  - 「特別施策ヒアリング可否」シートで×の施策は、自動整形行に「ヒアリングNG」を追記
  - トップのKPIカードの法人数（重複除く代理店数）と早見表見出しの月を更新
  - 基準データ data/srp_last.json（gitignore対象）に今回の内容を保存
"""
import datetime
import html
import json
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
PAGE = ROOT / "site" / "srp" / "index.html"
SNAP = ROOT / "data" / "srp_last.json"
STD = "通常ルール適用（金額NG / 専門用語NG）"
RED = "#B91C1C"
ORANGE = "#C2410C"


def to_md(v):
    if isinstance(v, (int, float)):
        v = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(v))
    if isinstance(v, (datetime.datetime, datetime.date)):
        return f"{v.month}/{v.day}"
    return str(v).strip() if v is not None else ""


def read_excel(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = next((w for w in wb if "DCM" in w.title), None)
    if ws is None:
        sys.exit("DCMシートが見つかりません: " + ", ".join(w.title for w in wb))
    # 見出し行（「代理店」「備考」）を探す
    hdr = None
    for r in ws.iter_rows(min_row=1, max_row=10):
        vals = {str(c.value).strip(): c.column for c in r if c.value is not None}
        if "代理店" in vals and "備考" in vals:
            hdr = (r[0].row, vals)
            break
    if not hdr:
        sys.exit("DCMシートに「代理店」「備考」の見出し行がありません")
    hrow, cols = hdr
    c_kb = cols.get("広域/地場")
    c_ag = cols["代理店"]
    c_ty = c_ag + 1  # 代理店の右隣が 直営/二次/-
    c_dt = cols.get("ヒアリング可否")
    c_nt = cols["備考"]
    month = None
    for r in ws.iter_rows(min_row=1, max_row=2):
        for c in r:
            m = re.search(r"【(\d+)/1", str(c.value or ""))
            if m:
                month = int(m.group(1))
    rows, kb, ag = [], "", ""
    for r in range(hrow + 1, ws.max_row + 1):
        g = lambda col: ws.cell(r, col).value if col else None
        if g(c_kb):
            kb = str(g(c_kb)).strip()
        if g(c_ag):
            ag = str(g(c_ag)).strip()
        ty = g(c_ty)
        if ty is None and g(c_dt) is None:
            continue
        rows.append({"kb": kb, "agency": ag, "type": str(ty or "-").strip(),
                     "date": to_md(g(c_dt)), "note": str(g(c_nt) or "").strip()})
    # 特別施策ヒアリング可否（DCM分）
    ng = {}
    sp = next((w for w in wb if "特別施策" in w.title), None)
    if sp is not None:
        head = None
        for r in sp.iter_rows(min_row=1, max_row=10):
            vals = {str(c.value).strip(): c.column for c in r if c.value is not None}
            if "代理店" in vals:
                head = (r[0].row, vals)
                break
        if head:
            hr, hv = head
            ca = hv["代理店"]
            items = [(k, v) for k, v in hv.items() if v > ca + 1]  # 代理店握り施策 / FFF8施策 / その他
            # DCM/KDDIで同名代理店（ラネット等）があるため、最初に出現した組（DCMが先）を優先
            first, a = {}, ""
            for r in range(hr + 1, sp.max_row + 1):
                if sp.cell(r, ca).value:
                    a = str(sp.cell(r, ca).value).strip()
                t = str(sp.cell(r, ca + 1).value or "-").strip()
                if (a, t) in first:
                    continue
                first[(a, t)] = [k for k, col in items
                                 if str(sp.cell(r, col).value or "").strip() == "×"]
            ng = first
    for x in rows:
        x["ng"] = ng.get((x["agency"], x["type"]), [])
    return month, rows


def who(note):
    m = re.search(r"ヒアリング対象：([^\s　\n]+)", note)
    t = m.group(1) if m else ""
    if not t and re.search(r"ヒアリングは店長のみ", note):
        t = "店長のみ"
    if t in ("店長/副店長", "制限なし"):
        return f'<span style="color:{RED}; font-weight:800;">店長 / 副店長</span>'
    if t in ("店長のみ", "店長"):
        return f'<span style="color:{RED}; font-weight:800;">⭐ {t}</span>'
    return "常勤FSM誰でも"


HL = [(r"金額の?(?:話|コミュニケーション)は(?:絶対)?NG", RED),
      (r"SRP(?:等|など|といった)の?専門用語での(?:ヒアリング|コミュニケーション)は(?:NG|不可)", ORANGE)]


def fmt_note(x):
    note = x["note"]
    if not note:
        body = STD
    else:
        body = html.escape(note).replace("\n", "<br>")
        for pat, col in HL:
            body = re.sub(pat, lambda m: f'<b style="color:{col};">{m.group(0)}</b>', body)
    if x["ng"]:
        body += f'<br><b style="color:{RED};">※{"・".join(x["ng"])}はヒアリングNG</b>（施策自体は実施中）'
    return body


def gen_row(x):
    return (f'            <tr><td>{x["kb"]}</td><td><b>{html.escape(x["agency"])}</b></td>'
            f'<td>{x["type"]}</td><td>{x["date"]}</td><td>{who(x["note"])}</td>'
            f'<td>{fmt_note(x)}</td></tr>\n')


def page_rows(h):
    s = h.index('id="agencyTable"')
    b = h.index("<tbody>\n", s) + len("<tbody>\n")
    e = h.index("</tbody>", b)
    rows = re.findall(r"[ \t]*<tr>.*?</tr>\n", h[b:e])
    out = {}
    for r in rows:
        c = re.findall(r"<td>(.*?)</td>", r)
        out[(re.sub("<.*?>", "", c[1]), c[2])] = r
    return b, e, out


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dry, init = "--dry-run" in sys.argv, "--init" in sys.argv
    if not args:
        sys.exit(__doc__)
    month, rows = read_excel(args[0])
    snap_rows = [{k: x[k] for k in ("kb", "agency", "type", "date", "note", "ng")} for x in rows]
    if init:
        SNAP.parent.mkdir(exist_ok=True)
        SNAP.write_text(json.dumps({"month": month, "rows": snap_rows}, ensure_ascii=False, indent=1), "utf-8")
        print(f"基準データを保存: {SNAP}（{len(rows)}行）")
        return
    if not SNAP.exists():
        sys.exit("基準データ data/srp_last.json がありません。先に前月Excelで --init してください。")
    prev = {(x["agency"], x["type"]): x for x in json.loads(SNAP.read_text("utf-8"))["rows"]}
    h = PAGE.read_text("utf-8")
    b, e, old = page_rows(h)

    rep = {"そのまま（日付のみ更新）": [], "【要確認】新規追加": [], "【要確認】備考変更": [], "削除": []}
    out = []
    for x in rows:
        k = (x["agency"], x["type"])
        p = prev.get(k)
        if k in old and p and p["note"] == x["note"] and p.get("ng", []) == x["ng"]:
            r = re.sub(r"(<td>[^<]*</td><td><b>.*?</b></td><td>[^<]*</td><td>)[^<]*(</td>)",
                       lambda m: m.group(1) + x["date"] + m.group(2), old[k], count=1)
            out.append(r)
            rep["そのまま（日付のみ更新）"].append(k)
        else:
            out.append(gen_row(x))
            rep["【要確認】新規追加" if k not in old else "【要確認】備考変更"].append(k)
    keys = {(x["agency"], x["type"]) for x in rows}
    rep["削除"] = [k for k in old if k not in keys]
    odd = [x for x in rows if not re.fullmatch(r"\d+/\d+", x["date"])]

    n_ag = len({x["agency"] for x in rows})
    nh = h[:b] + "".join(out) + "          " + h[e:]
    if month:
        nh = re.sub(r"📋 \d+月SRP ドコモ代理店", f"📋 {month}月SRP ドコモ代理店", nh)
        nh = re.sub(r"<!-- 🏢 \d+月SRP 代理店別 早見表", f"<!-- 🏢 {month}月SRP 代理店別 早見表", nh)
    nh = re.sub(r'(<div class="kpi-val">)\d+ 法人', rf"\g<1>{n_ag} 法人", nh)
    nh = re.sub(r"ヒアリング対象代理店（全\d+社）", f"ヒアリング対象代理店（全{n_ag}社）", nh)

    print(f"== {month}月 DCM: {len(rows)}行 / {n_ag}法人 ==")
    for t, ks in rep.items():
        print(f"{t}: {len(ks)}")
        if not t.startswith("そのまま"):
            for k in ks:
                print("   -", k[0], k[1])
    for x in odd:
        print(f"【要確認】開始日が日付でない: {x['agency']} {x['type']} = {x['date']!r}")
    if dry:
        print("(dry-run: ファイルは変更していません)")
        return
    PAGE.write_text(nh, "utf-8", newline="")
    SNAP.write_text(json.dumps({"month": month, "rows": snap_rows}, ensure_ascii=False, indent=1), "utf-8")
    print(f"更新しました: {PAGE}\n基準データ更新: {SNAP}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
