# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "openpyxl",
# ]
# ///
"""把「表2-2-歷年中華民國國民出國目的地人次統計.xlsx」整理成 docs/data.js。

網頁以 <script src="data.js"> 載入，資料掛在 window.OUTBOUND_DATA，
雙擊 index.html（file://）即可使用，不需要伺服器。
原始 xlsx 僅以唯讀方式開啟，不會被修改。
"""

import json
import re
from datetime import date
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "表2-2-歷年中華民國國民出國目的地人次統計.xlsx"
OUTPUT = ROOT / "docs" / "data.js"

CHECK_YEAR_ROC = 112
CHECK_TOTAL = 11_795_834


def split_name(text) -> tuple[str, str]:
    """把「香港\\nHong Kong」這類表頭拆成（中文名, 英文名）"""
    if not text:
        return "", ""
    s = " ".join(str(text).split())
    zh = "".join(re.findall(r"[一-龥]+", s))
    en = re.sub(r"[一-龥]+", " ", s).strip()
    en = " ".join(en.split())
    return (zh or en), en


def to_int(val):
    if val is None or val == "":
        return None
    try:
        return int(val)
    except (ValueError, TypeError):
        return None


def read_sheet(path: Path):
    ws = openpyxl.load_workbook(path, data_only=True, read_only=True).worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    header_region, header_country = rows[1], rows[2]

    # 欄位 → (洲別, 國家)；第 2 列洲別為跨欄合併，往右延續
    columns = {}
    region = ("其他", "Others")
    for c in range(1, len(header_country)):
        if header_region[c]:
            region = split_name(header_region[c])
        name = split_name(header_country[c] or header_region[c])
        if name[0]:
            columns[c] = (region, name)

    def is_total(name):
        return bool(re.search("合計|小計|總計", name))

    # 「X洲其他地區」與「總計」是公式欄，檔案沒有快取值 → 由各洲合計回推
    others = {}
    for c, (reg, (name, _)) in columns.items():
        if "其他地區" in name:
            same = [k for k, (r, _) in columns.items() if r == reg]
            total_col = next((k for k in same if is_total(columns[k][1][0])), None)
            if total_col is None:
                raise ValueError(f"找不到 {reg[0]} 的合計欄")
            others[c] = (total_col, [k for k in same if k not in (c, total_col)])

    data = []  # (民國年, {欄: 人次})
    for row in rows[3:]:
        m = re.match(r"\s*(\d+)年", str(row[0] or ""))
        if not m:
            continue
        vals = {c: to_int(row[c]) for c in columns}
        for c, (total_col, listed) in others.items():
            if vals[c] is None and vals[total_col] is not None:
                vals[c] = vals[total_col] - sum(vals[k] or 0 for k in listed)
        data.append((int(m.group(1)), vals))

    region_totals = {c for c, (_, (n, _)) in columns.items() if is_total(n)}
    return columns, region_totals, data


def main():
    columns, total_cols, data = read_sheet(SOURCE)
    dest_cols = [c for c in columns if c not in total_cols]

    regions = []
    for c in dest_cols:
        reg = columns[c][0]
        if reg not in regions:
            regions.append(reg)

    years_roc = [y for y, _ in data]
    destinations = []
    for c in dest_cols:
        reg, (name, en) = columns[c]
        series = [vals[c] for _, vals in data]
        destinations.append({
            "name": name,
            "aliases": [en] if en else [],
            "region": regions.index(reg),
            "counts": series,
        })

    # 自我核對：每年各目的地加總 = 各洲合計 + 其他
    for y, vals in data:
        dest_sum = sum(vals[c] or 0 for c in dest_cols)
        expected = sum(vals[c] or 0 for c in total_cols) + sum(
            vals[c] or 0 for c in dest_cols if columns[c][0][0] == "其他"
        )
        assert dest_sum == expected, f"{y} 年加總不一致：{dest_sum} != {expected}"

    payload = {
        "meta": {
            "title": "歷年中華民國國民出國目的地人次統計",
            "source": SOURCE.name,
            "unit": "人次",
            "note": "目的地以出境班機第一站為準；「X洲其他地區」由該洲合計減去已列出國家回推。",
            "generated": date.today().isoformat(),
        },
        "years": [y + 1911 for y in years_roc],
        "yearsRoc": years_roc,
        "regions": [{"name": n, "aliases": [e] if e else []} for n, e in regions],
        "destinations": destinations,
    }

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    OUTPUT.write_text(f"window.OUTBOUND_DATA={body};\n", encoding="utf-8")

    idx = years_roc.index(CHECK_YEAR_ROC)
    total = sum(d["counts"][idx] or 0 for d in destinations)
    status = "符合" if total == CHECK_TOTAL else f"不符（差 {total - CHECK_TOTAL:+,}）"
    print(f"已產出 {OUTPUT.relative_to(ROOT)}（{OUTPUT.stat().st_size:,} bytes）")
    print(f"{len(years_roc)} 年（民國 {years_roc[0]}–{years_roc[-1]}）、"
          f"{len(regions)} 個洲別、{len(destinations)} 個目的地")
    print(f"核對 {CHECK_YEAR_ROC} 年總人次：{total:,}，預期 {CHECK_TOTAL:,} → {status}")


if __name__ == "__main__":
    main()
