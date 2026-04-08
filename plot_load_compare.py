#!/usr/bin/env python3
"""绘制市场披露预测统调负荷 vs 实际统调负荷（同图）。

- 预测：虚线
- 实际：实线

不依赖 pandas/matplotlib，仅使用 Python 标准库读取 xlsx 并生成 SVG。
"""

from __future__ import annotations

import argparse
import math
import re
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _col_to_idx(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref).group(0)
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def read_sheet_rows(xlsx_path: Path, sheet_name: str) -> list[list[str]]:
    with ZipFile(xlsx_path) as zf:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si", NS):
                text = "".join(t.text or "" for t in si.findall(".//a:t", NS))
                shared_strings.append(text)

        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

        target = None
        for sh in wb.findall("a:sheets/a:sheet", NS):
            if sh.attrib.get("name") == sheet_name:
                rid = sh.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
                target = "xl/" + rid_to_target[rid]
                break
        if not target:
            raise ValueError(f"未找到工作表: {sheet_name}")

        ws = ET.fromstring(zf.read(target))

        rows: list[list[str]] = []
        for row in ws.findall("a:sheetData/a:row", NS):
            vals: dict[int, str] = {}
            max_idx = -1
            for c in row.findall("a:c", NS):
                ref = c.attrib.get("r", "A1")
                col_idx = _col_to_idx(ref)
                max_idx = max(max_idx, col_idx)
                t = c.attrib.get("t")
                v = c.find("a:v", NS)
                if v is None:
                    vals[col_idx] = ""
                else:
                    raw = v.text or ""
                    vals[col_idx] = shared_strings[int(raw)] if t == "s" else raw
            if max_idx >= 0:
                rows.append([vals.get(i, "") for i in range(max_idx + 1)])

        return rows


def pick_series(rows: list[list[str]], key_col_idx: int, key_name: str) -> tuple[list[str], list[float]]:
    header = rows[0]
    time_cols = header[2:98]
    for row in rows[1:]:
        if len(row) > key_col_idx and row[key_col_idx].strip() == key_name:
            values = [float(x) for x in row[2:98]]
            return time_cols, values
    raise ValueError(f"未找到指标: {key_name}")


def draw_svg(times: list[str], pred: list[float], actual: list[float], output: Path) -> None:
    width, height = 1400, 700
    m_left, m_right, m_top, m_bottom = 90, 40, 60, 90
    plot_w = width - m_left - m_right
    plot_h = height - m_top - m_bottom

    y_min = min(min(pred), min(actual))
    y_max = max(max(pred), max(actual))

    span = y_max - y_min
    y_min -= span * 0.06
    y_max += span * 0.06

    def x_map(i: int) -> float:
        return m_left + i * plot_w / (len(times) - 1)

    def y_map(v: float) -> float:
        return m_top + (y_max - v) * plot_h / (y_max - y_min)

    def series_to_poly(vals: list[float]) -> str:
        return " ".join(f"{x_map(i):.2f},{y_map(v):.2f}" for i, v in enumerate(vals))

    y_ticks = 8
    x_tick_idx = list(range(0, len(times), 8))
    if x_tick_idx[-1] != len(times) - 1:
        x_tick_idx.append(len(times) - 1)

    svg = []
    svg.append(f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">')
    svg.append('<rect width="100%" height="100%" fill="white"/>')
    svg.append('<style>text{font-family:Arial,"Microsoft YaHei",sans-serif;} .grid{stroke:#e8e8e8;stroke-width:1;} .axis{stroke:#222;stroke-width:1.5;}</style>')

    # 网格 & y轴标签
    for i in range(y_ticks + 1):
        yv = y_min + (y_max - y_min) * i / y_ticks
        yp = y_map(yv)
        svg.append(f'<line class="grid" x1="{m_left}" y1="{yp:.2f}" x2="{width - m_right}" y2="{yp:.2f}"/>')
        svg.append(f'<text x="{m_left - 10}" y="{yp + 4:.2f}" font-size="12" text-anchor="end" fill="#444">{yv:.0f}</text>')

    # x刻度
    for i in x_tick_idx:
        xp = x_map(i)
        svg.append(f'<line class="grid" x1="{xp:.2f}" y1="{m_top}" x2="{xp:.2f}" y2="{height - m_bottom}"/>')
        svg.append(f'<text x="{xp:.2f}" y="{height - m_bottom + 22}" font-size="12" text-anchor="middle" fill="#444">{times[i]}</text>')

    # 坐标轴
    svg.append(f'<line class="axis" x1="{m_left}" y1="{m_top}" x2="{m_left}" y2="{height - m_bottom}"/>')
    svg.append(f'<line class="axis" x1="{m_left}" y1="{height - m_bottom}" x2="{width - m_right}" y2="{height - m_bottom}"/>')

    # 线条
    svg.append(f'<polyline fill="none" stroke="#1f77b4" stroke-width="2.5" stroke-dasharray="8 6" points="{series_to_poly(pred)}"/>')
    svg.append(f'<polyline fill="none" stroke="#d62728" stroke-width="2.5" points="{series_to_poly(actual)}"/>')

    # 标题
    svg.append(f'<text x="{width/2:.1f}" y="32" text-anchor="middle" font-size="20" font-weight="bold">市场披露统调负荷：预测 vs 实际</text>')
    svg.append(f'<text x="{width/2:.1f}" y="52" text-anchor="middle" font-size="12" fill="#666">预测(虚线) | 实际(实线)</text>')

    # 图例
    lx, ly = width - 290, 72
    svg.append(f'<rect x="{lx}" y="{ly}" width="250" height="58" rx="6" fill="#fff" stroke="#ddd"/>')
    svg.append(f'<line x1="{lx+14}" y1="{ly+20}" x2="{lx+74}" y2="{ly+20}" stroke="#1f77b4" stroke-width="2.5" stroke-dasharray="8 6"/>')
    svg.append(f'<text x="{lx+82}" y="{ly+24}" font-size="13">预测统调负荷</text>')
    svg.append(f'<line x1="{lx+14}" y1="{ly+42}" x2="{lx+74}" y2="{ly+42}" stroke="#d62728" stroke-width="2.5"/>')
    svg.append(f'<text x="{lx+82}" y="{ly+46}" font-size="13">实际统调负荷</text>')

    # 轴标题
    svg.append(f'<text x="{width/2:.1f}" y="{height-22}" text-anchor="middle" font-size="14">时间（15分钟）</text>')
    svg.append(f'<text x="24" y="{height/2:.1f}" transform="rotate(-90 24,{height/2:.1f})" text-anchor="middle" font-size="14">负荷 (MW)</text>')

    svg.append('</svg>')
    output.write_text("\n".join(svg), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pred-file", default="披露数据4.2更新/信息披露查询预测信息(2026-04-03).xlsx")
    parser.add_argument("--actual-file", default="披露数据4.2更新/信息披露查询实际信息(2026-04-01).xlsx")
    parser.add_argument("--output", default="统调负荷_预测_vs_实际.svg")
    args = parser.parse_args()

    pred_rows = read_sheet_rows(Path(args.pred_file), "负荷预测信息(2026-04-03)")
    actual_rows = read_sheet_rows(Path(args.actual_file), "机组出力情况(2026-04-01)")

    times, pred_series = pick_series(pred_rows, key_col_idx=1, key_name="统调负荷(MW)")
    _, actual_series = pick_series(actual_rows, key_col_idx=0, key_name="统调系统实际负荷(MW)")

    draw_svg(times, pred_series, actual_series, Path(args.output))
    print(f"已生成图表: {args.output}")


if __name__ == "__main__":
    main()
