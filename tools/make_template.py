#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Write docs/import_template.xlsx -- a ready-to-edit Excel import template."""

import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Side, Border
from openpyxl.utils import get_column_letter

HEADERS = ["姓名", "收藏", "手机", "邮箱", "微信", "QQ", "地址", "其他联系方式", "备注"]
ROWS = [
    ["李思远", "是", "13711112222", "lisy@fzu.edu.cn", "lisy_wx", "", "福州市大学城", "", "示例行，可删除"],
    ["周晓彤", "", "13733334444；0591-22860000", "zhoux@example.com", "", "20250101", "", "钉钉：zhoux", "一条联系方式用“；”分隔"],
]

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "通讯录"
    ws.append(HEADERS)

    head_fill = PatternFill("solid", fgColor="1F6FEB")
    thin = Side(style="thin", color="D0D7DE")
    for cell in ws[1]:
        cell.fill = head_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ROWS:
        ws.append(row)

    # 说明放在单独的 sheet，避免被当成联系人导入
    help_ws = wb.create_sheet("填写说明")
    notes = [
        "1. 一行一个联系人，“姓名”必填。",
        "2. “收藏”填“是”表示收藏，留空表示不收藏。",
        "3. 同一个人的多个同类联系方式用“；”分隔，例如两个手机号。",
        "4. 列的顺序可以打乱，程序按表头名称识别；多余的列会被忽略。",
        "5. 导入时按姓名去重：已存在的联系人只追加新的联系方式，不会覆盖原有数据。",
    ]
    for i, text in enumerate(notes, start=1):
        help_ws.cell(row=i, column=1, value=text).font = Font(color="33415C", size=11)
    help_ws.column_dimensions["A"].width = 80

    for idx, width in enumerate([12, 7, 22, 26, 14, 12, 24, 18, 26], start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"

    target = os.path.join(BASE, "docs", "import_template.xlsx")
    os.makedirs(os.path.dirname(target), exist_ok=True)
    wb.save(target)
    print("wrote", target)


if __name__ == "__main__":
    main()
