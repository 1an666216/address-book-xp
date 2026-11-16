#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Populate data/contacts.json with a few demo contacts (idempotent)."""

from app import load_contacts, save_contacts, normalize_contact

SAMPLE = [
    {
        "name": "林晓雨",
        "bookmarked": True,
        "methods": [
            {"type": "phone", "value": "13800138001"},
            {"type": "email", "value": "linxy@fzu.edu.cn"},
            {"type": "wechat", "value": "linxy_2024"},
            {"type": "address", "value": "福建省福州市大学城 福州大学旗山校区 12#302"},
        ],
        "note": "负责前端界面。",
    },
    {
        "name": "陈志强",
        "bookmarked": True,
        "methods": [
            {"type": "phone", "value": "13900139002"},
            {"type": "email", "value": "chenzq@example.com"},
            {"type": "qq", "value": "100100100"},
        ],
        "note": "课后答疑助教。",
    },
    {
        "name": "王梦琪",
        "bookmarked": False,
        "methods": [
            {"type": "wechat", "value": "wmq_design"},
            {"type": "email", "value": "wangmq@example.com"},
        ],
        "note": "负责页面视觉设计。",
    },
    {
        "name": "阿卜杜拉",
        "bookmarked": False,
        "methods": [
            {"type": "phone", "value": "0912-8888666"},
            {"type": "address", "value": "新疆维吾尔自治区乌鲁木齐市天山区"},
        ],
        "note": "",
    },
    {
        "name": "赵一鸣",
        "bookmarked": False,
        "methods": [
            {"type": "qq", "value": "88888888"},
            {"type": "other", "value": "钉钉：zhaoym_fzu"},
        ],
        "note": "Excel 导入导出的测试对象。",
    },
]


def main() -> None:
    contacts = load_contacts()
    existing = {c["name"] for c in contacts}
    added = 0
    for raw in SAMPLE:
        if raw["name"] in existing:
            continue
        contacts.append(normalize_contact(raw))
        added += 1
    save_contacts(contacts)
    print(f"seeded {added} contacts, total = {len(contacts)}")


if __name__ == "__main__":
    main()
