#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""End-to-end smoke test for the address book (no test framework needed).

    python tests/e2e_test.py [--base http://127.0.0.1:8000]

It exercises every graded feature over real HTTP: CRUD, bookmarks, multiple
contact methods, Excel export, Excel import (new + merge), and error paths.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.error
import urllib.request
from io import BytesIO

from openpyxl import load_workbook, Workbook

BASE = "http://127.0.0.1:8000"
PASSED = 0
FAILED: list[str] = []


def check(label: str, condition: bool, detail: str = "") -> None:
    global PASSED
    if condition:
        PASSED += 1
        print(f"  PASS  {label}")
    else:
        FAILED.append(label)
        print(f"  FAIL  {label}  {detail}")


def request(method: str, path: str, body: bytes | None = None, headers: dict | None = None):
    req = urllib.request.Request(BASE + path, data=body, method=method)
    for key, value in (headers or {}).items():
        req.add_header(key, value)
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.status, resp.read()
    except urllib.error.HTTPError as exc:
        return exc.code, exc.read()


def json_request(method: str, path: str, payload: dict | None = None):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
    status, raw = request(method, path, body, {"Content-Type": "application/json"})
    try:
        return status, json.loads(raw.decode("utf-8"))
    except Exception:
        return status, {}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--base", default="http://127.0.0.1:8000")
    args = parser.parse_args()
    global BASE
    BASE = args.base.rstrip("/")

    print(f"Testing {BASE}")
    created_ids: list[str] = []

    # 1. health + listing ------------------------------------------------
    status, payload = json_request("GET", "/api/health")
    check("health endpoint", status == 200 and payload.get("ok") is True, str(payload))

    status, payload = json_request("GET", "/api/contacts")
    baseline = len(payload.get("contacts", []))
    check("list contacts", status == 200 and baseline >= 0, str(status))

    # 2. create with several contact methods -----------------------------
    new_contact = {
        "name": "E2E 测试同学",
        "bookmarked": False,
        "note": "automated test",
        "methods": [
            {"type": "phone", "value": "13512345678"},
            {"type": "email", "value": "e2e@fzu.edu.cn"},
            {"type": "wechat", "value": "e2e_wx"},
            {"type": "address", "value": "福州市闽侯县学园路 2 号"},
        ],
    }
    status, payload = json_request("POST", "/api/contacts", new_contact)
    check("create contact (201)", status == 201 and payload.get("contact", {}).get("id"), str(payload))
    contact_id = payload.get("contact", {}).get("id", "")
    if contact_id:
        created_ids.append(contact_id)
    check("multiple contact methods stored",
          len(payload.get("contact", {}).get("methods", [])) == 4, str(payload))

    # 3. duplicate name is rejected --------------------------------------
    status, payload = json_request("POST", "/api/contacts", {"name": "E2E 测试同学", "methods": []})
    check("duplicate name rejected (400)", status == 400, str(status))

    # 4. empty name is rejected ------------------------------------------
    status, payload = json_request("POST", "/api/contacts", {"name": "  ", "methods": []})
    check("empty name rejected (400)", status == 400, str(status))

    # 5. search ----------------------------------------------------------
    status, payload = json_request("GET", "/api/contacts?q=%E6%B5%8B%E8%AF%95")
    check("search by name", status == 200 and any(
        c["name"] == "E2E 测试同学" for c in payload.get("contacts", [])), str(status))
    status, payload = json_request("GET", "/api/contacts?q=13512345678")
    check("search by phone", any(c["id"] == contact_id for c in payload.get("contacts", [])))

    # 6. bookmark toggle --------------------------------------------------
    status, payload = json_request("POST", f"/api/contacts/{contact_id}/toggle-bookmark")
    check("bookmark on", status == 200 and payload.get("contact", {}).get("bookmarked") is True, str(payload))
    status, payload = json_request("GET", "/api/contacts?bookmarked=1")
    check("bookmarked filter includes it", any(c["id"] == contact_id for c in payload.get("contacts", [])))
    status, payload = json_request("POST", f"/api/contacts/{contact_id}/toggle-bookmark")
    check("bookmark off", payload.get("contact", {}).get("bookmarked") is False, str(payload))

    # 7. update -----------------------------------------------------------
    updated = dict(new_contact)
    updated["note"] = "edited by e2e"
    updated["methods"] = new_contact["methods"] + [{"type": "qq", "value": "123456"}]
    status, payload = json_request("PUT", f"/api/contacts/{contact_id}", updated)
    check("update contact", status == 200 and payload.get("contact", {}).get("note") == "edited by e2e", str(status))
    check("update keeps methods", len(payload.get("contact", {}).get("methods", [])) == 5)

    # 8. Excel export -----------------------------------------------------
    status, blob = request("GET", "/api/export")
    check("export returns xlsx", status == 200 and blob[:2] == b"PK", f"status={status} len={len(blob)}")
    wb = load_workbook(BytesIO(blob))
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    check("export header row", headers[:3] == ["姓名", "收藏", "手机"], str(headers))
    rows = [[cell.value for cell in row] for row in ws.iter_rows(min_row=2)]
    names = [row[0] for row in rows]
    check("export contains the contact", "E2E 测试同学" in names)
    if "E2E 测试同学" in names:
        row = rows[names.index("E2E 测试同学")]
        check("export one contact per row", len(row) == len(headers))
        check("export merges same-kind values into one cell",
              "13512345678" in str(row[2]), str(row[2]))
    os.makedirs("exports", exist_ok=True)
    with open(os.path.join("exports", "address_book_export.xlsx"), "wb") as fh:
        fh.write(blob)
    print(f"  INFO  exported {len(rows)} rows -> exports/address_book_export.xlsx")

    # 9. Excel import (new row + merge into existing row) ------------------
    wb2 = Workbook()
    ws2 = wb2.active
    ws2.append(["姓名", "收藏", "手机", "邮箱", "微信", "QQ", "地址", "其他联系方式", "备注"])
    ws2.append(["E2E 导入同学", "是", "13600001111", "import@fzu.edu.cn", "", "", "福州市鼓楼区", "", "来自 Excel"])
    ws2.append(["E2E 测试同学", "", "", "", "", "", "", "钉钉：e2e_ding", "追加一条"])
    buffer = BytesIO()
    wb2.save(buffer)
    status, raw = request("POST", "/api/import", buffer.getvalue(),
                          {"Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})
    payload = json.loads(raw.decode("utf-8"))
    check("import succeeds", status == 200 and payload.get("ok") is True, str(payload))
    stats = payload.get("stats") or {}
    check("import counts new contact", stats.get("imported") == 1, str(stats))
    check("import merges existing contact", stats.get("updated") >= 1, str(stats))

    status, payload = json_request("GET", "/api/contacts?q=E2E")
    imported = next((c for c in payload.get("contacts", []) if c["name"] == "E2E 导入同学"), None)
    check("imported contact has methods", imported is not None and len(imported["methods"]) == 3, str(imported))
    check("imported bookmark flag", imported is not None and imported["bookmarked"] is True)
    if imported:
        created_ids.append(imported["id"])

    status, payload = json_request("GET", f"/api/contacts?q=%E6%B5%8B%E8%AF%95")
    merged = next((c for c in payload.get("contacts", []) if c["id"] == contact_id), None)
    check("merge appended method", merged is not None and any(
        m.get("type") == "other" for m in merged.get("methods", [])), str(merged))

    # 10. bad upload ------------------------------------------------------
    status, raw = request("POST", "/api/import", b"not an excel file",
                          {"Content-Type": "application/octet-stream"})
    check("garbage upload rejected (400)", status == 400, str(status))

    # 11. delete + cleanup ------------------------------------------------
    status, payload = json_request("DELETE", f"/api/contacts/{contact_id}")
    check("delete contact", status == 200 and payload.get("ok") is True, str(status))
    for cid in created_ids:
        json_request("DELETE", f"/api/contacts/{cid}")
    status, payload = json_request("GET", "/api/contacts")
    check("cleanup back to baseline", len(payload.get("contacts", [])) == baseline,
          f"{len(payload.get('contacts', []))} vs {baseline}")

    print(f"\n{PASSED} passed, {len(FAILED)} failed")
    if FAILED:
        print("Failed checks: " + ", ".join(FAILED))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
