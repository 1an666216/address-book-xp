#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""EE308FZ Extreme Programming assignment -- Address Book server.

Pure standard library + openpyxl.  Run with:

    python app.py            # http://127.0.0.1:8000
    python app.py --port 9000

Data is persisted in data/contacts.json so the app survives a restart,
which makes the demo reproducible without a database.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import threading
import time
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from urllib.parse import parse_qs, urlparse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
DATA_DIR = os.path.join(BASE_DIR, "data")
DATA_FILE = os.path.join(DATA_DIR, "contacts.json")

# ---------------------------------------------------------------------------
# Domain model
# ---------------------------------------------------------------------------

# Method kinds: key -> (Chinese column label, display label)
METHOD_TYPES = {
    "phone": ("手机", "手机"),
    "email": ("邮箱", "邮箱"),
    "wechat": ("微信", "微信"),
    "qq": ("QQ", "QQ"),
    "address": ("地址", "地址"),
    "other": ("其他联系方式", "其他"),
}

EXCEL_HEADERS = ["姓名", "收藏", "手机", "邮箱", "微信", "QQ", "地址", "其他联系方式", "备注"]

# Tolerant import aliases (lower-case, stripped).
HEADER_ALIASES = {
    "name": "name",
    "姓名": "name",
    "名字": "name",
    "收藏": "bookmarked",
    "bookmark": "bookmarked",
    "bookmarked": "bookmarked",
    "手机": "phone",
    "电话": "phone",
    "电话号码": "phone",
    "手机号": "phone",
    "phone": "phone",
    "mobile": "phone",
    "邮箱": "email",
    "电子邮箱": "email",
    "email": "email",
    "e-mail": "email",
    "微信": "wechat",
    "wechat": "wechat",
    "qq": "qq",
    "地址": "address",
    "住址": "address",
    "address": "address",
    "其他": "other",
    "其他联系方式": "other",
    "other": "other",
    "备注": "note",
    "note": "note",
    "notes": "note",
}

_lock = threading.Lock()


def _now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())


def _ensure_dirs() -> None:
    os.makedirs(DATA_DIR, exist_ok=True)


def load_contacts() -> list[dict]:
    if not os.path.exists(DATA_FILE):
        return []
    with open(DATA_FILE, "r", encoding="utf-8") as fh:
        try:
            payload = json.load(fh)
        except json.JSONDecodeError:
            return []
    if isinstance(payload, dict):
        return payload.get("contacts", [])
    return payload


def save_contacts(contacts: list[dict]) -> None:
    _ensure_dirs()
    tmp = DATA_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump({"contacts": contacts}, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, DATA_FILE)


def _clean_methods(raw) -> list[dict]:
    methods = []
    for item in raw or []:
        if not isinstance(item, dict):
            continue
        value = str(item.get("value", "")).strip()
        if not value:
            continue
        kind = str(item.get("type", "other")).strip().lower()
        if kind not in METHOD_TYPES:
            kind = "other"
        methods.append({"type": kind, "value": value})
    return methods


def normalize_contact(raw: dict, existing: dict | None = None) -> dict:
    """Validate + normalise a client payload into the stored shape."""
    name = str(raw.get("name", "")).strip()
    if not name:
        raise ValueError("姓名不能为空 / name is required")
    bookmarked = raw.get("bookmarked")
    if bookmarked is None and existing is not None:
        bookmarked = existing.get("bookmarked", False)
    contact = {
        "id": (existing or {}).get("id") or "c_" + uuid.uuid4().hex[:10],
        "name": name,
        "bookmarked": bool(bookmarked),
        "methods": _clean_methods(raw.get("methods")),
        "note": str(raw.get("note", "") or "").strip(),
        "createdAt": (existing or {}).get("createdAt") or _now(),
        "updatedAt": _now(),
    }
    return contact


def summarize(contacts: list[dict]) -> dict:
    total = len(contacts)
    bookmarked = sum(1 for c in contacts if c.get("bookmarked"))
    methods = sum(len(c.get("methods", [])) for c in contacts)
    return {"total": total, "bookmarked": bookmarked, "methods": methods}


# ---------------------------------------------------------------------------
# Excel export / import
# ---------------------------------------------------------------------------

_SPLIT_RE = re.compile(r"[；;]")


def _sheet_value(contact: dict, key: str) -> str:
    if key == "name":
        return contact.get("name", "")
    if key == "bookmarked":
        return "是" if contact.get("bookmarked") else ""
    if key == "note":
        return contact.get("note", "")
    values = [m["value"] for m in contact.get("methods", []) if m.get("type") == key]
    if not values and key == "other":
        values = [m["value"] for m in contact.get("methods", []) if m.get("type") == "other"]
    return "；".join(values)


def contacts_to_workbook(contacts: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "通讯录"

    header_fill = PatternFill("solid", fgColor="1F6FEB")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    thin = Side(style="thin", color="D0D7DE")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.append(EXCEL_HEADERS)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # 姓名列写实质字段名，方便导入端按列名映射
    ordered = ["name", "bookmarked", "phone", "email", "wechat", "qq", "address", "other", "note"]
    for contact in sorted(contacts, key=lambda c: (not c.get("bookmarked"), c.get("name", ""))):
        ws.append([_sheet_value(contact, key) for key in ordered])

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(EXCEL_HEADERS)):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    widths = [14, 7, 18, 26, 16, 14, 30, 20, 26]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(EXCEL_HEADERS))}{max(ws.max_row, 1)}"
    ws.row_dimensions[1].height = 22

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def _header_map(row) -> dict[str, int]:
    """Map a header row (cell objects or plain values) to field names."""
    mapping: dict[str, int] = {}
    for idx, cell in enumerate(row):
        value = cell.value if hasattr(cell, "value") else cell
        if value is None:
            continue
        key = str(value).strip().lower()
        field = HEADER_ALIASES.get(key)
        if field and field not in mapping:
            mapping[field] = idx
    return mapping


def workbook_to_contacts(blob: bytes) -> list[dict]:
    """Parse an uploaded .xlsx into raw contact dicts (not yet merged)."""
    try:
        wb = load_workbook(BytesIO(blob), data_only=True)
    except Exception as exc:  # noqa: BLE001 - any parse error is a client error
        raise ValueError(f"无法解析 Excel 文件，请上传 .xlsx：{exc}") from exc
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Excel 文件是空的 / empty workbook")

    header_index = 0
    mapping = _header_map(rows[0])
    if "name" not in mapping:
        # tolerate a leading title row: look at the first 5 rows
        for i, row in enumerate(rows[:5]):
            candidate = _header_map(row)
            if "name" in candidate:
                header_index, mapping = i, candidate
                break
    if "name" not in mapping:
        raise ValueError("找不到“姓名”列，请使用导出的模板格式 / missing 姓名 column")

    parsed: list[dict] = []
    for row in rows[header_index + 1:]:
        def get(field: str):
            idx = mapping.get(field)
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        name = str(get("name") or "").strip()
        if not name:
            continue
        bookmark_raw = str(get("bookmarked") or "").strip().lower()
        methods = []
        for field in ("phone", "email", "wechat", "qq", "address", "other"):
            cell = get(field)
            if cell is None:
                continue
            for piece in _SPLIT_RE.split(str(cell)):
                piece = piece.strip()
                if piece:
                    methods.append({"type": field, "value": piece})
        parsed.append({
            "name": name,
            "bookmarked": bookmark_raw in {"是", "y", "yes", "true", "1", "√", "✓", "收藏"},
            "methods": methods,
            "note": str(get("note") or "").strip(),
        })
    if not parsed:
        raise ValueError("Excel 中没有可导入的联系人 / no data rows found")
    return parsed


def merge_imported(contacts: list[dict], imported: list[dict]) -> tuple[list[dict], dict]:
    by_name = {c["name"]: c for c in contacts}
    stats = {"imported": 0, "updated": 0, "methodsAdded": 0}
    for item in imported:
        existing = by_name.get(item["name"])
        if existing is None:
            contact = normalize_contact(item)
            contacts.append(contact)
            by_name[contact["name"]] = contact
            stats["imported"] += 1
            continue
        stats["updated"] += 1
        have = {(m["type"], m["value"]) for m in existing.get("methods", [])}
        for method in item["methods"]:
            if (method["type"], method["value"]) not in have:
                existing.setdefault("methods", []).append(method)
                have.add((method["type"], method["value"]))
                stats["methodsAdded"] += 1
        if item.get("bookmarked"):
            existing["bookmarked"] = True
        if item.get("note") and not existing.get("note"):
            existing["note"] = item["note"]
        existing["updatedAt"] = _now()
    return contacts, stats


# ---------------------------------------------------------------------------
# HTTP layer
# ---------------------------------------------------------------------------

class Handler(BaseHTTPRequestHandler):
    server_version = "AddressBookXP/1.0"
    protocol_version = "HTTP/1.1"

    # -- helpers ----------------------------------------------------------
    def _send(self, status: int, body: bytes, content_type: str, extra: dict | None = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for key, value in (extra or {}).items():
            self.send_header(key, value)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _json(self, status: int, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self._send(status, body, "application/json; charset=utf-8")

    def _error(self, status: int, message: str) -> None:
        self._json(status, {"ok": False, "error": message})

    def _read_body(self) -> bytes:
        length = int(self.headers.get("Content-Length") or 0)
        return self.rfile.read(length) if length else b""

    def _read_json(self) -> dict:
        raw = self._read_body()
        if not raw:
            return {}
        try:
            data = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError(f"请求体不是合法 JSON: {exc}") from exc
        if not isinstance(data, dict):
            raise ValueError("请求体必须是 JSON 对象")
        return data

    def log_message(self, fmt: str, *args) -> None:  # quieter console
        sys.stderr.write("[%s] %s\n" % (time.strftime("%H:%M:%S"), fmt % args))

    def handle(self) -> None:
        """Swallow client-side disconnects (keep-alive sockets closed early)."""
        try:
            super().handle()
        except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError):
            pass

    # -- routing ----------------------------------------------------------
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        try:
            if path == "/api/health":
                return self._json(200, {"ok": True, "service": "address-book"})
            if path == "/api/contacts":
                return self._list_contacts(parse_qs(parsed.query))
            if path == "/api/stats":
                with _lock:
                    contacts = load_contacts()
                return self._json(200, {"ok": True, "stats": summarize(contacts)})
            if path == "/api/export":
                with _lock:
                    contacts = load_contacts()
                blob = contacts_to_workbook(contacts)
                stamp = time.strftime("%Y%m%d")
                return self._send(
                    200,
                    blob,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    {"Content-Disposition": f'attachment; filename="address_book_{stamp}.xlsx"'},
                )
            return self._serve_static(path)
        except ValueError as exc:
            return self._error(400, str(exc))
        except Exception as exc:  # noqa: BLE001 - keep the demo server alive
            return self._error(500, f"服务器内部错误: {exc}")

    def do_HEAD(self) -> None:
        self.do_GET()

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        try:
            if path == "/api/contacts":
                return self._create_contact()
            if path == "/api/import":
                return self._import_excel()
            match = re.fullmatch(r"/api/contacts/([\w-]+)/toggle-bookmark", path)
            if match:
                return self._toggle_bookmark(match.group(1))
            return self._error(404, "未知接口")
        except ValueError as exc:
            return self._error(400, str(exc))
        except Exception as exc:  # noqa: BLE001
            return self._error(500, f"服务器内部错误: {exc}")

    def do_PUT(self) -> None:
        match = re.fullmatch(r"/api/contacts/([\w-]+)", urlparse(self.path).path)
        if not match:
            return self._error(404, "未知接口")
        try:
            return self._update_contact(match.group(1))
        except ValueError as exc:
            return self._error(400, str(exc))
        except Exception as exc:  # noqa: BLE001
            return self._error(500, f"服务器内部错误: {exc}")

    def do_DELETE(self) -> None:
        match = re.fullmatch(r"/api/contacts/([\w-]+)", urlparse(self.path).path)
        if not match:
            return self._error(404, "未知接口")
        target = match.group(1)
        with _lock:
            contacts = load_contacts()
            remaining = [c for c in contacts if c["id"] != target]
            if len(remaining) == len(contacts):
                return self._error(404, "联系人不存在")
            save_contacts(remaining)
        return self._json(200, {"ok": True, "deleted": target})

    # -- handlers ---------------------------------------------------------
    def _list_contacts(self, query: dict) -> None:
        keyword = (query.get("q", [""])[0] or "").strip().lower()
        only_fav = (query.get("bookmarked", [""])[0] or "").lower() in {"1", "true", "yes"}
        with _lock:
            contacts = load_contacts()
        if only_fav:
            contacts = [c for c in contacts if c.get("bookmarked")]
        if keyword:
            def hit(contact: dict) -> bool:
                haystack = [contact.get("name", ""), contact.get("note", "")]
                for method in contact.get("methods", []):
                    haystack.append(method.get("value", ""))
                    # allow searching by the kind name as well ("微信", "邮箱", ...)
                    haystack.append(METHOD_TYPES.get(method.get("type"), ("其他",))[0])
                return any(keyword in str(item).lower() for item in haystack)
            contacts = [c for c in contacts if hit(c)]
        contacts.sort(key=lambda c: (not c.get("bookmarked"), c.get("name", "")))
        return self._json(200, {"ok": True, "contacts": contacts, "stats": summarize(load_contacts())})

    def _create_contact(self) -> None:
        payload = self._read_json()
        with _lock:
            contacts = load_contacts()
            contact = normalize_contact(payload)
            if any(c["name"] == contact["name"] for c in contacts):
                raise ValueError(f"联系人“{contact['name']}”已存在 / duplicate name")
            contacts.append(contact)
            save_contacts(contacts)
        return self._json(201, {"ok": True, "contact": contact})

    def _update_contact(self, contact_id: str) -> None:
        payload = self._read_json()
        with _lock:
            contacts = load_contacts()
            for idx, existing in enumerate(contacts):
                if existing["id"] == contact_id:
                    updated = normalize_contact(payload, existing=existing)
                    contacts[idx] = updated
                    save_contacts(contacts)
                    return self._json(200, {"ok": True, "contact": updated})
        raise ValueError("联系人不存在 / contact not found")

    def _toggle_bookmark(self, contact_id: str) -> None:
        with _lock:
            contacts = load_contacts()
            for contact in contacts:
                if contact["id"] == contact_id:
                    contact["bookmarked"] = not contact.get("bookmarked", False)
                    contact["updatedAt"] = _now()
                    save_contacts(contacts)
                    return self._json(200, {"ok": True, "contact": contact})
        return self._error(404, "联系人不存在")

    def _import_excel(self) -> None:
        blob = self._read_body()
        if not blob:
            raise ValueError("没有收到文件内容 / empty upload")
        imported = workbook_to_contacts(blob)
        with _lock:
            contacts = load_contacts()
            contacts, stats = merge_imported(contacts, imported)
            save_contacts(contacts)
        stats["total"] = len(contacts)
        return self._json(200, {"ok": True, "stats": stats})

    # -- static files -----------------------------------------------------
    def _serve_static(self, path: str) -> None:
        if path in ("/", ""):
            path = "/index.html"
        rel = os.path.normpath(path.lstrip("/"))
        if rel.startswith("..") or os.path.isabs(rel):
            return self._error(403, "禁止访问")
        full = os.path.join(STATIC_DIR, rel)
        if not os.path.isfile(full):
            return self._error(404, "文件不存在")
        ext = os.path.splitext(full)[1].lower()
        content_type = {
            ".html": "text/html; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
            ".svg": "image/svg+xml",
            ".png": "image/png",
            ".ico": "image/x-icon",
        }.get(ext, "application/octet-stream")
        with open(full, "rb") as fh:
            body = fh.read()
        return self._send(200, body, content_type)


def main() -> None:
    parser = argparse.ArgumentParser(description="EE308FZ XP assignment: address book")
    parser.add_argument("--port", type=int, default=int(os.environ.get("PORT", 8000)))
    parser.add_argument("--host", default="127.0.0.1")
    args = parser.parse_args()
    _ensure_dirs()
    httpd = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"Address Book running at http://{args.host}:{args.port}  (Ctrl+C to stop)")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nbye")
    finally:
        httpd.server_close()


if __name__ == "__main__":
    main()
