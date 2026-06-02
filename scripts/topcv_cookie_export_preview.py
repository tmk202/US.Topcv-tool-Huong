#!/usr/bin/env python3
import argparse
import base64
import csv
import io
import json
import os
import re
import sys
import tempfile
from pathlib import Path
from urllib.parse import quote, unquote, urlencode, urlsplit, urlunsplit, parse_qsl
from urllib import error, request

from topcv_excel_to_lark import (
    Candidate,
    candidate_from_row,
    load_env,
    normalize_header,
    read_candidates,
)


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_EXPORT_URL = "https://tuyendung-api.topcv.vn/api/v1/cv-management/cvs/export-excel?get_newest_cv=true"
JWT_RE = re.compile(r"eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+")


def normalize_cookie(cookie: str) -> str:
    cookie = cookie.strip()
    if cookie.lower().startswith("cookie:"):
        return cookie.split(":", 1)[1].strip()
    return cookie


def cookie_dict(cookie: str) -> dict[str, str]:
    pairs = {}
    for part in normalize_cookie(cookie).split(";"):
        if "=" not in part:
            continue
        key, value = part.split("=", 1)
        pairs[key.strip()] = unquote(value.strip())
    return pairs


def valid_token(value: str) -> bool:
    value = (value or "").strip()
    return bool(value) and value.lower() not in {"false", "null", "none", "undefined"}


def extract_jwt(value: str) -> str:
    value = (value or "").strip()
    if value.lower().startswith("bearer "):
        value = value.split(None, 1)[1].strip()
    match = JWT_RE.search(value)
    return match.group(0) if match else value


def jwt_payload(token: str) -> dict:
    try:
        payload = token.split(".")[1]
        payload += "=" * ((4 - len(payload) % 4) % 4)
        return json.loads(base64.urlsafe_b64decode(payload))
    except Exception:
        return {}


def access_token_from_cookie(cookie: str) -> str:
    cookies = cookie_dict(cookie)
    for key in (
        "cookie_employer_access_token_local",
        "cookie__token.refresh",
        "employer_access_token",
        "access_token",
    ):
        value = cookies.get(key, "")
        if valid_token(value):
            return extract_jwt(value)
    return ""


def csrf_token_from_cookie(cookie: str) -> str:
    cookies = cookie_dict(cookie)
    return cookies.get("XSRF-TOKEN") or cookies.get("xsrf-token") or ""


def export_url_with_token(export_url: str, cookie: str) -> str:
    if "___token=" in export_url:
        return export_url
    token = access_token_from_cookie(cookie)
    if not token:
        return export_url
    parts = urlsplit(export_url)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query.setdefault("get_newest_cv", "true")
    query["___token"] = token
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query, quote_via=quote), parts.fragment))


def build_headers(cookie: str, csrf_token: str = "") -> dict[str, str]:
    access_token = access_token_from_cookie(cookie)
    csrf_token = csrf_token or csrf_token_from_cookie(cookie)
    headers = {
        "Accept": "text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/json,*/*",
        "Cookie": normalize_cookie(cookie),
        "Origin": "https://tuyendung.topcv.vn",
        "Referer": "https://tuyendung.topcv.vn/",
        "User-Agent": "Mozilla/5.0",
    }
    if access_token:
        headers["Authorization"] = f"Bearer {access_token}"
    if csrf_token:
        headers["X-CSRF-TOKEN"] = csrf_token
        headers["X-XSRF-TOKEN"] = csrf_token
    return headers


def fetch_export(export_url: str, cookie: str, csrf_token: str = "") -> tuple[bytes, str]:
    export_url = export_url_with_token(export_url, cookie)
    req = request.Request(export_url, headers=build_headers(cookie, csrf_token))
    try:
        with request.urlopen(req, timeout=120) as resp:
            return resp.read(), resp.headers.get("content-type", "")
    except error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        if "Attention Required! | Cloudflare" in body:
            raise RuntimeError(
                "TopCV returned Cloudflare 403. Use browser-download mode in the TUI, "
                "because direct terminal HTTP is being blocked."
            ) from exc
        if "SESSION_TIMEOUT" in body:
            raise RuntimeError("TopCV session timed out. Export a fresh cookie JSON after logging in again.") from exc
        raise RuntimeError(f"TopCV returned HTTP {exc.code}: {body[:1000]}") from exc


def candidates_from_csv(data: bytes) -> list[Candidate]:
    text = data.decode("utf-8-sig", errors="replace")
    sample = text[:4096]
    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
    reader = csv.DictReader(io.StringIO(text), dialect=dialect)
    candidates: list[Candidate] = []
    for raw_row in reader:
        row = {
            normalize_header(header): (value or "").strip()
            for header, value in raw_row.items()
            if header is not None
        }
        candidate = candidate_from_row(row)
        if candidate:
            candidates.append(candidate)
    return candidates


def candidates_from_export(data: bytes, content_type: str) -> list[Candidate]:
    first_bytes = data[:500].lower()
    if b"<html" in first_bytes or b"<!doctype html" in first_bytes:
        raise RuntimeError("TopCV returned HTML, not CSV/XLSX. Cookie may be expired or export URL needs ___token.")

    if b"application/json" in content_type.encode() or first_bytes.startswith(b"{"):
        raise RuntimeError(f"TopCV returned JSON instead of CSV/XLSX: {data[:1000].decode('utf-8', errors='replace')}")

    if data[:2] == b"PK" or "spreadsheet" in content_type:
        temp = tempfile.NamedTemporaryFile(prefix="topcv-cookie-export-", suffix=".xlsx", delete=False)
        temp_path = Path(temp.name)
        temp.close()
        temp_path.write_bytes(data)
        return read_candidates(temp_path)

    return candidates_from_csv(data)


def print_table(candidates: list[Candidate], limit: int = 0) -> None:
    rows = candidates[:limit] if limit else candidates
    headers = ["#", "job_id", "job_title", "apply_at", "candidate_name", "candidate_email", "candidate_phone", "download_url"]
    widths = {header: len(header) for header in headers}
    payloads = []
    for idx, candidate in enumerate(rows, start=1):
        payload = {"#": str(idx), **candidate.payload()}
        payloads.append(payload)
        for header, value in payload.items():
            widths[header] = min(max(widths[header], len(value or "")), 48)

    def fit(value: str, width: int) -> str:
        value = value or ""
        if len(value) > width:
            return value[: width - 1] + "…"
        return value.ljust(width)

    print(" | ".join(fit(header, widths[header]) for header in headers))
    print("-+-".join("-" * widths[header] for header in headers))
    for payload in payloads:
        print(" | ".join(fit(payload.get(header, ""), widths[header]) for header in headers))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export TopCV CV list using Cookie and print parsed fields.")
    parser.add_argument("--cookie", help="Raw Cookie header value copied from Chrome DevTools.")
    parser.add_argument("--export-url", help="TopCV export URL. Defaults to TOPCV_EXPORT_URL or the base export endpoint.")
    parser.add_argument("--csrf-token", help="Optional CSRF/XSRF token header.")
    parser.add_argument("--json", action="store_true", help="Print JSON lines instead of a table.")
    parser.add_argument("--limit", type=int, default=20, help="Max rows to print. Use 0 for all rows.")
    parser.add_argument("--save-raw", help="Optional path to save the raw downloaded CSV/XLSX.")
    return parser.parse_args()


def main() -> int:
    load_env(ROOT / ".env")
    args = parse_args()
    cookie = (args.cookie or os.getenv("TOPCV_COOKIE", "")).strip()
    csrf_token = (args.csrf_token or os.getenv("TOPCV_CSRF_TOKEN", "")).strip()
    export_url = (args.export_url or os.getenv("TOPCV_EXPORT_URL", "") or DEFAULT_EXPORT_URL).strip()

    if not cookie:
        print("Missing cookie. Pass --cookie or set TOPCV_COOKIE in .env.", file=sys.stderr)
        return 2

    data, content_type = fetch_export(export_url, cookie, csrf_token)
    if args.save_raw:
        Path(args.save_raw).expanduser().write_bytes(data)
        print(f"Saved raw export to {args.save_raw}")

    candidates = candidates_from_export(data, content_type)
    print(f"Fetched {len(data)} bytes ({content_type or 'unknown content-type'}). Parsed {len(candidates)} candidates.")

    rows = candidates[: args.limit] if args.limit else candidates
    if args.json:
        for candidate in rows:
            print(json.dumps(candidate.payload(), ensure_ascii=False))
    else:
        print_table(candidates, args.limit)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
