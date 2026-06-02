#!/usr/bin/env python3
import argparse
import csv
import io
import json
import os
import sys
import time
import tempfile
import webbrowser
from dataclasses import dataclass
from pathlib import Path

from topcv_cookie_export_preview import (
    DEFAULT_EXPORT_URL,
    access_token_from_cookie,
    candidates_from_export,
    export_url_with_token,
    fetch_export,
    jwt_payload,
    print_table,
)
from topcv_excel_to_lark import load_env
from topcv_excel_to_lark import (
    candidate_from_row,
    load_state,
    normalize_header,
    post_lark,
    read_topcv_records,
    save_state,
)


ROOT = Path(__file__).resolve().parents[1]
DOWNLOAD_DIR = Path.home() / "Downloads"
CONFIG_FILE = ROOT / ".topcv-export-tui.json"


@dataclass
class CookieChoice:
    label: str
    cookie: str
    count: int


@dataclass
class RunSettings:
    cookie: str
    export_url: str
    limit: int
    output_mode: int
    fetch_mode: int
    save_raw: str
    send_lark: bool


def clear_screen() -> None:
    if sys.stdout.isatty():
        print("\033[2J\033[H", end="")


def pause() -> None:
    input("\nPress Enter to continue...")


def prompt_text(label: str, default: str = "") -> str:
    suffix = f" [{default}]" if default else ""
    value = input(f"{label}{suffix}: ").strip()
    return value or default


def prompt_choice(title: str, options: list[str], default_index: int = 0) -> int:
    while True:
        print(f"\n{title}")
        for idx, option in enumerate(options, start=1):
            marker = "*" if idx - 1 == default_index else " "
            print(f"  {idx}. {marker} {option}")
        raw = input(f"Choose 1-{len(options)} [{default_index + 1}]: ").strip()
        if not raw:
            return default_index
        if raw.isdigit() and 1 <= int(raw) <= len(options):
            return int(raw) - 1
        print("Invalid choice.")


def load_config() -> dict:
    if not CONFIG_FILE.exists():
        return {}
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def save_config(config: dict) -> None:
    CONFIG_FILE.write_text(json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def cookie_header_from_list(cookies: list[dict], domain_filter: str = "topcv.vn") -> str:
    pairs = []
    seen = set()
    for cookie in cookies:
        name = str(cookie.get("name", "")).strip()
        value = str(cookie.get("value", "")).strip()
        domain = str(cookie.get("domain", "")).strip()
        if not name or not value:
            continue
        if domain_filter and domain and domain_filter not in domain:
            continue
        if name in seen:
            continue
        seen.add(name)
        pairs.append(f"{name}={value}")
    return "; ".join(pairs)


def flatten_cookie_sources(data, label: str = "cookies") -> list[CookieChoice]:
    choices: list[CookieChoice] = []

    if isinstance(data, dict):
        if "cookies" in data and isinstance(data["cookies"], list):
            cookie = cookie_header_from_list(data["cookies"])
            if cookie:
                choices.append(CookieChoice(label, cookie, len(data["cookies"])))

        direct_pairs = []
        for key, value in data.items():
            if isinstance(value, str) and key not in {"url", "domain", "path", "expires", "sameSite"}:
                direct_pairs.append(f"{key}={value}")
        if direct_pairs:
            choices.append(CookieChoice(f"{label} mapping", "; ".join(direct_pairs), len(direct_pairs)))

        for key, value in data.items():
            if isinstance(value, (dict, list)) and key != "cookies":
                choices.extend(flatten_cookie_sources(value, str(key)))

    elif isinstance(data, list):
        if all(isinstance(item, dict) and "name" in item and "value" in item for item in data):
            cookie = cookie_header_from_list(data)
            if cookie:
                choices.append(CookieChoice(label, cookie, len(data)))
        else:
            for idx, value in enumerate(data, start=1):
                if isinstance(value, (dict, list)):
                    choices.extend(flatten_cookie_sources(value, f"{label} #{idx}"))
                elif isinstance(value, str) and "=" in value:
                    choices.append(CookieChoice(f"{label} #{idx}", value, value.count("=")))

    elif isinstance(data, str) and "=" in data:
        choices.append(CookieChoice(label, data, data.count("=")))

    return choices


def load_cookie_choices(path: Path) -> list[CookieChoice]:
    data = json.loads(path.read_text(encoding="utf-8"))
    choices = flatten_cookie_sources(data, path.name)
    deduped = []
    seen = set()
    for choice in choices:
        if choice.cookie in seen:
            continue
        seen.add(choice.cookie)
        deduped.append(choice)
    return deduped


def scan_cookie_json_files() -> list[tuple[Path, list[CookieChoice]]]:
    results: list[tuple[Path, list[CookieChoice]]] = []
    ignored_names = {"state.json", "excel-state.json", "package-lock.json", "package.json"}
    for path in sorted(ROOT.glob("*.json")):
        if path.name in ignored_names:
            continue
        try:
            choices = load_cookie_choices(path)
        except Exception:
            continue
        if choices:
            results.append((path, choices))
    return results


def choose_cookie_from_scan(scanned_files: list[tuple[Path, list[CookieChoice]]]) -> tuple[str, str, int]:
    file_index = prompt_choice(
        "Cookie JSON file",
        [f"{path.name} ({len(choices)} usable set{'s' if len(choices) != 1 else ''})" for path, choices in scanned_files],
    )
    path, choices = scanned_files[file_index]
    choice_index = prompt_choice(
        f"Cookie set in {path.name}",
        [f"{choice.label} ({choice.count} cookies, {len(choice.cookie)} chars)" for choice in choices],
    )
    return choices[choice_index].cookie, path.name, choice_index


def choose_cookie() -> tuple[str, dict]:
    env_cookie = os.getenv("TOPCV_COOKIE", "").strip()
    cookie_json = os.getenv("TOPCV_COOKIE_JSON", "").strip()
    scanned_files = scan_cookie_json_files()

    sources = []
    if scanned_files:
        sources.append((f"Auto-scan JSON files in {ROOT.name}", ""))
    if env_cookie:
        sources.append(("Use TOPCV_COOKIE from .env", env_cookie))
    if cookie_json:
        sources.append((f"Read cookie JSON from {cookie_json}", ""))
    sources.append(("Enter cookie JSON path", ""))
    sources.append(("Paste raw Cookie header", ""))

    choice = prompt_choice("Cookie source", [label for label, _ in sources])
    label, value = sources[choice]

    if label.startswith("Auto-scan"):
        cookie, file_name, set_index = choose_cookie_from_scan(scanned_files)
        return cookie, {
            "cookie_source": "auto_scan",
            "cookie_file": file_name,
            "cookie_set_index": set_index,
        }

    if label.startswith("Use TOPCV_COOKIE"):
        return value, {"cookie_source": "env_cookie"}

    if label.startswith("Paste"):
        return prompt_text("Paste Cookie header"), {"cookie_source": "paste"}

    path_text = cookie_json if label.startswith("Read cookie JSON") else prompt_text("Cookie JSON path")
    path = Path(path_text).expanduser()
    if not path.is_absolute():
        path = ROOT / path
    choices = load_cookie_choices(path)
    if not choices:
        raise RuntimeError(f"No usable TopCV cookies found in {path}")

    selected = prompt_choice(
        "Cookie set",
        [f"{choice.label} ({choice.count} cookies, {len(choice.cookie)} chars)" for choice in choices],
    )
    return choices[selected].cookie, {
        "cookie_source": "json_path",
        "cookie_json": str(path),
        "cookie_set_index": selected,
    }


def choose_export_url(cookie: str) -> str:
    env_url = os.getenv("TOPCV_EXPORT_URL", "").strip()
    options = []
    if env_url:
        options.append(("Use TOPCV_EXPORT_URL from .env", env_url))
    default_label = "Use default export endpoint"
    if access_token_from_cookie(cookie):
        default_label += " + ___token from cookie"
    options.append((default_label, DEFAULT_EXPORT_URL))
    options.append(("Paste full export URL", ""))
    choice = prompt_choice("Export endpoint", [label for label, _ in options])
    label, value = options[choice]
    if label.startswith("Paste"):
        return prompt_text("Export URL")
    return value


def find_latest_download(started_at: float, timeout_seconds: int = 180) -> Path:
    allowed_suffixes = {".xlsx", ".csv"}
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if DOWNLOAD_DIR.exists():
            candidates = [
                path
                for path in DOWNLOAD_DIR.iterdir()
                if path.is_file()
                and path.suffix.lower() in allowed_suffixes
                and path.stat().st_mtime >= started_at
                and not path.name.endswith(".crdownload")
            ]
            partials = [
                path
                for path in DOWNLOAD_DIR.iterdir()
                if path.is_file() and path.name.endswith(".crdownload") and path.stat().st_mtime >= started_at
            ]
            if candidates and not partials:
                return max(candidates, key=lambda path: path.stat().st_mtime)
        time.sleep(1)
    raise RuntimeError(f"No new .xlsx/.csv file found in {DOWNLOAD_DIR} after {timeout_seconds}s")


def parse_downloaded_file(path: Path):
    data = path.read_bytes()
    content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if path.suffix.lower() == ".csv":
        content_type = "text/csv"
    return data, content_type


def records_from_csv(data: bytes) -> list[dict[str, str]]:
    text = data.decode("utf-8-sig", errors="replace")
    dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
    reader = csv.DictReader(io.StringIO(text), dialect=dialect)
    records = []
    for row in reader:
        records.append({str(header or "").strip(): (value or "").strip() for header, value in row.items()})
    return records


def records_from_export(data: bytes, content_type: str) -> list[dict[str, str]]:
    first_bytes = data[:500].lower()
    if b"<html" in first_bytes or b"<!doctype html" in first_bytes:
        raise RuntimeError("TopCV returned HTML, not CSV/XLSX. Cookie may be expired or export URL needs ___token.")
    if b"application/json" in content_type.encode() or first_bytes.startswith(b"{"):
        raise RuntimeError(f"TopCV returned JSON instead of CSV/XLSX: {data[:1000].decode('utf-8', errors='replace')}")

    if data[:2] == b"PK" or "spreadsheet" in content_type:
        temp = tempfile.NamedTemporaryFile(prefix="topcv-export-records-", suffix=".xlsx", delete=False)
        temp_path = Path(temp.name)
        temp.close()
        temp_path.write_bytes(data)
        return read_topcv_records(temp_path)
    return records_from_csv(data)


def normalized_record(record: dict[str, str]) -> dict[str, str]:
    return {normalize_header(header): value for header, value in record.items()}


def flat_topcv_fields(record: dict[str, str]) -> dict[str, str]:
    flat = {}
    seen: dict[str, int] = {}
    for header, value in record.items():
        base = "topcv_" + normalize_header(header).replace(" ", "_")
        if base == "topcv_":
            base = "topcv_column"
        count = seen.get(base, 0) + 1
        seen[base] = count
        key = base if count == 1 else f"{base}_{count}"
        flat[key] = value
    return flat


def lark_payload_from_record(record: dict[str, str]) -> dict:
    candidate = candidate_from_row(normalized_record(record))
    payload = candidate.payload() if candidate else {
        "job_id": "",
        "job_title": "",
        "apply_at": "",
        "candidate_name": "",
        "candidate_email": "",
        "candidate_phone": "",
        "download_url": "",
    }
    payload.update(flat_topcv_fields(record))
    payload["topcv_fields"] = record
    return payload


def record_key(record: dict[str, str]) -> str:
    candidate = candidate_from_row(normalized_record(record))
    if candidate:
        return candidate.key()
    return json.dumps(record, ensure_ascii=False, sort_keys=True)


def send_records_to_lark(records: list[dict[str, str]]) -> None:
    webhook_url = os.getenv("LARK_WEBHOOK_URL", "").strip()
    if not webhook_url:
        raise RuntimeError("Missing LARK_WEBHOOK_URL in .env")

    state_path = Path(os.getenv("EXCEL_STATE_FILE", "excel-state.json")).expanduser()
    if not state_path.is_absolute():
        state_path = ROOT / state_path
    sent = load_state(state_path)
    sent_count = 0
    skipped_count = 0

    for record in records:
        key = record_key(record)
        if not key or key in sent:
            skipped_count += 1
            continue
        payload = lark_payload_from_record(record)
        label = payload.get("candidate_name") or payload.get("topcv_ho_ten") or key[:40]
        try:
            post_lark(webhook_url, payload)
        except Exception as exc:
            raise RuntimeError(f"Failed sending Lark record {label!r} with key {key!r}: {exc}") from exc
        sent.add(key)
        save_state(state_path, sent)
        sent_count += 1
        print(f"Sent Lark: {label}")
        time.sleep(float(os.getenv("LARK_SEND_DELAY_SECONDS", "1.5")))

    print(f"Lark done. Sent {sent_count}, skipped {skipped_count}, state={state_path.name}.")


def cookie_from_config(config: dict) -> str:
    source = config.get("cookie_source", "auto_scan")
    if source == "env_cookie":
        cookie = os.getenv("TOPCV_COOKIE", "").strip()
        if not cookie:
            raise RuntimeError("Saved config uses TOPCV_COOKIE, but TOPCV_COOKIE is empty.")
        return cookie

    if source == "json_path":
        path = Path(config.get("cookie_json", "")).expanduser()
        if not path.is_absolute():
            path = ROOT / path
        choices = load_cookie_choices(path)
        index = int(config.get("cookie_set_index", 0))
        return choices[index].cookie

    scanned_files = scan_cookie_json_files()
    if not scanned_files:
        raise RuntimeError(f"No usable cookie JSON file found in {ROOT}")

    total_choices = sum(len(choices) for _path, choices in scanned_files)
    if len(scanned_files) > 1 or total_choices > 1:
        print("\nMultiple cookie JSON choices found. Please choose which one to use.")
        cookie, file_name, set_index = choose_cookie_from_scan(scanned_files)
        config.update({
            "cookie_source": "auto_scan",
            "cookie_file": file_name,
            "cookie_set_index": set_index,
        })
        save_config(config)
        print(f"Saved cookie choice to {CONFIG_FILE.name}.")
        return cookie

    path, choices = scanned_files[0]
    config.update({
        "cookie_source": "auto_scan",
        "cookie_file": path.name,
        "cookie_set_index": 0,
    })
    save_config(config)
    return choices[0].cookie


def settings_from_config(config: dict) -> RunSettings:
    cookie = cookie_from_config(config)
    export_url = DEFAULT_EXPORT_URL
    if config.get("export_endpoint") == "env" and os.getenv("TOPCV_EXPORT_URL", "").strip():
        export_url = os.getenv("TOPCV_EXPORT_URL", "").strip()
    elif config.get("export_endpoint") == "custom" and config.get("export_url"):
        export_url = config["export_url"]

    return RunSettings(
        cookie=cookie,
        export_url=export_url,
        limit=int(config.get("limit", 0)),
        output_mode=int(config.get("output_mode", 0)),
        fetch_mode=int(config.get("fetch_mode", 0)),
        save_raw=str(config.get("save_raw", "")),
        send_lark=bool(config.get("send_lark", False)),
    )


def prompt_settings() -> RunSettings:
    cookie, cookie_config = choose_cookie()
    export_url = choose_export_url(cookie)
    limit = int(prompt_text("Rows to display, 0 for all", "0"))
    output_mode = prompt_choice("Output mode", ["Table", "JSON lines"])
    fetch_mode = prompt_choice(
        "Fetch mode",
        [
            "Open export URL in browser, then parse latest download",
            "Direct API from terminal",
        ],
    )
    save_raw = prompt_text("Save raw export path, blank to skip", "")
    send_lark = prompt_choice("Send to Lark webhook", ["Yes", "No"]) == 0

    export_config = {"export_endpoint": "default"}
    if export_url == os.getenv("TOPCV_EXPORT_URL", "").strip() and export_url:
        export_config = {"export_endpoint": "env"}
    elif export_url != DEFAULT_EXPORT_URL:
        export_config = {"export_endpoint": "custom", "export_url": export_url}

    save_config({
        **cookie_config,
        **export_config,
        "limit": limit,
        "output_mode": output_mode,
        "fetch_mode": fetch_mode,
        "save_raw": save_raw,
        "send_lark": send_lark,
    })
    print(f"\nSaved choices to {CONFIG_FILE.name}. Use --interactive to change them.")

    return RunSettings(cookie, export_url, limit, output_mode, fetch_mode, save_raw, send_lark)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="TopCV export preview terminal UI.")
    parser.add_argument("--interactive", action="store_true", help="Show menus and overwrite saved choices.")
    parser.add_argument("--no-save", action="store_true", help="Ignore saved choices and show menus for this run.")
    parser.add_argument("--preview-only", action="store_true", help="Do not send to Lark even if saved config enables it.")
    return parser.parse_args()


def run(settings: RunSettings) -> int:
    cookie = settings.cookie
    export_url = settings.export_url
    limit = settings.limit
    output_mode = settings.output_mode
    fetch_mode = settings.fetch_mode
    save_raw = settings.save_raw
    send_lark = settings.send_lark

    final_export_url = export_url_with_token(export_url, cookie)
    if final_export_url != export_url:
        token_meta = jwt_payload(access_token_from_cookie(cookie))
        tk_n = token_meta.get("tk_n", "unknown")
        exp = token_meta.get("exp", "unknown")
        print(f"\nUsing access token from cookie for ___token (tk_n={tk_n}, exp={exp}).")
    if fetch_mode == 0:
        print("Opening export URL in your browser...")
        started_at = time.time()
        if not webbrowser.open(final_export_url):
            print(f"Open this URL manually:\n{final_export_url}")
        print(f"Waiting for a new .xlsx/.csv file in {DOWNLOAD_DIR}...")
        downloaded_path = find_latest_download(started_at)
        print(f"Found download: {downloaded_path}")
        data, content_type = parse_downloaded_file(downloaded_path)
    else:
        print("Fetching TopCV export...")
        csrf_token = os.getenv("TOPCV_CSRF_TOKEN", "").strip()
        data, content_type = fetch_export(export_url, cookie, csrf_token)

    if save_raw:
        Path(save_raw).expanduser().write_bytes(data)
        print(f"Saved raw export to {save_raw}")

    candidates = candidates_from_export(data, content_type)
    records = records_from_export(data, content_type)
    print(
        f"Fetched {len(data)} bytes ({content_type or 'unknown content-type'}). "
        f"Parsed {len(candidates)} candidates, {len(records)} full records.\n"
    )

    rows = candidates[:limit] if limit else candidates
    if output_mode == 1:
        for candidate in rows:
            print(json.dumps(candidate.payload(), ensure_ascii=False))
    else:
        print_table(candidates, limit)

    if send_lark:
        print("\nSending full TopCV records to Lark...")
        send_records_to_lark(records)
    return 0


def main() -> int:
    load_env(ROOT / ".env")
    args = parse_args()
    clear_screen()
    print("TopCV Export Preview TUI")
    print("========================")

    try:
        config = load_config()
        if config and not args.interactive and not args.no_save:
            print(f"Using saved choices from {CONFIG_FILE.name}. Run with --interactive to change them.")
            settings = settings_from_config(config)
        else:
            settings = prompt_settings()
        if args.preview_only:
            settings.send_lark = False
        return run(settings)
    except KeyboardInterrupt:
        print("\nCancelled.")
        return 130
    except Exception as exc:
        print(f"\nError: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
