#!/usr/bin/env python3
import asyncio
import json
import os
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Optional
from urllib import request

from playwright.async_api import Page, async_playwright


ROOT = Path(__file__).resolve().parents[1]
SESSION_DIR = ROOT / ".topcv-session"
DEFAULT_STATE_FILE = ROOT / "state.json"

EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.I)
PHONE_RE = re.compile(r"(?:\+?84|0)(?:[\s.-]?\d){8,10}")
DATE_RE = re.compile(r"\b\d{2}/\d{2}/\d{4}(?:\s+\d{2}:\d{2}(?::\d{2})?)?\b")
ID_RE = re.compile(r"#?(\d{4,})")


@dataclass
class Candidate:
    job_id: str
    job_title: str
    apply_at: str
    candidate_name: str
    candidate_email: str
    candidate_phone: str
    download_url: str

    def key(self) -> str:
        parts = [
            self.job_id,
            self.candidate_email,
            digits_only(self.candidate_phone),
            self.candidate_name,
            self.apply_at,
        ]
        return "|".join(part.strip().lower() for part in parts if part and part.strip())

    def payload(self) -> dict[str, str]:
        return {
            "job_id": self.job_id,
            "job_title": self.job_title,
            "apply_at": self.apply_at,
            "candidate_name": self.candidate_name,
            "candidate_email": self.candidate_email,
            "candidate_phone": self.candidate_phone,
            "download_url": self.download_url,
        }


def load_env(path: Path) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


def digits_only(value: str) -> str:
    return re.sub(r"\D+", "", value or "")


def normalize_date(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return ""
    formats = [
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(value[:19], fmt)
            if fmt == "%d/%m/%Y":
                return dt.strftime("%Y-%m-%d 00:00:00")
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            pass
    return value


def first_text(obj: dict[str, Any], keys: tuple[str, ...]) -> str:
    for key in keys:
        value = deep_get(obj, key)
        if value is not None and str(value).strip():
            return str(value).strip()
    return ""


def deep_get(obj: Any, dotted_key: str) -> Any:
    current = obj
    for part in dotted_key.split("."):
        if isinstance(current, dict):
            current = current.get(part)
        else:
            return None
    return current


def walk_json(value: Any):
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from walk_json(child)
    elif isinstance(value, list):
        for child in value:
            yield from walk_json(child)


def candidate_from_json(obj: dict[str, Any]) -> Optional[Candidate]:
    download_url = first_text(
        obj,
        (
            "download_url",
            "cv_download_url",
            "downloadUrl",
            "download.url",
            "cv.url",
            "cv.file_url",
            "resume_url",
            "onetime_download_url",
        ),
    )
    if not download_url:
        for item in walk_json(obj):
            maybe_url = first_text(item, ("url", "link", "href"))
            if "onetime-download" in maybe_url or "cv-management" in maybe_url:
                download_url = maybe_url
                break

    email = first_text(obj, ("candidate_email", "email", "candidate.email", "profile.email", "user.email"))
    phone = first_text(
        obj,
        ("candidate_phone", "phone", "candidate.phone", "profile.phone", "mobile", "phone_number"),
    )
    name = first_text(
        obj,
        ("candidate_name", "name", "full_name", "fullname", "candidate.name", "profile.full_name", "user.name"),
    )

    if not any((download_url, email, phone)) or not name:
        return None

    job_title = first_text(
        obj,
        (
            "job_title",
            "job.title",
            "campaign_title",
            "campaign.title",
            "recruitment_campaign.title",
            "recruitment.title",
        ),
    )
    job_id = first_text(
        obj,
        ("job_id", "job.id", "campaign_id", "campaign.id", "recruitment_campaign.id", "recruitment.id", "id"),
    )
    apply_at = first_text(
        obj,
        ("apply_at", "applied_at", "created_at", "createdAt", "submitted_at", "updated_at"),
    )

    return Candidate(
        job_id=job_id,
        job_title=job_title,
        apply_at=normalize_date(apply_at),
        candidate_name=name,
        candidate_email=email,
        candidate_phone=phone,
        download_url=download_url,
    )


async def candidates_from_dom(page: Page) -> list[Candidate]:
    candidates: list[Candidate] = []
    rows = await page.locator("tr, [role=row], .candidate-item, .list-candidate-item").all()
    for row in rows:
        try:
            text = " ".join((await row.inner_text()).split())
        except Exception:
            continue
        if "@" not in text and not PHONE_RE.search(text):
            continue

        email_match = EMAIL_RE.search(text)
        phone_match = PHONE_RE.search(text)
        date_match = DATE_RE.search(text)
        id_match = ID_RE.search(text)

        links = await row.locator("a[href]").evaluate_all(
            "(els) => els.map((a) => a.href).filter(Boolean)"
        )
        download_url = ""
        for link in links:
            if "onetime-download" in link or "download" in link or "cv-management" in link:
                download_url = link
                break

        lines = [line.strip() for line in (await row.inner_text()).splitlines() if line.strip()]
        name = ""
        for line in lines:
            if EMAIL_RE.search(line) or PHONE_RE.search(line):
                continue
            if DATE_RE.search(line) or line.startswith("#"):
                continue
            if line.lower() in {"ung tuyen", "ứng tuyển", "tiep nhan", "tiếp nhận"}:
                continue
            name = line
            break

        if not name:
            continue

        candidates.append(
            Candidate(
                job_id=id_match.group(1) if id_match else "",
                job_title="",
                apply_at=normalize_date(date_match.group(0) if date_match else ""),
                candidate_name=name,
                candidate_email=email_match.group(0) if email_match else "",
                candidate_phone=phone_match.group(0) if phone_match else "",
                download_url=download_url,
            )
        )
    return candidates


def load_state(path: Path) -> set[str]:
    if not path.exists():
        return set()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return set()
    if isinstance(data, list):
        return set(str(item) for item in data)
    return set(str(item) for item in data.get("sent", []))


def save_state(path: Path, sent: set[str]) -> None:
    path.write_text(
        json.dumps({"sent": sorted(sent)}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def post_lark(webhook_url: str, payload: dict[str, str]) -> None:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = request.Request(
        webhook_url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with request.urlopen(req, timeout=30) as resp:
        body = resp.read().decode("utf-8", errors="replace")
        if resp.status >= 300:
            raise RuntimeError(f"Lark returned HTTP {resp.status}: {body[:500]}")


async def click_next_page(page: Page) -> bool:
    selectors = [
        "button:has-text('Sau')",
        "a:has-text('Sau')",
        "button:has-text('Next')",
        "a:has-text('Next')",
        ".pagination-next:not(.disabled)",
        "li.next:not(.disabled) a",
    ]
    for selector in selectors:
        locator = page.locator(selector).first
        try:
            if await locator.count() and await locator.is_visible() and await locator.is_enabled():
                await locator.click()
                await page.wait_for_load_state("networkidle", timeout=15000)
                return True
        except Exception:
            continue
    return False


async def open_browser(p, start_url: str, headless: bool):
    chrome_cdp_url = os.getenv("CHROME_CDP_URL", "").strip()
    if chrome_cdp_url:
        browser = await p.chromium.connect_over_cdp(chrome_cdp_url)
        context = browser.contexts[0] if browser.contexts else await browser.new_context()
        page = None
        for candidate_page in context.pages:
            if "topcv.vn" in candidate_page.url:
                page = candidate_page
                break
        if page is None:
            page = context.pages[0] if context.pages else await context.new_page()
        return browser, context, page, True

    context = await p.chromium.launch_persistent_context(
        str(SESSION_DIR),
        headless=headless,
        viewport={"width": 1440, "height": 900},
    )
    page = await context.new_page()
    await page.goto(start_url, wait_until="domcontentloaded")
    return context, context, page, False


async def main() -> int:
    load_env(ROOT / ".env")
    webhook_url = os.getenv("LARK_WEBHOOK_URL", "").strip()
    start_url = os.getenv("TOPCV_START_URL", "https://tuyendung.topcv.vn/").strip()
    state_file = Path(os.getenv("STATE_FILE", str(DEFAULT_STATE_FILE))).expanduser()
    if not state_file.is_absolute():
        state_file = ROOT / state_file
    headless = os.getenv("HEADLESS", "false").lower() in {"1", "true", "yes"}
    max_pages = int(os.getenv("MAX_PAGES", "0") or "0")
    send_existing = os.getenv("SEND_EXISTING", "true").lower() in {"1", "true", "yes"}

    if not webhook_url:
        print("Missing LARK_WEBHOOK_URL. Add it to .env first.", file=sys.stderr)
        return 2

    sent = load_state(state_file)
    found: dict[str, Candidate] = {}
    api_candidates: dict[str, Candidate] = {}

    async with async_playwright() as p:
        browser_or_context, context, page, attached_to_existing_chrome = await open_browser(
            p,
            start_url,
            headless,
        )

        async def handle_response(response):
            content_type = response.headers.get("content-type", "")
            if "application/json" not in content_type:
                return
            url = response.url.lower()
            if not any(token in url for token in ("cv", "candidate", "apply", "recruit")):
                return
            try:
                data = await response.json()
            except Exception:
                return
            for item in walk_json(data):
                candidate = candidate_from_json(item)
                if candidate:
                    api_candidates[candidate.key()] = candidate

        page.on("response", lambda response: asyncio.create_task(handle_response(response)))
        if attached_to_existing_chrome:
            print("Attached to existing Chrome via CHROME_CDP_URL.")
        else:
            print("Browser opened. Log in and navigate to the TopCV CV list if needed.")
        print("Press Enter here after the candidate list is visible...")
        await asyncio.to_thread(sys.stdin.readline)

        page_no = 0
        while True:
            page_no += 1
            await page.wait_for_load_state("networkidle", timeout=15000)
            await page.wait_for_timeout(1500)

            dom_candidates = await candidates_from_dom(page)
            for candidate in [*api_candidates.values(), *dom_candidates]:
                found[candidate.key()] = candidate

            print(f"Scanned page {page_no}: {len(found)} unique candidates collected.")

            if max_pages and page_no >= max_pages:
                break
            if not await click_next_page(page):
                break

        if attached_to_existing_chrome:
            await browser_or_context.close()
        else:
            await context.close()

    new_count = 0
    skipped_count = 0
    for candidate in found.values():
        key = candidate.key()
        if not key or key in sent:
            skipped_count += 1
            continue
        if not send_existing:
            sent.add(key)
            skipped_count += 1
            continue
        post_lark(webhook_url, candidate.payload())
        sent.add(key)
        new_count += 1
        print(f"Sent: {candidate.candidate_name} / {candidate.candidate_email}")

    save_state(state_file, sent)
    print(f"Done. Sent {new_count}, skipped {skipped_count}, total seen {len(found)}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
