#!/usr/bin/env python3
"""Export logged-in NCESNext reviews to an XLSX workbook.

Credentials are read from NCESNEXT_USERNAME and NCESNEXT_PASSWORD. The script
does not write credentials or cookies to disk. It uses the logged-in HTML pages
that are visible to the account and avoids write actions.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import datetime as dt
import html
import os
import re
import time
import zipfile
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable
from urllib.parse import urljoin
from xml.sax.saxutils import escape

import requests


BASE_URL = "https://ncesnext.com"
SIGNIN_URL = f"{BASE_URL}/signin/"
LATEST_REVIEWS_URL = f"{BASE_URL}/latest_reviews"
REQUEST_TIMEOUT = (8, 45)
USER_AGENT = "python-requests/2.32"
DEFAULT_WORKERS = 6

HEADERS = [
    "发布人",
    "评价课程名称",
    "课程代号",
    "教授姓名",
    "开设院系",
    "评分",
    "课程难度",
    "作业多少",
    "给分好坏",
    "收获大小",
    "评价内容",
    "评价发布时间",
    "课程链接",
    "点评链接/点评ID",
    "抓取状态/备注",
]


@dataclass
class ReviewIndex:
    review_id: str
    course_id: str
    course_url: str
    review_url: str


@dataclass
class CourseMeta:
    course_id: str
    course_name: str
    course_code: str
    professor_name: str
    department: str
    course_url: str
    note: str = ""


@dataclass
class ReviewRow:
    author: str
    course_name: str
    course_code: str
    professor_name: str
    department: str
    rating: str
    difficulty: str
    homework: str
    grading: str
    gain: str
    content: str
    publish_time: str
    course_url: str
    review_url: str
    note: str


class TextExtractor(HTMLParser):
    """Convert small HTML fragments into readable text."""

    BLOCK_TAGS = {"p", "div", "li", "br", "ul", "ol", "blockquote", "h1", "h2", "h3"}

    def __init__(self, include_links: bool = False) -> None:
        super().__init__()
        self.include_links = include_links
        self.parts: list[str] = []
        self.href_stack: list[str | None] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in self.BLOCK_TAGS:
            self.parts.append("\n")
        if tag == "a":
            attrs_dict = dict(attrs)
            self.href_stack.append(attrs_dict.get("href"))

    def handle_data(self, data: str) -> None:
        if data:
            self.parts.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag == "a" and self.href_stack:
            href = self.href_stack.pop()
            if self.include_links and href:
                self.parts.append(f" ({href})")
        if tag in self.BLOCK_TAGS:
            self.parts.append("\n")

    def text(self) -> str:
        value = html.unescape("".join(self.parts)).replace("\xa0", " ")
        lines = [" ".join(line.split()) for line in value.splitlines()]
        return "\n".join(line for line in lines if line).strip()


def html_to_text(fragment: str, include_links: bool = False) -> str:
    parser = TextExtractor(include_links=include_links)
    parser.feed(fragment)
    return parser.text()


def normalize_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def clean_cell(value: object) -> str:
    text = "" if value is None else str(value)
    # XML 1.0 does not allow most ASCII control characters.
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", text)
    if len(text) > 32767:
        text = text[:32750] + "\n[内容超过 Excel 单元格长度限制，已截断]"
    return text


def department_from_code(course_code: str) -> str:
    match = re.match(r"([A-Za-z]+)", course_code.strip())
    return match.group(1).upper() if match else "未知"


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})
    return session


def login(username: str, password: str) -> requests.Session:
    session = make_session()
    response = session.get(SIGNIN_URL, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    csrf_match = re.search(r'name="csrf_token"[^>]*value="([^"]+)"', response.text)
    if not csrf_match:
        raise RuntimeError("Cannot find login CSRF token")

    login_response = session.post(
        f"{SIGNIN_URL}?ajax=1",
        data={
            "csrf_token": csrf_match.group(1),
            "username": username,
            "password": password,
            "remember": "y",
        },
        headers={"Referer": SIGNIN_URL},
        timeout=REQUEST_TIMEOUT,
    )
    login_response.raise_for_status()
    if '"status":200' not in login_response.text:
        raise RuntimeError(f"Login did not succeed: {login_response.text[:200]}")

    check = session.get(f"{LATEST_REVIEWS_URL}?per_page=1", timeout=REQUEST_TIMEOUT)
    check.raise_for_status()
    if "/logout/" not in check.text:
        raise RuntimeError("Login check failed; logout link not present")
    return session


def fetch_review_index(session: requests.Session, per_page: int | None = None) -> tuple[int, dict[str, ReviewIndex]]:
    first = session.get(f"{LATEST_REVIEWS_URL}?per_page=1", timeout=REQUEST_TIMEOUT)
    first.raise_for_status()
    total_match = re.search(r"共\s*(\d+)\s*个点评", first.text)
    if not total_match:
        raise RuntimeError("Cannot find total review count on latest reviews page")
    total = int(total_match.group(1))
    requested = per_page or total + 100

    response = session.get(f"{LATEST_REVIEWS_URL}?per_page={requested}", timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    if "Making sure you" in response.text:
        raise RuntimeError("Latest reviews page returned a bot challenge")

    indexes: dict[str, ReviewIndex] = {}
    for course_id, review_id in re.findall(r'href="/course/(\d+)/#review-(\d+)"', response.text):
        course_url = f"{BASE_URL}/course/{course_id}/"
        review_url = f"{course_url}#review-{review_id}"
        indexes[review_id] = ReviewIndex(
            review_id=review_id,
            course_id=course_id,
            course_url=course_url,
            review_url=review_url,
        )
    return total, indexes


def parse_course_meta(course_id: str, page_html: str) -> CourseMeta:
    name_match = re.search(
        r'<(?:span|h1|h2)[^>]*class="[^"]*blue h[12] fw-bold[^"]*"[^>]*>(.*?)</(?:span|h1|h2)>',
        page_html,
        re.S,
    )
    course_name = html_to_text(name_match.group(1)) if name_match else "未知"

    code_match = re.search(r'<span class="badge bg-secondary mono-font">(.*?)</span>', page_html, re.S)
    course_code = html_to_text(code_match.group(1)) if code_match else "未知"
    if not course_code:
        course_code = "未知"

    department_match = re.search(r"开课单位：</strong>\s*([^<]+)", page_html)
    department = normalize_spaces(html.unescape(department_match.group(1))) if department_match else ""
    note = ""
    if not department or department == "未知":
        fallback = department_from_code(course_code)
        department = fallback
        note = "开课单位缺失，使用课程代号前缀" if fallback != "未知" else "开课单位和课程代号均缺失，标记为未知"

    professors = parse_professors(page_html)
    if not professors:
        professors = "未知"

    return CourseMeta(
        course_id=course_id,
        course_name=course_name,
        course_code=course_code,
        professor_name=professors,
        department=department,
        course_url=f"{BASE_URL}/course/{course_id}/",
        note=note,
    )


def parse_professors(page_html: str) -> str:
    side_match = re.search(r'<div class="col-md-4 rl-pd-lg">(.*?)(?:其他老师的|</body>)', page_html, re.S)
    names: list[str] = []
    if side_match:
        names = [
            html_to_text(name_html)
            for name_html in re.findall(r'<h3 class="blue mt-2"><a[^>]*>(.*?)</a></h3>', side_match.group(1), re.S)
        ]

    if not names:
        mobile_match = re.search(r'<span class="h5 blue mobile">\s*（(.*?)）\s*</span>', page_html, re.S)
        if mobile_match:
            names = [html_to_text(mobile_match.group(1))]

    seen: set[str] = set()
    unique_names: list[str] = []
    for name in names:
        if name and name not in seen:
            unique_names.append(name)
            seen.add(name)
    return ", ".join(unique_names)


def split_review_blocks(page_html: str) -> list[tuple[str, str]]:
    starts = list(
        re.finditer(
            r'<div class="card small-padding-card mb-3 shadow-sm review review-content"\s+id="review-(\d+)"',
            page_html,
        )
    )
    blocks: list[tuple[str, str]] = []
    for index, match in enumerate(starts):
        end = starts[index + 1].start() if index + 1 < len(starts) else len(page_html)
        blocks.append((match.group(1), page_html[match.start():end]))
    return blocks


def parse_rating_from_header(header_html: str) -> str:
    rating = 0
    for cls in re.findall(r'<span class="([^"]*fa-star[^"]*)"', header_html):
        if "fa-star-half-stroke" in cls:
            rating += 1
        elif "fa-solid" in cls and "fa-star" in cls:
            rating += 2
    return str(rating)


def parse_review_block(block: str, meta: CourseMeta, review_id: str) -> ReviewRow:
    note_parts: list[str] = []
    if meta.note:
        note_parts.append(meta.note)

    body_split = block.split('<div class="card-body">', 1)
    header_html = body_split[0]
    body_footer_html = body_split[1] if len(body_split) == 2 else ""
    body_html, _, footer_html = body_footer_html.partition('<div class="card-footer">')

    author_match = re.search(r'<span class="px16 no-underline">(.*?)</span>', header_html, re.S)
    author = html_to_text(author_match.group(1), include_links=False) if author_match else "未知"
    if not author:
        author = "未知"

    rating = parse_rating_from_header(header_html)

    metrics = {}
    for label in ("难度", "作业", "给分", "收获"):
        metric_match = re.search(label + r"：\s*([^<]+)", body_html)
        metrics[label] = normalize_spaces(html.unescape(metric_match.group(1))) if metric_match else ""
        if not metrics[label]:
            note_parts.append(f"{label}缺失")

    content_html = re.sub(
        r'<ul class="list-inline text-body-secondary">.*?</ul>',
        "",
        body_html,
        count=1,
        flags=re.S,
    )
    content = html_to_text(content_html, include_links=True)
    if not content:
        note_parts.append("评价内容为空")

    publish_match = re.search(r'<span class="small localtime"[^>]*>(.*?)</span>', footer_html, re.S)
    publish_time = html_to_text(publish_match.group(1)) if publish_match else ""
    if not publish_time:
        note_parts.append("评价发布时间缺失")

    return ReviewRow(
        author=author,
        course_name=meta.course_name,
        course_code=meta.course_code,
        professor_name=meta.professor_name,
        department=meta.department,
        rating=rating,
        difficulty=metrics["难度"],
        homework=metrics["作业"],
        grading=metrics["给分"],
        gain=metrics["收获"],
        content=content,
        publish_time=publish_time,
        course_url=meta.course_url,
        review_url=f"{meta.course_url}#review-{review_id}",
        note="; ".join(dict.fromkeys(note_parts)),
    )


def parse_course_page(course_id: str, page_html: str, wanted_review_ids: set[str]) -> tuple[CourseMeta, dict[str, ReviewRow]]:
    meta = parse_course_meta(course_id, page_html)
    rows: dict[str, ReviewRow] = {}
    for review_id, block in split_review_blocks(page_html):
        if review_id in wanted_review_ids:
            rows[review_id] = parse_review_block(block, meta, review_id)
    return meta, rows


def fetch_course_page(cookies: requests.cookies.RequestsCookieJar, course_id: str, retries: int = 2) -> str:
    url = f"{BASE_URL}/course/{course_id}/"
    for attempt in range(retries + 1):
        session = make_session()
        session.cookies.update(cookies)
        try:
            response = session.get(url, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
            if "Making sure you" in response.text:
                raise RuntimeError("bot challenge")
            return response.text
        except Exception:
            if attempt >= retries:
                raise
            time.sleep(0.8 * (attempt + 1))
    raise RuntimeError("unreachable")


def fetch_and_parse_courses(
    session: requests.Session,
    indexes: dict[str, ReviewIndex],
    workers: int,
) -> tuple[dict[str, ReviewRow], dict[str, str]]:
    by_course: dict[str, set[str]] = {}
    for review_id, index in indexes.items():
        by_course.setdefault(index.course_id, set()).add(review_id)

    rows: dict[str, ReviewRow] = {}
    failures: dict[str, str] = {}
    completed = 0
    total_courses = len(by_course)
    cookies = session.cookies.copy()

    def fetch_one(item: tuple[str, set[str]]) -> tuple[str, dict[str, ReviewRow], str | None]:
        course_id, review_ids = item
        try:
            page_html = fetch_course_page(cookies, course_id)
            _, parsed_rows = parse_course_page(course_id, page_html, review_ids)
            missing = sorted(review_ids - set(parsed_rows))
            note = f"missing reviews on course page: {', '.join(missing[:5])}" if missing else None
            return course_id, parsed_rows, note
        except Exception as exc:
            return course_id, {}, f"{exc.__class__.__name__}: {exc}"

    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, workers)) as executor:
        futures = [executor.submit(fetch_one, item) for item in by_course.items()]
        for future in concurrent.futures.as_completed(futures):
            course_id, parsed_rows, failure = future.result()
            rows.update(parsed_rows)
            if failure:
                failures[course_id] = failure
            completed += 1
            if completed % 100 == 0 or completed == total_courses:
                print(
                    f"Parsed course pages: {completed}/{total_courses}; reviews parsed: {len(rows)}",
                    flush=True,
                )

    return rows, failures


def make_placeholder_rows(
    indexes: dict[str, ReviewIndex],
    parsed_rows: dict[str, ReviewRow],
    failures: dict[str, str],
) -> dict[str, ReviewRow]:
    all_rows = dict(parsed_rows)
    for review_id, index in indexes.items():
        if review_id in all_rows:
            continue
        failure = failures.get(index.course_id, "详情页未解析到该点评")
        all_rows[review_id] = ReviewRow(
            author="未知",
            course_name="未知",
            course_code="未知",
            professor_name="未知",
            department="未知",
            rating="",
            difficulty="",
            homework="",
            grading="",
            gain="",
            content="",
            publish_time="",
            course_url=index.course_url,
            review_url=index.review_url,
            note=f"点评索引存在，但详情页抓取/解析失败: {failure}",
        )
    return all_rows


def excel_col_name(index: int) -> str:
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def cell_xml(row_index: int, col_index: int, value: object, numeric: bool = False) -> str:
    ref = f"{excel_col_name(col_index)}{row_index}"
    text = clean_cell(value)
    if numeric and text != "":
        return f'<c r="{ref}"><v>{escape(text)}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'


def build_sheet_xml(rows: list[ReviewRow]) -> str:
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
        'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
        '<cols>'
        '<col min="1" max="1" width="20" customWidth="1"/>'
        '<col min="2" max="2" width="28" customWidth="1"/>'
        '<col min="3" max="3" width="14" customWidth="1"/>'
        '<col min="4" max="5" width="24" customWidth="1"/>'
        '<col min="6" max="10" width="12" customWidth="1"/>'
        '<col min="11" max="11" width="80" customWidth="1"/>'
        '<col min="12" max="12" width="22" customWidth="1"/>'
        '<col min="13" max="14" width="42" customWidth="1"/>'
        '<col min="15" max="15" width="42" customWidth="1"/>'
        '</cols>',
        '<sheetData>',
    ]
    header_cells = "".join(cell_xml(1, idx, header) for idx, header in enumerate(HEADERS, start=1))
    lines.append(f'<row r="1">{header_cells}</row>')

    for row_index, row in enumerate(rows, start=2):
        values = [
            row.author,
            row.course_name,
            row.course_code,
            row.professor_name,
            row.department,
            row.rating,
            row.difficulty,
            row.homework,
            row.grading,
            row.gain,
            row.content,
            row.publish_time,
            row.course_url,
            row.review_url,
            row.note,
        ]
        cells = []
        for col_index, value in enumerate(values, start=1):
            cells.append(cell_xml(row_index, col_index, value, numeric=col_index == 6))
        lines.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    lines.extend(["</sheetData>", f'<autoFilter ref="A1:O{len(rows) + 1}"/>', "</worksheet>"])
    return "".join(lines)


def write_xlsx(rows: list[ReviewRow], output_path: Path) -> None:
    created = dt.datetime.now(dt.UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    sheet_xml = build_sheet_xml(rows)
    files = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>""",
        "docProps/app.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Python</Application>
</Properties>""",
        "docProps/core.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>Cursor Agent</dc:creator><cp:lastModifiedBy>Cursor Agent</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>
</cp:coreProperties>""",
        "xl/workbook.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="NCES点评" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
        "xl/_rels/workbook.xml.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""",
        "xl/styles.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>""",
        "xl/worksheets/sheet1.xml": sheet_xml,
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", default="ncesnext_reviews.xlsx", help="Output XLSX path")
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS, help="Concurrent course page workers")
    parser.add_argument("--per-page", type=int, default=None, help="Override latest_reviews per_page")
    parser.add_argument("--limit-courses", type=int, default=None, help="Debug only: limit number of courses fetched")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    username = os.environ.get("NCESNEXT_USERNAME")
    password = os.environ.get("NCESNEXT_PASSWORD")
    if not username or not password:
        raise SystemExit("Set NCESNEXT_USERNAME and NCESNEXT_PASSWORD before running")

    session = login(username, password)
    total, indexes = fetch_review_index(session, args.per_page)
    print(f"Indexed {len(indexes)} unique reviews; site total is {total}.", flush=True)
    if len(indexes) != total:
        print("Warning: indexed review count does not match site total.", flush=True)

    if args.limit_courses is not None:
        keep_courses = set(sorted({index.course_id for index in indexes.values()})[: args.limit_courses])
        indexes = {
            review_id: index
            for review_id, index in indexes.items()
            if index.course_id in keep_courses
        }
        print(f"Debug limit active: keeping {len(indexes)} reviews from {len(keep_courses)} courses.", flush=True)

    parsed_rows, failures = fetch_and_parse_courses(session, indexes, args.workers)
    all_rows = make_placeholder_rows(indexes, parsed_rows, failures)
    ordered_rows = [
        all_rows[review_id]
        for review_id in sorted(all_rows, key=lambda value: int(value), reverse=True)
    ]
    write_xlsx(ordered_rows, Path(args.output))
    print(
        f"Wrote {len(ordered_rows)} rows to {args.output}; "
        f"course failures/missing groups: {len(failures)}",
        flush=True,
    )


if __name__ == "__main__":
    main()
