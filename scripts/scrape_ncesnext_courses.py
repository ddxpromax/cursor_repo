#!/usr/bin/env python3
"""Scrape public NCES course metadata and export it to an XLSX file.

The script intentionally stays unauthenticated and avoids /api/* endpoints,
which are disallowed by the site's robots.txt. It reads the public course list
and each course detail page's public "开课单位" field.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import datetime as dt
import html
import re
import time
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterable
from urllib.parse import urljoin
from xml.sax.saxutils import escape

import requests


BASE_URL = "https://ncesnext.com"
COURSE_LIST_URL = f"{BASE_URL}/course/?per_page={{per_page}}"
DEFAULT_PER_PAGE = 6000
DEFAULT_WORKERS = 6
DETAIL_READ_LIMIT_BYTES = 160 * 1024
REQUEST_TIMEOUT = (8, 30)
# The site allows the default python-requests style user agent but challenges
# generic browser/bot-like strings.
USER_AGENT = "python-requests/2.32"

HEADERS = [
    "课程名称",
    "课程代号",
    "教授姓名",
    "开设院系",
    "总评分",
    "评分人数",
    "开设学期",
    "课程链接",
    "抓取状态/备注",
]


@dataclass
class CourseRow:
    course_name: str
    course_code: str
    professor_name: str
    department: str
    rating: str
    rating_count: int
    latest_term: str
    course_url: str
    note: str


def strip_tags(value: str) -> str:
    return html.unescape(re.sub(r"<[^>]+>", " ", value)).strip()


def normalize_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def department_from_code(course_code: str) -> str:
    match = re.match(r"([A-Za-z]+)", course_code.strip())
    return match.group(1).upper() if match else course_code.strip()


def parse_course_title(anchor_html: str) -> tuple[str, str, str]:
    code_match = re.search(r'<span[^>]*class="[^"]*badge[^"]*"[^>]*>(.*?)</span>', anchor_html, re.S)
    course_code = normalize_spaces(strip_tags(code_match.group(1))) if code_match else ""

    before_badge = anchor_html.split("<span", 1)[0]
    title_text = normalize_spaces(strip_tags(before_badge))

    teacher_match = re.match(r"^(.*?)\s*（(.*?)）\s*$", title_text)
    if teacher_match:
        course_name = normalize_spaces(teacher_match.group(1))
        professor_name = normalize_spaces(teacher_match.group(2))
    else:
        course_name = title_text
        professor_name = "未知"

    return course_name, course_code, professor_name


def parse_course_list(html_text: str) -> tuple[int, list[CourseRow]]:
    total_match = re.search(r"共\s*(\d+)\s*门课", html_text)
    total = int(total_match.group(1)) if total_match else 0

    anchor_pattern = re.compile(
        r'<a class="px16" href="(?P<href>/course/(?P<course_id>\d+)/)">'
        r"(?P<title>.*?)</a>",
        re.S,
    )

    courses: list[CourseRow] = []
    anchors = list(anchor_pattern.finditer(html_text))
    for anchor_index, anchor_match in enumerate(anchors):
        next_anchor_start = anchors[anchor_index + 1].start() if anchor_index + 1 < len(anchors) else len(html_text)
        card_html = html_text[anchor_match.start():next_anchor_start]

        course_name, course_code, professor_name = parse_course_title(anchor_match.group("title"))
        term_match = re.search(r'<span class="small text-body-secondary">\s*(.*?)</span>', card_html, re.S)
        latest_term = normalize_spaces(strip_tags(term_match.group(1))).replace("...", "") if term_match else ""
        rating_match = re.search(r'<span class="rl-pd-sm h4 mono-font">\s*([^<]+?)\s*</span>', card_html)
        count_match = re.search(r"\((\d+)\s*人评价\)", card_html)
        rating = normalize_spaces(strip_tags(rating_match.group(1))) if rating_match else ""
        rating_count = int(count_match.group(1)) if count_match else 0
        course_url = urljoin(BASE_URL, anchor_match.group("href"))
        fallback_department = department_from_code(course_code)

        note = ""
        if not rating:
            note = "暂无评价"

        courses.append(
            CourseRow(
                course_name=course_name,
                course_code=course_code,
                professor_name=professor_name,
                department=fallback_department,
                rating=rating,
                rating_count=rating_count,
                latest_term=latest_term,
                course_url=course_url,
                note=note,
            )
        )

    return total, courses


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})
    return session


def fetch_course_list(per_page: int) -> tuple[int, list[CourseRow]]:
    session = make_session()
    response = session.get(COURSE_LIST_URL.format(per_page=per_page), timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    return parse_course_list(response.text)


def parse_department_from_detail(detail_html: str) -> str | None:
    match = re.search(r"开课单位：</strong>\s*([^<]+)", detail_html)
    if not match:
        return None
    department = normalize_spaces(html.unescape(match.group(1)))
    return department if department and department != "未知" else None


def fetch_detail_department(course_url: str, retries: int = 2) -> tuple[str | None, str | None]:
    """Return (department, note). Only reads the beginning of the detail page."""
    for attempt in range(retries + 1):
        session = make_session()
        try:
            with session.get(course_url, stream=True, timeout=REQUEST_TIMEOUT) as response:
                response.raise_for_status()
                chunks: list[bytes] = []
                bytes_read = 0
                for chunk in response.iter_content(chunk_size=8192):
                    if not chunk:
                        continue
                    chunks.append(chunk)
                    bytes_read += len(chunk)
                    text = b"".join(chunks).decode(response.encoding or "utf-8", errors="ignore")
                    department = parse_department_from_detail(text)
                    if department:
                        return department, None
                    if bytes_read >= DETAIL_READ_LIMIT_BYTES:
                        break

            return None, "详情页未找到开课单位，使用课程代号前缀"
        except requests.RequestException as exc:
            if attempt >= retries:
                return None, f"详情页访问失败，使用课程代号前缀: {exc.__class__.__name__}"
            time.sleep(0.6 * (attempt + 1))

    return None, "详情页访问失败，使用课程代号前缀"


def enrich_departments(courses: list[CourseRow], workers: int) -> None:
    completed = 0
    total = len(courses)

    def enrich(index_and_course: tuple[int, CourseRow]) -> tuple[int, str | None, str | None]:
        index, course = index_and_course
        department, note = fetch_detail_department(course.course_url)
        return index, department, note

    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
        futures = [
            executor.submit(enrich, item)
            for item in enumerate(courses)
        ]
        for future in concurrent.futures.as_completed(futures):
            index, department, note = future.result()
            if department:
                courses[index].department = department
            if note:
                existing_note = courses[index].note
                courses[index].note = f"{existing_note}; {note}" if existing_note else note
            completed += 1
            if completed % 250 == 0 or completed == total:
                print(f"Enriched departments: {completed}/{total}", flush=True)


def excel_col_name(index: int) -> str:
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def cell_xml(row_index: int, col_index: int, value: object, numeric: bool = False) -> str:
    ref = f"{excel_col_name(col_index)}{row_index}"
    if value is None:
        value = ""
    text = str(value)
    if numeric and text != "":
        return f'<c r="{ref}"><v>{escape(text)}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'


def build_sheet_xml(rows: Iterable[CourseRow]) -> str:
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
        'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
        '<cols>'
        '<col min="1" max="1" width="28" customWidth="1"/>'
        '<col min="2" max="2" width="14" customWidth="1"/>'
        '<col min="3" max="3" width="24" customWidth="1"/>'
        '<col min="4" max="4" width="24" customWidth="1"/>'
        '<col min="5" max="6" width="12" customWidth="1"/>'
        '<col min="7" max="7" width="12" customWidth="1"/>'
        '<col min="8" max="8" width="36" customWidth="1"/>'
        '<col min="9" max="9" width="42" customWidth="1"/>'
        '</cols>',
        '<sheetData>',
    ]
    header_cells = "".join(cell_xml(1, idx, header) for idx, header in enumerate(HEADERS, start=1))
    lines.append(f'<row r="1">{header_cells}</row>')

    for row_index, row in enumerate(rows, start=2):
        values = [
            row.course_name,
            row.course_code,
            row.professor_name,
            row.department,
            row.rating,
            row.rating_count,
            row.latest_term,
            row.course_url,
            row.note,
        ]
        cells = []
        for col_index, value in enumerate(values, start=1):
            cells.append(cell_xml(row_index, col_index, value, numeric=col_index in {5, 6}))
        lines.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    last_row = sum(1 for _ in rows) + 1 if not isinstance(rows, list) else len(rows) + 1
    lines.extend(
        [
            '</sheetData>',
            f'<autoFilter ref="A1:I{last_row}"/>',
            '</worksheet>',
        ]
    )
    return "".join(lines)


def write_xlsx(rows: list[CourseRow], output_path: Path) -> None:
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
<sheets><sheet name="NCES课程" sheetId="1" r:id="rId1"/></sheets>
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
    parser.add_argument("--output", default="ncesnext_courses.xlsx", help="Output XLSX path")
    parser.add_argument("--per-page", type=int, default=DEFAULT_PER_PAGE, help="Course list per_page value")
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS, help="Concurrent detail page workers")
    parser.add_argument("--skip-details", action="store_true", help="Use course-code prefixes for departments")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    total, courses = fetch_course_list(args.per_page)
    print(f"Parsed {len(courses)} course rows from list page; site total is {total}.", flush=True)
    if total and total != len(courses):
        print("Warning: parsed row count does not match site total.", flush=True)

    if not args.skip_details:
        enrich_departments(courses, max(1, args.workers))

    output_path = Path(args.output)
    write_xlsx(courses, output_path)
    print(f"Wrote {len(courses)} rows to {output_path}", flush=True)


if __name__ == "__main__":
    main()
