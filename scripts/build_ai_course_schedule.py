import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
import posixpath

SCRIPT_DIR = Path(__file__).resolve().parent
ROOT_DIR = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR / ".vendor"))
import xlsxwriter  # type: ignore


INPUTS_DIR = ROOT_DIR / "inputs"
OUTPUTS_DIR = ROOT_DIR / "outputs"
COURSE_XLSX = INPUTS_DIR / "sustech_2026_fall_undergrad_courses.xlsx"
OUTPUT_REMAINING_TIMETABLE_XLSX = OUTPUTS_DIR / "ai_major_remaining_weekly_timetable_optimized.xlsx"

# Based on the user's screenshots and instruction:
# completed set = union of the two screenshots minus dropped EAP and CS306.
# Only professional required/elective plan courses are relevant here.
COMPLETED_PLAN_CODES = {
    "CS201",
    "CS203",
    "MA212",
    "COE201",
    "COE202",
    "CS405",
}


PLAN_COURSES = [
    {
        "plan_part": "必修",
        "module": "专业基础课",
        "code": "CS203",
        "name": "数据结构与算法分析",
        "credits": 3,
        "recommended_term": "2/秋",
        "prereq": "CS109",
        "alt_codes": [],
        "alt_names": [],
        "note": "",
    },
    {
        "plan_part": "必修",
        "module": "专业基础课",
        "code": "MA212",
        "name": "概率论与数理统计",
        "credits": 3,
        "recommended_term": "2/秋",
        "prereq": "MA127",
        "alt_codes": [],
        "alt_names": [],
        "note": "",
    },
    {
        "plan_part": "必修",
        "module": "专业基础课",
        "code": "COE201",
        "name": "人工智能大模型",
        "credits": 3,
        "recommended_term": "2/春",
        "prereq": "无",
        "alt_codes": ["AI201"],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "必修",
        "module": "专业基础课",
        "code": "CS405",
        "name": "机器学习",
        "credits": 3,
        "recommended_term": "2/春",
        "prereq": "MA212",
        "alt_codes": [],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "必修",
        "module": "专业基础课",
        "code": "CS208",
        "name": "算法设计与分析",
        "credits": 3,
        "recommended_term": "2/春",
        "prereq": "CS203",
        "alt_codes": [],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "必修",
        "module": "专业核心课",
        "code": "CS303",
        "name": "人工智能",
        "credits": 3,
        "recommended_term": "3/秋",
        "prereq": "CS203, MA212",
        "alt_codes": [],
        "alt_names": [],
        "note": "",
    },
    {
        "plan_part": "必修",
        "module": "专业核心课",
        "code": "CS324",
        "name": "深度学习",
        "credits": 3,
        "recommended_term": "3/秋",
        "prereq": "CS303",
        "alt_codes": [],
        "alt_names": [],
        "note": "开课表中被标成专业选修课，但在 AI 培养方案里属于核心课。",
    },
    {
        "plan_part": "必修",
        "module": "专业核心课",
        "code": "COE301",
        "name": "体系结构",
        "credits": 3,
        "recommended_term": "3/秋",
        "prereq": "PHY106",
        "alt_codes": ["CS331"],
        "alt_names": ["计算机体系结构"],
        "note": "2026秋识别到 CS331 计算机体系结构，需人工确认是否可作为 COE301 认定。",
    },
    {
        "plan_part": "必修",
        "module": "专业核心课",
        "code": "COE302",
        "name": "人工智能芯片与系统",
        "credits": 3,
        "recommended_term": "3/春",
        "prereq": "COE301",
        "alt_codes": ["AI302"],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "必修",
        "module": "专业核心课",
        "code": "CS310",
        "name": "自然语言处理",
        "credits": 3,
        "recommended_term": "3/春",
        "prereq": "CS303",
        "alt_codes": [],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "CS207",
        "name": "数字逻辑",
        "credits": 3,
        "recommended_term": "2/秋",
        "prereq": "无",
        "alt_codes": [],
        "alt_names": [],
        "note": "",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "CS307",
        "name": "数据库原理",
        "credits": 3,
        "recommended_term": "2/秋",
        "prereq": "CS109",
        "alt_codes": [],
        "alt_names": [],
        "note": "",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "EE211",
        "name": "机器人感知与智能",
        "credits": 3,
        "recommended_term": "2/秋",
        "prereq": "无",
        "alt_codes": [],
        "alt_names": [],
        "note": "",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "COE202",
        "name": "人工智能数学基础",
        "credits": 3,
        "recommended_term": "2/春",
        "prereq": "MA127",
        "alt_codes": ["AI202"],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "CS202",
        "name": "计算机组成原理",
        "credits": 3,
        "recommended_term": "2/春",
        "prereq": "CS207",
        "alt_codes": [],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "CS201",
        "name": "离散数学",
        "credits": 3,
        "recommended_term": "2/春",
        "prereq": "MA127, MA113",
        "alt_codes": [],
        "alt_names": [],
        "note": "2026秋有开设，但培养方案建议学期为春季。",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "CS306",
        "name": "数据挖掘",
        "credits": 3,
        "recommended_term": "2/春",
        "prereq": "CS203",
        "alt_codes": [],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "COE303",
        "name": "人工智能工程实践",
        "credits": 3,
        "recommended_term": "3/秋",
        "prereq": "无",
        "alt_codes": ["AI303"],
        "alt_names": [],
        "note": "2026秋识别到 AI303，但课表里没有完整时间和教师信息，已单独列出。",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "CS325",
        "name": "多智能体系统",
        "credits": 3,
        "recommended_term": "3/秋",
        "prereq": "CS203",
        "alt_codes": [],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "CS305",
        "name": "计算机网络",
        "credits": 3,
        "recommended_term": "3/秋",
        "prereq": "CS109",
        "alt_codes": [],
        "alt_names": [],
        "note": "",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "BMEB316",
        "name": "医学图像处理",
        "credits": 3,
        "recommended_term": "3/秋",
        "prereq": "无",
        "alt_codes": [],
        "alt_names": [],
        "note": "",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "COE403",
        "name": "人工智能科研实践",
        "credits": 3,
        "recommended_term": "3/春",
        "prereq": "无",
        "alt_codes": ["AI403"],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "CS308",
        "name": "计算机视觉",
        "credits": 3,
        "recommended_term": "3/春",
        "prereq": "CS203, MA127, MA113",
        "alt_codes": [],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "CS342",
        "name": "优化方法",
        "credits": 3,
        "recommended_term": "3/春",
        "prereq": "无",
        "alt_codes": [],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "CS340",
        "name": "计算伦理学",
        "credits": 3,
        "recommended_term": "3/春",
        "prereq": "CS303",
        "alt_codes": [],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "COE404",
        "name": "人工智能前沿实践",
        "credits": 3,
        "recommended_term": "3/春",
        "prereq": "无",
        "alt_codes": ["AI404"],
        "alt_names": [],
        "note": "2026秋未在当前课表中识别到。",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "CS323",
        "name": "编译原理",
        "credits": 3,
        "recommended_term": "4/秋",
        "prereq": "CS202",
        "alt_codes": [],
        "alt_names": [],
        "note": "2026秋有开设，但培养方案推荐学期更靠后。",
    },
    {
        "plan_part": "选修",
        "module": "专业选修课",
        "code": "CS328",
        "name": "分布与云计算",
        "credits": 3,
        "recommended_term": "4/秋",
        "prereq": "CS305",
        "alt_codes": [],
        "alt_names": [],
        "note": "2026秋有开设，但培养方案推荐学期更靠后。",
    },
]


DAY_MAP = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "日": 7, "天": 7}
DAY_NAME = {1: "周一", 2: "周二", 3: "周三", 4: "周四", 5: "周五", 6: "周六", 7: "周日"}
COURSE_COLORS = [
    "#C00000",
    "#1F4E79",
    "#548235",
    "#7F6000",
    "#7030A0",
    "#0F766E",
    "#9E480E",
    "#1D3557",
    "#2B6E3F",
    "#8A1538",
    "#3D405B",
    "#006D77",
]


def norm_text(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def load_workbook_rows(xlsx_path: Path):
    ns = {
        "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    with zipfile.ZipFile(xlsx_path) as zf:
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si", ns):
                shared_strings.append("".join(t.text or "" for t in si.iterfind(".//a:t", ns)))

        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

        target = None
        for sheet in wb.find("a:sheets", ns):
            if sheet.attrib["name"] == "整理版":
                rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
                target = posixpath.normpath(posixpath.join("xl", rel_map[rid])).lstrip("/")
                break
        if not target:
            raise RuntimeError("Cannot find 整理版 sheet.")

        root = ET.fromstring(zf.read(target))
        rows = root.findall(".//a:sheetData/a:row", ns)

        def cell_text(cell):
            ctype = cell.attrib.get("t")
            v = cell.find("a:v", ns)
            is_elem = cell.find("a:is", ns)
            if ctype == "s" and v is not None:
                return shared_strings[int(v.text)]
            if ctype == "inlineStr" and is_elem is not None:
                return "".join(t.text or "" for t in is_elem.iterfind(".//a:t", ns))
            return "" if v is None else (v.text or "")

        def col_name(ref):
            m = re.match(r"([A-Z]+)\d+", ref)
            return m.group(1) if m else ""

        parsed = []
        for row in rows[1:]:
            data = {}
            for cell in row.findall("a:c", ns):
                data[col_name(cell.attrib.get("r", ""))] = cell_text(cell)
            parsed.append(data)
        return parsed


def parse_weeks(token: str):
    token = token.replace("周", "").strip()
    parity = "all"
    if token.endswith("单"):
        parity = "odd"
        token = token[:-1]
    elif token.endswith("双"):
        parity = "even"
        token = token[:-1]
    elif token.endswith("单周"):
        parity = "odd"
        token = token[:-2]
    elif token.endswith("双周"):
        parity = "even"
        token = token[:-2]

    weeks = set()
    for part in token.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start, end = part.split("-", 1)
            rng = range(int(start), int(end) + 1)
        else:
            rng = range(int(part), int(part) + 1)
        for week in rng:
            if parity == "all":
                weeks.add(week)
            elif parity == "odd" and week % 2 == 1:
                weeks.add(week)
            elif parity == "even" and week % 2 == 0:
                weeks.add(week)
    return weeks, parity


TIME_RE = re.compile(
    r"(?P<weeks>\d+(?:-\d+)?(?:单周|双周|周)?),星期(?P<day>[一二三四五六日天])第(?P<start>\d+)(?:-(?P<end>\d+))?节(?:\s+(?P<location>.+))?"
)


def parse_time_segments(time_text: str):
    segments = []
    for raw_line in (time_text or "").splitlines():
        line = raw_line.strip()
        if not line or line == "-split-主任务：":
            continue
        m = TIME_RE.search(line)
        if not m:
            continue
        weeks_token = m.group("weeks")
        weeks, parity = parse_weeks(weeks_token)
        start = int(m.group("start"))
        end = int(m.group("end") or m.group("start"))
        day = DAY_MAP[m.group("day")]
        location = (m.group("location") or "").strip()
        segments.append(
            {
                "raw": line,
                "weeks_token": weeks_token,
                "weeks": weeks,
                "parity": parity,
                "day": day,
                "start": start,
                "end": end,
                "location": location,
                "slot_label": f"{weeks_token}-{DAY_NAME[day]}-{start}-{end}",
            }
        )
    return segments


def parse_time_parts(time_text: str):
    lines = [line.strip() for line in (time_text or "").splitlines() if line.strip()]
    if "-split-主任务：" in lines:
        split_idx = lines.index("-split-主任务：")
        group_lines = lines[:split_idx]
        lecture_lines = lines[split_idx + 1 :]
        return parse_time_segments("\n".join(lecture_lines)), parse_time_segments("\n".join(group_lines))
    return parse_time_segments("\n".join(lines)), []


def plan_lookup():
    return {course["code"]: course for course in PLAN_COURSES}


def match_plan_course(row):
    row_code = row.get("D", "").strip()
    row_name = row.get("F", "").strip()
    row_name_norm = norm_text(row_name)

    for course in PLAN_COURSES:
        if row_code == course["code"]:
            return course, "exact_code"
        if row_code in course["alt_codes"]:
            return course, "alt_code"
    for course in PLAN_COURSES:
        if row_name_norm == norm_text(course["name"]):
            return course, "exact_name"
        if any(row_name_norm == norm_text(alt) for alt in course["alt_names"]):
            return course, "alt_name"
    return None, ""


def conflict_between_segments(seg_a, seg_b):
    if seg_a["day"] != seg_b["day"]:
        return False
    if seg_a["end"] < seg_b["start"] or seg_b["end"] < seg_a["start"]:
        return False
    if not (seg_a["weeks"] & seg_b["weeks"]):
        return False
    return True


def conflict_between_sections(section_a, section_b):
    for seg_a in section_a["segments"]:
        for seg_b in section_b["segments"]:
            if conflict_between_segments(seg_a, seg_b):
                return True
    return False


def build_section_rows(raw_rows):
    matched = []
    for raw in raw_rows:
        course, match_type = match_plan_course(raw)
        if not course:
            continue
        time_text = raw.get("N", "")
        segments = parse_time_segments(time_text)
        lecture_segments, group_segments = parse_time_parts(time_text)
        matched.append(
            {
                "plan_code": course["code"],
                "plan_name": course["name"],
                "plan_part": course["plan_part"],
                "module": course["module"],
                "recommended_term": course["recommended_term"],
                "prereq": course["prereq"],
                "note": course["note"],
                "match_type": match_type,
                "offered_code": raw.get("D", ""),
                "class_name": raw.get("A", ""),
                "offered_name": raw.get("F", ""),
                "course_nature": raw.get("H", ""),
                "course_category": raw.get("I", ""),
                "credit": raw.get("K", ""),
                "teacher": raw.get("M", ""),
                "time_text": time_text,
                "requirements": raw.get("O", ""),
                "limit_obj": raw.get("P", ""),
                "remark": raw.get("Q", ""),
                "dept": raw.get("R", ""),
                "campus": raw.get("S", ""),
                "capacity": raw.get("T", ""),
                "selected": raw.get("W", ""),
                "task_id": raw.get("AA", ""),
                "segments": segments,
                "lecture_segments": lecture_segments,
                "group_segments": group_segments,
            }
        )
    return matched


def build_conflicts(section_rows):
    index_rows = [row for row in section_rows if row["segments"]]
    conflicts = defaultdict(list)
    matrix = []
    for i, row_a in enumerate(index_rows):
        matrix_row = []
        for j, row_b in enumerate(index_rows):
            if i == j:
                matrix_row.append("")
                continue
            hit = conflict_between_sections(row_a, row_b)
            matrix_row.append("冲突" if hit else "")
            if hit:
                conflicts[row_a["task_id"]].append(row_b)
        matrix.append(matrix_row)
    return index_rows, conflicts, matrix


def schedule_grid(section_rows):
    grid = {(day, period): [] for day in range(1, 7) for period in range(1, 11)}
    for row in section_rows:
        for seg in row["segments"]:
            for period in range(seg["start"], seg["end"] + 1):
                if seg["day"] in range(1, 7) and period in range(1, 11):
                    text = (
                        f"{row['plan_code']} {row['offered_code']} | {row['class_name']}\n"
                        f"{seg['weeks_token']} | {row['teacher'] or '待定'}"
                    )
                    grid[(seg["day"], period)].append(text)
    return grid


def build_summary_rows(section_rows):
    offered_by_plan = defaultdict(list)
    for row in section_rows:
        offered_by_plan[row["plan_code"]].append(row)

    rows = []
    for course in PLAN_COURSES:
        offered_rows = offered_by_plan.get(course["code"], [])
        offered_codes = sorted({r["offered_code"] for r in offered_rows if r["offered_code"]})
        class_count = len(offered_rows)
        has_time = any(r["segments"] for r in offered_rows)
        rows.append(
            {
                "plan_part": course["plan_part"],
                "module": course["module"],
                "code": course["code"],
                "name": course["name"],
                "credits": course["credits"],
                "recommended_term": course["recommended_term"],
                "prereq": course["prereq"],
                "offered_in_2026_fall": "是" if offered_rows else "否",
                "offered_codes": ", ".join(offered_codes),
                "offered_sections": class_count,
                "time_info_ready": "是" if has_time else ("无开设" if not offered_rows else "否"),
                "note": course["note"],
            }
        )
    return rows


def filter_remaining_plan_courses(section_rows):
    remaining_plan_codes = {course["code"] for course in PLAN_COURSES if course["code"] not in COMPLETED_PLAN_CODES}
    filtered_sections = [row for row in section_rows if row["plan_code"] in remaining_plan_codes]
    filtered_summary = [row for row in build_summary_rows(filtered_sections) if row["code"] in remaining_plan_codes]
    return filtered_sections, filtered_summary


def extract_class_label(class_name: str):
    match = re.search(r"(\d+班)", class_name)
    return match.group(1) if match else class_name


def extract_group_label(class_name: str):
    match = re.search(r"(\d+组)", class_name)
    return match.group(1) if match else ""


def strip_group_suffix(class_name: str):
    return re.sub(r"-\d+组$", "", class_name)


def build_timetable_entries(section_rows):
    lecture_entries = {}
    group_entries = {}

    for row in section_rows:
        if row["plan_code"] in COMPLETED_PLAN_CODES:
            continue

        class_label = extract_class_label(row["class_name"])
        group_label = extract_group_label(row["class_name"])
        base_class = strip_group_suffix(row["class_name"])

        for seg in row["segments"]:
            is_lab_group = "机房" in (seg.get("location") or "")
            if not is_lab_group:
                key = (row["plan_code"], base_class, seg["day"], seg["start"], seg["end"], seg["weeks_token"])
                lecture_entries[key] = {
                    "plan_code": row["plan_code"],
                    "plan_name": row["plan_name"],
                    "teacher": row["teacher"] or "待定",
                    "type": "大课",
                    "class_label": class_label,
                    "group_label": "",
                    "segment": seg,
                    "display": f"大 {row['plan_code']} {row['plan_name']} {class_label}",
                }
            else:
                effective_group = group_label or class_label
                display_label = f"{class_label}-{group_label}" if group_label else class_label
                key = (
                    row["plan_code"],
                    base_class,
                    effective_group,
                    seg["day"],
                    seg["start"],
                    seg["end"],
                    seg["weeks_token"],
                    seg.get("location", ""),
                )
                group_entries[key] = {
                    "plan_code": row["plan_code"],
                    "plan_name": row["plan_name"],
                    "teacher": row["teacher"] or "待定",
                    "type": "分组",
                    "class_label": class_label,
                    "group_label": effective_group,
                    "segment": seg,
                    "display": f"组 {row['plan_code']} {row['plan_name']} {display_label}",
                }

    entries = list(lecture_entries.values()) + list(group_entries.values())
    entries.sort(
        key=lambda x: (
            x["segment"]["day"],
            x["segment"]["start"],
            x["plan_code"],
            0 if x["type"] == "大课" else 1,
            x["class_label"],
            x["group_label"],
        )
    )
    return entries


def course_color_map(entries):
    codes = sorted({entry["plan_code"] for entry in entries})
    return {code: COURSE_COLORS[idx % len(COURSE_COLORS)] for idx, code in enumerate(codes)}


def write_optimized_timetable(workbook, entries, summary_rows):
    ws = workbook.add_worksheet("周时间表")

    base_border = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})
    header = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "align": "center", "valign": "vcenter"})
    note = workbook.add_format({"text_wrap": True, "valign": "top"})
    legend_header = workbook.add_format({"bold": True, "bg_color": "#E2F0D9", "border": 1})
    legend_cell = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})

    ws.freeze_panes(1, 1)
    ws.set_column(0, 0, 8)
    for day in range(1, 7):
        ws.write(0, day, DAY_NAME[day], header)
        ws.set_column(day, day, 34)
    for col in range(8, 12):
        ws.set_column(col, col, 22)

    pair_rows = [
        (1, 2, "第1-2节"),
        (3, 4, "第3-4节"),
        (5, 6, "第5-6节"),
        (7, 8, "第7-8节"),
        (9, 10, "第9-10节"),
    ]
    row_index_by_start = {start: idx + 1 for idx, (start, _, _) in enumerate(pair_rows)}

    for idx, (_, _, label) in enumerate(pair_rows, start=1):
        ws.write(idx, 0, label, header)
        ws.set_row(idx, 110)
        for day in range(1, 7):
            ws.write_blank(idx, day, None, base_border)

    grid = {(day, row_idx): [] for day in range(1, 7) for row_idx in range(1, 6)}
    for entry in entries:
        seg = entry["segment"]
        pair_start = seg["start"] if seg["start"] % 2 == 1 else seg["start"] - 1
        row_idx = row_index_by_start.get(pair_start)
        if row_idx and (seg["day"], row_idx) in grid:
            grid[(seg["day"], row_idx)].append(entry)

    color_map = course_color_map(entries)
    text_formats = {
        code: workbook.add_format({"font_color": color, "bold": True})
        for code, color in color_map.items()
    }

    for day in range(1, 7):
        for row_idx in range(1, 6):
            cell_entries = grid[(day, row_idx)]
            if not cell_entries:
                continue
            parts = []
            for idx, entry in enumerate(cell_entries):
                if idx:
                    parts.append("\n")
                parts.append(text_formats[entry["plan_code"]])
                parts.append(entry["display"])
            if len(parts) >= 3:
                ws.write_rich_string(row_idx, day, *parts, base_border)
            else:
                entry = cell_entries[0]
                single_fmt = workbook.add_format(
                    {"border": 1, "valign": "top", "text_wrap": True, "font_color": color_map[entry["plan_code"]], "bold": True}
                )
                ws.write(row_idx, day, entry["display"], single_fmt)

    ws.write(0, 8, "颜色图例", legend_header)
    ws.write(0, 9, "说明", legend_header)
    legend_codes = sorted(color_map)
    for idx, code in enumerate(legend_codes, start=1):
        course_name = next((row["name"] for row in summary_rows if row["code"] == code), code)
        fmt = workbook.add_format({"border": 1, "font_color": color_map[code], "bold": True})
        ws.write(idx, 8, code, fmt)
        ws.write(idx, 9, course_name, legend_cell)

    ws.write(8, 8, "读表规则", legend_header)
    ws.write(
        8,
        9,
        "大 = 大课，只标到班；组 = 实验/机房/分组课，标到组。"
        "\n如果同一班拆成多个组，大课只保留一次，不重复显示。"
        "\n时间轴按双节合并显示：1-2、3-4、5-6、7-8、9-10。"
        "\n周表主区域不再显示 1-16周 这类周次信息；若你后续需要，我可以再单独做一版仅用于核单双周冲突的表。",
        note,
    )

    ai303_rows = [row for row in summary_rows if row["code"] == "COE303"]
    ai303_note = "AI303 已单独处理：课表中出现但没有完整时间信息，因此未放入周时间表。"
    if ai303_rows:
        ai303_note += f"\n当前备注：{ai303_rows[0]['note']}"
    ws.write(10, 8, "AI303", legend_header)
    ws.write(10, 9, ai303_note, note)

    unopened = [row for row in summary_rows if row["offered_in_2026_fall"] == "否"]
    if unopened:
        ws.write(12, 8, "本学期未开", legend_header)
        ws.write(12, 9, "\n".join(f"{row['code']} {row['name']}" for row in unopened), note)

    no_time = [
        row for row in summary_rows if row["offered_in_2026_fall"] == "是" and row["time_info_ready"] == "否" and row["code"] != "COE303"
    ]
    if no_time:
        ws.write(14, 8, "已开但缺时间", legend_header)
        ws.write(14, 9, "\n".join(f"{row['code']} {row['name']}" for row in no_time), note)


def add_formats(workbook):
    return {
        "header": workbook.add_format(
            {
                "bold": True,
                "bg_color": "#D9EAF7",
                "border": 1,
                "valign": "top",
                "text_wrap": True,
            }
        ),
        "cell": workbook.add_format({"border": 1, "valign": "top", "text_wrap": True}),
        "cell_center": workbook.add_format({"border": 1, "valign": "top", "align": "center", "text_wrap": True}),
        "required": workbook.add_format({"border": 1, "valign": "top", "text_wrap": True, "bg_color": "#EAF4E2"}),
        "elective": workbook.add_format({"border": 1, "valign": "top", "text_wrap": True, "bg_color": "#FFF2CC"}),
        "warn": workbook.add_format({"border": 1, "valign": "top", "text_wrap": True, "bg_color": "#FCE4D6"}),
        "note": workbook.add_format({"valign": "top", "text_wrap": True}),
        "conflict": workbook.add_format({"border": 1, "align": "center", "valign": "vcenter", "bg_color": "#F4CCCC"}),
        "grid": workbook.add_format({"border": 1, "valign": "top", "text_wrap": True}),
        "title": workbook.add_format({"bold": True, "font_size": 14}),
    }


def write_summary_sheet(workbook, fmts, summary_rows):
    ws = workbook.add_worksheet("培养方案总表")
    headers = [
        "培养方案部分",
        "模块",
        "课程代码",
        "课程名称",
        "学分",
        "建议学期",
        "先修课程",
        "2026秋是否开设",
        "识别到的开课代码",
        "识别到的课次条数",
        "时间信息是否完整",
        "备注",
    ]
    ws.freeze_panes(1, 0)
    for col, header in enumerate(headers):
        ws.write(0, col, header, fmts["header"])

    widths = [10, 12, 12, 18, 8, 10, 18, 12, 16, 14, 14, 36]
    for idx, width in enumerate(widths):
        ws.set_column(idx, idx, width)

    for row_idx, row in enumerate(summary_rows, start=1):
        base_fmt = fmts["required"] if row["plan_part"] == "必修" else fmts["elective"]
        values = [
            row["plan_part"],
            row["module"],
            row["code"],
            row["name"],
            row["credits"],
            row["recommended_term"],
            row["prereq"],
            row["offered_in_2026_fall"],
            row["offered_codes"],
            row["offered_sections"],
            row["time_info_ready"],
            row["note"],
        ]
        for col, value in enumerate(values):
            fmt = base_fmt if col != 11 else (fmts["warn"] if "确认" in str(value) or "不完整" in str(value) else base_fmt)
            ws.write(row_idx, col, value, fmt)
    ws.autofilter(0, 0, len(summary_rows), len(headers) - 1)


def write_offerings_sheet(workbook, fmts, section_rows):
    ws = workbook.add_worksheet("2026秋开设节次")
    headers = [
        "培养方案代码",
        "培养方案课程名",
        "培养方案部分",
        "模块",
        "建议学期",
        "先修",
        "识别方式",
        "开课代码",
        "教学班",
        "教师",
        "学分",
        "课程性质(开课表)",
        "课程类别(开课表)",
        "上课信息原文",
        "标准化时间块",
        "限制对象",
        "备注",
        "开课院系",
        "容量",
        "已选",
        "任务号",
    ]
    ws.freeze_panes(1, 0)
    for col, header in enumerate(headers):
        ws.write(0, col, header, fmts["header"])
    widths = [12, 18, 10, 12, 10, 18, 10, 12, 28, 14, 8, 12, 14, 34, 34, 18, 18, 14, 8, 8, 24]
    for idx, width in enumerate(widths):
        ws.set_column(idx, idx, width)

    for row_idx, row in enumerate(section_rows, start=1):
        base_fmt = fmts["required"] if row["plan_part"] == "必修" else fmts["elective"]
        normalized = "\n".join(
            f"{DAY_NAME[seg['day']]} 第{seg['start']}-{seg['end']}节 | {seg['weeks_token']} | {seg['location']}"
            for seg in row["segments"]
        )
        values = [
            row["plan_code"],
            row["plan_name"],
            row["plan_part"],
            row["module"],
            row["recommended_term"],
            row["prereq"],
            row["match_type"],
            row["offered_code"],
            row["class_name"],
            row["teacher"],
            row["credit"],
            row["course_nature"],
            row["course_category"],
            row["time_text"],
            normalized,
            row["limit_obj"],
            row["remark"] or row["note"],
            row["dept"],
            row["capacity"],
            row["selected"],
            row["task_id"],
        ]
        for col, value in enumerate(values):
            fmt = base_fmt
            if col in (13, 14, 16) and ("确认" in str(value) or "不完整" in str(value)):
                fmt = fmts["warn"]
            ws.write(row_idx, col, value, fmt)
    ws.autofilter(0, 0, len(section_rows), len(headers) - 1)


def write_conflict_sheet(workbook, fmts, indexed_rows, conflicts):
    ws = workbook.add_worksheet("冲突摘要")
    headers = [
        "培养方案代码",
        "课程名",
        "教学班",
        "教师",
        "标准化时间块",
        "冲突节次数",
        "冲突对象",
    ]
    for col, header in enumerate(headers):
        ws.write(0, col, header, fmts["header"])
    widths = [12, 18, 28, 14, 34, 10, 54]
    for idx, width in enumerate(widths):
        ws.set_column(idx, idx, width)
    ws.freeze_panes(1, 0)

    for row_idx, row in enumerate(indexed_rows, start=1):
        base_fmt = fmts["required"] if row["plan_part"] == "必修" else fmts["elective"]
        normalized = "\n".join(
            f"{DAY_NAME[seg['day']]} 第{seg['start']}-{seg['end']}节 | {seg['weeks_token']}" for seg in row["segments"]
        )
        hit_rows = conflicts.get(row["task_id"], [])
        hit_desc = "\n".join(
            f"{hit['plan_code']} {hit['class_name']} | {hit['teacher'] or '待定'}" for hit in hit_rows
        )
        values = [
            row["plan_code"],
            row["plan_name"],
            row["class_name"],
            row["teacher"],
            normalized,
            len(hit_rows),
            hit_desc,
        ]
        for col, value in enumerate(values):
            fmt = fmts["warn"] if col >= 5 and hit_rows else base_fmt
            ws.write(row_idx, col, value, fmt)
    ws.autofilter(0, 0, len(indexed_rows), len(headers) - 1)


def write_matrix_sheet(workbook, fmts, indexed_rows, matrix):
    ws = workbook.add_worksheet("冲突矩阵")
    ws.freeze_panes(1, 1)
    labels = [f"{row['plan_code']} | {row['class_name']}" for row in indexed_rows]
    ws.write(0, 0, "节次", fmts["header"])
    for idx, label in enumerate(labels, start=1):
        ws.write(0, idx, label, fmts["header"])
        ws.write(idx, 0, label, fmts["header"])
        ws.set_column(idx, idx, 14)
    ws.set_column(0, 0, 28)
    for i, row_vals in enumerate(matrix, start=1):
        for j, value in enumerate(row_vals, start=1):
            fmt = fmts["conflict"] if value else fmts["cell_center"]
            ws.write(i, j, value, fmt)


def write_grid_sheet(workbook, fmts, section_rows):
    ws = workbook.add_worksheet("周课表视图")
    ws.write(0, 0, "节次", fmts["header"])
    for day in range(1, 7):
        ws.write(0, day, DAY_NAME[day], fmts["header"])
        ws.set_column(day, day, 32)
    ws.set_column(0, 0, 8)
    grid = schedule_grid(section_rows)
    for period in range(1, 11):
        ws.write(period, 0, f"第{period}节", fmts["header"])
        ws.set_row(period, 80)
        for day in range(1, 7):
            cell_text = "\n---\n".join(grid[(day, period)])
            ws.write(period, day, cell_text, fmts["grid"])

    ws.write(12, 0, "说明", fmts["header"])
    ws.merge_range(
        12,
        1,
        13,
        6,
        "这是把 2026 秋季所有与 AI 培养方案专业必修/专业选修可识别对应的开课节次，全部投影到周课表中的结果。"
        "\n同一格里出现多条，说明这些节次在该时间块重叠；是否真正冲突还要结合单双周和起止周，详见“冲突摘要”和“冲突矩阵”两张表。"
        "\nAI303 因缺少时间信息，没有出现在此表中。",
        fmts["note"],
    )


def write_ai303_sheet(workbook, fmts, section_rows):
    ws = workbook.add_worksheet("AI303单独说明")
    ws.set_column(0, 0, 16)
    ws.set_column(1, 1, 60)
    ws.write(0, 0, "字段", fmts["header"])
    ws.write(0, 1, "内容", fmts["header"])

    ai_rows = [row for row in section_rows if row["plan_code"] == "COE303"]
    if not ai_rows:
        ws.write(1, 0, "说明", fmts["cell"])
        ws.write(1, 1, "未在 2026 秋季开课表中识别到 AI303 / COE303。", fmts["cell"])
        return

    row = ai_rows[0]
    pairs = [
        ("培养方案课程", f"{row['plan_code']} {row['plan_name']}"),
        ("识别到的开课代码", row["offered_code"]),
        ("教学班", row["class_name"]),
        ("教师", row["teacher"] or "未给出"),
        ("上课信息", row["time_text"] or "未给出"),
        ("容量", row["capacity"] or "未给出"),
        ("说明", "该课已在开课表中出现，但当前缺少完整时间和教师信息，因此未参与冲突计算与周课表视图。"),
    ]
    for idx, (k, v) in enumerate(pairs, start=1):
        ws.write(idx, 0, k, fmts["cell"])
        ws.write(idx, 1, v, fmts["warn"] if idx == len(pairs) else fmts["cell"])


def build_workbook(output_path: Path, section_rows, summary_rows):
    indexed_rows, conflicts, matrix = build_conflicts(section_rows)

    workbook = xlsxwriter.Workbook(str(output_path))
    fmts = add_formats(workbook)

    write_summary_sheet(workbook, fmts, summary_rows)
    write_offerings_sheet(workbook, fmts, section_rows)
    write_conflict_sheet(workbook, fmts, indexed_rows, conflicts)
    write_matrix_sheet(workbook, fmts, indexed_rows, matrix)
    write_grid_sheet(workbook, fmts, indexed_rows)
    write_ai303_sheet(workbook, fmts, section_rows)

    workbook.close()
    print(f"Wrote {output_path}")


def build_optimized_timetable_workbook(output_path: Path, section_rows, summary_rows):
    workbook = xlsxwriter.Workbook(str(output_path))
    entries = build_timetable_entries(section_rows)
    write_optimized_timetable(workbook, entries, summary_rows)
    workbook.close()
    print(f"Wrote {output_path}")


def main():
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    raw_rows = load_workbook_rows(COURSE_XLSX)
    section_rows = build_section_rows(raw_rows)
    remaining_sections, remaining_summary = filter_remaining_plan_courses(section_rows)
    build_optimized_timetable_workbook(OUTPUT_REMAINING_TIMETABLE_XLSX, remaining_sections, remaining_summary)


if __name__ == "__main__":
    main()
