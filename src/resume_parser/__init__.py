import json
import re
import statistics
import sys
from dataclasses import dataclass
from datetime import date

import docx2txt
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTChar

_EMAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
_PHONE = re.compile(
    r"(?<!\d)(?:\+91[\-\s]?|91[\-\s]|0)?([6-9](?:[\-\s]?\d){9})(?!\d)"
)
_PHONE_US = re.compile(r"\(?\d{3}\)?[\s.\-]?\d{3}[\s.\-]?\d{4}")
_URL = re.compile(r"https?://[^\s|]+")
_SCHOOL = re.compile(r"\b(institute|university|college|school|iiit|iit|nit)\b", re.I)
_DEGREE = re.compile(
    r"\b(associate|bachelor|master|ph\.?\s?d|b\.?\s?tech|m\.?\s?tech|b\.?\s?e\b|m\.?\s?e\b|mba|b\.sc|m\.sc)\b",
    re.I,
)
_DATE = re.compile(
    r"(?:19|20)\d{2}|present|jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?",
    re.I,
)
_MONTH = (
    r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
    r"jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
)
_YEAR = r"(?:19|20)\d{2}"
_BOUND = rf"(?:{_MONTH}\.?\s+{_YEAR}|{_YEAR}\s*/\s*(?:0?[1-9]|1[0-2])|{_YEAR})"
_DATE_SPAN = re.compile(
    rf"\b{_BOUND}\s*[-–—\u2212]+\s*(?:present|current|ongoing|now|{_BOUND})\b",
    re.I,
)
_TITLE = re.compile(
    r"\b(engineer|developer|intern|analyst|manager|designer|consultant|architect|scientist|officer|lead)\b",
    re.I,
)
_BULLET = re.compile(r"^[•●▪◦\-\u2022]\s*")
_JUNK = re.compile(r"(\(demo\)|\(github\)|watch)", re.I)
_LETTERS_ONLY = re.compile(r"[A-Za-z\s.]+")
_SECTION_KEYWORDS = (
    "experience",
    "education",
    "project",
    "skill",
    "job",
    "course",
    "objective",
    "summary",
    "award",
    "honor",
    "talk",
    "accomplishment",
    "achievement",
    "presentation",
)


@dataclass
class Item:
    text: str
    x: float
    y: float
    bold: bool


@dataclass
class Line:
    items: list
    y: float

    def text(self, drop_junk=False):
        parts = []
        for item in self.items:
            if drop_junk and _JUNK.fullmatch(item.text.strip()):
                continue
            parts.append(item.text.strip())
        return re.sub(r"\s+", " ", " ".join(parts)).strip()


def _is_bold_font(fontname):
    base = fontname.lower().split("+")[-1]
    if any(token in base for token in ("bold", "semibold", "demi", "black", "heavy")):
        return True
    # LaTeX Computer Modern bold is CMBX / CMSSBX, not "Bold".
    return "cmbx" in base or "cmssbx" in base


def _flat(text):
    return re.sub(r"[^a-z]", "", text.lower())


def _canonical(title):
    low = title.lower()
    flat = _flat(title)
    if any(word in low or word in flat for word in ("accomplishment", "achievement", "award", "honor")):
        return "accomplishments"
    if "talk" in low or "presentation" in low or "talk" in flat or "presentation" in flat:
        return "talks"
    if "project" in low or "project" in flat:
        return "projects"
    if "skill" in low or "skill" in flat:
        return "skills"
    if any(word in low or word in flat for word in ("educat", "academic", "qualificat")):
        return "education"
    if "experience" in low or "experience" in flat or low.strip() == "employment":
        return "work"
    if "summary" in low or "objective" in low or "summary" in flat or "objective" in flat:
        return "summary"
    return "custom"


def _is_section_title(line, index):
    if index < 2 or len(line.items) != 1:
        return False
    item = line.items[0]
    text = item.text.strip()
    letters = re.sub(r"[^A-Za-z]", "", text)
    if item.bold and letters and letters.isupper() and re.fullmatch(r"[A-Z0-9 &/\-]+", text):
        return True
    words = [word for word in text.split() if word != "&"]
    if (
        len(words) <= 3
        and re.fullmatch(r"[A-Za-z &]+", text)
        and text[:1].isupper()
        and any(keyword in text.lower() or keyword in _flat(text) for keyword in _SECTION_KEYWORDS)
    ):
        return True
    return False


def _items_from_chars(chars, avg_width):
    items = []
    current = []

    def flush():
        if not current:
            return
        text = ""
        for char, space in current:
            if space and text:
                text += " "
            text += char.get_text()
        text = re.sub(r" +", " ", text).strip()
        if text:
            items.append(
                Item(
                    text=text,
                    x=current[0][0].x0,
                    y=current[0][0].y0,
                    bold=any(_is_bold_font(char.fontname) for char, _ in current),
                )
            )

    previous = None
    for char in chars:
        gap = None if previous is None else char.x0 - previous.x1
        if gap is not None and gap > avg_width * 3 and current:
            flush()
            current = []
            gap = None
        current.append((char, gap is not None and gap > 0.45))
        previous = char
    flush()
    return items


def _gutter_split(chars, page_width):
    """Return an x cutoff when the page has a full-height empty column gap."""
    if len(chars) < 40 or page_width <= 0:
        return None
    bin_width = 4
    bin_count = int(page_width // bin_width) + 1
    counts = [0] * bin_count
    for char in chars:
        index = int(max(char.x0, 0) // bin_width)
        if 0 <= index < bin_count:
            counts[index] += 1
    margin = int(page_width * 0.08 / bin_width)
    ys = [char.y0 for char in chars]
    span = max(ys) - min(ys) or 1
    best = None
    start = None
    limit = bin_count - 1 - margin
    for index, count in enumerate(counts):
        in_gutter = count == 0 and margin <= index <= limit
        if in_gutter:
            if start is None:
                start = index
            continue
        if start is not None:
            best = _wider_gutter(start, index, bin_width, chars, span, best)
            start = None
    if start is not None:
        best = _wider_gutter(start, limit + 1, bin_width, chars, span, best)
    return best


def _wider_gutter(start, end, bin_width, chars, span, best):
    width = (end - start) * bin_width
    if width < 12:
        return best
    midpoint = (start + end) * bin_width / 2
    left = [char for char in chars if char.x0 < midpoint]
    right = [char for char in chars if char.x0 >= midpoint]
    if len(left) < 0.12 * len(chars) or len(right) < 0.12 * len(chars):
        return best

    def height(side):
        side_ys = [char.y0 for char in side]
        return max(side_ys) - min(side_ys)

    if height(left) < 0.35 * span or height(right) < 0.35 * span:
        return best
    if best is None or width > best[0]:
        return (width, midpoint)
    return best


def _lines_from_chars(chars):
    chars = [char for char in chars if char.get_text()]
    if not chars:
        return []
    widths = [
        char.x1 - char.x0
        for char in chars
        if char.get_text().strip() and "Bold" not in char.fontname
    ]
    avg_width = (sum(widths) / len(widths)) if widths else 4.0
    ordered = sorted(chars, key=lambda char: (-char.y0, char.x0))
    groups = []
    for char in ordered:
        if groups and abs(char.y0 - groups[-1][0].y0) <= 2.2:
            groups[-1].append(char)
        else:
            groups.append([char])
    lines = []
    for group in groups:
        group.sort(key=lambda char: char.x0)
        items = _items_from_chars(group, avg_width)
        if items:
            y = sum(char.y0 for char in group) / len(group)
            lines.append(Line(items, y))
    return lines


def _read_pdf(path):
    lines = []
    for page in extract_pages(path):
        chars = []

        def walk(obj):
            if isinstance(obj, LTChar):
                chars.append(obj)
            elif hasattr(obj, "__iter__"):
                for child in obj:
                    walk(child)

        walk(page)
        split = _gutter_split(chars, page.width)
        if split is None:
            lines.extend(_lines_from_chars(chars))
            continue
        _width, cutoff = split
        lines.extend(_lines_from_chars([char for char in chars if char.x0 < cutoff]))
        lines.extend(_lines_from_chars([char for char in chars if char.x0 >= cutoff]))
    return lines


def _read_docx(path):
    text = docx2txt.process(path) or ""
    lines = []
    y = 0.0
    for raw in text.splitlines():
        raw = raw.strip()
        if not raw:
            continue
        y -= 12
        lines.append(Line([Item(raw, 0, y, False)], y))
    return lines


def _sections(lines):
    sections = []
    current_key = "profile"
    current_title = "profile"
    bucket = []
    for index, line in enumerate(lines):
        if _is_section_title(line, index):
            sections.append((current_key, current_title, bucket))
            current_title = line.text()
            current_key = _canonical(current_title)
            bucket = []
            continue
        bucket.append(line)
    sections.append((current_key, current_title, bucket))
    return sections


def _best_name(lines):
    best_text = None
    best_score = 0
    for line in lines:
        for item in line.items:
            text = item.text.strip()
            score = 0
            if _LETTERS_ONLY.fullmatch(text):
                score += 3
            if item.bold:
                score += 2
            letters = re.sub(r"[^A-Za-z]", "", text)
            if letters and letters.isupper() and len(letters) > 1:
                score += 2
            if "@" in text:
                score -= 4
            if re.search(r"\d", text):
                score -= 4
            if "," in text:
                score -= 4
            if "/" in text:
                score -= 4
            if score > best_score:
                best_score = score
                best_text = text
    return best_text


def _phone(text):
    match = _PHONE.search(text)
    if match:
        return "91" + re.sub(r"\D", "", match.group(1))
    us = _PHONE_US.search(text)
    if us:
        return re.sub(r"\D", "", us.group(0))
    return None


def _profile(lines):
    blob = "\n".join(line.text() for line in lines)
    emails = _EMAIL.findall(blob)
    urls = []
    for found in _URL.findall(blob):
        cleaned = found.rstrip(".,)")
        if cleaned not in urls:
            urls.append(cleaned)
    phone = _phone(blob)
    name = _best_name(lines)
    location = None
    for line in lines:
        text = line.text()
        if not text or text == name:
            continue
        if _EMAIL.search(text) or _URL.search(text) or _phone(text):
            continue
        location = text
        break
    return {
        "name": name,
        "email": emails[0] if emails else None,
        "phone": phone,
        "url": urls,
        "location": location,
    }


def _subsections(lines, is_header):
    if not lines:
        return []
    groups = [[]]
    for line in lines:
        if is_header(line) and groups[-1]:
            groups.append([line])
        else:
            groups[-1].append(line)
    return [group for group in groups if group]


def _line_gap_subsections(lines):
    if not lines:
        return []
    gaps = [lines[i].y - lines[i + 1].y for i in range(len(lines) - 1)]
    body = [gap for gap in gaps if 0 < gap <= 14]
    typical = statistics.median(body) if body else 10
    groups = [[lines[0]]]
    for line, gap in zip(lines[1:], gaps):
        if gap > typical * 1.4:
            groups.append([line])
        else:
            groups[-1].append(line)
    return groups


def _descriptions(lines):
    descriptions = []
    for line in lines:
        text = line.text(drop_junk=True)
        if not text:
            continue
        if _BULLET.match(text) or not descriptions:
            descriptions.append(_BULLET.sub("", text).strip())
        elif descriptions[-1].endswith("-"):
            descriptions[-1] = f"{descriptions[-1]}{text}".strip()
        else:
            descriptions[-1] = f"{descriptions[-1]} {text}".strip()
    return descriptions


def _is_edu_header(line):
    return any(
        item.bold and (_SCHOOL.search(item.text) or _DEGREE.search(item.text))
        for item in line.items
    )


def _education(lines):
    schools = []
    if any(_is_edu_header(line) for line in lines):
        groups = _subsections(lines, _is_edu_header)
    else:
        groups = _line_gap_subsections(lines)
    for group in groups:
        school = degree = date = None
        leftover = []
        for line in group:
            extras = []
            line_degree = None
            for item in line.items:
                text = item.text.strip()
                if school is None and _SCHOOL.search(text):
                    school = text
                elif line_degree is None and _DEGREE.search(text):
                    line_degree = text
                elif date is None and _DATE.search(text) and not _SCHOOL.search(text):
                    date = text
                else:
                    extras.append(text)
            if line_degree and degree is None:
                degree = ", ".join([line_degree, *extras]) if extras else line_degree
            else:
                leftover.extend(extras)
        if school or degree:
            schools.append(
                {
                    "school": school,
                    "degree": degree,
                    "date": date,
                    "descriptions": [text for text in leftover if text and text not in {school, degree, date}],
                }
            )
    return schools


def _split_title_company(text):
    if "," in text:
        left, right = text.split(",", 1)
        if _TITLE.search(left):
            return left.strip(), right.strip()
    return text.strip(), None


def _split_role(text):
    parts = [part.strip() for part in re.split(r"\s*\|\s*", text) if part.strip()]
    if len(parts) < 2:
        return _split_title_company(text)
    if _TITLE.search(parts[1]) and not _TITLE.search(parts[0]):
        return parts[1], parts[0]
    return parts[0], " | ".join(parts[1:])


def _peel_date(text):
    cleaned = text.strip()
    match = _DATE_SPAN.search(cleaned)
    if not match:
        return None, cleaned
    before = cleaned[: match.start()].strip(" |,;-–—")
    after = cleaned[match.end() :].strip(" |,;-–—")
    if before and after:
        return None, cleaned
    return match.group(0).strip(), before or after


def _is_job_header(line):
    text = line.text()
    if not text or _BULLET.match(text):
        return False
    span = _peel_date(text)[0]
    titled = bool(_TITLE.search(text))
    bold = any(item.bold for item in line.items)
    if span and (bold or titled):
        return True
    return bool(bold and titled and len(text.split()) <= 8)


def _looks_like_company(text):
    text = text.strip()
    if not text or _BULLET.match(text) or text.endswith(".") or len(text) > 60:
        return False
    if _DATE.search(text) or _TITLE.search(text):
        return False
    return True


def _header_fields(line):
    title = company = date = None
    for item in line.items:
        peeled_date, rest = _peel_date(item.text)
        if peeled_date and not rest:
            if date is None:
                date = peeled_date
            continue
        if peeled_date:
            if date is None:
                date = peeled_date
            if title is None:
                title, split_company = _split_role(rest)
                if company is None:
                    company = split_company
            elif company is None and _looks_like_company(rest):
                company = rest
            continue
        if title is None:
            title, split_company = _split_role(item.text)
            if company is None:
                company = split_company
        elif company is None and _looks_like_company(item.text):
            company = item.text.strip()
    return title, company, date


def _span_inside(inner, outer):
    if not inner or not outer:
        return False
    return outer[0] <= inner[0] and inner[1] <= outer[1]


def _nest_under_role(jobs, title, company, date, descriptions):
    """Fold a project line into the role it sits inside, when the resume lists both."""
    if not title or not company or _TITLE.search(title):
        return False
    span = _parse_span(date)
    for job in jobs:
        if not _TITLE.search(job.get("title") or ""):
            continue
        if (job.get("company") or "").strip().lower() != company.strip().lower():
            continue
        if not _span_inside(span, _parse_span(job.get("date"))):
            continue
        heading = f"{title} ({date})" if date else title
        job["descriptions"].append(heading)
        job["descriptions"].extend(descriptions)
        return True
    return False


def _work(lines):
    headers = [index for index, line in enumerate(lines) if _is_job_header(line)]
    if not headers:
        return []
    jobs = []
    for position, start in enumerate(headers):
        end = headers[position + 1] if position + 1 < len(headers) else len(lines)
        title, company, date = _header_fields(lines[start])
        if start > 0 and company is None and _looks_like_company(lines[start - 1].text()):
            company = lines[start - 1].text()
            if jobs and jobs[-1]["descriptions"] and jobs[-1]["descriptions"][-1] == company:
                jobs[-1]["descriptions"].pop()
        body_end = end - 1 if end < len(lines) and _looks_like_company(lines[end - 1].text()) else end
        body = lines[start + 1 : body_end]
        if not (title or company or date or body):
            continue
        descriptions = _descriptions(body)
        if _nest_under_role(jobs, title, company, date, descriptions):
            continue
        jobs.append(
            {
                "company": company,
                "title": title,
                "date": date,
                "descriptions": descriptions,
            }
        )
    return jobs


def _projects(lines):
    projects = []
    current = None
    for line in lines:
        text = line.text(drop_junk=True)
        if not text:
            continue
        bold = any(item.bold for item in line.items)
        if bold and not _DATE.fullmatch(text):
            if current:
                projects.append(current)
            date = None
            name = text
            for item in line.items:
                if _DATE.search(item.text) and not _TITLE.search(item.text):
                    date = item.text.strip()
                elif item.bold:
                    name = item.text.strip()
            current = {"name": name, "date": date, "descriptions": []}
            continue
        if current is None:
            current = {"name": text, "date": None, "descriptions": []}
            continue
        if _BULLET.match(text) or not current["descriptions"]:
            current["descriptions"].append(_BULLET.sub("", text).strip())
        elif current["descriptions"][-1].endswith("-"):
            current["descriptions"][-1] = f"{current['descriptions'][-1]}{text}".strip()
        else:
            current["descriptions"][-1] = f"{current['descriptions'][-1]} {text}".strip()
    if current:
        projects.append(current)
    return projects


def _plain_lines(lines, bullets_only=False):
    texts = []
    for line in lines:
        text = line.text(drop_junk=True)
        if text:
            texts.append(text)
    if bullets_only:
        bullets = [_BULLET.sub("", text).strip() for text in texts if _BULLET.match(text)]
        if bullets:
            return bullets
    return [_BULLET.sub("", text).strip() for text in texts]


_MONTHS = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}
_SKILL_HEADER = re.compile(
    r"^(programming languages|languages|frameworks(?:/libraries)?|libraries|databases|tools|others|soft skills|technical skills)$",
    re.I,
)
_GLANCE_SKIP = {"html", "css", "git"}


def _parse_bound(text, edge):
    cleaned = text.strip().lower().replace(".", " ")
    if re.search(r"\b(present|current|ongoing|now)\b", cleaned):
        today = date.today()
        return (today.year, today.month)
    numbered = re.search(r"((?:19|20)\d{2})\s*/\s*(0?[1-9]|1[0-2])\b", text.strip())
    if numbered:
        return (int(numbered.group(1)), int(numbered.group(2)))
    numbered = re.search(r"\b(0?[1-9]|1[0-2])\s*/\s*((?:19|20)\d{2})", text.strip())
    if numbered:
        return (int(numbered.group(2)), int(numbered.group(1)))
    year = re.search(r"(?:19|20)\d{2}", cleaned)
    if not year:
        return None
    month_name = re.search(r"[a-z]+", cleaned)
    if month_name and month_name.group() in _MONTHS:
        month = _MONTHS[month_name.group()]
    elif edge == "start":
        month = 7
    else:
        month = 12
    return (int(year.group()), month)


def _parse_span(text):
    if not text:
        return None
    parts = [part for part in re.split(r"\s*(?:–|—|-|to)\s*", text, flags=re.I) if part.strip()]
    if not parts:
        return None
    start = _parse_bound(parts[0], "start")
    end = _parse_bound(parts[-1], "end") if len(parts) > 1 else _parse_bound(parts[0], "end")
    if not start or not end or end < start:
        return None
    return (start, end)


def _merge_spans(spans):
    merged = []
    for start, end in sorted(spans):
        if not merged or start > merged[-1][1]:
            merged.append([start, end])
        elif end > merged[-1][1]:
            merged[-1][1] = end
    return [(start, end) for start, end in merged]


def _month_count(start, end):
    return (end[0] - start[0]) * 12 + (end[1] - start[1])


def _shift_month(bound, delta):
    index = bound[0] * 12 + (bound[1] - 1) + delta
    return (index // 12, index % 12 + 1)


def _duration_label(months):
    if months <= 0:
        return None
    years, rem = divmod(months, 12)
    parts = []
    if years:
        parts.append("1 year" if years == 1 else f"{years} years")
    if rem:
        parts.append("1 month" if rem == 1 else f"{rem} months")
    return " ".join(parts)


_MONTH_ABBR = (
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
)


def _month_label(bound):
    return f"{_MONTH_ABBR[bound[1] - 1]} {bound[0]}"


def _employment_spans(work):
    spans = []
    for job in work:
        span = _parse_span(job.get("date"))
        if span:
            spans.append(span)
    return _merge_spans(spans)


def _gap_label(months, start, end):
    when = _month_label(start) if start == end else f"{_month_label(start)} – {_month_label(end)}"
    return f"{_duration_label(months)} · {when}"


def _open_months(left, right):
    """Months strictly between two bounds. Adjacent months (Dec then Jan) are 0."""
    return _month_count(left, right) - 1


def _work_gaps(work):
    merged = _employment_spans(work)
    if not merged:
        return None
    gaps = []
    for (_, previous_end), (start, _) in zip(merged, merged[1:]):
        months = _open_months(previous_end, start)
        if months > 0:
            gaps.append(_gap_label(months, _shift_month(previous_end, 1), _shift_month(start, -1)))
    today = (date.today().year, date.today().month)
    latest_end = merged[-1][1]
    if latest_end < today:
        months = _month_count(latest_end, today)
        if months > 0:
            gaps.append(_gap_label(months, _shift_month(latest_end, 1), today))
    return gaps


def _experience_label(work):
    months = sum(_month_count(start, end) for start, end in _employment_spans(work))
    return _duration_label(months)


def _current_role(work):
    if not work:
        return None
    job = next((item for item in work if re.search(r"present|current", item.get("date") or "", re.I)), work[0])
    role = " · ".join(part for part in (job.get("title"), job.get("company")) if part)
    return role or None


def _short_school(name):
    if not name:
        return None
    paren = re.search(r"\(([^)]+)\)\s*,?\s*(.*)$", name)
    if paren and len(paren.group(1)) <= 16:
        short = paren.group(1).strip()
        tail = paren.group(2).strip(" ,")
        if tail and len(tail) <= 24:
            return f"{short} {tail}"
        return short
    return name


def _study_lines(education):
    lines = []
    for school in education[:2]:
        place = _short_school(school.get("school"))
        line = " · ".join(part for part in (school.get("degree"), place) if part)
        if line:
            lines.append(line)
    return lines


def _split_csv(text):
    parts = []
    buf = []
    depth = 0
    for char in text:
        if char == "(":
            depth += 1
        elif char == ")" and depth:
            depth -= 1
        if char == "," and depth == 0:
            piece = "".join(buf).strip()
            if piece:
                parts.append(piece)
            buf = []
        else:
            buf.append(char)
    tail = "".join(buf).strip(" ,")
    if tail:
        parts.append(tail)
    return parts


def _flatten_skill(token):
    match = re.fullmatch(r"(.+?)\s*\((.+)\)", token.strip())
    if not match:
        return [token.strip()]
    return [match.group(1).strip(), *_split_csv(match.group(2))]


def _tech(skills, per_line=6, limit=18):
    chips = []
    seen = set()
    for raw in skills:
        text = raw.strip().rstrip(",")
        if ":" in text:
            text = text.split(":", 1)[1].strip()
        if not text or ("," not in raw and _SKILL_HEADER.match(text)):
            continue
        kept = 0
        pieces = _split_csv(text) if "," in text else [text]
        for piece in pieces:
            for chip in _flatten_skill(piece):
                chip = chip.strip(" .")
                key = chip.lower()
                if not chip or key in _GLANCE_SKIP or key in seen:
                    continue
                seen.add(key)
                chips.append(chip)
                kept += 1
                if kept == per_line or len(chips) == limit:
                    break
            if kept == per_line or len(chips) == limit:
                break
        if len(chips) == limit:
            break
    return chips


def _brief(result):
    return {
        "experience": _experience_label(result["work"]),
        "gaps": _work_gaps(result["work"]),
        "role": _current_role(result["work"]),
        "education": _study_lines(result["education"]),
        "tech": _tech(result["skills"]),
    }


def parse(path):
    lowered = path.lower()
    if lowered.endswith(".pdf"):
        lines = _read_pdf(path)
    elif lowered.endswith(".docx"):
        lines = _read_docx(path)
    else:
        raise SystemExit("usage: resume-parser <resume.pdf|resume.docx>")
    sections = _sections(lines)
    result = {
        "profile": {"name": None, "email": None, "phone": None, "url": [], "location": None},
        "education": [],
        "work": [],
        "projects": [],
        "skills": [],
        "talks": [],
        "accomplishments": [],
    }
    for key, _title, bucket in sections:
        if key == "profile":
            result["profile"] = _profile(bucket)
        elif key == "summary":
            summary = " ".join(line.text() for line in bucket).strip()
            if summary:
                result["profile"]["summary"] = summary
        elif key == "education":
            result["education"].extend(_education(bucket))
        elif key == "work":
            result["work"].extend(_work(bucket))
        elif key == "projects":
            result["projects"].extend(_projects(bucket))
        elif key == "skills":
            result["skills"].extend(_plain_lines(bucket))
        elif key == "talks":
            result["talks"].extend(_plain_lines(bucket))
        elif key == "accomplishments":
            result["accomplishments"].extend(_plain_lines(bucket, bullets_only=True))
    result["brief"] = _brief(result)
    return result


def main() -> None:
    if len(sys.argv) != 2:
        print("usage: resume-parser <resume.pdf|resume.docx>", file=sys.stderr)
        raise SystemExit(1)
    print(json.dumps(parse(sys.argv[1]), indent=2))
