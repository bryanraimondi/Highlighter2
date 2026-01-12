import re
from dataclasses import dataclass
from datetime import date
from typing import List, Optional, Tuple

from dateutil.parser import parse as dt_parse


MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]
MONTH_RE = r"(January|February|March|April|May|June|July|August|September|October|November|December)"


@dataclass
class DocMetadata:
    work_date: date
    supervisor: str
    superintendent: str


def _clean_spaces(s: str) -> str:
    return re.sub(r"[ \t]+", " ", s).strip()


# ✅ UPDATED: allow optional hyphen between letters and the 2-char block (e.g., HK-10)
_ECS_BASE_RE = re.compile(r"\b(\d)\s*([A-Z]{2,3})\s*[- ]?\s*([0-9A-Z]{2})\s*([A-Z]{2})\b", re.IGNORECASE)

# items like 2292 or 0031.1
_ITEM_RE = re.compile(r"\b\d{4}(?:\.\d)?\b")


def normalize_ecs_base(raw: str) -> Optional[str]:
    """
    Accepts:
      '1 HNX 10 ST' -> '1HNX10ST'
      '1HK-10SE'    -> '1HK10SE'
      '1 HDD 0B ST' -> '1HDD0BST'
    """
    if not raw:
        return None
    s = _clean_spaces(raw.upper())

    m = _ECS_BASE_RE.search(s)
    if not m:
        return None
    return f"{m.group(1)}{m.group(2).upper()}{m.group(3).upper()}{m.group(4).upper()}"


def extract_text_from_docx(doc) -> str:
    """
    ✅ UPDATED: Extracts text from:
      - normal paragraphs
      - table cells (CRITICAL for Shift Report templates)
    """
    texts: List[str] = []

    # 1) Body paragraphs
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t:
            texts.append(t)

    # 2) Tables (most Shift Reports store content here)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    t = (p.text or "").strip()
                    if t:
                        texts.append(t)

    return "\n".join(texts)


def extract_metadata(full_text: str, assumed_year: int) -> DocMetadata:
    """
    - Date: finds patterns like '7th January' or '7 January' (with/without year)
    - Names: reads 'Signed (Supervisor)' and 'Signed (Superintendent)' if present
    """
    txt = full_text

    supervisor = ""
    superintendent = ""

    m1 = re.search(r"Signed\s*\(Supervisor\)\s*([A-Za-z][A-Za-z '\-]+)", txt, flags=re.IGNORECASE)
    if m1:
        supervisor = _clean_spaces(m1.group(1))

    m2 = re.search(r"Signed\s*\(Superintendent\)\s*([A-Za-z][A-Za-z '\-]+)", txt, flags=re.IGNORECASE)
    if m2:
        superintendent = _clean_spaces(m2.group(1))

    # Date like: "7th  January" or "7 January 2025"
    md = re.search(rf"\b(\d{{1,2}})(?:st|nd|rd|th)?\s+{MONTH_RE}\b(?:\s+(\d{{4}}))?",
                   txt, flags=re.IGNORECASE)
    if not md:
        # fallback: try parse a line after "Date"
        fallback = None
        mdate = re.search(r"Date\s*\n?\s*([^\n]+)", txt, flags=re.IGNORECASE)
        if mdate:
            fallback = mdate.group(1).strip()

        if fallback:
            dt = dt_parse(
                fallback,
                dayfirst=True,
                fuzzy=True,
                default=dt_parse(f"01/01/{assumed_year}", dayfirst=True)
            )
            return DocMetadata(work_date=dt.date(), supervisor=supervisor, superintendent=superintendent)

        # last resort
        return DocMetadata(work_date=date(assumed_year, 1, 1), supervisor=supervisor, superintendent=superintendent)

    day = int(md.group(1))
    month_name = md.group(2).capitalize()
    year = int(md.group(3)) if md.group(3) else assumed_year
    month_num = MONTHS.index(month_name) + 1

    return DocMetadata(
        work_date=date(year, month_num, day),
        supervisor=supervisor,
        superintendent=superintendent,
    )


def _clip_relevant_zone(full_text: str) -> str:
    """
    Tries to limit search zone to reduce false positives.
    Uses 'Today’s Tasks' -> 'Manpower' if available.
    """
    t = full_text
    start = 0
    end = len(t)

    # handle Today’s / Today's
    m_start = re.search(r"Today[’']?s\s+Tasks", t, flags=re.IGNORECASE)
    if m_start:
        start = m_start.start()

    m_end = re.search(r"\bManpower\b", t, flags=re.IGNORECASE)
    if m_end:
        end = m_end.start()

    return t[start:end]


def extract_ecs_rows(full_text: str) -> List[Tuple[str, str]]:
    """
    Returns list of tuples:
      (ecs_base, item_str) e.g. ('1HNX10ST', '2292'), ('1HPB0NST', '0031.1')

    Strategy:
      - find all ECS base occurrences
      - for each ECS base, scan until next ECS base
      - collect items in that chunk (dedup per ECS chunk, preserve order)
    """
    zone = _clip_relevant_zone(full_text)

    matches = list(_ECS_BASE_RE.finditer(zone))
    if not matches:
        return []

    rows: List[Tuple[str, str]] = []

    for i, m in enumerate(matches):
        ecs_base = f"{m.group(1)}{m.group(2).upper()}{m.group(3).upper()}{m.group(4).upper()}"
        chunk_start = m.end()
        chunk_end = matches[i + 1].start() if i + 1 < len(matches) else len(zone)
        chunk = zone[chunk_start:chunk_end]

        items = _ITEM_RE.findall(chunk)

        # dedup items for this ECS base, keep order
        seen = set()
        for it in items:
            if it not in seen:
                seen.add(it)
                rows.append((ecs_base, it))

    return rows
