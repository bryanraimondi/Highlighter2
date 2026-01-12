import re
from dataclasses import dataclass
from datetime import date
from typing import List, Optional, Tuple

from dateutil.parser import parse as dt_parse

MONTHS = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

MONTH_RE = r"(January|February|March|April|May|June|July|August|September|October|November|December)"


@dataclass
class DocMetadata:
    work_date: date
    supervisor: str
    superintendent: str


def _clean_spaces(s: str) -> str:
    return re.sub(r"[ \t]+", " ", s).strip()


def normalize_ecs_base(raw: str) -> Optional[str]:
    """
    Aceita variações tipo:
      '1 HNX 10 ST'  -> '1HNX10ST'
      '1HDD 0B ST'   -> '1HDD0BST'
      '1 HPB ON ST'  -> '1HPBONST' (se realmente vier 'ON')
    Regra:
      <1 dígito><2-3 letras><2 alfanum><2 letras>
    """
    if not raw:
        return None
    s = _clean_spaces(raw.upper())

    # tenta capturar com espaços opcionais
    m = re.search(r"\b(\d)\s*([A-Z]{2,3})\s*([0-9A-Z]{2})\s*([A-Z]{2})\b", s)
    if not m:
        return None
    return f"{m.group(1)}{m.group(2)}{m.group(3)}{m.group(4)}"


def extract_text_from_docx(doc) -> str:
    # junta parágrafos preservando quebra de linha
    paras = [p.text for p in doc.paragraphs if p.text and p.text.strip()]
    return "\n".join(paras)


def extract_metadata(full_text: str, assumed_year: int) -> DocMetadata:
    """
    - Data: procura algo como '7th  January' ou '7 January' (com/sem ano).
    - Supervisor / Superintendent: lê do bloco 'Signed (Supervisor)' etc.
    """
    txt = full_text

    # Supervisor / Superintendent
    sup = ""
    supint = ""

    m1 = re.search(r"Signed\s*\(Supervisor\)\s*([A-Za-z][A-Za-z '\-]+)", txt, flags=re.IGNORECASE)
    if m1:
        sup = _clean_spaces(m1.group(1))

    m2 = re.search(r"Signed\s*\(Superintendent\)\s*([A-Za-z][A-Za-z '\-]+)", txt, flags=re.IGNORECASE)
    if m2:
        supint = _clean_spaces(m2.group(1))

    # Data (dia + mês + opcional ano)
    # exemplos: "7th  January" / "7 January" / "7th January 2025"
    md = re.search(rf"\b(\d{{1,2}})(?:st|nd|rd|th)?\s+{MONTH_RE}\b(?:\s+(\d{{4}}))?", txt, flags=re.IGNORECASE)
    if not md:
        # fallback: tenta parser geral procurando "Date" perto
        # (melhor esforço)
        fallback = None
        mdate = re.search(r"Date\s*\n?\s*([^\n]+)", txt, flags=re.IGNORECASE)
        if mdate:
            fallback = mdate.group(1).strip()
        if fallback:
            dt = dt_parse(fallback, dayfirst=True, fuzzy=True, default=dt_parse(f"01/01/{assumed_year}", dayfirst=True))
            return DocMetadata(work_date=dt.date(), supervisor=sup, superintendent=supint)

        # último recurso: 01/01/assumed_year
        return DocMetadata(work_date=date(assumed_year, 1, 1), supervisor=sup, superintendent=supint)

    day = int(md.group(1))
    month_name = md.group(2).capitalize()
    year = int(md.group(3)) if md.group(3) else assumed_year
    month_num = MONTHS.index(month_name) + 1

    return DocMetadata(
        work_date=date(year, month_num, day),
        supervisor=sup,
        superintendent=supint,
    )


def _clip_relevant_zone(full_text: str) -> str:
    """
    Tenta limitar a zona do documento para reduzir falsos positivos.
    Pega entre 'Today' e 'Manpower' se existir.
    """
    t = full_text
    start = 0
    end = len(t)

    m_start = re.search(r"Today[’']?s\s+Tasks", t, flags=re.IGNORECASE)
    if m_start:
        start = m_start.start()

    m_end = re.search(r"\bManpower\b", t, flags=re.IGNORECASE)
    if m_end:
        end = m_end.start()

    return t[start:end]


def extract_ecs_rows(full_text: str) -> List[Tuple[str, str]]:
    """
    Retorna lista de tuplas:
      (ecs_base, item_str)  -> ex ('1HNX10ST', '2292'), ('1HPB0NST', '0031.1')
    Estratégia:
      - acha todas ocorrências de ECS base no texto
      - para cada ECS base, pega o trecho até o próximo ECS base
      - dentro do trecho, captura itens no formato 4 dígitos com opcional '.d'
    """
    zone = _clip_relevant_zone(full_text)

    ecs_pat = re.compile(r"\b(\d)\s*([A-Z]{2,3})\s*([0-9A-Z]{2})\s*([A-Z]{2})\b")
    item_pat = re.compile(r"\b\d{4}(?:\.\d)?\b")

    matches = list(ecs_pat.finditer(zone))
    if not matches:
        return []

    rows: List[Tuple[str, str]] = []

    for i, m in enumerate(matches):
        ecs_base = f"{m.group(1)}{m.group(2)}{m.group(3)}{m.group(4)}"
        chunk_start = m.end()
        chunk_end = matches[i + 1].start() if i + 1 < len(matches) else len(zone)
        chunk = zone[chunk_start:chunk_end]

        items = item_pat.findall(chunk)
        # remove duplicados mantendo ordem
        seen = set()
        uniq = []
        for it in items:
            if it not in seen:
                seen.add(it)
                uniq.append(it)

        for it in uniq:
            rows.append((ecs_base, it))

    return rows
