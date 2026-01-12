from __future__ import annotations

from io import BytesIO
from typing import Optional

import pandas as pd

COLUMNS = [
    "ECS_CODE_FULL",
    "ECS_BASE",
    "ITEM",
    "WORK_DATE",
    "SUPERVISOR",
    "SUPERINTENDENT",
    "SOURCE_FILE",
    "INGESTED_AT",
]

DEDUP_KEYS = ["ECS_CODE_FULL", "WORK_DATE"]


def read_master(master_bytes: Optional[bytes]) -> pd.DataFrame:
    if not master_bytes:
        return pd.DataFrame(columns=COLUMNS)

    bio = BytesIO(master_bytes)
    df = pd.read_excel(bio, sheet_name=0, engine="openpyxl")

    # garante colunas mínimas
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = pd.NA

    # ordena colunas
    df = df[COLUMNS]
    return df


def append_and_dedup(master_df: pd.DataFrame, new_df: pd.DataFrame) -> pd.DataFrame:
    combined = pd.concat([master_df, new_df], ignore_index=True)

    # normaliza datas como string ISO ou date; aqui tenta manter como datetime
    combined["WORK_DATE"] = pd.to_datetime(combined["WORK_DATE"], errors="coerce").dt.date

    # dedup: mantém primeira ocorrência
    combined = combined.drop_duplicates(subset=DEDUP_KEYS, keep="first")

    # ordena por data e ECS
    combined = combined.sort_values(by=["WORK_DATE", "ECS_CODE_FULL"], kind="stable").reset_index(drop=True)
    return combined


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "MASTER") -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return out.getvalue()
