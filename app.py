from __future__ import annotations

from datetime import datetime
from io import BytesIO
from typing import List

import pandas as pd
import streamlit as st
from docx import Document

from parser import extract_text_from_docx, extract_metadata, extract_ecs_rows
from excel_io import read_master, append_and_dedup, to_excel_bytes, COLUMNS


st.set_page_config(page_title="LB6 Shift Report → Master Excel", layout="wide")

st.title("LB6 Shift Reports (.docx) → Master Excel (append + dedup)")

st.markdown(
    """
**O que faz:**
- Upload de 1+ shift reports em Word (.docx)
- Extrai **ECS base + itens**, data, Supervisor e Superintendent
- Atualiza um **Master.xlsx** (opcional) de forma **crescente** e com **deduplicação**
"""
)

colA, colB = st.columns([1, 1])

with colA:
    docx_files = st.file_uploader(
        "Upload dos Shift Reports (.docx) — pode ser múltiplo",
        type=["docx"],
        accept_multiple_files=True,
    )

with colB:
    master_file = st.file_uploader(
        "Upload do Master.xlsx (opcional — se não enviar, ele cria um novo)",
        type=["xlsx"],
        accept_multiple_files=False,
    )

assumed_year = st.number_input(
    "Ano assumido (se o Word não tiver ano explícito)",
    min_value=2020,
    max_value=2035,
    value=datetime.utcnow().year,
    step=1,
)

process = st.button("Processar", type="primary", disabled=(not docx_files))

if process:
    if not docx_files:
        st.error("Envie pelo menos 1 arquivo .docx.")
        st.stop()

    master_bytes = master_file.getvalue() if master_file else None
    master_df = read_master(master_bytes)

    added_all: List[pd.DataFrame] = []
    errors = []

    for f in docx_files:
        try:
            doc = Document(BytesIO(f.getvalue()))
            text = extract_text_from_docx(doc)
            meta = extract_metadata(text, assumed_year=assumed_year)
            ecs_rows = extract_ecs_rows(text)

            if not ecs_rows:
                errors.append(f"⚠️ {f.name}: não encontrei ECS + itens nessa leitura.")
                continue

            now = datetime.utcnow().isoformat(timespec="seconds") + "Z"

            rows = []
            for ecs_base, item in ecs_rows:
                ecs_full = f"{ecs_base}{item}"
                rows.append(
                    {
                        "ECS_CODE_FULL": ecs_full,
                        "ECS_BASE": ecs_base,
                        "ITEM": item,
                        "WORK_DATE": meta.work_date,
                        "SUPERVISOR": meta.supervisor,
                        "SUPERINTENDENT": meta.superintendent,
                        "SOURCE_FILE": f.name,
                        "INGESTED_AT": now,
                    }
                )

            df_new = pd.DataFrame(rows, columns=COLUMNS)
            added_all.append(df_new)

        except Exception as e:
            errors.append(f"❌ {f.name}: erro ao processar — {e}")

    if not added_all:
        st.error("Não consegui extrair nenhuma linha dos documentos enviados.")
        for e in errors:
            st.write(e)
        st.stop()

    new_df = pd.concat(added_all, ignore_index=True)
    updated = append_and_dedup(master_df, new_df)

    # métricas
    before = len(master_df)
    after = len(updated)
    delta = after - before

    st.success(f"Master atualizado. Linhas antes: {before} | depois: {after} | adicionadas (líquido): {delta}")

    with st.expander("Prévia — últimas 50 linhas"):
        st.dataframe(updated.tail(50), use_container_width=True)

    if errors:
        st.warning("Avisos / erros em alguns arquivos:")
        for e in errors:
            st.write(e)

    out_bytes = to_excel_bytes(updated, sheet_name="MASTER")
    st.download_button(
        "⬇️ Baixar Master_updated.xlsx",
        data=out_bytes,
        file_name="Master_updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # opcional: baixar só o extraído desta execução (antes da dedup)
    out_extracted = to_excel_bytes(new_df, sheet_name="EXTRACTED")
    st.download_button(
        "⬇️ Baixar extracted_rows.xlsx (apenas desta execução)",
        data=out_extracted,
        file_name="extracted_rows.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
