# LB6 Shift Report ETL (DOCX → Master XLSX)

App Streamlit para:
- Upload de Shift Reports (.docx)
- Extração de ECS base + itens
- Captura de data + Supervisor + Superintendent
- Append em Master.xlsx (com dedup)

## Rodar local
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# mac/linux:
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
