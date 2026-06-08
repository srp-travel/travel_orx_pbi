# travel_orx_pbi

Outil Streamlit pour réconcilier un export Orchestra (`detailedSearchExport*.xlsx`) avec un export PowerBI.

## Structure

- `main.py` : lanceur d'application
- `src/app.py` : application Streamlit principale
- `requirements.txt` : dépendances Python

## Installation

```powershell
C:/Users/s.cherkaoui/AppData/Local/Microsoft/WindowsApps/python3.12.exe -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Exécution

```powershell
streamlit run main.py
```

## Dépendances

- streamlit
- pandas
- openpyxl
- xlsxwriter
