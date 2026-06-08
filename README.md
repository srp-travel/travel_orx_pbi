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

## Déploiement sur Streamlit Cloud

1. Pousse le dépôt sur GitHub si ce n'est pas déjà fait :

```powershell
git add .
git commit -m "Prépare le projet pour Streamlit Cloud"
git push origin main
```

2. Va sur https://share.streamlit.io
3. Connecte-toi avec ton compte GitHub
4. Clique sur "New app" et choisis :
   - GitHub repo : `srp-travel/travel_orx_pbi`
   - Branch : `main`
   - File path : `main.py`

5. Lance le déploiement. Streamlit Cloud utilisera automatiquement `requirements.txt`.

> `main.py` est déjà configuré comme point d'entrée et le projet lit les fichiers uploadés uniquement en mémoire.
