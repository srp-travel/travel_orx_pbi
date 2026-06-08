"""
Utilitaire Streamlit — Réconciliation Orchestra (ORX) ↔ PowerBI (PBI)
=====================================================================

Reporting simplifié autour de 4 questions clés :

  1. Le nombre de dossiers est-il identique entre ORX et PBI ?
  2. Le chiffre d'affaires est-il identique ?
  3. Quels dossiers ORX sont absents de PBI ?
  4. Quels dossiers PBI sont mal intégrés ?
     (TO manquant, Sale id manquant, prix incohérent vs ORX)

Lancement :
    pip install streamlit pandas openpyxl xlsxwriter
    streamlit run app.py
"""

from __future__ import annotations

import io
from datetime import date
from typing import Any

import pandas as pd
import streamlit as st


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

DEFAULT_ORX_COLS: dict[str, str] = {
    "key":      "Nº interne",
    "campaign": "Code campagne",
    "status":   "Statut dossier TO",
    "amount":   "TTC Prix",
    "date":     "Date de réservation",
}

DEFAULT_PBI_COLS: dict[str, str] = {
    "key":      "Order id (supplier)",
    "campaign": "Sale id",
    "status":   "Booking state",
    "amount":   "€ turnover (inc. VAT)",
    "date":     "Date",
}

ORX_ENRICH_COLS: dict[str, str] = {
    "channel":       "Canal de réservation",
    "product_name":  "Nom du produit",
    "product_code":  "Code interne produit",
    "tour_operator": "Tour Opérateur",
}

PBI_ENRICH_COLS: dict[str, str] = {
    "tour_operator": "Tour Operator",
    "product":       "Product",
    "product_id":    "Product id",
}

ORX_DEFAULT_HEADER_ROW = 1
PBI_DEFAULT_HEADER_ROW = 0

CANCELLED_TOKENS: set[str] = {"annulee", "annulée", "cancelled", "canceled", "cancel"}
PRICE_TOLERANCE_DEFAULT = 0.01


# ---------------------------------------------------------------------------
# Thème UI — palette grise sobre + accent #345A7A
# ---------------------------------------------------------------------------

CUSTOM_CSS = """
<style>
  :root {
    --grey-50:  #fafafa;
    --grey-100: #f2f2f3;
    --grey-200: #e5e5e7;
    --grey-300: #d1d1d4;
    --grey-500: #8a8a8f;
    --grey-700: #4a4a52;
    --grey-900: #1f1f23;
    --accent:        #345A7A;
    --accent-dark:   #264560;
    --accent-light:  #e8eef4;
  }
  .stApp { background-color: var(--grey-50); color: var(--grey-900); }
  h1, h2, h3, h4 { color: var(--grey-900); font-weight: 600; letter-spacing: -0.01em; }
  h1 { border-bottom: 2px solid var(--accent); padding-bottom: 8px; display: inline-block; }
  a { color: var(--accent); }

  /* Metrics */
  [data-testid="stMetric"] {
    background: #ffffff; border: 1px solid var(--grey-200);
    border-radius: 8px; padding: 14px 18px;
    box-shadow: 0 1px 2px rgba(0,0,0,0.03);
  }
  [data-testid="stMetricLabel"] {
    color: var(--grey-700) !important;
    font-size: 0.82rem; text-transform: uppercase; letter-spacing: 0.04em;
  }
  [data-testid="stMetricValue"] { color: var(--accent) !important; font-weight: 600; }
  [data-testid="stMetricDelta"] svg { display: none; }

  /* Boutons */
  .stButton > button {
    background: var(--accent); color: #fff; border: none;
    border-radius: 6px; padding: 0.5rem 1.1rem; font-weight: 500;
  }
  .stButton > button:hover { background: var(--accent-dark); }
  .stDownloadButton > button {
    background: #ffffff; color: var(--accent);
    border: 1px solid var(--accent); border-radius: 6px; font-weight: 500;
  }
  .stDownloadButton > button:hover { background: var(--accent-light); }

  /* Onglets */
  .stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 1px solid var(--grey-200); }
  .stTabs [data-baseweb="tab"] {
    background: transparent; color: var(--grey-700);
    padding: 8px 14px; border-radius: 6px 6px 0 0;
  }
  .stTabs [aria-selected="true"] {
    background: #ffffff; color: var(--accent);
    border: 1px solid var(--grey-200); border-top: 2px solid var(--accent);
    border-bottom: 1px solid #ffffff; font-weight: 600;
  }

  /* DataFrame */
  [data-testid="stDataFrame"] {
    border: 1px solid var(--grey-200); border-radius: 6px; overflow: hidden;
  }

  /* Sidebar */
  section[data-testid="stSidebar"] {
    background: var(--grey-100); border-right: 1px solid var(--grey-200);
  }
  section[data-testid="stSidebar"] h2,
  section[data-testid="stSidebar"] h3 { color: var(--accent); }

  /* Inputs focus */
  .stSelectbox > div > div:focus-within,
  .stNumberInput > div > div:focus-within {
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 1px var(--accent) !important;
  }

  /* Alertes */
  [data-testid="stAlert"][data-baseweb="notification"] {
    border-radius: 6px;
  }

  /* Guide box */
  .guide-box {
    background: #ffffff; border: 1px solid var(--grey-200);
    border-left: 4px solid var(--accent);
    border-radius: 6px; padding: 18px 22px; margin: 8px 0 20px 0;
  }
  .guide-box h4 { margin-top: 0; color: var(--accent); }
  .guide-box ol, .guide-box ul { margin: 6px 0 6px 18px; color: var(--grey-700); }
  .guide-box li { margin: 4px 0; }
  .guide-box strong { color: var(--grey-900); }
  .guide-box code {
    background: var(--accent-light); color: var(--accent-dark);
    padding: 1px 6px; border-radius: 4px; font-size: 0.9em;
  }
</style>
"""


GUIDE_MARKDOWN = """
<div class="guide-box">

#### 📖 Comment ça marche ? — Guide rapide

**Objectif** : comparer un export Orchestra à un export PowerBI pour vérifier
que les dossiers et le chiffre d'affaires sont bien remontés dans PBI.

##### En 4 étapes

1. **Charger les deux fichiers Excel** dans la barre latérale gauche
   _(Orchestra `detailedSearchExport*.xlsx` et l'export PowerBI)._
2. **Vérifier le mapping des colonnes** — les noms par défaut sont
   pré-remplis ; modifie uniquement si tes en-têtes diffèrent.
3. **Contrôler la cohérence des dates** — l'outil affiche les périodes
   couvertes par chaque fichier et bloque la comparaison en cas
   d'incohérence (override possible).
4. **Cliquer sur _Lancer la comparaison_** pour obtenir le rapport.

##### Ce que tu obtiens

- **Volumétrie** : nombre de dossiers ORX vs PBI + écart.
- **Chiffre d'affaires** : CA total ORX vs PBI + écart en €.
- **Onglet « ORX absents de PBI »** : liste des dossiers présents dans
  Orchestra mais manquants dans PowerBI _(avec Canal, Nom du produit,
  Code interne produit si dispos)._
- **Onglet « PBI mal intégrés »** : dossiers PBI avec Tour Opérateur
  manquant, Sale id manquant, ou prix divergent d'Orchestra _(toutes
  les colonnes d'origine PBI sont conservées, dont `Product` et `Product id`)._
- **Onglets de groupements** : volumétrie et CA agrégés par vente,
  statut et Tour Opérateur, des deux côtés.
- **Export Excel** multi-onglets téléchargeable.

##### Points importants

- Les en-têtes du fichier **Orchestra sont sur la ligne 2** (la ligne 1
  est une bannière). C'est le réglage par défaut.
- L'ORX est **dédupliqué automatiquement** sur les 5 colonnes mappées.
- Les dates ORX sont normalisées à minuit pour fiabiliser la
  comparaison et la déduplication.
- Tolérance d'écart de prix configurable dans la sidebar (défaut : 0,01 €).

</div>
"""


# ---------------------------------------------------------------------------
# Helpers I/O
# ---------------------------------------------------------------------------

def _read_excel(file: Any, sheet_name: str, header_row: int = 0) -> pd.DataFrame:
    return pd.read_excel(file, sheet_name=sheet_name, dtype=object, header=header_row)


def _list_sheets(file: Any) -> list[str]:
    return [str(s) for s in pd.ExcelFile(file).sheet_names]


def _autodetect(df: pd.DataFrame, candidates: dict[str, str]) -> dict[str, str]:
    norm_map = {str(c).strip().lower(): str(c) for c in df.columns}
    found: dict[str, str] = {}
    for key, default_name in candidates.items():
        actual = norm_map.get(default_name.strip().lower())
        if actual is not None:
            found[key] = actual
    return found


# ---------------------------------------------------------------------------
# Helpers de normalisation
# ---------------------------------------------------------------------------

def _normalize_key(series: pd.Series) -> pd.Series:
    def conv(v: Any) -> str | None:
        if pd.isna(v):
            return None
        if isinstance(v, bool):
            return str(v)
        if isinstance(v, int):
            return str(v)
        if isinstance(v, float):
            return str(int(v)) if float(v).is_integer() else str(v)
        s = str(v).strip()
        if s == "" or s.lower() == "nan":
            return None
        try:
            f = float(s.replace(",", "."))
            return str(int(f)) if f.is_integer() else str(f)
        except ValueError:
            return s
    return series.map(conv)


def _to_float(v: Any) -> float | None:
    if pd.isna(v):
        return None
    try:
        return float(str(v).replace(" ", "").replace(",", "."))
    except (ValueError, TypeError):
        return None


def _to_date_only(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=True).dt.normalize()


def _is_blank(v: Any) -> bool:
    if pd.isna(v):
        return True
    return str(v).strip().lower() in ("", "nan", "none", "null", "n/a", "na", "-", "—")


def _fmt_date(ts: Any) -> str:
    if ts is None or pd.isna(ts):
        return "∅"
    return pd.Timestamp(ts).date().isoformat()


def _fmt_eur(v: float | int) -> str:
    return f"{v:,.2f} €".replace(",", " ").replace(".", ",")


# ---------------------------------------------------------------------------
# Cohérence des dates
# ---------------------------------------------------------------------------

def check_date_coherence(
    df_orx: pd.DataFrame, col_orx_date: str,
    df_pbi: pd.DataFrame, col_pbi_date: str,
) -> dict[str, Any]:
    orx_dates = _to_date_only(df_orx[col_orx_date])
    pbi_dates = pd.to_datetime(df_pbi[col_pbi_date], errors="coerce", dayfirst=True)
    orx_min, orx_max = orx_dates.min(), orx_dates.max()
    pbi_min, pbi_max = pbi_dates.min(), pbi_dates.max()
    info: dict[str, Any] = {
        "orx_min": orx_min, "orx_max": orx_max,
        "pbi_min": pbi_min, "pbi_max": pbi_max,
        "overlap_ok": False, "coverage_pct": 0.0,
    }
    if not all(pd.notna(x) for x in (orx_min, orx_max, pbi_min, pbi_max)):
        return info
    orx_min_ts, orx_max_ts = pd.Timestamp(orx_min), pd.Timestamp(orx_max)
    pbi_min_ts, pbi_max_ts = pd.Timestamp(pbi_min), pd.Timestamp(pbi_max)
    overlap_min = max(orx_min_ts, pbi_min_ts)
    overlap_max = min(orx_max_ts, pbi_max_ts)
    info["overlap_ok"] = overlap_min <= overlap_max
    if info["overlap_ok"]:
        union_min = min(orx_min_ts, pbi_min_ts)
        union_max = max(orx_max_ts, pbi_max_ts)
        ovl_days = (overlap_max - overlap_min).days + 1
        union_days = (union_max - union_min).days + 1
        info["coverage_pct"] = round(100 * ovl_days / union_days, 1)
    return info


# ---------------------------------------------------------------------------
# Groupings
# ---------------------------------------------------------------------------

def _group_summary(df: pd.DataFrame, by: str, key_col: str, amount_col: str) -> pd.DataFrame:
    if not by or by not in df.columns:
        return pd.DataFrame(columns=[by or "groupe", "Nb dossiers", "Chiffre d'affaires (€)"])
    grouped = (
        df.groupby(df[by].fillna("(non renseigné)").astype(str), dropna=False)
          .agg(**{
              "Nb dossiers": (key_col, "nunique"),
              "Chiffre d'affaires (€)": (amount_col,
                  lambda s: float(pd.to_numeric(s, errors="coerce").fillna(0).sum())),
          })
          .reset_index()
          .sort_values("Chiffre d'affaires (€)", ascending=False)
    )
    return grouped


# ---------------------------------------------------------------------------
# Réconciliation
# ---------------------------------------------------------------------------

ORX_DEDUP_SIGNATURE = ("_key_norm", "_campaign_norm", "_status_norm", "_amount_norm", "_date_norm")


def reconcile(
    df_orx_raw: pd.DataFrame, orx_cols: dict[str, str], orx_enrich: dict[str, str],
    df_pbi_raw: pd.DataFrame, pbi_cols: dict[str, str], pbi_enrich: dict[str, str],
    price_tolerance: float = PRICE_TOLERANCE_DEFAULT,
) -> dict[str, Any]:
    df_orx = df_orx_raw.copy()
    df_orx["_key_norm"]      = _normalize_key(df_orx[orx_cols["key"]])
    df_orx["_campaign_norm"] = _normalize_key(df_orx[orx_cols["campaign"]])
    df_orx["_status_norm"]   = df_orx[orx_cols["status"]].astype(object)
    df_orx["_amount_norm"]   = df_orx[orx_cols["amount"]].map(_to_float)
    df_orx["_date_norm"]     = _to_date_only(df_orx[orx_cols["date"]])
    df_orx = df_orx.dropna(subset=["_key_norm"])
    nb_orx_raw = len(df_orx)
    df_orx_dedup = df_orx.drop_duplicates(list(ORX_DEDUP_SIGNATURE)).reset_index(drop=True)
    nb_dup_removed = nb_orx_raw - len(df_orx_dedup)

    df_pbi = df_pbi_raw.copy()
    df_pbi["_key_norm"]      = _normalize_key(df_pbi[pbi_cols["key"]])
    df_pbi["_campaign_norm"] = _normalize_key(df_pbi[pbi_cols["campaign"]])
    df_pbi["_status_norm"]   = df_pbi[pbi_cols["status"]].astype(object)
    df_pbi["_amount_norm"]   = df_pbi[pbi_cols["amount"]].map(_to_float)
    df_pbi = df_pbi.dropna(subset=["_key_norm"])

    has_to = "tour_operator" in pbi_enrich

    kpis = {
        "nb_orx":   df_orx_dedup["_key_norm"].nunique(),
        "nb_pbi":   df_pbi["_key_norm"].nunique(),
        "ca_orx":   float(df_orx_dedup["_amount_norm"].fillna(0).sum()),
        "ca_pbi":   float(df_pbi["_amount_norm"].fillna(0).sum()),
        "nb_dup_orx_removed": nb_dup_removed,
    }
    kpis["nb_delta"] = kpis["nb_pbi"] - kpis["nb_orx"]
    kpis["ca_delta"] = kpis["ca_pbi"] - kpis["ca_orx"]

    pbi_keys = set(df_pbi["_key_norm"].dropna())
    missing_in_pbi = df_orx_dedup[~df_orx_dedup["_key_norm"].isin(pbi_keys)].copy()
    view_cols = [orx_cols["key"], orx_cols["campaign"], orx_cols["status"],
                 orx_cols["amount"], orx_cols["date"]]
    for k in ("channel", "product_name", "product_code"):
        if k in orx_enrich:
            view_cols.append(orx_enrich[k])
    view_cols = [c for c in view_cols if c in missing_in_pbi.columns]
    missing_in_pbi_view = missing_in_pbi[view_cols].copy() if view_cols else missing_in_pbi.copy()

    df_pbi["_anom_to_missing"]   = df_pbi[pbi_enrich["tour_operator"]].map(_is_blank) if has_to else False
    df_pbi["_anom_sale_missing"] = df_pbi[pbi_cols["campaign"]].map(_is_blank)
    orx_price_map = df_orx_dedup.set_index("_key_norm")["_amount_norm"]
    df_pbi["_orx_amount"]   = df_pbi["_key_norm"].map(orx_price_map)
    df_pbi["_amount_delta"] = df_pbi["_amount_norm"].fillna(0) - df_pbi["_orx_amount"].fillna(0)
    df_pbi["_anom_price_diff"] = (
        df_pbi["_key_norm"].isin(pbi_keys & set(df_orx_dedup["_key_norm"]))
        & (df_pbi["_amount_delta"].abs() > price_tolerance)
    )
    df_pbi["_has_anomaly"] = (
        df_pbi["_anom_to_missing"] | df_pbi["_anom_sale_missing"] | df_pbi["_anom_price_diff"]
    )
    badly = df_pbi[df_pbi["_has_anomaly"]].copy()
    badly["Anomalie - TO manquant"]        = badly["_anom_to_missing"]
    badly["Anomalie - Sale id manquant"]   = badly["_anom_sale_missing"]
    badly["Anomalie - Prix différent ORX"] = badly["_anom_price_diff"]
    badly["Prix ORX (référence)"]          = badly["_orx_amount"]
    badly["Écart prix (PBI - ORX)"]        = badly["_amount_delta"]
    badly = badly.drop(columns=[c for c in badly.columns if c.startswith("_")])

    orx_to_col = orx_enrich.get("tour_operator", "")
    pbi_to_col = pbi_enrich.get("tour_operator", "")
    groupings = {
        "orx_by_sale":   _group_summary(df_orx_dedup, orx_cols["campaign"], orx_cols["key"], orx_cols["amount"]),
        "orx_by_status": _group_summary(df_orx_dedup, orx_cols["status"],   orx_cols["key"], orx_cols["amount"]),
        "orx_by_to":     _group_summary(df_orx_dedup, orx_to_col,           orx_cols["key"], orx_cols["amount"]) if orx_to_col else pd.DataFrame(),
        "pbi_by_sale":   _group_summary(df_pbi,        pbi_cols["campaign"], pbi_cols["key"], pbi_cols["amount"]),
        "pbi_by_status": _group_summary(df_pbi,        pbi_cols["status"],   pbi_cols["key"], pbi_cols["amount"]),
        "pbi_by_to":     _group_summary(df_pbi,        pbi_to_col,           pbi_cols["key"], pbi_cols["amount"]) if pbi_to_col else pd.DataFrame(),
    }
    return {
        "kpis": kpis,
        "missing_in_pbi": missing_in_pbi_view,
        "badly_integrated": badly,
        **groupings,
    }


# ---------------------------------------------------------------------------
# Rapport Excel
# ---------------------------------------------------------------------------

def build_report_xlsx(results: dict[str, Any], date_info: dict[str, Any]) -> bytes:
    buf = io.BytesIO()
    k = results["kpis"]
    with pd.ExcelWriter(buf, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        pd.DataFrame([
            ["1. Nb dossiers ORX",           k["nb_orx"]],
            ["1. Nb dossiers PBI",           k["nb_pbi"]],
            ["1. Écart (PBI - ORX)",         k["nb_delta"]],
            ["2. CA ORX (€)",                round(k["ca_orx"], 2)],
            ["2. CA PBI (€)",                round(k["ca_pbi"], 2)],
            ["2. Écart CA (PBI - ORX) (€)",  round(k["ca_delta"], 2)],
            ["Doublons ORX supprimés",       k["nb_dup_orx_removed"]],
            ["Dates ORX",                    f"{date_info['orx_min']} → {date_info['orx_max']}"],
            ["Dates PBI",                    f"{date_info['pbi_min']} → {date_info['pbi_max']}"],
            ["Recouvrement temporel (%)",    date_info["coverage_pct"]],
        ], columns=["Indicateur", "Valeur"]).to_excel(writer, sheet_name="1_Synthese", index=False)

        results["missing_in_pbi"].to_excel(writer,   sheet_name="2_ORX_absents_de_PBI", index=False)
        results["badly_integrated"].to_excel(writer, sheet_name="3_PBI_mal_integres",   index=False)
        results["orx_by_sale"].to_excel(writer,      sheet_name="ORX_par_vente",   index=False)
        results["orx_by_status"].to_excel(writer,    sheet_name="ORX_par_statut",  index=False)
        if not results["orx_by_to"].empty:
            results["orx_by_to"].to_excel(writer, sheet_name="ORX_par_TO", index=False)
        results["pbi_by_sale"].to_excel(writer,      sheet_name="PBI_par_vente",   index=False)
        results["pbi_by_status"].to_excel(writer,    sheet_name="PBI_par_statut",  index=False)
        if not results["pbi_by_to"].empty:
            results["pbi_by_to"].to_excel(writer, sheet_name="PBI_par_TO", index=False)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

def _pick(label: str, options: list[str], default: str) -> str:
    idx = options.index(default) if default in options else 0
    return st.selectbox(label, options, index=idx)


def run() -> None:
    st.set_page_config(page_title="Réconciliation ORX vs PBI", layout="wide")
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    st.title("Réconciliation Orchestra ↔ PowerBI")
    st.caption("Quatre contrôles clés : volumétrie, chiffre d'affaires, "
               "dossiers absents de PBI, dossiers mal intégrés.")

    with st.expander("📖 Comment ça marche ? — Guide rapide", expanded=True):
        st.markdown(GUIDE_MARKDOWN, unsafe_allow_html=True)

    with st.sidebar:
        st.header("Fichiers")
        orx_file = st.file_uploader("Fichier Orchestra (ORX)", type=["xlsx", "xls"], key="orx")
        pbi_file = st.file_uploader("Fichier PowerBI (PBI)",   type=["xlsx", "xls"], key="pbi")

        st.header("Paramètres")
        price_tol = st.number_input("Tolérance écart de prix (€)",
                                    min_value=0.0, value=PRICE_TOLERANCE_DEFAULT,
                                    step=0.01, format="%.2f")
        st.markdown("**Ligne d'en-têtes**")
        orx_header_excel = st.number_input("ORX (numéro Excel)", min_value=1,
                                           value=ORX_DEFAULT_HEADER_ROW + 1, step=1,
                                           help="Par défaut : ligne 2.")
        pbi_header_excel = st.number_input("PBI (numéro Excel)", min_value=1,
                                           value=PBI_DEFAULT_HEADER_ROW + 1, step=1)

    if not orx_file or not pbi_file:
        st.info("Charge les deux fichiers dans la barre latérale pour démarrer.")
        st.stop()

    col1, col2 = st.columns(2)
    with col1:
        orx_sheet = st.selectbox("Feuille ORX", _list_sheets(orx_file), index=0)
    with col2:
        pbi_sheet = st.selectbox("Feuille PBI", _list_sheets(pbi_file), index=0)

    df_orx = _read_excel(orx_file, orx_sheet, header_row=int(orx_header_excel) - 1)
    df_pbi = _read_excel(pbi_file, pbi_sheet, header_row=int(pbi_header_excel) - 1)

    st.caption(f"ORX : {len(df_orx):,} lignes × {len(df_orx.columns)} colonnes  •  "
               f"PBI : {len(df_pbi):,} lignes × {len(df_pbi.columns)} colonnes")

    st.subheader("Mapping des colonnes de comparaison")
    st.caption("Seuls les champs strictement nécessaires à la comparaison sont affichés. "
               "Les colonnes d'enrichissement (Canal, Produit, Tour Opérateur…) sont "
               "détectées automatiquement si elles sont présentes dans tes fichiers.")

    orx_opts = [str(c) for c in df_orx.columns]
    pbi_opts = [str(c) for c in df_pbi.columns]

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Orchestra (ORX)**")
        orx_cols = {
            "key":      _pick("Nº interne",          orx_opts, DEFAULT_ORX_COLS["key"]),
            "campaign": _pick("Code campagne",       orx_opts, DEFAULT_ORX_COLS["campaign"]),
            "status":   _pick("Statut dossier TO",   orx_opts, DEFAULT_ORX_COLS["status"]),
            "amount":   _pick("TTC Prix",            orx_opts, DEFAULT_ORX_COLS["amount"]),
            "date":     _pick("Date de réservation", orx_opts, DEFAULT_ORX_COLS["date"]),
        }
    with c2:
        st.markdown("**PowerBI (PBI)**")
        pbi_cols = {
            "key":      _pick("Order id (supplier)",   pbi_opts, DEFAULT_PBI_COLS["key"]),
            "campaign": _pick("Sale id",               pbi_opts, DEFAULT_PBI_COLS["campaign"]),
            "status":   _pick("Booking state",         pbi_opts, DEFAULT_PBI_COLS["status"]),
            "amount":   _pick("€ turnover (inc. VAT)", pbi_opts, DEFAULT_PBI_COLS["amount"]),
            "date":     _pick("Date",                  pbi_opts, DEFAULT_PBI_COLS["date"]),
        }

    orx_enrich = _autodetect(df_orx, ORX_ENRICH_COLS)
    pbi_enrich = _autodetect(df_pbi, PBI_ENRICH_COLS)
    auto_detected = []
    if orx_enrich:
        auto_detected.append("ORX → " + ", ".join(orx_enrich.values()))
    if pbi_enrich:
        auto_detected.append("PBI → " + ", ".join(pbi_enrich.values()))
    if auto_detected:
        st.caption("🔎 Colonnes d'enrichissement détectées : " + " ; ".join(auto_detected))

    st.subheader("Cohérence des périodes de réservation")
    date_info = check_date_coherence(df_orx, orx_cols["date"], df_pbi, pbi_cols["date"])
    dc1, dc2, dc3 = st.columns(3)
    dc1.metric("Fenêtre ORX",
               f"{_fmt_date(date_info['orx_min'])} → {_fmt_date(date_info['orx_max'])}")
    dc2.metric("Fenêtre PBI",
               f"{_fmt_date(date_info['pbi_min'])} → {_fmt_date(date_info['pbi_max'])}")
    dc3.metric("Recouvrement", f"{date_info['coverage_pct']} %",
               delta="OK" if date_info["overlap_ok"] else "KO", delta_color="off")

    if not date_info["overlap_ok"]:
        st.error("Périodes incohérentes — risque de faux positifs.")
        bypass = st.checkbox("Lancer quand même")
    elif date_info["coverage_pct"] < 80:
        st.warning(f"Recouvrement partiel ({date_info['coverage_pct']} %).")
        bypass = st.checkbox("Continuer", value=True)
    else:
        bypass = True
    if not bypass:
        st.stop()

    st.subheader("Lancer la réconciliation")
    if not st.button("Lancer la comparaison", type="primary"):
        return

    with st.spinner("Comparaison en cours…"):
        results = reconcile(df_orx, orx_cols, orx_enrich,
                            df_pbi, pbi_cols, pbi_enrich,
                            price_tolerance=price_tol)

    k = results["kpis"]

    st.subheader("1. Volumétrie")
    a1, a2, a3 = st.columns(3)
    a1.metric("Dossiers ORX", f"{k['nb_orx']:,}".replace(",", " "))
    a2.metric("Dossiers PBI", f"{k['nb_pbi']:,}".replace(",", " "))
    a3.metric("Écart (PBI − ORX)", f"{k['nb_delta']:+,}".replace(",", " "),
              delta="Identique" if k["nb_delta"] == 0 else "Écart détecté",
              delta_color="off" if k["nb_delta"] == 0 else "inverse")
    if k["nb_delta"] == 0:
        st.success("Le nombre de dossiers est identique entre ORX et PBI.")
    else:
        st.warning(f"Écart de {abs(k['nb_delta']):,} dossier(s).".replace(",", " "))

    st.subheader("2. Chiffre d'affaires")
    b1, b2, b3 = st.columns(3)
    b1.metric("CA ORX", _fmt_eur(k["ca_orx"]))
    b2.metric("CA PBI", _fmt_eur(k["ca_pbi"]))
    b3.metric("Écart CA", _fmt_eur(k["ca_delta"]),
              delta="Identique" if abs(k["ca_delta"]) < 0.01 else "Écart détecté",
              delta_color="off" if abs(k["ca_delta"]) < 0.01 else "inverse")

    tabs = st.tabs([
        "3. ORX absents de PBI",
        "4. PBI mal intégrés",
        "ORX — Groupements",
        "PBI — Groupements",
    ])

    with tabs[0]:
        st.markdown(f"**{len(results['missing_in_pbi'])} dossier(s)** présents dans Orchestra "
                    "mais absents de PowerBI.")
        st.dataframe(results["missing_in_pbi"], use_container_width=True, hide_index=True)

    with tabs[1]:
        bi = results["badly_integrated"]
        st.markdown(f"**{len(bi)} dossier(s)** mal intégrés dans PowerBI "
                    "(TO manquant, Sale id manquant, ou prix divergent d'ORX).")
        st.caption("Toutes les colonnes d'origine PBI sont conservées (dont Product et Product id).")
        st.dataframe(bi, use_container_width=True, hide_index=True)

    with tabs[2]:
        st.markdown("##### Par vente (Code campagne)")
        st.dataframe(results["orx_by_sale"], use_container_width=True, hide_index=True)
        st.markdown("##### Par statut")
        st.dataframe(results["orx_by_status"], use_container_width=True, hide_index=True)
        if not results["orx_by_to"].empty:
            st.markdown("##### Par Tour Opérateur")
            st.dataframe(results["orx_by_to"], use_container_width=True, hide_index=True)

    with tabs[3]:
        st.markdown("##### Par vente (Sale id)")
        st.dataframe(results["pbi_by_sale"], use_container_width=True, hide_index=True)
        st.markdown("##### Par statut (Booking state)")
        st.dataframe(results["pbi_by_status"], use_container_width=True, hide_index=True)
        if not results["pbi_by_to"].empty:
            st.markdown("##### Par Tour Opérateur")
            st.dataframe(results["pbi_by_to"], use_container_width=True, hide_index=True)

    st.subheader("Export")
    xlsx_bytes = build_report_xlsx(results, date_info)
    st.download_button(
        "Télécharger le rapport Excel",
        data=xlsx_bytes,
        file_name=f"reconciliation_ORX_PBI_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    run()
