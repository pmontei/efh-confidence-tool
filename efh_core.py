"""
Lógica central do EFH Confidence Assessment Tool (EMODnet Seabed Habitats).
Separada da interface para poder ser reutilizada (Streamlit, scripts, testes).

Authors: Pedro Monteiro, Frederico Oliveira, Luís Bentes, Jorge M. Gonçalves — CCMAR (Centre of Marine Sciences), EMODnet Seabed Habitats project
"""
from datetime import datetime
from io import BytesIO

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

APP_TITLE = "EFH Confidence Assessment Tool"
APP_VERSION = "1.1"
AUTHOR = "Pedro Monteiro, Frederico Oliveira, Luís Bentes, Jorge M. Gonçalves"
INSTITUTION = "CCMAR – Centre of Marine Sciences, University of Algarve"
PROJECT = "EMODnet Seabed Habitats"
EMODNET_SBH_URL = "https://emodnet.ec.europa.eu/en/seabed-habitats"
GUIDE_URL = "https://emodnet.ec.europa.eu/sites/emodnet.ec.europa.eu/files/public/emodnet_guide-for-submitting-efh_v3.pdf"
GUIDE_REF = ("Gonçalves, J., Monteiro, P., Oliveira, F., Bentes, L. (CCMAR), Lillis, H., Duncan, G., "
             "Manca, E. (JNCC), McGrath, F. (MI), Vasquez, M. (Ifremer). Guidance for submitting "
             "Essential Fish Habitat (EFH) data v3.0. EMODnet Seabed Habitats, Deliverable D2.11.")

# ---------------------------------------------------------------------------
# Critérios: (nome, [(pontuação, descrição), ...])
# Os máximos de cada secção são calculados automaticamente.
# ---------------------------------------------------------------------------
DATA_ACQUISITION = [
    ("Acquisition Technique", [(3, "Highly Appropriate"), (2, "Moderately Appropriate"), (1, "Slightly Appropriate")]),
    ("Sampling Density", [(3, "Good Coverage"), (2, "Moderate Coverage"), (1, "Poor Coverage")]),
    ("Source Data", [(3, "Fisheries Independent"), (2, "Fisheries Dependent"), (1, "Interviews/Questionnaires")]),
    ("Standard", [(3, "Standardized Methods"), (1, "Non-Standardized Methods")]),
    ("Study Vintage", [(3, "< 5 years"), (2, "5-10 years"), (1, "> 10 years")]),
    ("Time Series", [(3, "> 5 years"), (2, "1-5 years"), (1, "< 1 year")]),
]

DATA_SELECTION = [
    ("Data Adequacy", [(3, "Detailed Analysis"), (2, "Abundance Data"), (1, "Presence/Absence")]),
    ("Interpretation", [(3, "Appropriate Technique"), (2, "Moderately Appropriate"), (1, "Inappropriate Technique")]),
    ("Map Validation", [(3, "High Accuracy (External)"), (2, "High Accuracy (Internal)"), (1, "Low Accuracy or None")]),
]

# ---------------------------------------------------------------------------
# Scoring guidelines — from "Guidance for submitting EFH data v3.0",
# section 5, Tables 1–3.
# ---------------------------------------------------------------------------
GUIDE_INTRO = (
    "The Confidence Assessment Tool (CAT) assigns a quality and suitability score to each EFH map. "
    "It assesses the quality and suitability of the **mapping process**, not the accuracy of the map. "
    "The overall score combines a **Data Acquisition** score (*How good was the data acquisition?*) "
    "and a **Data Selection & Interpretation** score (*How good is the interpretation?*). "
    "The CAT was originally developed for Seabed Habitat maps in the EU MESH project and adapted "
    "for Essential Fish Habitats."
)

GUIDE = {
    "Acquisition Technique": {
        "question": "Were the survey techniques used appropriate for the species type?",
        "definition": "Data quality depends strongly on choosing an appropriate sampling method. "
                      "E.g. pelagic trawling suits pelagic species but is less suitable for benthic "
                      "species; very large mesh sizes may under-represent juveniles.",
    },
    "Sampling Density": {
        "question": "What density of sampling points did the survey cover?",
        "definition": "Point density obtained for the species under assessment and, to a lesser "
                      "extent, for any environmental parameters used in the EFH modelling.",
    },
    "Source Data": {
        "question": "What is the data source?",
        "definition": "Research cruises (fisheries-independent) are more reliable than "
                      "fisheries-dependent surveys, which are in turn more reliable than "
                      "interviews/questionnaires on fishers' perceptions.",
    },
    "Standard": {
        "question": "Were standards applied to the collection of the data?",
        "definition": "Standardised sampling: the same appropriate method, with the same "
                      "characteristics, in the same season and in the same way over time, keeping "
                      "sampling bias and precision constant.",
    },
    "Study Vintage": {
        "question": "How recent is the data?",
        "definition": "Age of the data collection, not the date the EFH map was produced.",
    },
    "Time Series": {
        "question": "How long is the time series?",
        "definition": "Length (in years) of the time series used to produce the EFH map.",
    },
    "Data Adequacy": {
        "question": "Is the data adequate?",
        "definition": "Level of detail of the data used to define the EFH. 3 = suitable "
                      "biological criteria (e.g. highest concentrations of individuals below size "
                      "at first maturity for nurseries, or of mature females for spawning areas); "
                      "2 = at least abundance data; 1 = presence/absence only.",
    },
    "Interpretation": {
        "question": "Were the data appropriately interpreted?",
        "definition": "Type of scientific technique or model used to define the EFH, ranging from "
                      "species distribution models with statistical classification to fishers' "
                      "anecdotal information.",
    },
    "Map Validation": {
        "question": "How accurate is the resulting map?",
        "definition": "Level of accuracy obtained during model development (external validation "
                      "with independent data scores higher than internal validation).",
    },
}

GUIDE_SUBMISSION = (
    "**Submission package (per EFH map):** EFH map (.shp categorical or .tif continuous), "
    "Study area (.shp), this **Confidence assessment score sheet (.xlsx)**, Metadata (INSPIRE .xml) "
    "and Data Provider Agreement (.docx).\n\n"
    "The **Overall Confidence Score** is the value to enter in the `SUM_CONF` field of the Study Area "
    "shapefile. All file names should start with the 8-character **GUI** code "
    "(2-letter ISO 3166-1 country code + 6 digits, e.g. `PT000001`)."
)

# ---------------------------------------------------------------------------
# Campos descritivos
# ---------------------------------------------------------------------------
EFH_TYPES = ["Spawning grounds", "Nursery grounds", "Feeding grounds", "Migration corridors"]

EFH_TYPE_HELP = {
    "Spawning grounds": "Areas with clear evidence of spawning.",
    "Nursery grounds": "Areas with clear evidence of heavy concentrations of fish below minimum conservation reference size.",
    "Feeding grounds": "Areas with clear evidence of feeding aggregations.",
    "Migration corridors": "Routes used by highly mobile species during their life cycle.",
}

LIFE_STAGES = ["Eggs", "Larvae", "Juveniles", "Sub-adults", "Adults", "Spawners / Mature",
               "Multiple stages", "Unknown"]

SEA_BASINS = [
    "Baltic Sea", "North Sea", "Celtic Seas", "Bay of Biscay and Iberian Coast", "Macaronesia",
    "Western Mediterranean Sea", "Adriatic Sea", "Ionian Sea and Central Mediterranean Sea",
    "Aegean-Levantine Sea", "Black Sea", "Arctic Ocean", "Norwegian Sea", "Other",
]

COUNTRIES = [
    "Albania", "Belgium", "Bulgaria", "Croatia", "Cyprus", "Denmark", "Estonia", "Faroe Islands",
    "Finland", "France", "Georgia", "Germany", "Greece", "Iceland", "Ireland", "Israel", "Italy",
    "Latvia", "Lithuania", "Malta", "Montenegro", "Morocco", "Netherlands", "Norway", "Poland",
    "Portugal", "Romania", "Slovenia", "Spain", "Sweden", "Tunisia", "Türkiye", "Ukraine",
    "United Kingdom", "Other",
]

# Ordem das colunas no Excel
METADATA_FIELDS = ["GUI", "Species (scientific name)", "AphiaID", "EFH Type", "Life Stage",
                   "Sea Basin", "Country", "Institution", "Assessor", "Notes"]


def option_label(score: int, text: str) -> str:
    return f"{score} - {text}"


def valid_gui(gui: str) -> bool:
    """GUI = 2 letras (ISO 3166-1) + 6 dígitos, ex. PT000001."""
    gui = gui.strip()
    return len(gui) == 8 and gui[:2].isalpha() and gui[:2].isupper() and gui[2:].isdigit()


def section_score(criteria, selections: dict) -> float:
    """Percentagem: soma das pontuações escolhidas / soma dos máximos."""
    total = sum(selections[name] for name, _ in criteria)
    maximum = sum(max(s for s, _ in opts) for _, opts in criteria)
    return total / maximum * 100


def compute_scores(selections: dict) -> dict:
    acq = section_score(DATA_ACQUISITION, selections)
    sel = section_score(DATA_SELECTION, selections)
    return {
        "Data Acquisition Score (%)": round(acq),
        "Data Selection & Interpretation Score (%)": round(sel),
        "Overall Confidence Score (%)": round((acq + sel) / 2),
    }


def confidence_class(overall: float) -> str:
    """Classe qualitativa (ajustar os limites ao protocolo do projeto)."""
    if overall >= 75:
        return "High"
    if overall >= 50:
        return "Medium"
    return "Low"


def build_record(metadata: dict, selections: dict) -> dict:
    scores = compute_scores(selections)
    labels = {}
    for name, opts in DATA_ACQUISITION + DATA_SELECTION:
        labels[name] = option_label(selections[name], dict(opts)[selections[name]])
    return {
        "Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
        **{f: metadata.get(f, "") for f in METADATA_FIELDS},
        **labels,
        **scores,
        "Confidence Class": confidence_class(scores["Overall Confidence Score (%)"]),
    }


def guidelines_table() -> pd.DataFrame:
    rows = []
    for section, crit in (("Data Acquisition", DATA_ACQUISITION),
                          ("Data Selection & Interpretation", DATA_SELECTION)):
        for name, opts in crit:
            rows.append({
                "Section": section,
                "Criterion": name,
                "Question": GUIDE[name]["question"],
                "Scoring guidelines": "\n".join(option_label(s, t) for s, t in opts),
                "Definition": GUIDE[name]["definition"],
            })
    return pd.DataFrame(rows)


def about_table() -> pd.DataFrame:
    return pd.DataFrame({
        "Item": ["Tool", "Version", "Project", "Authors", "Institution", "Portal", "Reference",
                 "Guide (PDF)", "Exported"],
        "Value": [APP_TITLE, APP_VERSION, PROJECT, AUTHOR, INSTITUTION, EMODNET_SBH_URL, GUIDE_REF, GUIDE_URL,
                  datetime.now().strftime("%Y-%m-%d %H:%M")],
    })


def read_existing(file) -> pd.DataFrame:
    """Lê um Excel exportado antes por esta ferramenta (para continuar a acrescentar)."""
    return pd.read_excel(file, sheet_name="Assessments", dtype={"AphiaID": str}).fillna("")


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Gera o .xlsx com três folhas: Assessments, Scoring Guidelines e About."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Assessments", index=False)
        guidelines_table().to_excel(writer, sheet_name="Scoring Guidelines", index=False)
        about_table().to_excel(writer, sheet_name="About", index=False)

        header_fill = PatternFill("solid", fgColor="0A3D62")
        for ws in writer.book.worksheets:
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width + 2, 10), 60)
            ws.freeze_panes = "A2"
    return buffer.getvalue()
