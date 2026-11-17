"""
EFH Confidence Assessment Tool — EMODnet Seabed Habitats (Streamlit web app).

Authors: Pedro Monteiro, Frederico Oliveira, Luís Bentes, Jorge M. Gonçalves — CCMAR, EMODnet Seabed Habitats project

Executar localmente:   streamlit run app.py
Publicar online:       Streamlit Community Cloud (ver README.md)
"""
from pathlib import Path

import pandas as pd
import streamlit as st

from efh_core import (APP_TITLE, APP_VERSION, AUTHOR, COUNTRIES, DATA_ACQUISITION,
                      DATA_SELECTION, EFH_TYPE_HELP, EFH_TYPES, EMODNET_SBH_URL, GUIDE,
                      GUIDE_INTRO, GUIDE_REF, GUIDE_URL, GUIDE_SUBMISSION, INSTITUTION, LIFE_STAGES, PROJECT,
                      SEA_BASINS, build_record, compute_scores, confidence_class, option_label,
                      read_existing, to_excel_bytes, valid_gui)

ASSETS = Path(__file__).parent / "assets"
LOGO = ASSETS / "emodnet_sbh_logo.png"

st.set_page_config(page_title=f"{APP_TITLE} | EMODnet Seabed Habitats",
                   page_icon=str(LOGO) if LOGO.exists() else "🐟", layout="wide")

if "records" not in st.session_state:
    st.session_state.records = []

# ---------------------------------------------------------------- Cabeçalho
if LOGO.exists():
    st.logo(str(LOGO), link=EMODNET_SBH_URL)
    head_logo, head_text = st.columns([1, 5], vertical_alignment="center")
    head_logo.image(str(LOGO), use_container_width=True)
else:
    head_text = st.container()

with head_text:
    st.title(APP_TITLE)
    st.markdown(
        f"Confidence scoring for **Essential Fish Habitat (EFH)** maps submitted to "
        f"[**EMODnet Seabed Habitats**]({EMODNET_SBH_URL})."
    )

# ---------------------------------------------------------------- Sidebar
with st.sidebar:
    st.header("Continue an existing file")
    uploaded = st.file_uploader("Load a previously exported .xlsx", type=["xlsx"])
    if uploaded is not None and st.session_state.get("loaded_name") != uploaded.name:
        try:
            existing = read_existing(uploaded)
            st.session_state.records = existing.to_dict("records") + st.session_state.records
            st.session_state.loaded_name = uploaded.name
            st.success(f"{len(existing)} assessments loaded.")
        except Exception as e:
            st.error(f"Could not read file: {e}")

    if st.button("Clear all assessments"):
        st.session_state.records = []
        st.session_state.pop("loaded_name", None)
        st.rerun()

    st.divider()
    st.markdown(
        f"**{PROJECT}**  \n[{EMODNET_SBH_URL.removeprefix('https://')}]({EMODNET_SBH_URL})\n\n"
        f"**Authors:** {AUTHOR}  \n**Institution:** {INSTITUTION}  \n**Version:** {APP_VERSION}"
    )

show_guide = st.toggle("📘 Scoring guidelines", value=False,
                       help="Show the EMODnet guidance next to the form")

if show_guide:
    main, side = st.columns([3, 2], gap="large")
else:
    main, side = st.container(), None

# ---------------------------------------------------------------- Painel lateral de orientações
if side is not None:
    with side:
        with st.container(border=True, height=900):
            st.subheader("📘 Scoring guidelines")
            st.markdown(GUIDE_INTRO)
            st.markdown("**Overall score** = average of the two section scores. "
                        "Each section score = sum of points ÷ maximum possible × 100.")
            for section, crit in (("Data Acquisition — How good is the data acquisition?", DATA_ACQUISITION),
                                  ("Data Selection & Interpretation — How good is the selection and interpretation?",
                                   DATA_SELECTION)):
                st.markdown(f"#### {section}")
                for name, opts in crit:
                    with st.expander(name):
                        st.markdown(f"*{GUIDE[name]['question']}*")
                        st.markdown("\n".join(f"- **{s}** = {t}" for s, t in opts))
                        st.caption(GUIDE[name]["definition"])
            st.markdown("#### EFH types")
            st.markdown("\n".join(f"- **{k}** – {v}" for k, v in EFH_TYPE_HELP.items()))
            st.markdown("#### Submission")
            st.markdown(GUIDE_SUBMISSION)
            st.caption(f"Source: {GUIDE_REF} [Download the guide (PDF)]({GUIDE_URL})")

# ---------------------------------------------------------------- Formulário
with main:
    st.subheader("1. Map information")
    metadata = {}
    c1, c2, c3 = st.columns(3)
    metadata["GUI"] = c1.text_input("GUI (EFH map code) *", placeholder="e.g. PT000001",
                                    help="2-letter ISO country code + 6 digits, provided by the "
                                         "national project partner.").strip().upper()
    metadata["Species (scientific name)"] = c2.text_input("Species (scientific name)",
                                                          placeholder="e.g. Merluccius merluccius")
    metadata["AphiaID"] = c3.text_input("AphiaID", placeholder="e.g. 126484",
                                        help="WoRMS identifier: marinespecies.org/aphia.php")

    c1, c2, c3 = st.columns(3)
    efh_type = c1.selectbox("EFH Type *", EFH_TYPES, index=None, placeholder="Select…",
                            help="\n".join(f"{k}: {v}" for k, v in EFH_TYPE_HELP.items()))
    life_stage = c2.selectbox("Life Stage", LIFE_STAGES, index=None, placeholder="Select…")
    sea_basin = c3.selectbox("Sea Basin", SEA_BASINS, index=None, placeholder="Select…")
    metadata["EFH Type"] = efh_type or ""
    metadata["Life Stage"] = life_stage or ""
    metadata["Sea Basin"] = sea_basin or ""

    c1, c2, c3 = st.columns(3)
    metadata["Country"] = c1.selectbox("Country", COUNTRIES, index=None, placeholder="Select…") or ""
    metadata["Institution"] = c2.text_input("Institution")
    metadata["Assessor"] = c3.text_input("Assessor (who made the assessment) *")
    metadata["Notes"] = st.text_area("Notes", height=80)

    def criteria_inputs(criteria, key_prefix):
        chosen = {}
        for name, opts in criteria:
            labels = [option_label(s, t) for s, t in opts]
            choice = st.selectbox(name, labels, key=f"{key_prefix}_{name}",
                                  help=f"{GUIDE[name]['question']}\n\n{GUIDE[name]['definition']}")
            chosen[name] = int(choice.split(" - ")[0])
        return chosen

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("2. Data Acquisition")
        selections = criteria_inputs(DATA_ACQUISITION, "acq")
    with col_b:
        st.subheader("3. Data Selection & Interpretation")
        selections |= criteria_inputs(DATA_SELECTION, "sel")

    # ------------------------------------------------------------ Resultados (tempo real)
    scores = compute_scores(selections)
    st.subheader("4. Confidence scores")
    m1, m2, m3 = st.columns(3)
    m1.metric("Data Acquisition", f"{scores['Data Acquisition Score (%)']}%")
    m2.metric("Data Selection & Interpretation", f"{scores['Data Selection & Interpretation Score (%)']}%")
    overall = scores["Overall Confidence Score (%)"]
    m3.metric("Overall Confidence (SUM_CONF)", f"{overall}%", confidence_class(overall), delta_color="off")

    if st.button("➕ Add this assessment to the table", type="primary"):
        missing = [f for f, v in (("GUI", metadata["GUI"]), ("EFH Type", metadata["EFH Type"]),
                                  ("Assessor", metadata["Assessor"])) if not v]
        if missing:
            st.warning(f"Please fill in: {', '.join(missing)}.")
        else:
            if not valid_gui(metadata["GUI"]):
                st.info("Note: the GUI should be 2 letters + 6 digits (e.g. PT000001). Saved anyway.")
            st.session_state.records.append(build_record(metadata, selections))
            st.success(f"Assessment for '{metadata['GUI']}' added.")

    # ------------------------------------------------------------ Tabela + exportação
    st.subheader(f"5. Assessments in this session ({len(st.session_state.records)})")
    if st.session_state.records:
        df = pd.DataFrame(st.session_state.records)
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Download Excel (.xlsx)",
            data=to_excel_bytes(df),
            file_name="EFH_confidence_assessments.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("No assessments yet. Fill in the form and click “Add this assessment”.")

# ---------------------------------------------------------------- Rodapé
st.divider()
st.caption(
    f"{APP_TITLE} v{APP_VERSION} · Developed by **{AUTHOR}** ({INSTITUTION}) for the "
    f"[{PROJECT}]({EMODNET_SBH_URL}) project. "
    "Methodology adapted from the MESH confidence assessment; see [*Guidance for submitting "
    f"Essential Fish Habitat (EFH) data v3.0*]({GUIDE_URL})."
)
