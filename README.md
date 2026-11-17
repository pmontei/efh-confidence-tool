# EFH Confidence Assessment Tool — EMODnet Seabed Habitats

Web app to score the confidence of Essential Fish Habitat (EFH) maps submitted to
[EMODnet Seabed Habitats](https://emodnet.ec.europa.eu/en/seabed-habitats) and export the
results to Excel (.xlsx).

**Authors:** Pedro Monteiro, Frederico Oliveira, Luís Bentes, Jorge M. Gonçalves · **Institution:** CCMAR – Centre of Marine Sciences, University of Algarve ·
**Project:** EMODnet Seabed Habitats

Methodology: [*Guidance for submitting Essential Fish Habitat (EFH) data v3.0*](https://emodnet.ec.europa.eu/sites/emodnet.ec.europa.eu/files/public/emodnet_guide-for-submitting-efh_v3.pdf), section 5
(adapted from the MESH confidence assessment).

## Online version
Open the tool in the browser: `https://<your-user>.github.io/efh-confidence-tool/`
(no installation needed; works offline if you download `index.html`).

## Files
```
index.html                  Stand-alone web version (GitHub Pages)
app.py                      Streamlit interface
efh_core.py                 Scoring, guidelines text, Excel export
requirements.txt
assets/emodnet_sbh_logo.png EMODnet Seabed Habitats logo (included)
```

## Logo
The official EMODnet logo is included in `assets/emodnet_sbh_logo.png`.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Share with partners (free, browser only)
1. Create a GitHub repository with the files above (including `assets/`).
2. Go to https://share.streamlit.io, sign in with GitHub → **Create app**.
3. Choose the repository, branch `main`, file `app.py` → **Deploy**.
4. Send the link (`https://<name>.streamlit.app`) to the partners.

Each partner downloads their own `.xlsx` (sheets: *Assessments*, *Scoring Guidelines*, *About*).
To continue later, upload the same file in the sidebar and keep adding maps.
The Overall Confidence Score is the value for the `SUM_CONF` field of the Study Area shapefile.

## Customise
Edit `efh_core.py`: criteria (`DATA_ACQUISITION`, `DATA_SELECTION`), guideline texts (`GUIDE`),
lists (`EFH_TYPES`, `SEA_BASINS`, `COUNTRIES`, `LIFE_STAGES`) and class thresholds
(`confidence_class()`).
