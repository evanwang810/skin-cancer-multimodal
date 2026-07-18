# Multimodal Skin Cancer Detection (mini project)

A small homework project: does adding basic patient info (age, sex, lesion
location) to a skin-lesion image help a model tell **benign from malignant**? It
trains an image-only model and an image+metadata model on **HAM10000** and
compares them.

Nothing to install or download by hand — open the notebook in Google Colab, set
the runtime to **GPU**, and run it top to bottom.

Live site: **https://evanwang810.github.io/skin-cancer-multimodal/**

## Files
- `skin_cancer_multimodal.ipynb` — the whole thing.
- `outputs/` — model checkpoints (created when you run `save_model(...)`).
- `report/` + `skin_cancer_report.zip` — the exported test-case results: images,
  Grad-CAM overlays, predictions (`data.json`/`data.js`), and the mined findings
  (`insights.json`/`insights.js`). Produced by the notebook's export cell.
- `index.html` — the **Explorer**: an annotated diagram of the two-tower
  network, then the ~500 test lesions browsable as a 2D feature-space scatter,
  a rotatable 3D point cloud, or a sortable/filterable table, with CSV/JSON
  export of whatever's currently filtered. Click a case for the photo,
  Grad-CAM, and prediction. The 3D depth axis is a real third PCA component once
  the notebook is rerun (it exports 3-component PCA); until then it falls back
  to P(malignant), labeled as such.
- `insights.html` — real patterns mined from the 500 cases: how malignancy rate
  and recall shift with age, which anatomical sites the model struggles with,
  a sex-based recall gap, and how often the model is confidently wrong.
- `presentation.html` — a 14-slide walkthrough (arrow keys / click to advance)
  covering the method, a real PCA feature-space scatter, real example cases,
  and the mined findings (age, location, sex, calibration).
- `presentation.pptx` — the same deck as an actual PowerPoint file (native
  charts, embedded case images), for anyone who wants to present it without
  a browser or drop it into Google Slides / Keynote.

All three pages read `report/*.js` (embedded as JS globals), so they work via a
plain double-click with no server — `report/*.json` is the same data for anyone
who'd rather load it over http or parse it directly.

## What it does
1. Loads HAM10000 and buckets the 7 diagnoses into benign / malignant.
2. Splits by `lesion_id` so the same lesion never lands in both train and test.
3. Trains image-only vs image+metadata (encoder set by `IMAGE_ENCODER`).
4. Reports AUC / accuracy / precision / recall, tunes the threshold on validation,
   and bootstraps whether the metadata actually helped.
5. Grad-CAM heatmaps + a downloadable results bundle for the showcase page.

## Handy knobs
- `N_SAMPLES` — how much data to load.
- `IMAGE_ENCODER` — `"resnet50"` (default) / `"vit"` / `"cnn"`.
- `EARLY_STOP` — off by default.
- `DISPLAY_N` — how many test cases to export.

> Just a learning project — not a real diagnostic tool.
