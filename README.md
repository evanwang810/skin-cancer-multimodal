# Multimodal Skin Cancer Detection (mini project)

A small homework project: does adding basic patient info (age, sex, lesion
location) to a skin-lesion image help a model tell **benign from malignant**? It
trains an image-only model and an image+metadata model on **HAM10000** and
compares them.

Nothing to install or download by hand — open the notebook in Google Colab, set
the runtime to **GPU**, and run it top to bottom.

## Files
- `skin_cancer_multimodal.ipynb` — the whole thing.
- `outputs/` — model checkpoints (created when you run `save_model(...)`).
- `report/` + `skin_cancer_report.zip` — the exported test-case results (images,
  Grad-CAM overlays, predictions). Produced by the last cell.
- `index.html` — the interactive showcase (project root, so it can serve as a
  GitHub Pages site). Open it after running the last cell to browse the ~500 test
  lesions on a feature-space scatter: hover for a preview, click for the photo,
  Grad-CAM, and prediction. It reads `report/data.js` (+ images), which the last
  cell writes, so a plain double-click works with no server.

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
