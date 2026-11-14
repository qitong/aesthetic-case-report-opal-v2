---
name: aesthetic-case-report-opal-v2
description: Staged, fast-track medical aesthetics clinical case report PPTX generator (v2) with interactive Express Key-Slides (5-7 pages) vs. Full Deck (14-20 pages) selection, staged feedback checkpoints, single-pass photo de-identification, and declarative zero-overflow JSON-to-PPTX compilation. Use when the user asks to create a medical aesthetics case report slide deck, convert clinical patient records or before-and-after photos into a presentation, synthesize a clinical paper with a case report, or run aesthetic-case-report-opal-v2.
---

# Aesthetic Case Report Opal v2 (Staged & Fast-Track Edition)

Generate clinical-grade, high-aesthetic PowerPoint (`.pptx`) presentations for medical aesthetics case reports **without monolithic 1-hour black-box runs**.

> **Core Architecture Principle**: Break long deck-building tasks into **4 staged sub-tasks**, deliver concrete intermediate outputs at every stage, and establish an explicit **Validation & Feedback Loop (Feedback Checkpoint)** before advancing to the next stage.

---

## Staged Workflow & Feedback Checkpoints

### Stage 0: Interactive Scope, Deck Tier & Privacy Alignment (MANDATORY FIRST STEP)

Before generating any images or writing any slide data, inspect the workspace files (using `scripts/inspect_case_assets.py <photo_dir>` if a photo folder exists so you only read **1 contact sheet** instead of 10+ raw photos) and ask a concise, structured question set to lock in the scope:

1. **Deck Scope & Time-Budget Tier (必选档位确认 — 现场时间短可做精简关键页)**:
   - **Option A — Express Key-Slides Mode (极速核心对照版 · 5–7 页 · 约 2–5 分钟)**:
     Focuses strictly on the highest-impact clinical slides for rapid meetings or chairside review:
     1. Cover Slide (封面与术者信息)
     2. Baseline Multi-Axis Assessment OR Anatomical Atlas (基线评估或整页解剖靶点图)
     3. Layered Treatment Plan & Bilateral L/R Rheology (分层注射方案与双侧剂量)
     4. Before & After Comparison — Primary View (正位静态术前术后对照)
     5. Before & After Comparison — Secondary / Dynamic View (45°斜位/动态对照)
     6. Longitudinal Patient Journey OR Key Clinical Pearls (纵向随访时间轴或核心要点总结)
   - **Option B — Full Comprehensive Deck Mode (完整学术/临床报告版 · 14–20 页)**:
     Includes section dividers (`01`–`05`), full anatomical mapping/consensus tables, Pre- & Post-treatment 5-axis assessment slides, objective 5-axis shift bar chart (`chart`), multi-angle Before & After gallery, longitudinal journey, safety follow-up table, and clinical pearls/disclosure.

2. **Input Sources & Simulated Photo Strategy**:
   - Confirm patient demographics, chief complaint, product rheology, bilateral L/R dosage, and needle/cannula specs.
   - If real photos are absent and AI-simulated photos are requested:
     - **Express Mode** generates at most **2–4 photos** (e.g., 1 Baseline + 1 Post-Tx Frontal, plus optional 45° pair).
     - **Full Mode** generates **4–6 photos** (anchoring all follow-ups to the approved Baseline portrait).
     - Always state platform image-generation credit usage upfront before generating.

3. **Template Palette & De-identification Confirmation**:
   - Choose **Editorial Luxury (`editorial`, Champagne Ivory + Rose Gold)**, **Academic Clinical (`academic`, Warm Alabaster + Clinical Slate)**, or a custom `design.md`.
   - Confirm that `scripts/apply_eye_mask.py` will mask the orbital region (`#1E2022` opaque horizontal band, `--opacity 1.0 --radius 0`) on all patient/simulated portraits unless explicit written consent to keep eyes unmasked is provided.

> **🔄 Feedback Checkpoint 0**: Wait for the user to confirm **Express vs. Full Mode** and input preferences before executing Stage 1. *(Only skip if the user explicitly specified both the mode and all inputs in their initial prompt.)*

---

### Stage 1: Clinical Photo Preprocessing / Anchor-First Generation & Privacy Gate

Execute image preparation as a self-contained sub-task and verify before moving to slide authoring:

- **Case A — User Supplied Existing Photos**:
  1. Run `python3 scripts/inspect_case_assets.py <photo_dir> --out <work_dir>/_contact_sheet.jpg` and inspect `_contact_sheet.jpg` once to classify angles (Frontal, 45° Oblique, 90° Profile) and timepoints.
  2. Apply mandatory eye-band de-identification in a single pass (preserving untouched originals in `_unmasked/`):
     ```bash
     python3 scripts/apply_eye_mask.py <photo_dir> --in-place --backup-dir <photo_dir>/_unmasked
     ```

- **Case B — AI-Simulated Patient Journey Photos**:
  1. **Anchor First (先单张锚点确认)**: Generate **ONLY 1 Baseline (`T0`) portrait first**. Inspect it to confirm framing, lighting, and clinical severity.
  2. Generate the remaining follow-up / after views using `T0` as the strict identity/pose reference so lighting, clothing/hairline, and Frankfurt horizontal plane stay consistent.
  3. **Single-Pass Trim + Mask + Resize**: Never run separate watermark-cropping and eye-masking loops. Run one command to trim any bottom watermark strip, cap DSLR dimensions (`--max-dim 2400`), and apply the solid eye mask simultaneously:
     ```bash
     python3 scripts/apply_eye_mask.py <raw_sim_dir> --output-dir <work_dir>/photos --suffix "" --crop-watermark-ratio 0.06
     ```

> **🔄 Feedback Checkpoint 1**: Present the masked photo summary (or contact sheet preview) confirming: (1) zero iris/eyelid exposure, (2) target anatomy (e.g., midface/ tear trough / neck lines) completely unobstructed, and (3) realistic clinical progression across timepoints. If working in an interactive session with newly simulated photos, briefly confirm user satisfaction or proceed directly if pre-approved.

---

### Stage 2: Structured Case Spec (`deck_spec.json`) & Slide Blueprint Verification

**DO NOT write custom 400-line `python-pptx` wrapper scripts from scratch.** `scripts/generate_case_pptx.py` v2 natively implements all 11 canonical medical aesthetics slide archetypes and auto-fits text sizes to prevent footer collisions.

1. Read `references/default_template_spec.md` for the JSON schema and slide blueprint corresponding to the user's chosen tier (**Express 5–7 slides** vs. **Full 14–20 slides**).
2. Author `<work_dir>/deck_spec.json` containing:
   - `theme_mode` (`"editorial"` or `"academic"`), `output_path`, and `default_footer` (including `眼部已脱敏 / Eyes masked`).
   - `slides`: array of declarative slide objects (`cover`, `divider`, `fullbleed_atlas`, `anatomy_split`, `assessment`, `chart`, `treatment_plan`, `before_after`, `patient_journey`, `consensus_table`, `summary`).
3. Verify clinical precision in the JSON:
   - Every `treatment_plan` step specifies exact anatomical layer/compartment (e.g., `DMCFP`, `SOOF`, `Supraperiosteal`, `Subdermal`) and **Left (L) vs. Right (R) mL volumes**.
   - Every `before_after` slide labels both **View angle** (`正位 Frontal` / `45°斜位 Oblique`) and **Expression state** (`静态 Static` / `动态 Dynamic`).

> **🔄 Feedback Checkpoint 2**: Output a concise **Slide-by-Slide Outline Table** (Page # · Slide Type · Core Title & Dosage Summary) for visibility, or pause for user sign-off if clinical parameters were inferred. Editing a title or dosage at this stage takes 1 second in `deck_spec.json`.

---

### Stage 3: Instant Zero-Overflow PPTX Compilation & Fast QA Loop

1. Compile `<work_dir>/deck_spec.json` directly into `.pptx` and trigger the built-in millisecond static layout validator:
   ```bash
   python3 scripts/generate_case_pptx.py <work_dir>/deck_spec.json
   ```
2. **Built-in Fast QA (Replaces 15-Minute LibreOffice Render-and-Fix Loops)**:
   - `generate_case_pptx.py` automatically scales `clinical_notes` (10.5pt → 8.5pt), compacts 5-axis score rows when 4+ notes are present, auto-fits 3–4 treatment steps, applies high-contrast alternating table row fills (`#FFFDFB` / `#F1E9DF`), and checks every shape against `FOOTER_TOP_IN = 7.02 in`.
   - Confirm the terminal reports `[QA PASS] Zero footer collisions, zero unresolved placeholders, all text auto-fitted.` (If any `[QA WARN]` is printed, fix the flagged slide in `deck_spec.json` and re-run in <2 seconds).
   - *Optional Visual Spot-Check*: Only run LibreOffice/PDF visual rendering if the user supplies a custom `design.md` override or explicitly requests a PNG thumbnail overview sheet.

> **🔄 Feedback Checkpoint 3 (Final Delivery & Iterative Refinement)**:
> Deliver the `.pptx` link (plus optional case markdown summary) and remind the user they can request instant incremental edits—such as *"expand this 6-page Express deck into the 18-page Full version"*, *"adjust L/R dosage on Slide 3"*, or *"swap theme from `editorial` to `academic`"*—which re-compiles from `deck_spec.json` in 2 seconds.

---

## Gotchas

- **Never Hand-Roll Custom `python-pptx` Scripts Unless Necessary**: Hand-coding shape coordinates caused >30 minutes of coding and visual-overflow debugging in v1. Always use `deck_spec.json` + `scripts/generate_case_pptx.py` first.
- **Clinical Photo Distortion**: Never stretch aspect ratios. `OpalCaseReportDeck._add_image_fitted` automatically reads EXIF orientation, center-crops (`crop_to_fill=True`) or letterboxes (`crop_to_fill=False`), and downscales oversized DSLR files in memory.
- **Identifiable Eyes in Clinical Photos**: Never place an unmasked eye region into a shareable case deck—including AI-simulated portraits. Always run `scripts/apply_eye_mask.py` (`--opacity 1.0 --radius 0`) and include `眼部已脱敏 / Eyes masked` in `default_footer`.
- **Mask Geometry vs. Target Anatomy**: Verify on one frontal and one oblique photo that the 33%–49% band covers brows/eyes without obscuring the chief-complaint region (override `--top` / `--bottom` for tight crops or 90° profiles).
- **Bilateral Dosage & Dynamic Expression Clarity**: Always distinguish Left (L) vs. Right (R) mL volumes and label Static (`静态`) vs. Dynamic (`微笑/动态`) photography.
