# Status

## Current State

- Streamlit app calculates norm tables from an uploaded survey Excel workbook.
- App shell now uses the VN / Smart Tables visual theme and branded header.
- Workbook setup and norm mapping now live on the Survey Question Audit page rather than the sidebar.
- Each app page includes Back and Next page buttons for workflow navigation.
- Saved-norm review mode shows only the Next path from the upload page and moves between Norm tables and Saved datasets.
- Survey Question Audit prompts for missing project metadata fields, but those fields are not included in norm table output.
- Survey Question Audit includes question type, suggested norm/benchmark mapping, editable norm dropdowns, and NA exclusion.
- Raw respondent rows are no longer previewed on the audit page.
- Standard survey platform/system columns default to `NA` in the audit table.
- Scale-like response options use Smart Tables-style ordering in audit summaries and norm tables.
- Survey Question Audit can add optional T2B, T3B, B2B, and B3B rows for scale norms.
- Survey Question Audit has bulk buttons for applying or clearing box-score selections across Scale / Likert rows.
- Saved or suggested norm/benchmark names that include T2B, T3B, B2B, or B3B pre-check the matching box score.
- Saved dataset rules are checked during audit so matching questions reuse prior box-score selections by source variable, question label, or norm/benchmark label.
- Audit box-score selections that conflict with previous saved norms are flagged and require confirmation before the audit mapping or dataset can be saved.
- Selected box-score rows replace the full scale in norm tables for that question.
- Norm audit decisions are stored in `norm_mapping_settings.json`.
- Saved NA aliases are stored in `na_alias_settings.json` and reused on future uploads.
- Question type audit decisions are stored in `question_type_settings.json`.
- Box-score audit decisions are stored in `box_score_settings.json`.
- Uploaded Excel data is treated as respondent-level rows with question columns.
- Survey Question Audit defaults to the BLS / Smart Tables layout with variable names in row 1 and question labels in row 2.
- Response-option counting now follows the Smart Tables exact-match plus common multi-select delimiter matching pattern.
- Norm calculations run at the response-option level.
- Norm tables use audited norm/benchmark variable labels as table headings rather than QNR question text.
- Norm tables always show the saved norms database; active uploads only add save-to-database controls.
- Norm tables are limited to Response option, Control, Test, Lift, and Significance result.
- Control and Test percentages are rounded to whole percentages with no decimals.
- Lift is rounded to whole percentage points with no decimals.
- On-screen norm tables render as BLS-style white tables with black text.
- Norm tables include one denominator-labeled Base size row per metric and do not include response-option selected-count columns.
- Norm tables can be filtered by saved respondent-level fields such as Project, Brand, Client, Industry, Country, Year, Quarter, Gender, and Age when available, with a reset button to return to total control vs test. Metadata filters use exact metadata fields, so Brand does not use brand sentiment.
- Saved Norm tables recalculate from all filtered saved respondent-level datasets as one appended read, summing counts and denominators before percentages, lift, and significance.
- Norm tables page includes an explicit save-to-norms-database step that writes local dataset workbooks, a manifest, and an aggregate saved norms workbook.
- Saved norms storage is anchored to the app directory by default and can be moved to shared persistent storage with `BLS_NORMS_DATA_DIR`.
- Saved datasets are also backed up in `uploaded_datasets/` with separate folders for raw uploads, app-ready norm workbooks, and norm-setting backups.
- If the working `norm_database/` is missing or partial on startup/refresh, the app restores missing data from `uploaded_datasets/`.
- GitHub autocommit can push `uploaded_datasets/` after dataset saves and saved-rule updates when `[github_autocommit]` Streamlit secrets are configured.
- Dataset saves and saved-rule edits use a write lock and atomic manifest writes for cross-session protection.
- Duplicate uploads are detected by respondent ID overlap when an ID field is available; 80% or higher overlap is flagged and requires replacement rather than duplicate saving.
- Upload and Saved datasets pages allow the respondent ID variable to be selected manually when auto-detection misses it.
- Duplicate datasets are flagged during upload with a Dataset already added to norms notification.
- Saved datasets page can edit previous dataset rules and regenerate saved norm tables when standards change.
- Newly saved dataset workbooks include cleaned respondent data and rule sheets for recalculation.
- Response-label sheets are auto-detected when the workbook includes recognizable label columns.
- The app no longer shows a separate Response labels tab.
- Significance output uses simple labels: Significant, Not significant, or Not tested.
- Denominator decisions are stored in `denominator_settings.json`.
- Denominator settings display variable/mapped norm labels instead of QNR question text.

## Denominator Settings

<!-- denominator-settings:start -->
| Norm question | Denominator |
| --- | --- |
| Age | Total answering |
| Brand Advocacy | Total answering |
| Brand Retention | Total answering |
| Brand Sentiment | Total answering |
| Category Interest | Total answering |
| Gender | Total answering |
| OM Feature Awareness | Total answering |
| P_Culture | Total answering |
| P_Tech Leadership | Total answering |
| P_Tech Style | Total answering |
| Purchase Intent | Total answering |
| WRBLS Awareness | Total answering |
| WRBLS BR Awareness | Total answering |
| WRBLS Consideration | Total sample |
| WRBLS Leadership | Total answering |
| awareness | Total answering |
| consideration | Total sample |
<!-- denominator-settings:end -->
