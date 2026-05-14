# BLS Norms Database (Meta)

Streamlit app for calculating response-option norm tables from Excel survey data.

## Run

```bash
streamlit run app.py
```

## Calculation Assumptions

- Norms are calculated at the response-option level.
- Uploaded Excel files are interpreted as respondent-level data: each row is one respondent, and each column is one question or respondent field.
- If project metadata fields are missing, the app prompts for `brand`, `industry`, `client`, `quarter`, `year`, `methodology`, `project`, `country`, and `c_key`; each field can be typed as one workbook-level value or sourced from an existing uploaded variable.
- Project metadata fields are excluded from norm-question audit rows and are not included in norm table output.
- Workbook setup and question mapping are handled on the Survey Question Audit page rather than in the sidebar.
- Each app page includes Back and Next page buttons for moving through the workflow.
- The Survey Question Audit can start saved-norm review mode without uploading a workbook; after selection, Next proceeds to Norm tables and review navigation stays focused on Norm tables and Saved datasets.
- The Survey Question Audit lists uploaded questions, question type, the app's suggested closest norm/benchmark, a dropdown of available norms, an NA checkbox, and optional T2B/T3B/B2B/B3B box-score rows.
- Norm / benchmark audit decisions are stored in `norm_mapping_settings.json` and reused later.
- NA audit decisions also store normalized variable/question-text aliases in `na_alias_settings.json`, so future uploads can default matching questions back to `NA`.
- Question type audit decisions are stored in `question_type_settings.json` and reused later.
- T2B/T3B/B2B/B3B audit decisions are stored in `box_score_settings.json` and reused later.
- If a saved or suggested norm/benchmark name already includes T2B, T3B, B2B, or B3B, the matching box-score option is checked by default unless a saved audit decision overrides it.
- The audit also checks saved dataset rules by source variable, question label, and norm/benchmark label so prior T2B/T3B/B2B/B3B choices are reused for matching questions.
- If a new audit selection conflicts with prior saved norm rules, the app flags the mismatch and requires confirmation before saving the audit mapping or saving the dataset.
- Confirmed standards changes should be followed by updating previous saved datasets so each metric is measured consistently.
- Audit buttons can apply T2B, T3B, B2B, B3B, or clear box selections across all non-NA Scale / Likert questions.
- The default data layout is the BLS / Smart Tables layout: row 1 contains variable names, row 2 contains question labels, and following rows contain respondent-level data.
- During upload, the app reviews file structure for repeated question text with `- response choice` suffixes and can combine those option columns into one audited question with a selected question type.
- Grouped upload questions can keep a custom grouped variable name or use a dropdown of existing saved database variable names for easier future matching.
- Combined multi-select option columns use the response-choice suffix as the answer option and ignore common non-selection values such as `-99`, `0`, `false`, and `Not Selected`.
- Standard survey platform/system columns such as StartDate, EndDate, IPAddress, Progress, Duration, and ResponseId stay visible in the audit table but default to `NA`.
- Each included norm question corresponds to one audited source column in the selected Excel data sheet.
- Each distinct non-empty value in a mapped question column is treated as a response option.
- Scale-like response options are ordered with the same Smart Tables-style positive-to-negative logic when response labels are available.
- Saved `Scale / Likert` question types force Smart Tables-style scale ordering for later audits and norm calculations.
- T2B and T3B use the first two or three ordered non-exclusive scale options; B2B and B3B use the last two or three ordered non-exclusive scale options.
- When any T2B/T3B/B2B/B3B option is selected for a question, the norm table shows only the selected box-score rows for that question instead of the full scale.
- Norm tables can be filtered by Project, Brand, Client, Industry, Country, Year, Quarter, Gender, and Age when those variables or labels exist in the uploaded workbook.
- Norm table filters apply to the saved database view. Project, Brand, Client, Industry, Country, Year, and Quarter filters use exact metadata-variable matches so metric columns such as brand sentiment are not used as filters.
- Reset to total control vs test clears all norm table filters and returns to all saved datasets.
- Norm tables always show the saved norms database. When a workbook is active, the upload is used for the save-to-database step rather than replacing the saved database view.
- Norm tables are headed by the audited norm/benchmark variable label rather than QNR question text.
- Norm table output is limited to Response option, Control, Test, Lift, and Significance result.
- Control and Test percentages are displayed as rounded whole percentages with no decimals.
- On-screen norm tables use a BLS-style white table with black text.
- Each metric includes one Base size row labeled with the denominator setting; response-option rows do not include selected-count columns.
- Response-option counting follows the Smart Tables matching pattern: exact response match first, then common multi-select delimiters such as `;` and `,`.
- If the workbook includes a recognizable response-label sheet, the app auto-detects it and maps coded values to readable response labels.
- The response labels sheet is expected to have one row per response option, with columns for question/variable, response value/code, and response label.
- If available, a question label column can also display readable question text while preserving the source column name.
- Control and test samples are selected from one uploaded group-label column.
- Control % = selected response option count within the control group divided by the selected control denominator, displayed as a rounded whole percentage with no decimals.
- Test % = selected response option count within the test group divided by the selected test denominator, displayed as a rounded whole percentage with no decimals.
- Lift = Test % minus Control %, displayed as rounded whole percentage points with no decimals.
- Default denominator = Total answering.
- Total answering means only people in the selected group who answered that specific norm question are included in the denominator. This matches the Smart Tables Total Answering section.
- Total sample means all people in the selected group are included in the denominator. This matches the Smart Tables Total Base section.
- Denominator decisions are stored in `denominator_settings.json` and reused later.
- Denominator changes saved in the app are tracked in `CHANGELOG.md` and reflected in `status.md`.
- Saving to the norms database is an explicit action on the Norm tables page after setup and review.
- Saved norm datasets are stored in persistent app-level storage, with one workbook per saved dataset and an aggregate `saved_norm_tables.xlsx` workbook.
- By default, saved norms live under the app folder at `norm_database/`, so restarting Streamlit from a different working directory does not make the database look empty.
- Each saved dataset is also backed up under `uploaded_datasets/`: raw uploads go to `raw_uploads/`, app-ready saved norm workbooks go to `norm_workbooks/`, and manifest/rule backups go to `norm_settings/`.
- Saved dataset changes create restore points under `uploaded_datasets/norm_history/`, including an initial baseline plus saves, replacements, rule edits, deletions, and restores.
- Saved datasets can be removed from the Saved datasets page, and previous states can be restored from version history.
- On startup or refresh, if the working `norm_database/` is missing or partial, the app restores missing manifest and dataset files from `uploaded_datasets/`.
- On Streamlit Cloud, GitHub autocommit must be configured to save `uploaded_datasets/` back to the repository after each dataset save, edit, delete, or restore. Without it, Streamlit's filesystem can reset and recent saved datasets may disappear after an app restart.
- Configure GitHub autocommit in Streamlit secrets with a token that has repository Contents read/write access:

```toml
[github_autocommit]
enabled = true
repo = "OWNER/REPO"
branch = "main"
token = "github_pat_or_classic_token"
data_path = "uploaded_datasets"
```

- If GitHub autocommit is not configured, files written to `uploaded_datasets/` still need to be committed or otherwise synced to GitHub to survive redeploys.
- Set `BLS_NORMS_DATA_DIR` to an absolute shared/persistent directory when hosting the app for multiple devices or on a server with mounted persistent storage.
- Saves and saved-rule updates use a database write lock and atomic manifest writes so overlapping browser sessions do not overwrite each other.
- Duplicate-upload safeguards use respondent ID overlap when an ID field such as `ResponseId` is available; 80% or higher overlap is flagged as a possible duplicate.
- If the app cannot detect the respondent ID field, users can select the respondent ID variable during upload. Saved datasets can also be edited to set or change the respondent ID variable.
- If no respondent ID field is available, duplicate safeguards fall back to exact cleaned-data matching rather than file name.
- Duplicate datasets are flagged during workbook upload with a simple Dataset already added to norms notification.
- The Saved datasets page can edit saved dataset rules and regenerate saved norm tables when standards change.
- Newly saved dataset workbooks store cleaned respondent data and rule sheets so previous calculations can be rebuilt.
- The Norm tables page rebuilds saved norms from the stored respondent-level workbooks as one appended read. Filters apply to respondent rows first, then control/test denominators, counts, percentages, lift, and significance are recalculated from the combined data.
- Significance uses a two-proportion z-test with a 0.05 threshold.
- Significance output is limited to Significant, Not significant, or Not tested.
- Rows with missing group labels, missing sample sizes, invalid denominators, or unavailable percentage logic are shown as Not tested and/or Not available rather than forcing a result.
- Norm tables can be downloaded as an Excel workbook with an All Norms sheet plus one sheet per mapped question.
