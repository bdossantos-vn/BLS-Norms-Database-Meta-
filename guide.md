# Guide

## Norm Tables

Norms are calculated at the response-option level. For every mapped norm question, the app creates one row per response option found in the control or test sample.

Norm tables use the audited norm/benchmark variable label as the table heading. The table itself is intentionally compact: Response option, Control, Test, Lift, and Significance result. Control/Test percentages and Lift are rounded to whole points with no decimals.

On-screen norm tables use a BLS-style white table with black text for readability.

The Norm tables page shows the saved norms database. When a workbook is active, that upload is used for the save-to-database step and does not replace the saved database view.

Each metric includes one Base size row labeled with the denominator setting. Response-option rows show Control and Test as percentages; the app does not show selected-count columns for each response option.

Use the filter buttons above the Norm tables to limit the saved database rows being reviewed. Available filters are Project, Brand, Client, Industry, Country, Year, Quarter, Gender, and Age when those fields exist in the saved database view. Metadata filters use exact variable or label matches, so the Brand filter uses the Brand field rather than metric columns such as brand sentiment.

The Reset to total control vs test button clears all filter selections and returns to all saved datasets.

## Excel Input

Use the Survey Question Audit page to upload an Excel workbook, choose control/test labels, and map questions to norms or benchmarks. Intake controls are kept on that page instead of the sidebar so workbook assumptions and question decisions can be reviewed together.

Use the Back and Next page buttons at the top and bottom of each page to move through the workflow.

Use the saved-norm review option when you want to inspect existing saved norms without uploading a workbook. After selecting it, use Next to proceed to Norm tables. Review mode keeps Back/Next navigation focused on Norm tables and Saved datasets.

The selected respondent data sheet is interpreted as respondent-level data:

- Each row is one respondent.
- Each column is one survey question or respondent field.

The default layout is the BLS / Smart Tables layout: the first row contains variable names, the second row contains question labels, and subsequent rows contain respondent-level data. Obvious metadata rows are removed before calculation.

After upload, select the sheet containing respondent-level data. The app shows respondent row/column counts before norms are calculated. Raw respondent rows are not previewed in the app.

Standard survey platform/system columns such as StartDate, EndDate, IPAddress, Progress, Duration, and ResponseId remain visible in the audit table for review, but they default to `NA` and are excluded from norm calculations unless the user changes them.

The app checks for project metadata fields: brand, industry, client, quarter, year, methodology, project, country, and c_key. If any are missing, the Survey Question Audit page prompts for a value. Entered values are added to the working dataset for that workbook and excluded from norm-question audit rows. They are not included in norm table output.

## Survey Question Audit

The audit table is modeled after the BLS Smart Tables Survey Question Audit. Each row is an uploaded source question.

The app fills Suggested norm/benchmark using the closest available norm. For the first project, the suggestion defaults to the source variable name. Use the Norm / benchmark dropdown to switch the mapping when needed. Check NA to exclude a question from the norms database.

Question Type follows the BLS Smart Tables audit categories: Single-Select, Multi-Select, Scale / Likert, Numeric Data, Open-End Text, and Ignore. The app detects a starting type, and the audit dropdown can be edited when needed.

For scale questions, use the T2B, T3B, B2B, and B3B columns to use top-box or bottom-box rows in the norm output. These rows use the ordered response choices shown in the audit table.

The buttons above the audit table can apply one box-score choice to all non-NA Scale / Likert rows, or clear box-score choices from those rows.

Saved norm mapping decisions are stored in `norm_mapping_settings.json` and reused later. Saved NA decisions also store normalized variable/question-text aliases in `na_alias_settings.json`, so a matching question in a future upload starts as NA. Saved question type decisions are stored in `question_type_settings.json`. Saved box-score decisions are stored in `box_score_settings.json`.

The audit also reads saved dataset rules. When a new upload has a matching source variable, question label, or norm/benchmark label, prior T2B/T3B/B2B/B3B choices are reused. If the current audit changes those choices, the app flags the mismatch, asks for confirmation, and recommends updating previous saved datasets so the metric stays consistently measured.

If a saved or suggested norm/benchmark name already includes T2B, T3B, B2B, or B3B, the app treats that as a prior audit signal and checks the matching box-score option by default. A saved manual audit decision takes precedence.

If responses are coded and the workbook includes a recognizable response-label sheet, the app auto-detects it. The labels sheet should contain one row per response option and columns for:

- Question or variable name matching the data sheet column.
- Response value or code matching the coded data value.
- Response label to show in norm tables.

The app also supports an optional question label column for audit review and scale detection. Norm table output uses the audited norm/benchmark variable label rather than QNR question text.

If no labels sheet is auto-detected, the app uses the values found in the data sheet as the response options.

Response-option counting follows the Smart Tables matching pattern: the app first checks for an exact response match, then checks common multi-select delimiters such as `;` and `,`.

Scale-like response options use the Smart Tables ordering logic where possible, so labels such as Strongly agree, Somewhat agree, Neutral, Somewhat disagree, and Strongly disagree appear in a consistent positive-to-negative order. Exclusive options such as None, Other, Don't know, and Prefer not to say are kept at the end.

When a question is saved as Scale / Likert, that saved type is reused to preserve scale ordering and box-score base order in future audits.

T2B and T3B are calculated from the first two or three ordered non-exclusive options. B2B and B3B are calculated from the last two or three ordered non-exclusive options. For example, on a five-point agree scale ordered positive to negative, T2B is Strongly agree plus Somewhat agree, and B2B is Somewhat disagree plus Strongly disagree.

When any box-score option is selected for a question, the norm table shows only the selected box-score rows for that question. It does not also show the full five-point scale.

## Saving Norm Datasets

After the Survey Question Audit and denominator settings are reviewed, use Save dataset to norms database on the Norm tables page. This is an explicit save step so uploads are not silently added to the database.

The save step uses the full audited upload, not temporary Norm table filters. By default, it writes one dataset workbook under the app folder at `norm_database/datasets/`, updates `norm_database/manifest.json`, and refreshes the aggregate workbook at `norm_database/saved_norm_tables.xlsx`.

For shared deployments, set `BLS_NORMS_DATA_DIR` to an absolute shared or mounted persistent directory. This keeps saved datasets available after app restarts and lets users on different devices write to the same norms database. Saves and saved-rule updates use a write lock and atomic manifest writes to protect the database from overlapping browser sessions.

The app checks each upload against saved respondent IDs when an ID field such as ResponseId is available. If 80% or more of unique respondent IDs overlap with a saved dataset in either direction, the upload page shows Dataset already added to norms and the save step requires Replace saved dataset instead of creating a duplicate. If no respondent ID field is available, the app falls back to exact cleaned-data matching rather than file name.

Use the Saved datasets page to edit previous dataset rules if standards change. The editor can update whether a source variable is included, its norm/benchmark mapping, denominator, question type, T2B/T3B/B2B/B3B selections, and the control/test setup. Saving updated rules regenerates that dataset workbook and the aggregate saved norms workbook.

Datasets saved before this editable-rule infrastructure may need to be replaced from the original upload before their rules can be edited, because older saved workbooks did not include cleaned respondent data and rule sheets.

## Denominator Choices

Each norm question has its own denominator setting.

- Total answering: percentages are based only on people in the selected group who answered that specific question. This matches the Smart Tables Total Answering section.
- Total sample: percentages are based on all people in the selected group. This matches the Smart Tables Total Base section.

The Denominator settings page uses variable and mapped norm labels, not the QNR question text, so the labels match the audit mapping.

The default is Total answering. Saved denominator decisions are reused in later runs through `denominator_settings.json`.

If the app cannot apply the selected denominator because data is missing, percentages and lift are shown as Not available.

## Significance Method

The app compares the control percentage and test percentage with a two-proportion z-test. In plain language, this checks whether the difference between two percentages is larger than expected from normal sample variation.

The output is intentionally simple:

- Significant: the difference passes the statistical test.
- Not significant: the result remained stable.
- Not tested: the app did not have the sample sizes, group labels, or valid denominators needed to run the test.
