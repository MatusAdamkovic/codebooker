# Codebooker

Turn a survey export into a SCORE-style codebook — in the browser, on any
operating system, with no installation required for end users.

- **Input:** Qualtrics `.qsf` (recommended), SPSS `.sav` dictionaries/data,
  PsyToolkit syntax (`.txt`/`.md`), questionnaire or Qualtrics text exports
  (`.docx`/`.txt`), and/or tabular data files (`.csv`/`.xlsx`) whose column
  structure is reconciled with the survey.
- **Output:** an editable SCORE-format codebook (Variable, Variable name,
  Measurement/unit, Values, Description, Note) exported as CSV, XLSX, or
  Markdown.
- **AI refinement (optional):** users can plug in their own Anthropic or
  OpenAI API key and choose from a curated model dropdown. The deterministic
  draft requires no key at all — files never leave the browser during that
  step.

## For end users (researchers)

Nothing to install. Open the hosted URL in any modern browser on Windows,
macOS, or Linux. Upload a survey export, click "Build draft codebook",
optionally refine with AI, edit, and download.

## Supported sources

- **Best source:** Qualtrics `.qsf`, because it preserves export tags, question
  types, matrix structure, and response options.
- **SPSS `.sav`:** reads the file dictionary in-browser, including long
  variable names, variable labels, value labels, and basic measurement hints.
  When paired with a questionnaire `.docx`, the richer `.sav` metadata is used
  as the codebook backbone.
- **Questionnaire `.docx` / `.txt`:** groups likely questions with nearby answer
  choices. If no export tags are present, generated variable names are used.
- **PsyToolkit `.txt` / `.md`:** reads labels, prompts, item lists, and scale
  definitions.
- **Tabular `.csv` / `.xlsx`:** creates one row per column from observed data
  and merges with matching survey metadata when possible.

## How it works

Codebooker has two stages: a deterministic browser-only draft, then an
optional AI cleanup pass.

### What you need

For the best result, upload at least one survey-definition file and, if
available, the dataset exported from the same survey:

- **Survey definition:** ideally Qualtrics `.qsf`; otherwise questionnaire
  `.docx`/`.txt`, PsyToolkit syntax, or a similar text export.
- **Data or dictionary:** `.sav`, `.csv`, or `.xlsx`.
- **Optional setting:** enable full wording if you want item wording preserved
  in the draft. Leave it off when you prefer shorter descriptions or need to
  avoid copying copyrighted item text.

You can upload only one file, but paired sources are better. For example, a
`.qsf` explains what each variable means, while a `.sav` or `.csv` confirms
which columns actually appear in the dataset.

### How information is extracted

Everything in the draft stage runs locally in the browser:

- **Qualtrics `.qsf`:** reads question IDs/export tags, question types, matrix
  statements, sliders, text-entry fields, multiple-choice options, and response
  scales. Descriptive text blocks are skipped because they are not variables.
- **SPSS `.sav`:** reads the SPSS dictionary metadata, not just raw values:
  long variable names, variable labels, value labels, text/numeric type, and
  simple measurement hints.
- **Questionnaire `.docx`/`.txt`:** extracts raw text, detects likely numbered
  questions, and groups nearby answer choices with each question. If no export
  tags are available, Codebooker generates provisional variable names.
- **PsyToolkit syntax:** reads labels, prompts, item lists, and scale
  definitions where they are present in the syntax file.
- **CSV/XLSX data:** reads column headers and samples observed values from
  each column to infer simple measurement types and examples.

Each parser turns its source into SCORE-style draft rows with the same six
fields: `Variable`, `Variable name`, `Measurement/unit`, `Values`,
`Description`, and `Note`.

### How survey exports and data are combined

After parsing, Codebooker reconciles rows from all uploaded files:

- Rows are grouped mainly by normalized variable name. If a questionnaire has
  weak generated names but the text label matches another source, matching
  labels can also be used.
- When survey metadata and data evidence match, the row is merged. The survey
  source usually supplies the variable name, response options, and description;
  the data source supplies observed structure, measurement hints, and notes.
- `.sav` dictionary metadata is treated as strong data metadata. When a `.sav`
  is paired with a questionnaire `.docx`, the `.sav` variables form the
  backbone and questionnaire text is used only where it can be matched.
- CSV/XLSX-only columns still become draft rows, but they are marked as
  auto-generated from tabular structure so they can be reviewed.
- The final draft preserves the first-appearance order as much as possible and
  records merge decisions in the `Note` field.

This merge step is intentionally conservative: it prefers keeping uncertain
variables visible with notes over silently deleting or inventing structure.

### How AI refinement works

AI refinement is optional and starts only after the deterministic draft exists.
The app sends the selected provider a compact prompt containing:

- Short summaries of the uploaded files, such as variable names, inferred
  measurement types, and example values.
- The draft codebook rows as JSON.
- Instructions to return only SCORE-format JSON, preserve raw variable names,
  avoid inventing response options or skip logic, and put uncertainty in
  `Note`.

The model can improve labels, shorten or clarify descriptions, harmonize value
labels, and optionally condense obvious item batteries into a single scale row.
If a model response is too long or invalid, Codebooker retries with smaller
batches. If a single row still fails, the deterministic draft row is kept and
marked in the note.

API keys are not stored by the app. On the hosted GitHub Pages version, keys
are kept only in browser memory and sent directly from the browser to
Anthropic or OpenAI. The built-in keyless Claude option works only when the app
is running inside Claude.

## AI model choices

The model field is a dropdown rather than free text. Current choices are:

- **Built-in Claude (inside Claude only):** Claude Sonnet 5.
- **Anthropic API:** Claude Sonnet 5, Claude Haiku 4.5, Claude Opus 4.8,
  Claude Fable 5.
- **OpenAI API:** GPT-5.4 mini, GPT-5.5, GPT-5.4, GPT-5.4 nano.

Model names were refreshed against the provider docs in July 2026. If a
provider changes access or availability, update the `PROVIDERS` list in
`src/codebooker.jsx`.

## For the maintainer: run locally

Requires only [Node.js](https://nodejs.org) (LTS). Works identically on
Windows (PowerShell/CMD) and macOS/Linux (Terminal):

```bash
npm install
npm run dev
```

Then open the printed `http://localhost:5173` address.

## Deploy as a website (recommended distribution)

```bash
npm run build
```

This produces a `dist/` folder of plain static files (no server code, no
database). Host it anywhere:

- **GitHub Pages** — push the repo, enable Pages, publish the `dist/` output
  (e.g., with the `actions/deploy-pages` workflow or the `gh-pages` package).
- **Netlify / Vercel / Cloudflare Pages** — drag-and-drop the `dist/` folder
  or connect the repo; all have free tiers.
- **A university web server** — copy the contents of `dist/` into any
  directory; `base: "./"` in `vite.config.js` means it works from subfolders.

## Notes

- The "Built-in Claude (no key)" provider only functions when the app runs
  inside Claude as an artifact; on a standalone deployment users should
  select the Anthropic or OpenAI provider and enter their own key. Keys are
  held in memory only and sent directly from the browser to the provider.
  OpenAI refinement uses the Responses API for the current GPT model family.
- PDF survey exports are not parsed in-browser; export as `.docx`, `.txt`,
  or ideally `.qsf` instead.
- Always review the generated codebook before sharing: copyrighted item
  wording, matrix questions, branching logic, and derived variables need
  human judgment.

## Automatic deployment (GitHub Pages)

This repo ships with `.github/workflows/deploy.yml`. One-time setup:

1. Push this project to a GitHub repository (branch `main`).
2. In the repo: Settings → Pages → Source: **GitHub Actions**.

Every push to `main` then builds the app and publishes it at
`https://<username>.github.io/<repo>/` automatically.
