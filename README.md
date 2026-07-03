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
