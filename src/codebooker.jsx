import React, { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import mammoth from "mammoth";
import {
  Upload, FileText, Table2, Sparkles, Download, RotateCcw,
  Trash2, Plus, X, BookOpen, ChevronDown, ChevronUp, FlaskConical
} from "lucide-react";

/* ============================================================
   SCORE structure
   ============================================================ */
const SCORE_COLUMNS = [
  "Variable",
  "Variable name",
  "Measurement/unit",
  "Values",
  "Description",
  "Note",
];

const emptyRow = () =>
  Object.fromEntries(SCORE_COLUMNS.map((c) => [c, ""]));

/* ============================================================
   Small utilities
   ============================================================ */
const stripMarkup = (t) =>
  String(t || "")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/\s+/g, " ")
    .trim();

const normalizeId = (v) => String(v || "").toLowerCase().replace(/[^a-z0-9]+/g, "");

const prettify = (id) => {
  let s = String(id || "")
    .replace(/(?<=[a-z0-9])(?=[A-Z])/g, " ")
    .replace(/(?<=[A-Z])(?=[A-Z][a-z])/g, " ")
    .replace(/[_\-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!s) return id;
  return s[0].toUpperCase() + s.slice(1);
};

const firstSentence = (text) => {
  const cleaned = stripMarkup(text);
  if (!cleaned) return "";
  let s = cleaned.split(/(?<=[.!?])\s+/)[0].trim();
  if (s.length > 140) {
    s = s.split(/\s+/).slice(0, 18).join(" ").replace(/[,;:]+$/, "") + "…";
  }
  return s;
};

const truncate = (t, n) => {
  const c = stripMarkup(t);
  return c.length <= n ? c : c.slice(0, n - 1).trimEnd() + "…";
};

const inferMeasurement = (values) => {
  const nonNull = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
  if (!nonNull.length) return "Unspecified";
  const asNum = nonNull.map((v) => Number(v));
  if (asNum.every((n) => !Number.isNaN(n))) {
    const uniq = new Set(asNum).size;
    if (uniq <= 2 && asNum.every((n) => n === 0 || n === 1)) return "Binary (0/1)";
    return uniq > 10 ? "Numeric" : "Categorical numeric";
  }
  const uniq = new Set(nonNull.map(String)).size;
  if (uniq <= 2) return "Binary / categorical";
  return "Text / categorical";
};

const mkRow = (o) => ({
  Variable: o.variable || "",
  "Variable name": o.name || "",
  "Measurement/unit": o.measurement || "",
  Values: o.values || "",
  Description: o.description || "",
  Note: o.note || "",
  _role: o.role || "",
  _origin: o.origin || "",
});

/* ============================================================
   Parser: Qualtrics QSF (JSON survey export) — most reliable
   ============================================================ */
function looksLikeQSF(text) {
  try {
    const j = JSON.parse(text);
    return Array.isArray(j.SurveyElements);
  } catch {
    return false;
  }
}

function parseQSF(name, text, includeWording) {
  const doc = JSON.parse(text);
  const rows = [];
  const skipped = [];
  for (const el of doc.SurveyElements || []) {
    if (el.Element !== "SQ") continue;
    const p = el.Payload || {};
    const qtype = p.QuestionType || "";
    if (qtype === "DB") continue; // descriptive text blocks, not variables
    const tag = p.DataExportTag || p.QuestionID || "";
    const qtext = stripMarkup(p.QuestionText || "");
    const selector = p.Selector || "";

    const fmtChoices = (obj) =>
      obj
        ? Object.entries(obj).map(([k, v]) => `${k}=${stripMarkup(v.Display || "")}`)
        : [];
    const choices = fmtChoices(p.Choices);
    const answers = fmtChoices(p.Answers);

    const desc = (extra) =>
      includeWording
        ? [qtext, extra].filter(Boolean).join(" — ")
        : truncate(extra || qtext, 90);

    if (qtype === "Matrix") {
      // Each statement (Choices) becomes a sub-variable; Answers hold the scale.
      for (const [k, v] of Object.entries(p.Choices || {})) {
        const stmt = stripMarkup(v.Display || "");
        rows.push(
          mkRow({
            variable: firstSentence(stmt || qtext) || prettify(`${tag}_${k}`),
            name: `${tag}_${k}`,
            measurement: "Ordinal (Likert-type)",
            values: answers.join(", "),
            description: desc(stmt),
            note: `Matrix question \`${tag}\` from Qualtrics QSF; statement ${k}.`,
            role: "survey",
            origin: name,
          })
        );
      }
    } else if (qtype === "TE") {
      rows.push(
        mkRow({
          variable: firstSentence(qtext) || prettify(tag),
          name: tag,
          measurement: "Text (open entry)",
          values: "Free text",
          description: desc(""),
          note: "Text-entry question from Qualtrics QSF.",
          role: "survey",
          origin: name,
        })
      );
    } else if (qtype === "Slider") {
      const min = p.Configuration?.CSSliderMin;
      const max = p.Configuration?.CSSliderMax;
      rows.push(
        mkRow({
          variable: firstSentence(qtext) || prettify(tag),
          name: tag,
          measurement: "Numeric (slider)",
          values: min !== undefined && max !== undefined ? `${min}–${max}` : "",
          description: desc(""),
          note: "Slider question from Qualtrics QSF.",
          role: "survey",
          origin: name,
        })
      );
    } else if (qtype === "MC") {
      const multi = /MAVR|MAHR|MACOL|MSB/.test(selector);
      rows.push(
        mkRow({
          variable: firstSentence(qtext) || prettify(tag),
          name: tag,
          measurement: multi
            ? "Categorical (multiple selection)"
            : "Categorical (single choice)",
          values: choices.join(", "),
          description: desc(""),
          note: multi
            ? `Multiple-answer question \`${tag}\`; data export typically creates one binary column per option.`
            : "Multiple-choice question from Qualtrics QSF.",
          role: "survey",
          origin: name,
        })
      );
    } else {
      rows.push(
        mkRow({
          variable: firstSentence(qtext) || prettify(tag),
          name: tag,
          measurement: "Unspecified",
          values: [...choices, ...answers].join(", "),
          description: desc(""),
          note: `Question type \`${qtype}/${selector}\` from Qualtrics QSF; review manually.`,
          role: "survey",
          origin: name,
        })
      );
      skipped.push(qtype);
    }
  }
  return rows;
}

/* ============================================================
   Parser: PsyToolkit survey syntax
   ============================================================ */
function parsePsytoolkit(name, text, includeWording) {
  const scaleMap = {};
  let currentScale = null;
  const blocks = [];
  let cur = null;

  for (const raw of text.split("\n")) {
    const line = raw.trim();
    if (!line) continue;
    const sm = line.match(/^scale:\s*([A-Za-z0-9_]+)/);
    if (sm) {
      currentScale = sm[1];
      scaleMap[currentScale] = [];
      continue;
    }
    if (currentScale && line.startsWith("- ")) {
      scaleMap[currentScale].push(line.slice(2).trim());
      continue;
    }
    if (/^l:\s*[A-Za-z0-9_]+/.test(line)) {
      currentScale = null;
      if (cur) blocks.push(cur);
      cur = { label: line.replace(/^l:\s*/, "").trim(), type: "", prompt: "", items: [] };
      continue;
    }
    if (cur && /^t:\s*/.test(line)) { cur.type = line.replace(/^t:\s*/, "").trim(); continue; }
    if (cur && /^q:\s*/.test(line)) { cur.prompt = line.replace(/^q:\s*/, "").trim(); continue; }
    if (cur && line.startsWith("- ")) { cur.items.push(line.slice(2).trim()); continue; }
    if (cur && cur.prompt) cur.prompt = `${cur.prompt} ${line}`.trim();
  }
  if (cur) blocks.push(cur);

  const rows = [];
  for (const b of blocks) {
    const prompt = stripMarkup(b.prompt);
    let values = "";
    let scaleName = "";
    if (b.type.startsWith("scale ")) {
      scaleName = b.type.slice(6).trim();
      values = (scaleMap[scaleName] || []).map((v, i) => `${i + 1}=${v}`).join(", ");
    }
    const meas = values ? "Ordinal (Likert-type)" : "Categorical";
    if (b.items.length) {
      b.items.forEach((item, i) => {
        const id = `${b.label}_${i + 1}`;
        const description = [
          includeWording ? prompt : "",
          includeWording ? item : "",
          scaleName ? `Response scale: ${scaleName}.` : "",
        ].filter(Boolean).join(" ").trim();
        rows.push(
          mkRow({
            variable: firstSentence(includeWording ? item : prompt) || prettify(id),
            name: id,
            measurement: meas,
            values,
            description,
            note: `Item ${i + 1} of PsyToolkit block \`${b.label}\`.`,
            role: "survey",
            origin: name,
          })
        );
      });
    } else {
      const description = [
        includeWording ? prompt : "",
        scaleName ? `Response scale: ${scaleName}.` : "",
      ].filter(Boolean).join(" ").trim();
      rows.push(
        mkRow({
          variable: firstSentence(prompt) || prettify(b.label),
          name: b.label,
          measurement: meas,
          values,
          description,
          note: "Derived from PsyToolkit syntax.",
          role: "survey",
          origin: name,
        })
      );
    }
  }
  return rows;
}

/* ============================================================
   Parser: Qualtrics printed/text export (heuristic fallback)
   ============================================================ */
const looksLikeVariableId = (token) => {
  if (["start", "strongly", "agree", "disagree", "neither"].includes(token.toLowerCase())) return false;
  if (token.includes("_")) return true;
  if (/^[a-z]/.test(token)) return true;
  if (token === token.toUpperCase()) return false;
  const rest = token.slice(1);
  return /[A-Z]/.test(rest) && /[a-z]/.test(rest);
};

const looksLikeOptionLabel = (label) => {
  const words = label.split(/\s+/).filter(Boolean);
  return words.length > 0 && words.length <= 7;
};

function parseQualtricsText(name, text, includeWording) {
  const rows = [];
  let currentId = null;
  let currentQuestion = "";
  let options = [];
  let optionCodes = [];

  const flush = () => {
    if (!currentId) return;
    rows.push(
      mkRow({
        variable: firstSentence(currentQuestion) || prettify(currentId),
        name: currentId,
        measurement: options.length ? "Categorical" : "Text / numeric",
        values: options.join(", "),
        description: includeWording ? currentQuestion : truncate(currentQuestion, 90),
        note: "Derived from survey export text; review matrix items and labels.",
        role: "survey",
        origin: name,
      })
    );
    currentId = null;
    currentQuestion = "";
    options = [];
    optionCodes = [];
  };

  for (const raw of text.split("\n")) {
    const line = raw.replace(/\u2028/g, " ").replace(/\s+/g, " ").trim();
    if (!line) continue;
    const low = line.toLowerCase();
    if (low.startsWith("start of block:") || low.startsWith("end of block:") || low.startsWith("koniec bloku:")) {
      flush();
      continue;
    }

    const coded = line.match(/(.+?)\s+\(([^)]+)\)\s*$/);
    if (coded) {
      const label = coded[1].replace(/•/g, "").trim();
      const code = coded[2].trim();
      const numeric = /^\d+$/.test(code) ? parseInt(code, 10) : null;
      const descendingIntoItems =
        optionCodes.length && numeric !== null && numeric <= Math.min(...optionCodes);
      if (!descendingIntoItems && (line.startsWith("•") || looksLikeOptionLabel(label))) {
        options.push(`${code}=${label}`);
        if (numeric !== null) optionCodes.push(numeric);
        continue;
      }
    }

    const qm = line.match(/^([A-Za-z][A-Za-z0-9_]+)\s+(.+)$/);
    if (qm && !line.startsWith("http") && !line.startsWith("www")) {
      if (looksLikeVariableId(qm[1])) {
        flush();
        currentId = qm[1];
        currentQuestion = qm[2].trim();
        continue;
      }
    }

    const matrixItem = line.match(/^(.+?)\s+\((\d+)\)\s*$/);
    if (currentId && matrixItem && matrixItem[1].split(/\s+/).length > 2) {
      const itemText = matrixItem[1].trim();
      const itemId = `${currentId}_${matrixItem[2]}`;
      rows.push(
        mkRow({
          variable: firstSentence(itemText) || prettify(itemId),
          name: itemId,
          measurement: options.length ? "Ordinal (Likert-type)" : "Categorical",
          values: options.join(", "),
          description: includeWording ? itemText : currentQuestion,
          note: `Generated from matrix-style question \`${currentId}\`.`,
          role: "survey",
          origin: name,
        })
      );
      continue;
    }

    if (currentId && !line.startsWith("•")) {
      currentQuestion = `${currentQuestion} ${line}`.trim();
    }
  }
  flush();
  return rows;
}

/* ============================================================
   Parser: generic text fallback
   ============================================================ */
function parseGenericText(name, text) {
  const paragraphs = text.split("\n").map((l) => l.trim()).filter(Boolean);
  return paragraphs.slice(0, 20).map((p, i) =>
    mkRow({
      variable: firstSentence(p) || `Field ${i + 1}`,
      name: `field_${i + 1}`,
      measurement: "Unspecified",
      values: "",
      description: truncate(p, 240),
      note: "Fallback extraction from plain text; replace with exact metadata.",
      role: "survey",
      origin: name,
    })
  );
}

function parseSurveyText(name, text, includeWording) {
  const cleaned = text.replace(/\r/g, "");
  const lower = cleaned.toLowerCase();
  let rows = [];
  if (lower.includes("scale:") && /^\s*l:\s*/m.test(cleaned)) {
    rows = rows.concat(parsePsytoolkit(name, cleaned, includeWording));
  }
  if (lower.includes("start of block:")) {
    rows = rows.concat(parseQualtricsText(name, cleaned, includeWording));
  }
  if (!rows.length) rows = parseGenericText(name, cleaned);
  return rows;
}

/* ============================================================
   Parser: tabular data (one row per column)
   ============================================================ */
function parseTable(originName, headers, records) {
  return headers.map((h) => {
    const values = records.map((r) => r[h]);
    const nonNull = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
    const examples = [...new Set(nonNull.map(String))].slice(0, 8).join(", ");
    return mkRow({
      variable: `${prettify(String(h))} observed in uploaded data`,
      name: String(h),
      measurement: inferMeasurement(values),
      values: examples,
      description: `Values observed in the uploaded dataset for \`${h}\`.`,
      note: "Auto-generated from tabular structure; review wording and value labels.",
      role: "data",
      origin: originName,
    });
  });
}

/* ============================================================
   Merge / reconciliation layer
   ============================================================ */
function mergeRows(rows) {
  const groups = new Map();
  for (const r of rows) {
    const key = normalizeId(r["Variable name"]);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  }

  const merged = [];
  for (const key of groups.keys()) {
    const g = groups.get(key);
    const surveyRows = g.filter((r) => r._role === "survey");
    const dataRows = g.filter((r) => r._role === "data");

    const name =
      surveyRows.find((r) => r["Variable name"])?.["Variable name"] ||
      g.reduce((a, b) => (b["Variable name"].length > (a?.length || 0) ? b["Variable name"] : a), "") ||
      "variable";

    const descCand = [...surveyRows, ...dataRows, ...g].find((r) => {
      const d = r.Description.trim();
      return d && !d.toLowerCase().startsWith("values observed");
    });
    const description = (descCand || g.find((r) => r.Description.trim()) || { Description: "" }).Description.trim();

    const values =
      surveyRows.find((r) => r.Values.trim())?.Values.trim() ||
      g.find((r) => r.Values.trim())?.Values.trim() ||
      "";

    const measurement =
      dataRows.find((r) => r["Measurement/unit"])?.["Measurement/unit"] ||
      g.find((r) => r["Measurement/unit"])?.["Measurement/unit"] ||
      "Unspecified";

    const notes = [];
    const seen = new Set();
    if (surveyRows.length && dataRows.length) {
      notes.push("Merged survey metadata with observed data structure.");
    }
    for (const r of g) {
      const n = r.Note.trim();
      if (n && !seen.has(n)) { seen.add(n); notes.push(n); }
    }

    const variable =
      firstSentence(description) ||
      (dataRows.length ? `${prettify(name)} observed in uploaded data` : prettify(name));

    merged.push(
      mkRow({
        variable,
        name,
        measurement,
        values,
        description,
        note: notes.join(" "),
        role: surveyRows.length && dataRows.length ? "both" : g[0]._role,
      })
    );
  }
  // Preserve first-appearance order of variables
  const order = new Map();
  rows.forEach((r, i) => {
    const k = normalizeId(r["Variable name"]);
    if (!order.has(k)) order.set(k, i);
  });
  merged.sort(
    (a, b) => order.get(normalizeId(a["Variable name"])) - order.get(normalizeId(b["Variable name"]))
  );
  return merged;
}

/* ============================================================
   File reading
   ============================================================ */
async function readUpload(file, roleOverride, includeWording) {
  const ext = (file.name.split(".").pop() || "").toLowerCase();
  const notes = [];
  let rows = [];
  let role = roleOverride;
  let summary = "";

  const readText = () => file.text();
  const readBuffer = () => file.arrayBuffer();

  try {
    if (ext === "qsf" || ext === "json") {
      const text = await readText();
      if (looksLikeQSF(text)) {
        role = role || "survey";
        rows = parseQSF(file.name, text, includeWording);
        summary = truncate(text, 6000);
      } else {
        notes.push(`${file.name}: JSON file is not a Qualtrics QSF export; skipped.`);
        role = role || "unknown";
      }
    } else if (ext === "csv" || ext === "tsv") {
      const text = await readText();
      const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
      const headers = parsed.meta.fields || [];
      role = role || "data";
      rows = parseTable(file.name, headers, parsed.data);
      summary = summarizeTable(file.name, headers, parsed.data);
    } else if (ext === "xlsx" || ext === "xlsm" || ext === "xls") {
      const buf = await readBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      role = role || "data";
      const parts = [];
      for (const sheetName of wb.SheetNames) {
        const ws = wb.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (!json.length) continue;
        const headers = Object.keys(json[0]);
        rows = rows.concat(parseTable(`${file.name}::${sheetName}`, headers, json));
        parts.push(summarizeTable(`${file.name}::${sheetName}`, headers, json));
      }
      summary = parts.join("\n\n");
    } else if (ext === "docx") {
      const buf = await readBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer: buf });
      role = role || "survey";
      rows = parseSurveyText(file.name, result.value || "", includeWording);
      summary = truncate(result.value || "", 6000);
    } else if (ext === "txt" || ext === "md") {
      const text = await readText();
      role = role || "survey";
      rows = parseSurveyText(file.name, text, includeWording);
      summary = truncate(text, 6000);
    } else if (ext === "pdf") {
      notes.push(
        `${file.name}: PDF parsing isn't supported in this in-browser version — export the survey as .docx, .txt, or (best) .qsf instead.`
      );
      role = role || "survey";
    } else {
      notes.push(`${file.name}: unsupported file type.`);
      role = role || "unknown";
    }
  } catch (err) {
    notes.push(`Could not parse ${file.name}: ${err.message}`);
  }

  // If the user overrides a tabular file to "survey", treat its columns as survey-defined variables
  if (roleOverride === "survey") {
    rows = rows.map((r) => ({ ...r, _role: "survey" }));
  }
  if (roleOverride === "data") {
    rows = rows.map((r) => ({ ...r, _role: "data" }));
  }

  return { rows, notes, role: role || "unknown", summary };
}

function summarizeTable(label, headers, records) {
  const lines = [`Table: ${label}`, `Columns: ${headers.length}`];
  for (const h of headers.slice(0, 40)) {
    const values = records.map((r) => r[h]);
    const nonNull = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
    const ex = [...new Set(nonNull.map(String))].slice(0, 6).join(", ");
    lines.push(`- ${h} | ${inferMeasurement(values)} | e.g. ${ex}`);
  }
  if (headers.length > 40) lines.push(`… ${headers.length - 40} more columns`);
  return lines.join("\n");
}

/* ============================================================
   Claude refinement (no API key needed inside Claude artifacts)
   ============================================================ */
const AI_SYSTEM = `You create high-quality SCORE-style codebook rows for survey research.

Rules:
- Use only evidence from the provided file summaries and the deterministic draft rows.
- Keep "Variable name" as the raw field identifier when available.
- Make "Variable" a clear, human-readable label — more informative than "Variable name" but concise.
- When both survey and data evidence exist, prioritize the survey export for "Variable name", "Values", and descriptions; use data evidence for measurement hints and notes.
- Do not invent response options, scales, skip logic, or derived variables not supported by the evidence.
- If item wording may be copyrighted, paraphrase rather than quote verbatim.
- Keep descriptions concise but informative — under about 30 words each. Put uncertainty or merge decisions in "Note", briefly.
- Output compact JSON: no extra whitespace, no repetition of unchanged long text when a tighter phrasing works.
- Respond ONLY with a JSON object of the form {"rows":[{"Variable":"...","Variable name":"...","Measurement/unit":"...","Values":"...","Description":"...","Note":"..."}]} — no markdown, no preamble, no code fences.`;

const CONDENSE_RULE = `- CONDENSE: when several rows clearly belong to one scale or battery (e.g. numbered items scale_1..scale_10 sharing a response scale), merge them into ONE row: use the base variable name plus the item range (e.g. "rses_1 – rses_10"), summarize the construct in Description, list the shared response options in Values, and record the condensing decision and item count in Note. Do not condense genuinely distinct variables just because they share a response scale.`;
const NO_CONDENSE_RULE = `- Keep one output row per input draft row. Do not merge rows.`;

function chunkRowsByScale(rows, maxPerChunk) {
  // Keep same-base-name items (scale_1, scale_2, …) inside the same chunk.
  const groups = [];
  let currentBase = null;
  for (const r of rows) {
    const m = String(r["Variable name"]).match(/^(.*?)[_-]\d+$/);
    const base = m ? m[1] : `__solo__${groups.length}`;
    if (base === currentBase && groups.length) {
      groups[groups.length - 1].push(r);
    } else {
      groups.push([r]);
      currentBase = base;
    }
  }
  const chunks = [];
  let cur = [];
  for (const g of groups) {
    if (cur.length && cur.length + g.length > maxPerChunk) {
      chunks.push(cur);
      cur = [];
    }
    cur = cur.concat(g);
    if (cur.length >= maxPerChunk) {
      chunks.push(cur);
      cur = [];
    }
  }
  if (cur.length) chunks.push(cur);
  return chunks;
}

function extractJSON(text) {
  let t = text.replace(/```json/gi, "").replace(/```/g, "").trim();
  try {
    return JSON.parse(t);
  } catch {
    const start = t.indexOf("{");
    const end = t.lastIndexOf("}");
    if (start >= 0 && end > start) return JSON.parse(t.slice(start, end + 1));
    throw new Error("The model did not return valid JSON.");
  }
}

const PROVIDERS = {
  builtin: {
    label: "Built-in Claude (no key needed)",
    defaultModel: "claude-sonnet-4-6",
    models: ["claude-sonnet-4-6"],
    needsKey: false,
  },
  anthropic: {
    label: "Anthropic API (your key)",
    defaultModel: "claude-sonnet-4-6",
    models: ["claude-sonnet-4-6", "claude-haiku-4-5", "claude-opus-4-8"],
    needsKey: true,
  },
  openai: {
    label: "OpenAI API (your key)",
    defaultModel: "gpt-5-mini",
    models: ["gpt-5-mini", "gpt-5", "gpt-4.1-mini"],
    needsKey: true,
  },
};

async function callModel({ provider, model, apiKey, system, userMsg }) {
  if (provider === "openai") {
    let resp;
    try {
      resp = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${apiKey}`,
        },
        body: JSON.stringify({
          model,
          messages: [
            { role: "system", content: system },
            { role: "user", content: userMsg },
          ],
        }),
      });
    } catch {
      throw new Error(
        "Could not reach the OpenAI API. Note: inside claude.ai, browser security rules only allow calls to the Anthropic API — the OpenAI option works when this app is hosted standalone."
      );
    }
    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`OpenAI request failed (${resp.status}): ${truncate(errText, 200)}`);
    }
    const data = await resp.json();
    return {
      text: data.choices?.[0]?.message?.content || "",
      truncated: data.choices?.[0]?.finish_reason === "length",
    };
  }

  // Anthropic — either the built-in keyless endpoint or the user's own key
  const headers = { "Content-Type": "application/json" };
  if (provider === "anthropic") {
    headers["x-api-key"] = apiKey;
    headers["anthropic-version"] = "2023-06-01";
    headers["anthropic-dangerous-direct-browser-access"] = "true";
  }
  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers,
    body: JSON.stringify({
      model,
      max_tokens: provider === "builtin" ? 1000 : 4000,
      system,
      messages: [{ role: "user", content: userMsg }],
    }),
  });
  if (!resp.ok) {
    if (provider === "builtin") {
      throw new Error(
        "The built-in keyless option only works when this app runs inside Claude. Switch the provider in AI settings to Anthropic or OpenAI and enter your own API key."
      );
    }
    const errText = await resp.text();
    throw new Error(`Anthropic request failed (${resp.status}): ${truncate(errText, 200)}`);
  }
  const data = await resp.json();
  return {
    text: (data.content || [])
      .filter((b) => b.type === "text")
      .map((b) => b.text)
      .join("\n"),
    truncated: data.stop_reason === "max_tokens",
  };
}

/* Recover complete row objects from a truncated JSON array. */
function salvageRows(text) {
  const start = text.indexOf("[");
  if (start < 0) return [];
  const objs = [];
  let depth = 0;
  let objStart = -1;
  let inStr = false;
  let esc = false;
  for (let i = start + 1; i < text.length; i++) {
    const ch = text[i];
    if (inStr) {
      if (esc) esc = false;
      else if (ch === "\\") esc = true;
      else if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') { inStr = true; continue; }
    if (ch === "{") { if (depth === 0) objStart = i; depth++; }
    else if (ch === "}") {
      depth--;
      if (depth === 0 && objStart >= 0) {
        try { objs.push(JSON.parse(text.slice(objStart, i + 1))); } catch { /* skip broken object */ }
        objStart = -1;
      }
    } else if (ch === "]" && depth === 0) break;
  }
  return objs;
}

async function refineCodebook({ rows, fileSummaries, condense, ai, onProgress }) {
  const clean = rows.map((r) =>
    Object.fromEntries(SCORE_COLUMNS.map((c) => [c, r[c] || ""]))
  );
  // The built-in endpoint has a smaller response budget, so use smaller batches there.
  const chunks = chunkRowsByScale(clean, ai.provider === "builtin" ? 6 : 25);
  const summaryBudget = ai.provider === "builtin" ? 2000 : 4000;
  const globalSummary = fileSummaries
    .map((f) => `FILE ${f.name} (role: ${f.role})\n${truncate(f.summary, Math.floor(summaryBudget / Math.max(fileSummaries.length, 1)))}`)
    .join("\n\n");
  const system = AI_SYSTEM + "\n" + (condense ? CONDENSE_RULE : NO_CONDENSE_RULE);

  const toRows = (arr) =>
    (arr || []).map((r) => ({
      ...emptyRow(),
      ...Object.fromEntries(SCORE_COLUMNS.map((c) => [c, String(r[c] ?? "")])),
    }));

  // Refine one batch; if the reply is truncated or unparseable, split the
  // batch in half and retry each half, down to single rows.
  const refineBatch = async (batch) => {
    const userMsg =
      `Improve these draft codebook rows. Prefer cleaning and enriching the draft over inventing structure.\n\n` +
      `UPLOADED FILE SUMMARIES:\n${globalSummary}\n\n` +
      `DRAFT ROWS (JSON):\n${JSON.stringify(batch)}`;

    const { text, truncated } = await callModel({
      provider: ai.provider,
      model: ai.model,
      apiKey: ai.apiKey,
      system,
      userMsg,
    });

    if (!truncated) {
      try {
        const parsed = extractJSON(text);
        const newRows = toRows(parsed.rows);
        if (newRows.length) return newRows;
      } catch { /* fall through to splitting/salvage */ }
    }

    if (batch.length > 1) {
      const mid = Math.ceil(batch.length / 2);
      onProgress(`Reply was too long for one batch — splitting into smaller batches…`);
      const first = await refineBatch(batch.slice(0, mid));
      const second = await refineBatch(batch.slice(mid));
      return [...first, ...second];
    }

    // Single row and still failing: salvage anything complete, else keep the draft row.
    const salvaged = toRows(salvageRows(text));
    if (salvaged.length) return salvaged;
    return toRows([{ ...batch[0], Note: `${batch[0].Note} [AI refinement failed for this row; draft kept.]`.trim() }]);
  };

  const out = [];
  let done = 0;
  for (const chunk of chunks) {
    onProgress(`Refining rows ${done + 1}–${done + chunk.length} of ${clean.length}…`);
    const refined = await refineBatch(chunk);
    out.push(...refined);
    done += chunk.length;
  }
  if (!out.length) throw new Error("The model returned no usable rows.");
  return out;
}

/* ============================================================
   Exports
   ============================================================ */
const cleanForExport = (rows) =>
  rows.map((r) => Object.fromEntries(SCORE_COLUMNS.map((c) => [c, r[c] || ""])));

function toCSV(rows) {
  const esc = (v) => {
    const s = String(v ?? "");
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [SCORE_COLUMNS.map(esc).join(",")];
  for (const r of cleanForExport(rows)) {
    lines.push(SCORE_COLUMNS.map((c) => esc(r[c])).join(","));
  }
  return lines.join("\n");
}

function toMarkdown(rows) {
  const header = `| ${SCORE_COLUMNS.join(" | ")} |`;
  const divider = `| ${SCORE_COLUMNS.map(() => "---").join(" | ")} |`;
  const lines = [header, divider];
  for (const r of cleanForExport(rows)) {
    lines.push(
      `| ${SCORE_COLUMNS.map((c) =>
        String(r[c] ?? "").replace(/\n/g, " ").replace(/\|/g, "\\|")
      ).join(" | ")} |`
    );
  }
  return lines.join("\n");
}

function downloadBlob(content, filename, type) {
  const blob = content instanceof Blob ? content : new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function downloadXLSX(rows) {
  const ws = XLSX.utils.json_to_sheet(cleanForExport(rows), { header: SCORE_COLUMNS });
  ws["!cols"] = [{ wch: 34 }, { wch: 18 }, { wch: 20 }, { wch: 38 }, { wch: 52 }, { wch: 40 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "codebook");
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  downloadBlob(
    new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    "codebook.xlsx"
  );
}

/* ============================================================
   Sample data (PsyToolkit demo)
   ============================================================ */
const SAMPLE_PSYTOOLKIT = `scale: agree5
- Strongly disagree
- Disagree
- Neither agree nor disagree
- Agree
- Strongly agree

l: rses
t: scale agree5
q: Please indicate how much you agree with each statement about yourself.
- I feel that I am a person of worth.
- I feel that I have a number of good qualities.
- All in all, I am inclined to feel that I am a failure.
- I am able to do things as well as most other people.
- I feel I do not have much to be proud of.

l: age
t: textline
q: What is your age in years?

l: gender
t: radio
q: What is your gender?
- Woman
- Man
- Non-binary
- Prefer not to say
`;

/* ============================================================
   UI
   ============================================================ */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,650;9..144,750&family=Source+Sans+3:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

:root {
  --paper: #F6F4EC;
  --card: #FFFDF8;
  --ink: #17222B;
  --muted: #5B6A76;
  --green: #1F6F5B;
  --green-deep: #14513F;
  --green-wash: #E7F1EC;
  --slate: #36536B;
  --line: rgba(23,34,43,0.14);
  --line-soft: rgba(23,34,43,0.08);
  --warn-bg: #FBF0E3;
  --warn-ink: #8A5A1E;
}
.cbk * { box-sizing: border-box; }
.cbk {
  min-height: 100vh;
  background:
    radial-gradient(1100px 500px at 85% -10%, rgba(31,111,91,0.08), transparent 60%),
    var(--paper);
  color: var(--ink);
  font-family: 'Source Sans 3', system-ui, sans-serif;
  font-size: 15px;
  padding: 28px clamp(14px, 4vw, 44px) 64px;
}
.cbk-header { display: flex; align-items: baseline; gap: 14px; flex-wrap: wrap; margin-bottom: 6px; }
.cbk-wordmark {
  font-family: 'Fraunces', serif; font-weight: 750; font-size: clamp(30px, 4.5vw, 42px);
  letter-spacing: -0.01em; margin: 0; color: var(--ink);
}
.cbk-wordmark em { font-style: italic; color: var(--green); }
.cbk-tagline { color: var(--muted); font-size: 15px; max-width: 62ch; margin: 2px 0 20px; }
.cbk-tagline strong { color: var(--green-deep); }

.cbk-steps { display: flex; gap: 0; flex-wrap: wrap; border: 1px solid var(--line); border-radius: 12px; overflow: hidden; background: var(--card); margin-bottom: 22px; }
.cbk-step { flex: 1 1 150px; padding: 10px 14px; border-right: 1px solid var(--line-soft); display: flex; gap: 9px; align-items: center; min-width: 150px; }
.cbk-step:last-child { border-right: none; }
.cbk-step .num { font-family: 'IBM Plex Mono', monospace; font-size: 12px; color: var(--muted); border: 1px solid var(--line); border-radius: 999px; width: 22px; height: 22px; display: grid; place-items: center; flex: none; }
.cbk-step.done .num { background: var(--green); border-color: var(--green); color: #fff; }
.cbk-step.active .num { border-color: var(--green); color: var(--green); font-weight: 600; }
.cbk-step .lbl { font-size: 13px; font-weight: 600; color: var(--muted); }
.cbk-step.done .lbl, .cbk-step.active .lbl { color: var(--ink); }

.cbk-layout { display: grid; grid-template-columns: 330px minmax(0, 1fr); gap: 20px; align-items: start; }
@media (max-width: 940px) { .cbk-layout { grid-template-columns: 1fr; } }

.cbk-card { background: var(--card); border: 1px solid var(--line); border-radius: 14px; padding: 18px; box-shadow: 0 10px 30px rgba(23,34,43,0.05); }
.cbk-card h2 { font-family: 'Fraunces', serif; font-weight: 650; font-size: 18px; margin: 0 0 4px; }
.cbk-card .sub { color: var(--muted); font-size: 13px; margin: 0 0 12px; }

.cbk-drop {
  border: 1.5px dashed rgba(31,111,91,0.45); border-radius: 12px; background: var(--green-wash);
  padding: 22px 14px; text-align: center; cursor: pointer; transition: background 0.15s, border-color 0.15s;
}
.cbk-drop:hover, .cbk-drop.over { background: #DCEBE3; border-color: var(--green); }
.cbk-drop:focus-visible { outline: 2px solid var(--green); outline-offset: 2px; }
.cbk-drop .big { font-weight: 600; color: var(--green-deep); }
.cbk-drop .fmt { font-family: 'IBM Plex Mono', monospace; font-size: 11.5px; color: var(--muted); margin-top: 6px; }

.cbk-filelist { margin-top: 12px; display: flex; flex-direction: column; gap: 8px; }
.cbk-file { display: flex; align-items: center; gap: 8px; border: 1px solid var(--line-soft); border-radius: 10px; padding: 7px 9px; background: #fff; }
.cbk-file .fname { font-family: 'IBM Plex Mono', monospace; font-size: 12px; flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.cbk-file select {
  font-family: 'IBM Plex Mono', monospace; font-size: 11.5px; border: 1px solid var(--line); border-radius: 7px;
  padding: 3px 5px; background: var(--green-wash); color: var(--green-deep);
}
.cbk-iconbtn { border: none; background: transparent; color: var(--muted); cursor: pointer; padding: 3px; border-radius: 6px; display: grid; place-items: center; }
.cbk-iconbtn:hover { color: var(--ink); background: var(--line-soft); }

.cbk-opt { display: flex; align-items: flex-start; gap: 9px; padding: 9px 0; border-top: 1px solid var(--line-soft); cursor: pointer; }
.cbk-opt input { margin-top: 3px; accent-color: var(--green); width: 15px; height: 15px; }
.cbk-opt .t { font-weight: 600; font-size: 13.5px; }
.cbk-opt .d { color: var(--muted); font-size: 12.5px; }

.cbk-btn {
  display: inline-flex; align-items: center; justify-content: center; gap: 8px;
  border-radius: 10px; padding: 9px 14px; font-weight: 600; font-size: 14px; cursor: pointer;
  border: 1px solid var(--line); background: #fff; color: var(--ink);
  transition: transform 0.06s, background 0.15s;
  font-family: 'Source Sans 3', system-ui, sans-serif;
}
.cbk-btn:hover { background: var(--green-wash); }
.cbk-btn:active { transform: translateY(1px); }
.cbk-btn:disabled { opacity: 0.5; cursor: not-allowed; }
.cbk-btn.primary { background: var(--green); border-color: var(--green); color: #fff; }
.cbk-btn.primary:hover { background: var(--green-deep); }
.cbk-btn.ghost { border-color: transparent; background: transparent; color: var(--muted); }
.cbk-btn.ghost:hover { color: var(--ink); background: var(--line-soft); }
.cbk-btn:focus-visible { outline: 2px solid var(--green); outline-offset: 2px; }

.cbk-note { background: var(--warn-bg); color: var(--warn-ink); border-radius: 9px; padding: 8px 11px; font-size: 13px; margin-top: 10px; }

.cbk-ai { border-top: 1px solid var(--line-soft); margin-top: 6px; padding-top: 12px; }
.cbk-ai-title { font-family: 'IBM Plex Mono', monospace; font-size: 11px; letter-spacing: 0.1em; text-transform: uppercase; color: var(--muted); margin-bottom: 8px; }
.cbk-field { display: flex; flex-direction: column; gap: 3px; margin-bottom: 9px; }
.cbk-field > span { font-size: 12px; font-weight: 600; color: var(--muted); }
.cbk-field select, .cbk-field input {
  border: 1px solid var(--line); border-radius: 8px; padding: 7px 9px; font-size: 13px;
  background: #fff; color: var(--ink); font-family: 'Source Sans 3', system-ui, sans-serif; width: 100%;
}
.cbk-field input[list], .cbk-field input[type="password"] { font-family: 'IBM Plex Mono', monospace; font-size: 12.5px; }
.cbk-field select:focus-visible, .cbk-field input:focus-visible { outline: 2px solid var(--green); outline-offset: 1px; }
.cbk-ai-hint { font-size: 12px; color: var(--muted); line-height: 1.45; }

.cbk-tablehead { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; margin-bottom: 10px; }
.cbk-tablehead h2 { margin: 0; }
.cbk-stamp {
  font-family: 'IBM Plex Mono', monospace; font-size: 11px; letter-spacing: 0.14em; font-weight: 600;
  border: 1.5px solid currentColor; border-radius: 4px; padding: 3px 9px; transform: rotate(-2deg);
  text-transform: uppercase;
}
.cbk-stamp.draft { color: var(--slate); }
.cbk-stamp.ai { color: var(--green); }
.cbk-count { color: var(--muted); font-size: 13px; font-family: 'IBM Plex Mono', monospace; }
.cbk-spacer { flex: 1; }

.cbk-tablewrap { overflow-x: auto; border: 1px solid var(--line); border-radius: 12px; background: #fff; }
.cbk-table { border-collapse: collapse; width: 100%; min-width: 960px; }
.cbk-table th {
  text-align: left; font-family: 'IBM Plex Mono', monospace; font-size: 11px; letter-spacing: 0.08em;
  text-transform: uppercase; color: var(--muted); background: var(--paper);
  padding: 9px 10px; border-bottom: 1px solid var(--line); position: sticky; top: 0;
}
.cbk-table td { border-bottom: 1px solid var(--line-soft); vertical-align: top; padding: 0; }
.cbk-table tr:last-child td { border-bottom: none; }
.cbk-table td:first-child { border-left: 3px solid transparent; }
.cbk-table tr:hover td:first-child { border-left-color: var(--green); }
.cbk-cell {
  width: 100%; border: none; background: transparent; resize: none; font-size: 13.5px;
  font-family: 'Source Sans 3', system-ui, sans-serif; color: var(--ink);
  padding: 8px 10px; min-height: 38px; line-height: 1.4;
}
.cbk-cell.mono { font-family: 'IBM Plex Mono', monospace; font-size: 12.5px; color: var(--green-deep); }
.cbk-cell:focus { outline: 2px solid rgba(31,111,91,0.5); outline-offset: -2px; background: var(--green-wash); border-radius: 6px; }
.cbk-rowdel { padding: 8px 6px !important; }

.cbk-actions { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 14px; align-items: center; }
.cbk-status { font-size: 13.5px; color: var(--green-deep); display: flex; align-items: center; gap: 8px; }
.cbk-error { color: #9B2C2C; font-size: 13.5px; }
.cbk-spin { width: 14px; height: 14px; border: 2px solid var(--green-wash); border-top-color: var(--green); border-radius: 50%; animation: cbkspin 0.8s linear infinite; flex: none; }
@keyframes cbkspin { to { transform: rotate(360deg); } }
@media (prefers-reduced-motion: reduce) { .cbk-spin { animation-duration: 2s; } }

.cbk-md { margin-top: 12px; }
.cbk-md pre {
  background: #1E2A33; color: #E8EEF2; border-radius: 10px; padding: 14px; overflow-x: auto;
  font-family: 'IBM Plex Mono', monospace; font-size: 12px; line-height: 1.5; max-height: 320px;
}
.cbk-empty { text-align: center; color: var(--muted); padding: 44px 20px; }
.cbk-empty svg { opacity: 0.4; }
.cbk-footer { margin-top: 22px; color: var(--muted); font-size: 12.5px; max-width: 78ch; }
`;

export default function Codebooker() {
  const [fileEntries, setFileEntries] = useState([]); // {id, file, role}
  const [includeWording, setIncludeWording] = useState(true);
  const [condense, setCondense] = useState(false);
  const [rows, setRows] = useState([]);
  const [draftRows, setDraftRows] = useState([]);
  const [fileSummaries, setFileSummaries] = useState([]);
  const [notes, setNotes] = useState([]);
  const [mode, setMode] = useState("none"); // none | draft | ai
  const [busy, setBusy] = useState("");
  const [error, setError] = useState("");
  const [showMd, setShowMd] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [provider, setProvider] = useState("builtin");
  const [model, setModel] = useState(PROVIDERS.builtin.defaultModel);
  const [apiKey, setApiKey] = useState("");
  const inputRef = useRef(null);
  const idRef = useRef(0);

  const addFiles = (list) => {
    const entries = Array.from(list).map((file) => ({
      id: ++idRef.current,
      file,
      role: "", // auto
    }));
    setFileEntries((prev) => [...prev, ...entries]);
    setMode("none");
  };

  const removeFile = (id) => {
    setFileEntries((prev) => prev.filter((f) => f.id !== id));
    setMode("none");
  };

  const setRole = (id, role) => {
    setFileEntries((prev) => prev.map((f) => (f.id === id ? { ...f, role } : f)));
    setMode("none");
  };

  const loadSample = () => {
    const file = new File([SAMPLE_PSYTOOLKIT], "sample_psytoolkit_survey.txt", { type: "text/plain" });
    addFiles([file]);
  };

  const buildDraft = useCallback(async () => {
    setError("");
    setBusy("Parsing files…");
    try {
      let allRows = [];
      let allNotes = [];
      const summaries = [];
      for (const entry of fileEntries) {
        const result = await readUpload(entry.file, entry.role || null, includeWording);
        allRows = allRows.concat(result.rows);
        allNotes = allNotes.concat(result.notes);
        summaries.push({ name: entry.file.name, role: result.role, summary: result.summary });
      }
      const merged = mergeRows(allRows);
      setNotes(allNotes);
      setFileSummaries(summaries);
      setDraftRows(merged);
      setRows(merged.map((r) => ({ ...r })));
      setMode(merged.length ? "draft" : "none");
      if (!merged.length && !allNotes.length) {
        setError("No variables could be detected in the uploaded files.");
      }
    } catch (err) {
      setError(err.message);
    } finally {
      setBusy("");
    }
  }, [fileEntries, includeWording]);

  const runAI = async () => {
    setError("");
    if (PROVIDERS[provider].needsKey && !apiKey.trim()) {
      setError(`Enter your ${provider === "openai" ? "OpenAI" : "Anthropic"} API key in AI settings first, or switch to the built-in provider.`);
      return;
    }
    setBusy("Preparing refinement…");
    try {
      const refined = await refineCodebook({
        rows,
        fileSummaries,
        condense,
        ai: { provider, model: model.trim() || PROVIDERS[provider].defaultModel, apiKey: apiKey.trim() },
        onProgress: (msg) => setBusy(msg),
      });
      setRows(refined);
      setMode("ai");
    } catch (err) {
      setError(`AI refinement failed: ${err.message}`);
    } finally {
      setBusy("");
    }
  };

  const resetDraft = () => {
    setRows(draftRows.map((r) => ({ ...r })));
    setMode("draft");
    setError("");
  };

  const updateCell = (i, col, value) => {
    setRows((prev) => {
      const next = [...prev];
      next[i] = { ...next[i], [col]: value };
      return next;
    });
  };

  const deleteRow = (i) => setRows((prev) => prev.filter((_, j) => j !== i));
  const addRow = () => setRows((prev) => [...prev, emptyRow()]);

  const hasFiles = fileEntries.length > 0;
  const hasRows = rows.length > 0;
  const stepState = (n) => {
    const progress = mode === "ai" ? 3 : hasRows ? 2 : hasFiles ? 1 : 0;
    if (n <= progress) return "done";
    if (n === progress + 1) return "active";
    return "";
  };

  return (
    <div className="cbk">
      <style>{CSS}</style>

      <header className="cbk-header">
        <h1 className="cbk-wordmark">Code<em>booker</em></h1>
      </header>
      <p className="cbk-tagline">
        Turn a survey export into a <strong>SCORE-style codebook</strong>. A deterministic
        draft is built first, entirely in your browser; an AI model can then refine wording
        and structure — with the built-in Claude option, <strong>no API key is needed</strong>.
        Or plug in your own Anthropic or OpenAI key and pick a model. Review, edit, export.
      </p>

      <div className="cbk-steps" aria-label="Workflow">
        {["Upload sources", "Draft codebook", "Refine with Claude", "Review & export"].map((lbl, i) => (
          <div key={lbl} className={`cbk-step ${stepState(i + 1)}`}>
            <span className="num">{i + 1}</span>
            <span className="lbl">{lbl}</span>
          </div>
        ))}
      </div>

      <div className="cbk-layout">
        {/* ------------ Left panel ------------ */}
        <div className="cbk-card">
          <h2>Sources</h2>
          <p className="sub">Survey exports, data files, or both. Matching variables are reconciled.</p>

          <div
            className={`cbk-drop ${dragOver ? "over" : ""}`}
            role="button"
            tabIndex={0}
            onClick={() => inputRef.current?.click()}
            onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") inputRef.current?.click(); }}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => { e.preventDefault(); setDragOver(false); addFiles(e.dataTransfer.files); }}
          >
            <Upload size={20} style={{ color: "var(--green)" }} />
            <div className="big">Drop files or click to browse</div>
            <div className="fmt">.qsf · .csv · .xlsx · .docx · .txt · .md · .json</div>
          </div>
          <input
            ref={inputRef}
            type="file"
            multiple
            accept=".qsf,.json,.csv,.tsv,.xlsx,.xlsm,.xls,.docx,.txt,.md,.pdf"
            style={{ display: "none" }}
            onChange={(e) => { addFiles(e.target.files); e.target.value = ""; }}
          />

          {fileEntries.length > 0 && (
            <div className="cbk-filelist">
              {fileEntries.map((entry) => (
                <div className="cbk-file" key={entry.id}>
                  <FileText size={14} style={{ color: "var(--slate)", flex: "none" }} />
                  <span className="fname" title={entry.file.name}>{entry.file.name}</span>
                  <select
                    aria-label={`Role for ${entry.file.name}`}
                    value={entry.role}
                    onChange={(e) => setRole(entry.id, e.target.value)}
                  >
                    <option value="">auto</option>
                    <option value="survey">survey</option>
                    <option value="data">data</option>
                  </select>
                  <button className="cbk-iconbtn" aria-label={`Remove ${entry.file.name}`} onClick={() => removeFile(entry.id)}>
                    <X size={14} />
                  </button>
                </div>
              ))}
            </div>
          )}

          <div style={{ marginTop: 14 }}>
            <label className="cbk-opt">
              <input type="checkbox" checked={includeWording} onChange={(e) => { setIncludeWording(e.target.checked); setMode("none"); }} />
              <span>
                <span className="t">Include full item wording</span>
                <div className="d">Turn off when wording should stay short or may be copyrighted.</div>
              </span>
            </label>
            <label className="cbk-opt">
              <input type="checkbox" checked={condense} onChange={(e) => setCondense(e.target.checked)} />
              <span>
                <span className="t">Condense scales (AI step)</span>
                <div className="d">One row per scale (e.g. rses_1–rses_10) instead of one row per item.</div>
              </span>
            </label>
          </div>

          <div className="cbk-ai">
            <div className="cbk-ai-title">AI settings</div>
            <label className="cbk-field">
              <span>Provider</span>
              <select
                value={provider}
                onChange={(e) => {
                  const p = e.target.value;
                  setProvider(p);
                  setModel(PROVIDERS[p].defaultModel);
                }}
              >
                {Object.entries(PROVIDERS).map(([k, v]) => (
                  <option key={k} value={k}>{v.label}</option>
                ))}
              </select>
            </label>
            <label className="cbk-field">
              <span>Model</span>
              <input
                type="text"
                value={model}
                list="cbk-model-suggestions"
                onChange={(e) => setModel(e.target.value)}
                spellCheck={false}
              />
              <datalist id="cbk-model-suggestions">
                {PROVIDERS[provider].models.map((m) => (
                  <option key={m} value={m} />
                ))}
              </datalist>
            </label>
            {PROVIDERS[provider].needsKey && (
              <label className="cbk-field">
                <span>API key</span>
                <input
                  type="password"
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                  placeholder={provider === "openai" ? "sk-…" : "sk-ant-…"}
                  autoComplete="off"
                />
              </label>
            )}
            <div className="cbk-ai-hint">
              {provider === "builtin"
                ? "Runs on Claude with no key or cost — available while this app is used inside Claude."
                : provider === "anthropic"
                ? "Your key is kept in memory only and sent directly to Anthropic from your browser."
                : "Your key is kept in memory only and sent directly to OpenAI. Inside claude.ai this call is blocked by browser security rules; it works when the app is hosted standalone."}
            </div>
          </div>

          <div style={{ display: "flex", gap: 8, marginTop: 14, flexWrap: "wrap" }}>
            <button className="cbk-btn primary" disabled={!hasFiles || !!busy} onClick={buildDraft}>
              <Table2 size={15} /> Build draft codebook
            </button>
            <button className="cbk-btn ghost" disabled={!!busy} onClick={loadSample}>
              <FlaskConical size={15} /> Try a sample
            </button>
          </div>

          {notes.map((n, i) => (
            <div className="cbk-note" key={i}>{n}</div>
          ))}
        </div>

        {/* ------------ Right panel ------------ */}
        <div className="cbk-card">
          {!hasRows ? (
            <div className="cbk-empty">
              <BookOpen size={38} />
              <p style={{ fontFamily: "'Fraunces', serif", fontSize: 18, margin: "12px 0 4px" }}>
                Your codebook will appear here
              </p>
              <p style={{ fontSize: 13.5, margin: 0 }}>
                Add at least one source file on the left, then build the draft.
              </p>
            </div>
          ) : (
            <>
              <div className="cbk-tablehead">
                <h2>Codebook</h2>
                <span className={`cbk-stamp ${mode === "ai" ? "ai" : "draft"}`}>
                  {mode === "ai" ? "AI-refined" : "Deterministic draft"}
                </span>
                <span className="cbk-count">{rows.length} rows</span>
                <span className="cbk-spacer" />
                <button className="cbk-btn" disabled={!!busy} onClick={runAI} title={`Uses ${model}`}>
                  <Sparkles size={15} /> Refine with AI
                </button>
                {mode === "ai" && (
                  <button className="cbk-btn ghost" disabled={!!busy} onClick={resetDraft}>
                    <RotateCcw size={14} /> Back to draft
                  </button>
                )}
              </div>

              {busy && (
                <div className="cbk-status" role="status">
                  <span className="cbk-spin" /> {busy}
                </div>
              )}
              {error && <div className="cbk-error" role="alert">{error}</div>}

              <div className="cbk-tablewrap" style={{ marginTop: 10 }}>
                <table className="cbk-table">
                  <thead>
                    <tr>
                      {SCORE_COLUMNS.map((c) => (
                        <th key={c} style={c === "Description" ? { minWidth: 220 } : c === "Values" ? { minWidth: 180 } : {}}>
                          {c}
                        </th>
                      ))}
                      <th aria-label="Row actions" style={{ width: 34 }} />
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((r, i) => (
                      <tr key={i}>
                        {SCORE_COLUMNS.map((c) => (
                          <td key={c}>
                            <textarea
                              className={`cbk-cell ${c === "Variable name" ? "mono" : ""}`}
                              value={r[c] || ""}
                              rows={1}
                              onChange={(e) => updateCell(i, c, e.target.value)}
                              aria-label={`${c}, row ${i + 1}`}
                            />
                          </td>
                        ))}
                        <td className="cbk-rowdel">
                          <button className="cbk-iconbtn" aria-label={`Delete row ${i + 1}`} onClick={() => deleteRow(i)}>
                            <Trash2 size={14} />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="cbk-actions">
                <button className="cbk-btn ghost" onClick={addRow}>
                  <Plus size={14} /> Add row
                </button>
                <span className="cbk-spacer" />
                <button className="cbk-btn" onClick={() => downloadBlob(toCSV(rows), "codebook.csv", "text/csv")}>
                  <Download size={15} /> CSV
                </button>
                <button className="cbk-btn" onClick={() => downloadXLSX(rows)}>
                  <Download size={15} /> XLSX
                </button>
                <button className="cbk-btn" onClick={() => downloadBlob(toMarkdown(rows), "codebook.md", "text/markdown")}>
                  <Download size={15} /> Markdown
                </button>
                <button className="cbk-btn ghost" onClick={() => setShowMd((s) => !s)}>
                  {showMd ? <ChevronUp size={14} /> : <ChevronDown size={14} />} Preview markdown
                </button>
              </div>

              {showMd && (
                <div className="cbk-md">
                  <pre>{toMarkdown(rows)}</pre>
                </div>
              )}
            </>
          )}
        </div>
      </div>

      <p className="cbk-footer">
        Everything runs in your browser; files never leave it during the deterministic draft.
        The AI step sends only compact file summaries and the draft rows to the selected
        provider; API keys are held in memory and never stored. Always review the
        result before sharing — check copyrighted item wording, matrix questions, branching
        logic, and derived variables. Structure follows the SCORE data-dictionary format
        (Source and analysis-relevance columns omitted, as they are analysis-specific).
      </p>
    </div>
  );
}
