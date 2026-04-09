/* ===============================
QA LENGTH AUDIT (STANDALONE)
=============================== */

const TBX_DEFAULT_AD_SHEET = "Ad Copy";
const TBX_REPORT_SHEET = "QA Length Audit";
const TBX_FIX_LOG_SHEET = "QA Length Fix Log";
const TBX_HEADLINE_MAX = 30;
const TBX_DESCRIPTION_MAX = 90;
const TBX_MODE_KEY = "QAL_LENGTH_FIX_MODE";
const TBX_MODE_DETERMINISTIC = "deterministic";
const TBX_MODE_AI_FALLBACK = "ai_fallback";
const TBX_MODE_COMPARE = "compare_best";

function qaLengthAuditRun() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adSheetName = qalGetTextSetting_("AD_SHEET_NAME", TBX_DEFAULT_AD_SHEET);
  const adSheet = ss.getSheetByName(adSheetName);
  if (!adSheet) {
    SpreadsheetApp.getUi().alert("Ad sheet not found: " + adSheetName);
    return;
  }

  const data = adSheet.getDataRange().getValues();
  const report = qalGetOrCreateReportSheet_(ss);
  report.clearContents();
  report.appendRow([
    "Time", "Row", "AdGroup", "AssetType", "AssetIndex",
    "Column", "Length", "MaxAllowed", "Text"
  ]);

  const now = new Date();
  let findings = 0;

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const adGroup = qalToText_(data[i][1]);
    if (!adGroup) continue;

    const headlines = data[i].slice(3, 18).map(qalToText_);
    const descriptions = data[i].slice(18, 22).map(qalToText_);

    for (let h = 0; h < headlines.length; h++) {
      const text = headlines[h];
      if (!text) continue;
      if (text.length > TBX_HEADLINE_MAX) {
        report.appendRow([
          now, row, adGroup, "Headline", h + 1,
          qalColToLetter_(4 + h), text.length, TBX_HEADLINE_MAX, text
        ]);
        findings++;
      }
    }

    for (let d = 0; d < descriptions.length; d++) {
      const text = descriptions[d];
      if (!text) continue;
      if (text.length > TBX_DESCRIPTION_MAX) {
        report.appendRow([
          now, row, adGroup, "Description", d + 1,
          qalColToLetter_(19 + d), text.length, TBX_DESCRIPTION_MAX, text
        ]);
        findings++;
      }
    }
  }

  if (findings) {
    report.autoResizeColumns(1, 9);
  }
  SpreadsheetApp.getUi().alert("Length audit complete. Found " + findings + " over-limit assets.");
}

function qaLengthFixAllTooLong() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adSheetName = qalGetTextSetting_("AD_SHEET_NAME", TBX_DEFAULT_AD_SHEET);
  const keywordSheetName = qalGetTextSetting_("KEYWORD_SHEET_NAME", "List Of Keywords");
  const adSheet = ss.getSheetByName(adSheetName);
  const keywordSheet = ss.getSheetByName(keywordSheetName);
  if (!adSheet) {
    SpreadsheetApp.getUi().alert("Ad sheet not found: " + adSheetName);
    return;
  }
  const mode = qalGetFixMode_();

  const data = adSheet.getDataRange().getValues();
  const keywordData = keywordSheet ? keywordSheet.getDataRange().getValues() : [];
  const keywordHeaders = keywordData[0] || [];
  const logSheet = qalGetOrCreateFixLogSheet_(ss);
  const runId = "QALFIX_" + Date.now();
  let fixes = 0;

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const adGroup = qalToText_(data[i][1]);
    if (!adGroup) continue;

    for (let h = 0; h < 15; h++) {
      const col = 4 + h;
      const original = qalToText_(data[i][3 + h]);
      if (!original || original.length <= TBX_HEADLINE_MAX) continue;
      const keywords = qalGetKeywordsForAdGroup_(adGroup, keywordHeaders, keywordData);
      const selected = qalSelectBestFix_(original, "headline", TBX_HEADLINE_MAX, adGroup, keywords, mode);
      const fixed = selected.text;
      if (!fixed || fixed === original) continue;
      adSheet.getRange(row, col).setValue(fixed).setBackground("#d9ead3");
      logSheet.appendRow([new Date(), runId, row, adGroup, "Headline", h + 1, qalColToLetter_(col), original, fixed, original.length, fixed.length, mode, selected.method, selected.score]);
      fixes++;
    }

    for (let d = 0; d < 4; d++) {
      const col = 19 + d;
      const original = qalToText_(data[i][18 + d]);
      if (!original || original.length <= TBX_DESCRIPTION_MAX) continue;
      const keywords = qalGetKeywordsForAdGroup_(adGroup, keywordHeaders, keywordData);
      const selected = qalSelectBestFix_(original, "description", TBX_DESCRIPTION_MAX, adGroup, keywords, mode);
      const fixed = selected.text;
      if (!fixed || fixed === original) continue;
      adSheet.getRange(row, col).setValue(fixed).setBackground("#d9ead3");
      logSheet.appendRow([new Date(), runId, row, adGroup, "Description", d + 1, qalColToLetter_(col), original, fixed, original.length, fixed.length, mode, selected.method, selected.score]);
      fixes++;
    }
  }

  SpreadsheetApp.getUi().alert("Length fix complete. Updated " + fixes + " over-limit assets. Mode: " + mode);
}

function qaLengthFixSetMode() {
  const ui = SpreadsheetApp.getUi();
  const current = qalGetFixMode_();
  const resp = ui.prompt(
    "Set Length Fix Mode",
    "Enter one mode:\n- deterministic\n- ai_fallback\n- compare_best\n\nCurrent: " + current,
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const next = qalNormalizeMode_(resp.getResponseText());
  if (!next) {
    ui.alert("Invalid mode. Use deterministic, ai_fallback, or compare_best.");
    return;
  }
  PropertiesService.getScriptProperties().setProperty(TBX_MODE_KEY, next);
  ui.alert("Length fix mode set to: " + next);
}

function qaLengthFixShowMode() {
  SpreadsheetApp.getUi().alert("Current Length Fix mode: " + qalGetFixMode_());
}

function qaLengthFixModeDeterministic() {
  PropertiesService.getScriptProperties().setProperty(TBX_MODE_KEY, TBX_MODE_DETERMINISTIC);
  SpreadsheetApp.getUi().alert("Length Fix mode set to deterministic.");
}

function qaLengthFixModeAiFallback() {
  PropertiesService.getScriptProperties().setProperty(TBX_MODE_KEY, TBX_MODE_AI_FALLBACK);
  SpreadsheetApp.getUi().alert("Length Fix mode set to ai_fallback.");
}

function qaLengthFixModeCompareBest() {
  PropertiesService.getScriptProperties().setProperty(TBX_MODE_KEY, TBX_MODE_COMPARE);
  SpreadsheetApp.getUi().alert("Length Fix mode set to compare_best.");
}

function qalShortenToLimit_(text, maxLen, type) {
  const original = qalToText_(text);
  let t = original;
  if (!t) return t;
  if (t.length <= maxLen) return qalRestoreCaseStyle_(t, original, type);

  t = t.replace(/([a-z])([A-Z])/g, "$1 $2");
  t = t.replace(/([A-Za-z])\/([A-Za-z])/g, "$1 $2");

  const replacements = [
    [/private guided tours?/gi, "private tours"],
    [/private guided/gi, "private"],
    [/custom itineraries/gi, "custom plans"],
    [/custom itinerary/gi, "custom plan"],
    [/exclusive access/gi, "insider access"],
    [/tailored itinerary/gi, "tailored plan"],
    [/guided experiences?/gi, "guided tours"],
    [/experiences/gi, "tours"],
    [/journeys/gi, "trips"],
    [/journey/gi, "trip"]
  ];
  for (let i = 0; i < replacements.length; i++) {
    t = t.replace(replacements[i][0], replacements[i][1]);
  }
  t = t.replace(/\s+/g, " ").trim();
  if (t.length <= maxLen) return qalRestoreCaseStyle_(t, original, type);

  t = t.replace(/\b(the|a|an|very|really|that|which|just|simply|completely|fully|carefully)\b/gi, "");
  t = t.replace(/\s+/g, " ").trim();
  if (t.length <= maxLen) return qalRestoreCaseStyle_(t, original, type);

  t = t.split(/\s+/).map(function(w) {
    if (w.length > 4 && /s$/i.test(w) && !/(ss|us|is)$/i.test(w)) return w.slice(0, -1);
    return w;
  }).join(" ");
  t = t.replace(/\s+/g, " ").trim();
  if (t.length <= maxLen) return qalRestoreCaseStyle_(t, original, type);

  const sliced = t.slice(0, maxLen + 1);
  const lastSpace = sliced.lastIndexOf(" ");
  if (lastSpace > Math.floor(maxLen * 0.6)) return qalRestoreCaseStyle_(sliced.slice(0, lastSpace).trim(), original, type);
  return qalRestoreCaseStyle_(t.slice(0, maxLen).trim(), original, type);
}

function qalSelectBestFix_(original, type, maxLen, adGroup, keywords, mode) {
  const det = qalShortenToLimit_(original, maxLen, type);
  const detScore = qalHumanScore_(det, type, adGroup, original, maxLen);
  if (mode === TBX_MODE_DETERMINISTIC) {
    return { text: det, method: "deterministic", score: detScore };
  }

  const ai = qalAiRewriteToLimit_(original, type, maxLen, adGroup, keywords);
  const aiScore = qalHumanScore_(ai, type, adGroup, original, maxLen);

  if (mode === TBX_MODE_AI_FALLBACK) {
    if (detScore >= 60) return { text: det, method: "deterministic", score: detScore };
    return { text: ai || det, method: ai ? "ai" : "deterministic", score: ai ? aiScore : detScore };
  }

  // compare_best
  if (aiScore > detScore) return { text: ai, method: "ai", score: aiScore };
  return { text: det, method: "deterministic", score: detScore };
}

function qalAiRewriteToLimit_(original, type, maxLen, adGroup, keywords) {
  const apiKey = qalGetTextSetting_("OPENAI_API_KEY", "");
  if (!apiKey) return "";
  const model = qalGetTextSetting_("OPENAI_MODEL", "gpt-4o-mini");
  const banned = qalGetBlockedTerms_("BANNED_WORDS");
  const offShelf = qalGetBlockedTerms_("OFF_SHELF_WORDS");
  const prompt = `
Rewrite this ${type} for a Google Ads asset.
Keep meaning and place context.

Original:
${original}

Ad group:
${adGroup}

Keywords:
${(keywords || []).join(", ")}

Rules:
- Natural human phrasing
- Keep it specific, not generic
- No abbreviations like w/, svc
- If the event/context is Gemini, use "celebrate", "commemorate", or "milestone"
- For Gemini context, shift tone to festive and joyous
- Prefer "Celebrate your Gemini milestone" over neutral booking phrasing
- Avoid banned words: ${banned.join(", ")}
- Avoid off-shelf phrases: ${offShelf.join(", ")}
- Max ${maxLen} characters including spaces
- Provide 3 alternatives
- If place name cannot fit naturally, use a clean phrasing without place
- Never end with dangling prepositions like in, at, to, for, with, from

Return ONLY JSON:
{"alternatives":["alt1","alt2","alt3"]}
`;
  try {
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + apiKey },
      payload: JSON.stringify({
        model: model,
        messages: [{ role: "user", content: prompt }],
        temperature: 0.2
      }),
      muteHttpExceptions: true
    });
    const json = JSON.parse(response.getContentText());
    const out = json && json.choices && json.choices[0] && json.choices[0].message
      ? String(json.choices[0].message.content || "")
      : "";
    const candidates = qalExtractAlternatives_(out);
    if (!candidates.length) return "";
    let best = "";
    let bestScore = -1;
    for (let i = 0; i < candidates.length; i++) {
      let c = candidates[i];
      if (c.length > maxLen) c = qalShortenToLimit_(c, maxLen, type);
      c = qalRestoreCaseStyle_(c, original, type);
      const s = qalHumanScore_(c, type, adGroup, original, maxLen);
      if (s > bestScore) {
        bestScore = s;
        best = c;
      }
    }
    return best;
  } catch (_e) {
    return "";
  }
}

function qalExtractAlternatives_(raw) {
  const text = String(raw || "").replace(/```json/gi, "").replace(/```/g, "").trim();
  const alternatives = [];
  try {
    const parsed = JSON.parse(text);
    if (parsed && Array.isArray(parsed.alternatives)) {
      for (let i = 0; i < parsed.alternatives.length; i++) {
        const c = qalNormalizeCandidate_(parsed.alternatives[i]);
        if (c) alternatives.push(c);
      }
    }
  } catch (_e) {}
  if (alternatives.length) return alternatives;

  const lines = text.split(/\r?\n/).map(qalNormalizeCandidate_).filter(Boolean);
  for (let i = 0; i < lines.length && alternatives.length < 3; i++) {
    alternatives.push(lines[i]);
  }
  return alternatives;
}

function qalNormalizeCandidate_(value) {
  let t = String(value == null ? "" : value);
  t = t.replace(/^[-*]\s+/, "").replace(/^["']|["']$/g, "").replace(/\s+/g, " ").trim();
  return t;
}

function qalRestoreCaseStyle_(candidate, original, type) {
  let t = qalToText_(candidate);
  if (!t) return t;
  t = qalStripDanglingEnding_(t);
  if (!t) return t;
  const src = qalToText_(original);
  if (type === "headline" && src && /^[A-Z]/.test(src)) {
    t = t.charAt(0).toUpperCase() + t.slice(1);
  }
  return t;
}

function qalStripDanglingEnding_(text) {
  let t = qalToText_(text);
  if (!t) return t;
  const trailing = /\b(in|at|to|for|with|from|of|on|by|via)\s*$/i;
  let guard = 0;
  while (trailing.test(t) && guard < 3) {
    t = t.replace(trailing, "").trim();
    guard++;
  }
  return t;
}

function qalHumanScore_(text, type, adGroup, original, maxLen) {
  const t = qalToText_(text);
  if (!t) return 0;
  const banned = qalGetBlockedTerms_("BANNED_WORDS");
  const offShelf = qalGetBlockedTerms_("OFF_SHELF_WORDS");
  let score = 100;
  if (t.length > maxLen) score -= 60;
  if (t === qalToText_(original)) score -= 20;
  if (/[|:{}[\]]/.test(t)) score -= 30;
  if (/\b(in|at|to|for|with|from|of|on)\s*$/i.test(t)) score -= 25;
  if (/\bw\/\b/i.test(t) || /\bsvc\b/i.test(t)) score -= 25;
  if (/\b(location|destination|region)\b/i.test(t.toLowerCase())) score -= 25;
  if (qalContainsAnyTerm_(t, banned)) score -= 35;
  if (qalContainsAnyTerm_(t, offShelf)) score -= 30;
  if (type === "headline" && t.split(/\s+/).length < 3) score -= 20;
  if (type === "description" && t.length < 55) score -= 20;
  const adWords = qalToText_(adGroup).toLowerCase().split(/\s+/).filter(w => w.length > 2);
  if (adWords.length && !adWords.some(w => t.toLowerCase().includes(w))) score -= 20;
  return Math.max(score, 0);
}

function qalGetOrCreateReportSheet_(ss) {
  let sheet = ss.getSheetByName(TBX_REPORT_SHEET);
  if (!sheet) sheet = ss.insertSheet(TBX_REPORT_SHEET);
  return sheet;
}

function qalGetOrCreateFixLogSheet_(ss) {
  let sheet = ss.getSheetByName(TBX_FIX_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TBX_FIX_LOG_SHEET);
    sheet.appendRow([
      "Time", "RunId", "Row", "AdGroup", "AssetType", "AssetIndex",
      "Column", "Old", "New", "OldLength", "NewLength", "Mode", "Method", "HumanScore"
    ]);
  }
  return sheet;
}

function qalGetKeywordsForAdGroup_(adGroup, headers, data) {
  const normalized = qalToText_(adGroup).toLowerCase();
  const col = headers.findIndex(h => qalToText_(h).toLowerCase() === normalized);
  if (col === -1) return [];
  return data.slice(1).map(r => qalToText_(r[col])).filter(Boolean);
}

function qalGetFixMode_() {
  return qalNormalizeMode_(qalGetTextSetting_(TBX_MODE_KEY, TBX_MODE_COMPARE)) || TBX_MODE_COMPARE;
}

function qalNormalizeMode_(raw) {
  const v = qalToText_(raw).toLowerCase();
  if (v === TBX_MODE_DETERMINISTIC) return TBX_MODE_DETERMINISTIC;
  if (v === TBX_MODE_AI_FALLBACK || v === "ai" || v === "fallback") return TBX_MODE_AI_FALLBACK;
  if (v === TBX_MODE_COMPARE || v === "compare" || v === "best") return TBX_MODE_COMPARE;
  return "";
}

function qalGetTextSetting_(key, fallback) {
  const val = PropertiesService.getScriptProperties().getProperty(key);
  return val === null || val === "" ? fallback : String(val);
}

function qalGetBlockedTerms_(propertyKey) {
  const fallback = propertyKey === "BANNED_WORDS"
    ? "elevate,elevated,elevating,seamless,journey,journeys,architect,legendary,discover,bespoke,reunion,packages"
    : "packages,deals,travel packages,travel deals,reunion packages,luxury travel,travel experts,experiences";
  return qalGetTextSetting_(propertyKey, fallback)
    .split(/[\n,]/)
    .map(function(v) { return String(v || "").toLowerCase().trim(); })
    .filter(Boolean);
}

function qalContainsAnyTerm_(text, terms) {
  const lower = qalToText_(text).toLowerCase();
  for (let i = 0; i < terms.length; i++) {
    if (terms[i] && lower.includes(terms[i])) return true;
  }
  return false;
}

function qalToText_(v) {
  return (v == null ? "" : String(v)).trim();
}

function qalColToLetter_(column) {
  let temp = "";
  let letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
