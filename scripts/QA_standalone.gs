/* ===============================
QA STANDALONE (WEAK ASSET FIXER)
=============================== */

const QAS_DEFAULT_AD_SHEET = "Ad Copy";
const QAS_DEFAULT_KEYWORD_SHEET = "List Of Keywords";
const QAS_LOG_SHEET = "QA Standalone Log";
const QAS_BRAND_REPORT_SHEET = "QA Brand Compliance";
const QAS_HEADLINE_MAX = 30;
const QAS_DESCRIPTION_MAX = 90;
const QAS_MAX_FIX_ATTEMPTS = 2;
const QAS_ERROR_FLAG_COL = "QA Error";
const QAS_ERROR_MSG_COL = "QA Error Details";

function qaStandaloneFixAssets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adSheetName = qasGetTextSetting_("AD_SHEET_NAME", QAS_DEFAULT_AD_SHEET);
  const keywordSheetName = qasGetTextSetting_("KEYWORD_SHEET_NAME", QAS_DEFAULT_KEYWORD_SHEET);
  const adSheet = ss.getSheetByName(adSheetName);
  if (!adSheet) {
    SpreadsheetApp.getUi().alert("Ad sheet not found: " + adSheetName);
    return;
  }
  const keywordSheet = ss.getSheetByName(keywordSheetName);
  const keywordData = keywordSheet ? keywordSheet.getDataRange().getValues() : [];
  const keywordHeaders = keywordData[0] || [];

  const data = adSheet.getDataRange().getValues();
  if (!data.length) return;
  const errorCols = qasEnsureErrorColumns_(adSheet, data[0] || []);
  const logSheet = qasGetOrCreateLogSheet_(ss);
  const runId = "QAS_" + Date.now();

  let changes = 0;
  let skippedErrors = 0;
  let newErrors = 0;

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const rowData = data[i];
    const adGroup = qasToText_(rowData[1]);
    if (!adGroup) continue;

    const alreadyError = adSheet.getRange(row, errorCols.flagCol).getValue() === true;
    if (alreadyError) {
      skippedErrors++;
      continue;
    }

    const keywords = qasGetKeywordsForAdGroup_(adGroup, keywordHeaders, keywordData);
    let rowHadHardFailure = false;
    let rowErrorMsg = "";

    // Headlines D:R
    for (let h = 0; h < 15; h++) {
      const col = 4 + h;
      const original = qasToText_(rowData[3 + h]);
      if (!original) continue;
      const result = qasTryFixAsset_(original, "headline", QAS_HEADLINE_MAX, adGroup, keywords);
      if (result.status === "ok" && result.text !== original) {
        adSheet.getRange(row, col).setValue(result.text).setBackground("#d9ead3");
        logSheet.appendRow([new Date(), runId, row, adGroup, "Headline", h + 1, qasColToLetter_(col), original, result.text, original.length, result.text.length, result.method, result.score]);
        changes++;
      } else if (result.status === "error") {
        rowHadHardFailure = true;
        rowErrorMsg = result.message;
      }
    }

    // Descriptions S:V
    for (let d = 0; d < 4; d++) {
      const col = 19 + d;
      const original = qasToText_(rowData[18 + d]);
      if (!original) continue;
      const result = qasTryFixAsset_(original, "description", QAS_DESCRIPTION_MAX, adGroup, keywords);
      if (result.status === "ok" && result.text !== original) {
        adSheet.getRange(row, col).setValue(result.text).setBackground("#d9ead3");
        logSheet.appendRow([new Date(), runId, row, adGroup, "Description", d + 1, qasColToLetter_(col), original, result.text, original.length, result.text.length, result.method, result.score]);
        changes++;
      } else if (result.status === "error") {
        rowHadHardFailure = true;
        rowErrorMsg = result.message;
      }
    }

    if (rowHadHardFailure) {
      adSheet.getRange(row, errorCols.flagCol).setValue(true);
      adSheet.getRange(row, errorCols.msgCol).setValue(rowErrorMsg || "Unable to produce acceptable fix.");
      newErrors++;
    }
  }

  SpreadsheetApp.getUi().alert(
    "QA Standalone complete.\n" +
    "Updated assets: " + changes + "\n" +
    "New error rows: " + newErrors + "\n" +
    "Skipped error rows: " + skippedErrors
  );
}

function qaStandaloneUndoLastRun() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adSheetName = qasGetTextSetting_("AD_SHEET_NAME", QAS_DEFAULT_AD_SHEET);
  const adSheet = ss.getSheetByName(adSheetName);
  const logSheet = ss.getSheetByName(QAS_LOG_SHEET);
  if (!adSheet || !logSheet) {
    SpreadsheetApp.getUi().alert("Nothing to undo. Missing sheet(s).");
    return;
  }

  const values = logSheet.getDataRange().getValues();
  if (values.length < 2) {
    SpreadsheetApp.getUi().alert("Nothing to undo.");
    return;
  }

  const lastRunId = qasFindLastRunId_(values);
  if (!lastRunId) {
    SpreadsheetApp.getUi().alert("Nothing to undo.");
    return;
  }

  let reverted = 0;
  for (let i = values.length - 1; i >= 1; i--) {
    const row = values[i];
    if (row[1] !== lastRunId) continue;
    const sheetRow = Number(row[2]);
    const colLetter = qasToText_(row[6]);
    const oldText = qasToText_(row[7]);
    const col = qasLetterToCol_(colLetter);
    if (sheetRow > 1 && col > 0) {
      adSheet.getRange(sheetRow, col).setValue(oldText).setBackground(null);
      reverted++;
    }
  }

  SpreadsheetApp.getUi().alert("Undo complete. Reverted " + reverted + " cells from run " + lastRunId + ".");
}

function qaStandaloneClearColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adSheetName = qasGetTextSetting_("AD_SHEET_NAME", QAS_DEFAULT_AD_SHEET);
  const adSheet = ss.getSheetByName(adSheetName);
  if (!adSheet) {
    SpreadsheetApp.getUi().alert("Ad sheet not found: " + adSheetName);
    return;
  }
  const lastRow = adSheet.getLastRow();
  if (lastRow < 2) return;
  adSheet.getRange(2, 4, lastRow - 1, 19).setBackground(null);
  SpreadsheetApp.getUi().alert("Cleared QA highlight colors.");
}

function qaStandaloneBrandComplianceAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adSheetName = qasGetTextSetting_("AD_SHEET_NAME", QAS_DEFAULT_AD_SHEET);
  const adSheet = ss.getSheetByName(adSheetName);
  if (!adSheet) {
    SpreadsheetApp.getUi().alert("Ad sheet not found: " + adSheetName);
    return;
  }

  const data = adSheet.getDataRange().getValues();
  if (!data.length) return;
  const report = qasGetOrCreateBrandReportSheet_(ss);
  report.clearContents();
  report.appendRow([
    "Time", "Row", "AdGroup", "AssetType", "AssetIndex", "Column", "Issue", "Fixable", "CurrentText", "Suggested"
  ]);

  const rules = qasGetBrandRules_(ss, data);
  const now = new Date();
  let findings = 0;

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const adGroup = qasToText_(data[i][1]);
    if (!adGroup) continue;
    const headlines = data[i].slice(3, 18).map(qasToText_);
    const descriptions = data[i].slice(18, 22).map(qasToText_);
    const seenHeadline = {};

    for (let h = 0; h < 15; h++) {
      const text = headlines[h];
      if (!text) continue;
      const issues = qasGetBrandIssuesForAsset_(text, "headline", h, adGroup, rules);
      const key = qasNormalizeTextKey_(text);
      if (seenHeadline[key]) issues.push("duplicate_headline_exact");
      seenHeadline[key] = true;
      for (let j = 0; j < issues.length; j++) {
        report.appendRow([
          now, row, adGroup, "Headline", h + 1, qasColToLetter_(4 + h),
          issues[j], qasIsBrandIssueFixable_(issues[j]), text, qasGetSuggestedFix_(text, "headline", h, adGroup, rules)
        ]);
        findings++;
      }
    }

    for (let d = 0; d < 4; d++) {
      const text = descriptions[d];
      if (!text) continue;
      const issues = qasGetBrandIssuesForAsset_(text, "description", d, adGroup, rules);
      for (let j = 0; j < issues.length; j++) {
        report.appendRow([
          now, row, adGroup, "Description", d + 1, qasColToLetter_(19 + d),
          issues[j], qasIsBrandIssueFixable_(issues[j]), text, qasGetSuggestedFix_(text, "description", d, adGroup, rules)
        ]);
        findings++;
      }
    }

    const ctaIssue = qasGetRowLevelCtaIssue_(descriptions, rules);
    if (ctaIssue) {
      report.appendRow([now, row, adGroup, "Row", "", "", ctaIssue, false, descriptions.join(" | "), ""]);
      findings++;
    }
  }

  if (findings) report.autoResizeColumns(1, 10);
  SpreadsheetApp.getUi().alert("Brand compliance audit complete. Findings: " + findings);
}

function qaStandaloneBrandComplianceFix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adSheetName = qasGetTextSetting_("AD_SHEET_NAME", QAS_DEFAULT_AD_SHEET);
  const keywordSheetName = qasGetTextSetting_("KEYWORD_SHEET_NAME", QAS_DEFAULT_KEYWORD_SHEET);
  const adSheet = ss.getSheetByName(adSheetName);
  if (!adSheet) {
    SpreadsheetApp.getUi().alert("Ad sheet not found: " + adSheetName);
    return;
  }
  const keywordSheet = ss.getSheetByName(keywordSheetName);
  const keywordData = keywordSheet ? keywordSheet.getDataRange().getValues() : [];
  const keywordHeaders = keywordData[0] || [];
  const data = adSheet.getDataRange().getValues();
  if (!data.length) return;
  const logSheet = qasGetOrCreateLogSheet_(ss);
  const runId = "QASBRAND_" + Date.now();
  const rules = qasGetBrandRules_(ss, data);

  let updates = 0;
  let unresolved = 0;

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const adGroup = qasToText_(data[i][1]);
    if (!adGroup) continue;
    const keywords = qasGetKeywordsForAdGroup_(adGroup, keywordHeaders, keywordData);
    const seenHeadline = {};

    // Headlines
    for (let h = 0; h < 15; h++) {
      const col = 4 + h;
      const original = qasToText_(adSheet.getRange(row, col).getValue());
      if (!original) continue;
      const issues = qasGetBrandIssuesForAsset_(original, "headline", h, adGroup, rules);
      const originalKey = qasNormalizeTextKey_(original);
      const duplicateIssue = !!seenHeadline[originalKey];
      if (duplicateIssue) issues.push("duplicate_headline_exact");
      if (!issues.length) {
        seenHeadline[originalKey] = true;
        continue;
      }

      const fixedFromRule = qasGetSuggestedFix_(original, "headline", h, adGroup, rules);
      let nextText = fixedFromRule && fixedFromRule !== original ? fixedFromRule : "";
      if (!nextText) {
        if (duplicateIssue) {
          nextText = qasRewriteDuplicateHeadline_(original, adGroup, keywords, Object.keys(seenHeadline));
        } else {
          const attempt = qasTryFixAsset_(original, "headline", QAS_HEADLINE_MAX, adGroup, keywords);
          if (attempt.status === "ok" && attempt.text) nextText = attempt.text;
        }
      }
      if (!nextText || nextText === original) {
        unresolved++;
        seenHeadline[originalKey] = true;
        continue;
      }
      if (seenHeadline[qasNormalizeTextKey_(nextText)]) {
        unresolved++;
        seenHeadline[originalKey] = true;
        continue;
      }
      const recheck = qasGetBrandIssuesForAsset_(nextText, "headline", h, adGroup, rules);
      if (recheck.length) {
        unresolved++;
        seenHeadline[originalKey] = true;
        continue;
      }
      adSheet.getRange(row, col).setValue(nextText).setBackground("#d9ead3");
      logSheet.appendRow([new Date(), runId, row, adGroup, "Headline", h + 1, qasColToLetter_(col), original, nextText, original.length, nextText.length, "brand_compliance", 100]);
      updates++;
      seenHeadline[qasNormalizeTextKey_(nextText)] = true;
    }

    // Descriptions
    for (let d = 0; d < 4; d++) {
      const col = 19 + d;
      const original = qasToText_(adSheet.getRange(row, col).getValue());
      if (!original) continue;
      const issues = qasGetBrandIssuesForAsset_(original, "description", d, adGroup, rules);
      if (!issues.length) continue;

      const fixedFromRule = qasGetSuggestedFix_(original, "description", d, adGroup, rules);
      let nextText = fixedFromRule && fixedFromRule !== original ? fixedFromRule : "";
      if (!nextText) {
        const attempt = qasTryFixAsset_(original, "description", QAS_DESCRIPTION_MAX, adGroup, keywords);
        if (attempt.status === "ok" && attempt.text) nextText = attempt.text;
      }
      if (!nextText || nextText === original) {
        unresolved++;
        continue;
      }
      const recheck = qasGetBrandIssuesForAsset_(nextText, "description", d, adGroup, rules);
      if (recheck.length) {
        unresolved++;
        continue;
      }
      adSheet.getRange(row, col).setValue(nextText).setBackground("#d9ead3");
      logSheet.appendRow([new Date(), runId, row, adGroup, "Description", d + 1, qasColToLetter_(col), original, nextText, original.length, nextText.length, "brand_compliance", 100]);
      updates++;
    }
  }

  SpreadsheetApp.getUi().alert(
    "Brand compliance fix complete.\n" +
    "Updated assets: " + updates + "\n" +
    "Unresolved assets: " + unresolved
  );
}

function qasTryFixAsset_(original, type, maxLen, adGroup, keywords) {
  const text = qasToText_(original);
  if (!text) return { status: "skip", text: text, method: "none", score: 0 };

  const issue = qasAnalyzeAsset_(text, type, maxLen);
  if (!issue.hasIssue) {
    return { status: "ok", text: text, method: "none", score: 100 };
  }

  let best = "";
  let bestScore = -1;
  for (let attempt = 1; attempt <= QAS_MAX_FIX_ATTEMPTS; attempt++) {
    const deterministic = qasShortenToLimit_(text, maxLen, type);
    const detScore = qasHumanScore_(deterministic, type, adGroup, text, maxLen);
    if (detScore > bestScore) {
      best = deterministic;
      bestScore = detScore;
    }

    const ai = qasAiRewriteToLimit_(text, type, maxLen, adGroup, keywords);
    if (ai) {
      const aiScore = qasHumanScore_(ai, type, adGroup, text, maxLen);
      if (aiScore > bestScore) {
        best = ai;
        bestScore = aiScore;
      }
    }

    if (qasIsAcceptableRewrite_(best, text, type, maxLen)) {
      return { status: "ok", text: best, method: "det+ai", score: bestScore };
    }
  }

  return {
    status: "error",
    text: text,
    method: "failed",
    score: bestScore,
    message: "Could not create acceptable " + type + " within " + maxLen + " chars."
  };
}

function qasAnalyzeAsset_(text, type, maxLen) {
  const t = qasToText_(text);
  if (!t) return { hasIssue: false };
  const issues = [];
  const banned = qasGetBlockedTerms_("BANNED_WORDS");
  const offShelf = qasGetBlockedTerms_("OFF_SHELF_WORDS");
  if (t.length > maxLen) issues.push("too_long");
  if (/\b(in|at|to|for|with|from|of|on|by|via)\s*$/i.test(t)) issues.push("dangling_ending");
  if (/\bw\/\b/i.test(t) || /\bsvc\b/i.test(t)) issues.push("bad_abbrev");
  if (qasContainsAnyTerm_(t, banned)) issues.push("banned_word");
  if (qasContainsAnyTerm_(t, offShelf)) issues.push("off_shelf_phrase");
  if (type === "headline" && t.split(/\s+/).length < 2) issues.push("too_short");
  return { hasIssue: issues.length > 0, issues: issues };
}

function qasIsAcceptableRewrite_(candidate, original, type, maxLen) {
  const c = qasToText_(candidate);
  const o = qasToText_(original);
  const banned = qasGetBlockedTerms_("BANNED_WORDS");
  const offShelf = qasGetBlockedTerms_("OFF_SHELF_WORDS");
  if (!c) return false;
  if (c.length > maxLen) return false;
  if (/\b(in|at|to|for|with|from|of|on|by|via)\s*$/i.test(c)) return false;
  if (/\bw\/\b/i.test(c) || /\bsvc\b/i.test(c)) return false;
  if (qasContainsAnyTerm_(c, banned)) return false;
  if (qasContainsAnyTerm_(c, offShelf)) return false;
  if (c === o && o.length > maxLen) return false;
  return true;
}

function qasShortenToLimit_(text, maxLen, type) {
  const original = qasToText_(text);
  let t = original;
  if (!t) return t;
  if (t.length <= maxLen) return qasRestoreCaseStyle_(t, original, type);

  t = t.replace(/([a-z])([A-Z])/g, "$1 $2");
  t = t.replace(/([A-Za-z])\/([A-Za-z])/g, "$1 $2");
  t = t.replace(/\b(the|a|an|very|really|that|which|just|simply|completely|fully|carefully)\b/gi, "");
  t = t.replace(/\s+/g, " ").trim();
  if (t.length <= maxLen) return qasRestoreCaseStyle_(t, original, type);

  t = t.split(/\s+/).map(function(w) {
    if (w.length > 4 && /s$/i.test(w) && !/(ss|us|is)$/i.test(w)) return w.slice(0, -1);
    return w;
  }).join(" ");
  t = t.replace(/\s+/g, " ").trim();
  if (t.length <= maxLen) return qasRestoreCaseStyle_(t, original, type);

  const sliced = t.slice(0, maxLen + 1);
  const lastSpace = sliced.lastIndexOf(" ");
  if (lastSpace > Math.floor(maxLen * 0.6)) return qasRestoreCaseStyle_(sliced.slice(0, lastSpace).trim(), original, type);
  return qasRestoreCaseStyle_(t.slice(0, maxLen).trim(), original, type);
}

function qasAiRewriteToLimit_(original, type, maxLen, adGroup, keywords) {
  const apiKey = qasGetTextSetting_("OPENAI_API_KEY", "");
  if (!apiKey) return "";
  const model = qasGetTextSetting_("OPENAI_MODEL", "gpt-4o-mini");
  const banned = qasGetBlockedTerms_("BANNED_WORDS");
  const offShelf = qasGetBlockedTerms_("OFF_SHELF_WORDS");
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
    const candidates = qasExtractAlternatives_(out);
    if (!candidates.length) return "";

    let best = "";
    let bestScore = -1;
    for (let i = 0; i < candidates.length; i++) {
      let c = candidates[i];
      if (c.length > maxLen) c = qasShortenToLimit_(c, maxLen, type);
      c = qasRestoreCaseStyle_(c, original, type);
      const s = qasHumanScore_(c, type, adGroup, original, maxLen);
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

function qasExtractAlternatives_(raw) {
  const text = String(raw || "").replace(/```json/gi, "").replace(/```/g, "").trim();
  const alternatives = [];
  try {
    const parsed = JSON.parse(text);
    if (parsed && Array.isArray(parsed.alternatives)) {
      for (let i = 0; i < parsed.alternatives.length; i++) {
        const c = qasNormalizeCandidate_(parsed.alternatives[i]);
        if (c) alternatives.push(c);
      }
    }
  } catch (_e) {}
  if (alternatives.length) return alternatives;

  const lines = text.split(/\r?\n/).map(qasNormalizeCandidate_).filter(Boolean);
  for (let i = 0; i < lines.length && alternatives.length < 3; i++) {
    alternatives.push(lines[i]);
  }
  return alternatives;
}

function qasNormalizeCandidate_(value) {
  let t = String(value == null ? "" : value);
  t = t.replace(/^[-*]\s+/, "").replace(/^["']|["']$/g, "").replace(/\s+/g, " ").trim();
  return t;
}

function qasRestoreCaseStyle_(candidate, original, type) {
  let t = qasToText_(candidate);
  if (!t) return t;
  t = qasStripDanglingEnding_(t);
  if (!t) return t;
  const src = qasToText_(original);
  if (type === "headline" && src && /^[A-Z]/.test(src)) {
    t = t.charAt(0).toUpperCase() + t.slice(1);
  }
  return t;
}

function qasStripDanglingEnding_(text) {
  let t = qasToText_(text);
  if (!t) return t;
  const trailing = /\b(in|at|to|for|with|from|of|on|by|via)\s*$/i;
  let guard = 0;
  while (trailing.test(t) && guard < 3) {
    t = t.replace(trailing, "").trim();
    guard++;
  }
  return t;
}

function qasHumanScore_(text, type, adGroup, original, maxLen) {
  const t = qasToText_(text);
  if (!t) return 0;
  const banned = qasGetBlockedTerms_("BANNED_WORDS");
  const offShelf = qasGetBlockedTerms_("OFF_SHELF_WORDS");
  let score = 100;
  if (t.length > maxLen) score -= 60;
  if (t === qasToText_(original)) score -= 20;
  if (/[|:{}[\]]/.test(t)) score -= 30;
  if (/\b(in|at|to|for|with|from|of|on)\s*$/i.test(t)) score -= 25;
  if (/\bw\/\b/i.test(t) || /\bsvc\b/i.test(t)) score -= 25;
  if (/\b(location|destination|region)\b/i.test(t.toLowerCase())) score -= 25;
  if (qasContainsAnyTerm_(t, banned)) score -= 35;
  if (qasContainsAnyTerm_(t, offShelf)) score -= 30;
  if (type === "headline" && t.split(/\s+/).length < 3) score -= 20;
  if (type === "description" && t.length < 55) score -= 20;
  const adWords = qasToText_(adGroup).toLowerCase().split(/\s+/).filter(function(w) { return w.length > 2; });
  if (adWords.length && !adWords.some(function(w) { return t.toLowerCase().includes(w); })) score -= 20;
  return Math.max(score, 0);
}

function qasGetBrandRules_(ss, adData) {
  const specificHeadlineCount = qasGetNumberSetting_("SPECIFIC_HEADLINE_COUNT", 5, 3, 5);
  const specificDescriptionCount = qasGetNumberSetting_("SPECIFIC_DESCRIPTION_COUNT", 2, 1, 3);
  const ctaMode = qasGetEnumSetting_("CTA_MODE", "balanced", ["branded", "experience", "balanced"]);
  const ctaBranded = qasParseList_(qasGetTextSetting_("CTA_BRANDED_LIST", "Request a Custom Proposal"));
  const ctaExperience = qasParseList_(qasGetTextSetting_("CTA_EXPERIENCE_LIST", "Start Your Private Guided Experience"));
  const ctaPhrases = ctaMode === "branded"
    ? ctaBranded
    : ctaMode === "experience"
      ? ctaExperience
      : ctaBranded.concat(ctaExperience).filter(Boolean);
  const fixedAssets = qasReadFixedAssetsConfig_(ss, qasGetTextSetting_("FIXED_ASSETS_SHEET_NAME", "Brand Fixed Assets"));
  const angleLibrary = qasReadAngleLibrary_(ss, qasGetTextSetting_("ANGLE_LIBRARY_SHEET_NAME", "Brand Angles"));
  const selectedHeadlineAngles = qasParseIdList_(qasGetTextSetting_("SELECTED_HEADLINE_ANGLES", ""), 4);
  const selectedDescriptionAngles = qasParseIdList_(qasGetTextSetting_("SELECTED_DESCRIPTION_ANGLES", ""), 3);
  const angleAssignmentsHeadlines = qasParseAssignmentMap_(qasGetTextSetting_("ANGLE_SLOT_ASSIGNMENTS_HEADLINES", ""), 15);
  const angleAssignmentsDescriptions = qasParseAssignmentMap_(qasGetTextSetting_("ANGLE_SLOT_ASSIGNMENTS_DESCRIPTIONS", ""), 4);
  const sharedHeadlineBaseline = qasGetSharedBaseline_(adData, 3, 18, specificHeadlineCount, 15);
  const sharedDescriptionBaseline = qasGetSharedBaseline_(adData, 18, 22, specificDescriptionCount, 4);
  return {
    specificHeadlineCount: specificHeadlineCount,
    specificDescriptionCount: specificDescriptionCount,
    lockSharedHeadlines: qasGetBoolSetting_("LOCK_SHARED_HEADLINES", false),
    lockSharedDescriptions: qasGetBoolSetting_("LOCK_SHARED_DESCRIPTIONS", false),
    fixedHeadlineByPosition: fixedAssets.headlineByPosition || {},
    fixedDescriptionByPosition: fixedAssets.descriptionByPosition || {},
    angleLibrary: angleLibrary,
    selectedHeadlineAngles: selectedHeadlineAngles,
    selectedDescriptionAngles: selectedDescriptionAngles,
    angleAssignmentsHeadlines: angleAssignmentsHeadlines,
    angleAssignmentsDescriptions: angleAssignmentsDescriptions,
    socialProofGoogle5StarEnabled: qasGetBoolSetting_("SOCIAL_PROOF_GOOGLE_5STAR_ENABLED", false),
    socialProofGoogle5StarText: qasGetTextSetting_("SOCIAL_PROOF_GOOGLE_5STAR_TEXT", "Rated 5 stars across Google"),
    sharedHeadlineBaseline: sharedHeadlineBaseline,
    sharedDescriptionBaseline: sharedDescriptionBaseline,
    ctaPhrases: ctaPhrases
  };
}

function qasGetBrandIssuesForAsset_(text, type, index, adGroup, rules) {
  const issues = [];
  const maxLen = type === "headline" ? QAS_HEADLINE_MAX : QAS_DESCRIPTION_MAX;
  const base = qasAnalyzeAsset_(text, type, maxLen);
  if (base && base.hasIssue) {
    for (let i = 0; i < base.issues.length; i++) issues.push(base.issues[i]);
  }

  const slot = index + 1;
  if (type === "headline") {
    const expected = rules.fixedHeadlineByPosition[slot];
    if (expected && qasToText_(text) !== qasToText_(expected)) issues.push("fixed_position_mismatch");
    if (rules.lockSharedHeadlines && slot > rules.specificHeadlineCount) {
      const sharedExpected = rules.sharedHeadlineBaseline[slot - rules.specificHeadlineCount];
      if (sharedExpected && qasToText_(text) !== qasToText_(sharedExpected)) issues.push("shared_headline_mismatch");
    }
    if (slot <= rules.specificHeadlineCount && !qasLooksSpecificToAdGroup_(text, adGroup)) {
      issues.push("not_specific_to_ad_group");
    }
    const assignedAngleId = qasFindAssignedAngleIdForSlot_(rules.angleAssignmentsHeadlines, slot);
    if (assignedAngleId && !qasTextMatchesAngle_(text, qasGetAngleById_(rules.angleLibrary, assignedAngleId))) {
      issues.push("angle_assignment_mismatch");
    }
  } else {
    const expected = rules.fixedDescriptionByPosition[slot];
    if (expected && qasToText_(text) !== qasToText_(expected)) issues.push("fixed_position_mismatch");
    if (rules.lockSharedDescriptions && slot > rules.specificDescriptionCount) {
      const sharedExpected = rules.sharedDescriptionBaseline[slot - rules.specificDescriptionCount];
      if (sharedExpected && qasToText_(text) !== qasToText_(sharedExpected)) issues.push("shared_description_mismatch");
    }
    if (slot <= rules.specificDescriptionCount && !qasLooksSpecificToAdGroup_(text, adGroup)) {
      issues.push("not_specific_to_ad_group");
    }
    const assignedAngleId = qasFindAssignedAngleIdForSlot_(rules.angleAssignmentsDescriptions, slot);
    if (assignedAngleId && !qasTextMatchesAngle_(text, qasGetAngleById_(rules.angleLibrary, assignedAngleId))) {
      issues.push("angle_assignment_mismatch");
    }
  }

  return qasUnique_(issues);
}

function qasGetRowLevelCtaIssue_(descriptions, rules) {
  const lines = (descriptions || []).map(qasToText_).filter(Boolean);
  if (!lines.length || !rules.ctaPhrases || !rules.ctaPhrases.length) return "";
  const lower = lines.join(" || ").toLowerCase();
  const hasCta = rules.ctaPhrases.some(function(p) {
    const t = qasToText_(p).toLowerCase();
    return t && lower.includes(t);
  });
  if (!hasCta) return "missing_preferred_cta_phrase";
  if (rules.socialProofGoogle5StarEnabled) {
    const proofNeedle = qasToText_(rules.socialProofGoogle5StarText).toLowerCase();
    const hasSocialProof = proofNeedle && lower.includes(proofNeedle);
    if (!hasSocialProof) return "missing_google_5star_social_proof";
  }
  return "";
}

function qasGetSuggestedFix_(text, type, index, adGroup, rules) {
  const slot = index + 1;
  if (type === "headline") {
    const fixed = rules.fixedHeadlineByPosition[slot];
    if (fixed) return fixed;
    const assignedAngleId = qasFindAssignedAngleIdForSlot_(rules.angleAssignmentsHeadlines, slot);
    const assignedAngleText = qasGetAngleFirstExample_(qasGetAngleById_(rules.angleLibrary, assignedAngleId));
    if (assignedAngleText) return assignedAngleText;
    if (rules.lockSharedHeadlines && slot > rules.specificHeadlineCount) {
      const shared = rules.sharedHeadlineBaseline[slot - rules.specificHeadlineCount];
      if (shared) return shared;
    }
  } else {
    const fixed = rules.fixedDescriptionByPosition[slot];
    if (fixed) return fixed;
    const assignedAngleId = qasFindAssignedAngleIdForSlot_(rules.angleAssignmentsDescriptions, slot);
    const assignedAngleText = qasGetAngleFirstExample_(qasGetAngleById_(rules.angleLibrary, assignedAngleId));
    if (assignedAngleText) return assignedAngleText;
    if (rules.lockSharedDescriptions && slot > rules.specificDescriptionCount) {
      const shared = rules.sharedDescriptionBaseline[slot - rules.specificDescriptionCount];
      if (shared) return shared;
    }
  }
  return "";
}

function qasIsBrandIssueFixable_(issue) {
  const autoFixable = {
    too_long: true,
    dangling_ending: true,
    bad_abbrev: true,
    banned_word: true,
    off_shelf_phrase: true,
    fixed_position_mismatch: true,
    shared_headline_mismatch: true,
    shared_description_mismatch: true,
    not_specific_to_ad_group: true,
    angle_assignment_mismatch: true,
    duplicate_headline_exact: true
  };
  return !!autoFixable[String(issue || "")];
}

function qasEnsureErrorColumns_(sheet, headerRow) {
  let headers = headerRow.slice();
  let flagIndex = headers.indexOf(QAS_ERROR_FLAG_COL);
  let msgIndex = headers.indexOf(QAS_ERROR_MSG_COL);
  let lastCol = sheet.getLastColumn();

  if (flagIndex === -1) {
    lastCol++;
    sheet.getRange(1, lastCol).setValue(QAS_ERROR_FLAG_COL);
    headers.push(QAS_ERROR_FLAG_COL);
    flagIndex = headers.length - 1;
  }
  if (msgIndex === -1) {
    lastCol++;
    sheet.getRange(1, lastCol).setValue(QAS_ERROR_MSG_COL);
    headers.push(QAS_ERROR_MSG_COL);
    msgIndex = headers.length - 1;
  }
  return { flagCol: flagIndex + 1, msgCol: msgIndex + 1 };
}

function qasGetOrCreateLogSheet_(ss) {
  let sheet = ss.getSheetByName(QAS_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(QAS_LOG_SHEET);
    sheet.appendRow([
      "Time", "RunId", "Row", "AdGroup", "AssetType", "AssetIndex",
      "Column", "Old", "New", "OldLength", "NewLength", "Method", "HumanScore"
    ]);
  }
  return sheet;
}

function qasGetKeywordsForAdGroup_(adGroup, headers, data) {
  const normalized = qasToText_(adGroup).toLowerCase();
  const col = headers.findIndex(function(h) { return qasToText_(h).toLowerCase() === normalized; });
  if (col === -1) return [];
  return data.slice(1).map(function(r) { return qasToText_(r[col]); }).filter(Boolean);
}

function qasFindLastRunId_(logValues) {
  for (let i = logValues.length - 1; i >= 1; i--) {
    const runId = qasToText_(logValues[i][1]);
    if (runId) return runId;
  }
  return "";
}

function qasGetTextSetting_(key, fallback) {
  const val = PropertiesService.getScriptProperties().getProperty(key);
  return val === null || val === "" ? fallback : String(val);
}

function qasGetBoolSetting_(key, fallback) {
  const val = PropertiesService.getScriptProperties().getProperty(key);
  if (val === null) return !!fallback;
  return String(val).toLowerCase() === "true";
}

function qasGetNumberSetting_(key, fallback, min, max) {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  let n = Number(raw);
  if (!isFinite(n)) n = fallback;
  if (typeof min === "number" && n < min) n = min;
  if (typeof max === "number" && n > max) n = max;
  return n;
}

function qasGetEnumSetting_(key, fallback, allowed) {
  const val = String(qasGetTextSetting_(key, fallback) || "").toLowerCase().trim();
  return allowed.indexOf(val) >= 0 ? val : fallback;
}

function qasParseList_(text) {
  return String(text == null ? "" : text)
    .split(/[\n,]/)
    .map(function(v) { return String(v || "").trim(); })
    .filter(Boolean);
}

function qasNormalizeTextKey_(text) {
  return qasToText_(text).toLowerCase().replace(/\s+/g, " ").trim();
}

function qasRewriteDuplicateHeadline_(headline, adGroup, keywords, usedKeys) {
  const used = {};
  (usedKeys || []).forEach(function(k) { used[String(k || "").toLowerCase()] = true; });
  const base = qasToText_(headline);
  if (!base) return "";

  const adTokens = qasToText_(adGroup).split(/\s+/).filter(function(t) { return t.length > 2; });
  for (let i = 0; i < adTokens.length; i++) {
    const candidate = (base + " " + adTokens[i]).replace(/\s+/g, " ").trim();
    if (candidate.length > QAS_HEADLINE_MAX) continue;
    if (!used[qasNormalizeTextKey_(candidate)]) return candidate;
  }

  const ai = qasAiRewriteToLimit_(base, "headline", QAS_HEADLINE_MAX, adGroup, keywords);
  const aiKey = qasNormalizeTextKey_(ai);
  if (ai && !used[aiKey]) return ai;
  return "";
}

function qasParseIdList_(text, maxCount) {
  const seen = {};
  const out = [];
  const parts = String(text == null ? "" : text).split(/[\n,]/);
  for (let i = 0; i < parts.length; i++) {
    const id = String(parts[i] || "").toLowerCase().trim();
    if (!id || seen[id]) continue;
    seen[id] = true;
    out.push(id);
    if (typeof maxCount === "number" && out.length >= maxCount) break;
  }
  return out;
}

function qasParseAssignmentMap_(text, maxSlot) {
  const out = {};
  const entries = String(text == null ? "" : text).split(/[\n,]/);
  for (let i = 0; i < entries.length; i++) {
    const token = String(entries[i] || "").trim();
    if (!token) continue;
    const m = token.match(/^([a-z0-9 _-]+)\s*:\s*(\d+)$/i);
    if (!m) continue;
    const id = String(m[1] || "").toLowerCase().trim();
    const slot = Number(m[2]);
    if (!id || !isFinite(slot)) continue;
    if (slot < 1 || slot > maxSlot) continue;
    out[id] = slot;
  }
  return out;
}

function qasReadAngleLibrary_(ss, sheetName) {
  const name = String(sheetName || "Brand Angles");
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headers = (values[0] || []).map(function(v) { return qasToText_(v).toLowerCase(); });
  const idxEnabled = qasFindHeaderIndex_(headers, ["enabled", "active", "apply"]);
  const idxId = qasFindHeaderIndex_(headers, ["angleid", "angle id", "id"]);
  const idxName = qasFindHeaderIndex_(headers, ["anglename", "angle name", "name"]);
  const idxType = qasFindHeaderIndex_(headers, ["type"]);
  const idxFocus = qasFindHeaderIndex_(headers, ["focus"]);
  const idxExamples = qasFindHeaderIndex_(headers, ["examples", "example"]);
  if (idxId < 0) return [];

  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const enabled = idxEnabled < 0 ? true : qasToBoolCell_(row[idxEnabled]);
    if (!enabled) continue;
    const id = qasToText_(row[idxId]).toLowerCase();
    if (!id) continue;
    out.push({
      id: id,
      name: idxName >= 0 ? qasToText_(row[idxName]) : id,
      type: idxType >= 0 ? qasToText_(row[idxType]).toLowerCase() : "both",
      focus: idxFocus >= 0 ? qasToText_(row[idxFocus]) : "",
      examples: idxExamples >= 0 ? qasToText_(row[idxExamples]) : ""
    });
  }
  return out;
}

function qasFindAssignedAngleIdForSlot_(assignmentMap, slot) {
  const target = Number(slot);
  const keys = Object.keys(assignmentMap || {});
  for (let i = 0; i < keys.length; i++) {
    const id = keys[i];
    if (Number(assignmentMap[id]) === target) return id;
  }
  return "";
}

function qasGetAngleById_(library, angleId) {
  const id = qasToText_(angleId).toLowerCase();
  if (!id) return null;
  for (let i = 0; i < (library || []).length; i++) {
    if (qasToText_(library[i].id).toLowerCase() === id) return library[i];
  }
  return null;
}

function qasGetAngleFirstExample_(angle) {
  if (!angle) return "";
  const lines = String(angle.examples == null ? "" : angle.examples)
    .split(/[\n|]/)
    .map(function(v) { return qasToText_(v); })
    .filter(Boolean);
  return lines.length ? lines[0] : "";
}

function qasTextMatchesAngle_(text, angle) {
  if (!angle) return true;
  const lower = qasToText_(text).toLowerCase();
  if (!lower) return false;
  const tokens = []
    .concat(qasToText_(angle.id).toLowerCase().split(/[-\s_]+/))
    .concat(qasToText_(angle.name).toLowerCase().split(/\s+/))
    .concat(qasToText_(angle.focus).toLowerCase().split(/\s+/))
    .filter(function(t) { return t && t.length > 3; });
  if (!tokens.length) return true;
  return tokens.some(function(t) { return lower.includes(t); });
}

function qasGetOrCreateBrandReportSheet_(ss) {
  let sheet = ss.getSheetByName(QAS_BRAND_REPORT_SHEET);
  if (!sheet) sheet = ss.insertSheet(QAS_BRAND_REPORT_SHEET);
  return sheet;
}

function qasReadFixedAssetsConfig_(ss, sheetName) {
  const name = String(sheetName || "Brand Fixed Assets");
  const sheet = ss.getSheetByName(name);
  if (!sheet) return { headlineByPosition: {}, descriptionByPosition: {} };

  const values = sheet.getDataRange().getValues();
  const headers = (values[0] || []).map(function(v) { return qasToText_(v).toLowerCase(); });
  const idxEnabled = qasFindHeaderIndex_(headers, ["enabled", "active", "apply"]);
  const idxType = qasFindHeaderIndex_(headers, ["type", "asset type"]);
  const idxPosition = qasFindHeaderIndex_(headers, ["position", "slot", "fixed position"]);
  const idxText = qasFindHeaderIndex_(headers, ["text", "asset", "copy"]);
  if (idxType < 0 || idxPosition < 0 || idxText < 0) return { headlineByPosition: {}, descriptionByPosition: {} };

  const headlineByPosition = {};
  const descriptionByPosition = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const enabled = idxEnabled < 0 ? true : qasToBoolCell_(row[idxEnabled]);
    if (!enabled) continue;
    const type = qasToText_(row[idxType]).toLowerCase();
    const position = Number(row[idxPosition]);
    const text = qasToText_(row[idxText]);
    if (!text || !isFinite(position)) continue;
    if (type === "headline" && position >= 1 && position <= 15 && text.length <= QAS_HEADLINE_MAX) {
      headlineByPosition[position] = text;
    } else if (type === "description" && position >= 1 && position <= 4 && text.length <= QAS_DESCRIPTION_MAX) {
      descriptionByPosition[position] = text;
    }
  }
  return { headlineByPosition: headlineByPosition, descriptionByPosition: descriptionByPosition };
}

function qasGetSharedBaseline_(adData, startColIdx, endColIdx, specificCount, totalSlots) {
  const start = Math.max(0, Math.min(totalSlots - 1, Number(specificCount || 0)));
  for (let i = 1; i < adData.length; i++) {
    const assets = adData[i].slice(startColIdx, endColIdx).map(qasToText_);
    const shared = assets.slice(start, totalSlots);
    if (shared.length && shared.every(Boolean)) {
      const out = {};
      for (let p = start + 1; p <= totalSlots; p++) {
        out[p - start] = shared[p - start - 1];
      }
      return out;
    }
  }
  return {};
}

function qasLooksSpecificToAdGroup_(text, adGroup) {
  const t = qasToText_(text).toLowerCase();
  const words = qasToText_(adGroup).toLowerCase().split(/\s+/).filter(function(w) { return w.length > 2; });
  if (!words.length) return true;
  return words.some(function(w) { return t.includes(w); });
}

function qasFindHeaderIndex_(headers, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const idx = headers.indexOf(String(aliases[i] || "").toLowerCase());
    if (idx >= 0) return idx;
  }
  return -1;
}

function qasToBoolCell_(v) {
  if (v === true) return true;
  if (v === false) return false;
  const s = qasToText_(v).toLowerCase();
  if (!s) return true;
  if (s === "0" || s === "false" || s === "no" || s === "n" || s === "off") return false;
  return true;
}

function qasUnique_(arr) {
  const seen = {};
  const out = [];
  for (let i = 0; i < arr.length; i++) {
    const k = String(arr[i] || "");
    if (!k || seen[k]) continue;
    seen[k] = true;
    out.push(k);
  }
  return out;
}

function qasGetBlockedTerms_(propertyKey) {
  const fallback = propertyKey === "BANNED_WORDS"
    ? "elevate,elevated,elevating,seamless,journey,journeys,architect,legendary,discover,bespoke,reunion,packages"
    : "packages,deals,travel packages,travel deals,reunion packages,luxury travel,travel experts,experiences";
  return qasGetTextSetting_(propertyKey, fallback)
    .split(/[\n,]/)
    .map(function(v) { return String(v || "").toLowerCase().trim(); })
    .filter(Boolean);
}

function qasContainsAnyTerm_(text, terms) {
  const lower = qasToText_(text).toLowerCase();
  for (let i = 0; i < terms.length; i++) {
    if (terms[i] && lower.includes(terms[i])) return true;
  }
  return false;
}

function qasToText_(v) {
  return (v == null ? "" : String(v)).trim();
}

function qasColToLetter_(column) {
  let temp = "";
  let letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function qasLetterToCol_(letters) {
  const s = qasToText_(letters).toUpperCase();
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);
    if (code < 65 || code > 90) return 0;
    n = n * 26 + (code - 64);
  }
  return n;
}
