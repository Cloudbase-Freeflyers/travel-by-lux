/* ===============================
QA STANDALONE (UNIVERSAL BRAND FIXER)
=============================== */

const QAS_DEFAULT_AD_SHEET = "Search";
const QAS_DEFAULT_KEYWORD_SHEET = "List Of Keywords";
const QAS_DEFAULT_LANDING_PAGE_SHEET = "Landing Pages";
const QAS_DEFAULT_AD_TYPE_TABS = "Search,pmax";
const QAS_DEFAULT_AD_TYPE_SHEET_MAP = "search:Search,pmax:pmax";
const QAS_LOG_SHEET = "QA Standalone Log";
const QAS_BRAND_REPORT_SHEET = "QA Brand Compliance";
const QAS_HEADLINE_MAX = 30;
const QAS_DESCRIPTION_MAX = 90;
const QAS_MAX_FIX_ATTEMPTS = 2;
const QAS_ERROR_FLAG_COL = "QA Error";
const QAS_ERROR_MSG_COL = "QA Error Details";
const QAS_HARD_BLOCKED_TERMS = ["join", "quora", "blog", "blogs", "competitors"];

function qaStandaloneFixAssets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adSheets = qasGetTargetAdSheets_(ss);
  if (!adSheets.length) {
    SpreadsheetApp.getUi().alert("No configured ad sheets found.");
    return;
  }
  const keywordSheetName = qasGetTextSetting_("KEYWORD_SHEET_NAME", QAS_DEFAULT_KEYWORD_SHEET);
  const keywordSheet = ss.getSheetByName(keywordSheetName);
  const keywordData = keywordSheet ? keywordSheet.getDataRange().getValues() : [];
  const keywordHeaders = keywordData[0] || [];
  const useLandingPageContext = qasGetBoolSetting_("USE_LANDING_PAGE_CONTEXT", false);
  const landingPageSheetName = qasGetTextSetting_("LANDING_PAGE_SHEET_NAME", QAS_DEFAULT_LANDING_PAGE_SHEET);
  const landingSheet = useLandingPageContext ? ss.getSheetByName(landingPageSheetName) : null;
  const landingData = landingSheet ? landingSheet.getDataRange().getValues() : [];
  const landingCache = { byAdGroup: {}, byUrl: {} };
  const logSheet = qasGetOrCreateLogSheet_(ss);
  const runId = "QAS_" + Date.now();
  let changes = 0;
  let retriedErrorRows = 0;
  let newErrors = 0;
  let stopped = false;

  for (let s = 0; s < adSheets.length; s++) {
    if (qasStopRequested_()) { stopped = true; break; }
    const adSheet = adSheets[s];
    const data = adSheet.getDataRange().getValues();
    if (!data.length) continue;
    const schema = qasDetectAssetColumns_(data[0] || []);
    if (schema.adGroupCol < 0 || !schema.headlineCols.length || !schema.descriptionCols.length) continue;
    const errorCols = qasEnsureErrorColumns_(adSheet, data[0] || []);

    for (let i = 1; i < data.length; i++) {
      if (qasStopRequested_()) { stopped = true; break; }
      const row = i + 1;
      const adGroup = qasToText_(data[i][schema.adGroupCol]);
      if (!adGroup) continue;
      const alreadyError = adSheet.getRange(row, errorCols.flagCol).getValue() === true;
      if (alreadyError) retriedErrorRows++;
      const keywords = qasGetKeywordsForAdGroup_(adGroup, keywordHeaders, keywordData);
      const landingContext = useLandingPageContext ? qasGetLandingPageContextForAdGroup_(adGroup, landingData, landingCache) : "";
      let rowHadHardFailure = false;
      let rowErrorMsg = "";

      for (let h = 0; h < schema.headlineCols.length; h++) {
        const meta = schema.headlineCols[h];
        const col = meta.col + 1;
        const original = qasToText_(data[i][meta.col]);
        if (!original) continue;
        const result = qasTryFixAsset_(original, "headline", QAS_HEADLINE_MAX, adGroup, keywords, [], landingContext);
        if (result.status === "ok" && result.text !== original) {
          adSheet.getRange(row, col).setValue(result.text).setBackground("#d9ead3");
          logSheet.appendRow([new Date(), runId, row, adSheet.getName() + " | " + adGroup, "Headline", meta.slot, qasColToLetter_(col), original, result.text, original.length, result.text.length, result.method, result.score]);
          changes++;
        } else if (result.status === "error") {
          rowHadHardFailure = true;
          rowErrorMsg = result.message;
        }
      }

      for (let d = 0; d < schema.descriptionCols.length; d++) {
        const meta = schema.descriptionCols[d];
        const col = meta.col + 1;
        const original = qasToText_(data[i][meta.col]);
        if (!original) continue;
        const result = qasTryFixAsset_(original, "description", QAS_DESCRIPTION_MAX, adGroup, keywords, [], landingContext);
        if (result.status === "ok" && result.text !== original) {
          adSheet.getRange(row, col).setValue(result.text).setBackground("#d9ead3");
          logSheet.appendRow([new Date(), runId, row, adSheet.getName() + " | " + adGroup, "Description", meta.slot, qasColToLetter_(col), original, result.text, original.length, result.text.length, result.method, result.score]);
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
      } else if (alreadyError) {
        adSheet.getRange(row, errorCols.flagCol).setValue(false);
        adSheet.getRange(row, errorCols.msgCol).setValue("");
      }
    }
    if (stopped) break;
  }

  SpreadsheetApp.getUi().alert(
    "QA Standalone complete.\n" +
    "Updated assets: " + changes + "\n" +
    "New error rows: " + newErrors + "\n" +
    "Retried previously errored rows: " + retriedErrorRows +
    (stopped ? "\nStopped early by user request." : "")
  );
}

function qaStandaloneBrandComplianceAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adSheets = qasGetTargetAdSheets_(ss);
  if (!adSheets.length) {
    SpreadsheetApp.getUi().alert("No configured ad sheets found.");
    return;
  }
  const report = qasGetOrCreateBrandReportSheet_(ss);
  report.clearContents();
  report.appendRow(["Time", "Sheet", "Row", "AdGroup", "AssetType", "AssetIndex", "Column", "Issue", "Fixable", "CurrentText", "Suggested"]);
  const now = new Date();
  let findings = 0;
  let scannedSheets = 0;
  let stopped = false;

  for (let s = 0; s < adSheets.length; s++) {
    if (qasStopRequested_()) { stopped = true; break; }
    const adSheet = adSheets[s];
    const data = adSheet.getDataRange().getValues();
    if (!data.length) continue;
    const schema = qasDetectAssetColumns_(data[0] || []);
    if (schema.adGroupCol < 0 || !schema.headlineCols.length || !schema.descriptionCols.length) continue;
    scannedSheets++;
    const rules = qasGetBrandRules_(ss);
    for (let i = 1; i < data.length; i++) {
      if (qasStopRequested_()) { stopped = true; break; }
      const row = i + 1;
      const adGroup = qasToText_(data[i][schema.adGroupCol]);
      if (!adGroup) continue;
      const seen = {};
      for (let h = 0; h < schema.headlineCols.length; h++) {
        const meta = schema.headlineCols[h];
        const text = qasToText_(data[i][meta.col]);
        if (!text) continue;
        const issues = qasGetBrandIssuesForAsset_(text, "headline", h, adGroup, rules);
        const key = qasNormalizeTextKey_(text);
        if (seen[key]) issues.push("duplicate_headline_exact");
        seen[key] = true;
        for (let j = 0; j < issues.length; j++) {
          report.appendRow([now, adSheet.getName(), row, adGroup, "Headline", meta.slot, qasColToLetter_(meta.col + 1), issues[j], qasIsBrandIssueFixable_(issues[j]), text, qasGetSuggestedFix_(text, "headline", h, adGroup, rules)]);
          findings++;
        }
      }
      for (let d = 0; d < schema.descriptionCols.length; d++) {
        const meta = schema.descriptionCols[d];
        const text = qasToText_(data[i][meta.col]);
        if (!text) continue;
        const issues = qasGetBrandIssuesForAsset_(text, "description", d, adGroup, rules);
        for (let j = 0; j < issues.length; j++) {
          report.appendRow([now, adSheet.getName(), row, adGroup, "Description", meta.slot, qasColToLetter_(meta.col + 1), issues[j], qasIsBrandIssueFixable_(issues[j]), text, qasGetSuggestedFix_(text, "description", d, adGroup, rules)]);
          findings++;
        }
      }
    }
    if (stopped) break;
  }
  if (findings) report.autoResizeColumns(1, 11);
  SpreadsheetApp.getUi().alert(
    "Brand compliance audit complete.\nSheets scanned: " + scannedSheets + "\nFindings: " + findings +
    (stopped ? "\nStopped early by user request." : "")
  );
}

function qaStandaloneBrandComplianceFix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adSheets = qasGetTargetAdSheets_(ss);
  if (!adSheets.length) {
    SpreadsheetApp.getUi().alert("No configured ad sheets found.");
    return;
  }
  const keywordSheetName = qasGetTextSetting_("KEYWORD_SHEET_NAME", QAS_DEFAULT_KEYWORD_SHEET);
  const keywordSheet = ss.getSheetByName(keywordSheetName);
  const keywordData = keywordSheet ? keywordSheet.getDataRange().getValues() : [];
  const keywordHeaders = keywordData[0] || [];
  const useLandingPageContext = qasGetBoolSetting_("USE_LANDING_PAGE_CONTEXT", false);
  const landingPageSheetName = qasGetTextSetting_("LANDING_PAGE_SHEET_NAME", QAS_DEFAULT_LANDING_PAGE_SHEET);
  const landingSheet = useLandingPageContext ? ss.getSheetByName(landingPageSheetName) : null;
  const landingData = landingSheet ? landingSheet.getDataRange().getValues() : [];
  const landingCache = { byAdGroup: {}, byUrl: {} };
  const logSheet = qasGetOrCreateLogSheet_(ss);
  const runId = "QASBRAND_" + Date.now();
  const rules = qasGetBrandRules_(ss);
  const overrideMode = qasGetBoolSetting_("RUN_OVERRIDE_MODE", false);
  let updates = 0;
  let unresolved = 0;
  let processedSheets = 0;
  let stopped = false;

  for (let s = 0; s < adSheets.length; s++) {
    if (qasStopRequested_()) { stopped = true; break; }
    const adSheet = adSheets[s];
    const data = adSheet.getDataRange().getValues();
    if (!data.length) continue;
    const schema = qasDetectAssetColumns_(data[0] || []);
    if (schema.adGroupCol < 0 || !schema.headlineCols.length || !schema.descriptionCols.length) continue;
    processedSheets++;

    for (let i = 1; i < data.length; i++) {
      if (qasStopRequested_()) { stopped = true; break; }
      const row = i + 1;
      const adGroup = qasToText_(data[i][schema.adGroupCol]);
      if (!adGroup) continue;
      const keywords = qasGetKeywordsForAdGroup_(adGroup, keywordHeaders, keywordData);
      const landingContext = useLandingPageContext ? qasGetLandingPageContextForAdGroup_(adGroup, landingData, landingCache) : "";
      const seen = {};

      for (let h = 0; h < schema.headlineCols.length; h++) {
        const meta = schema.headlineCols[h];
        const col = meta.col + 1;
        const original = qasToText_(adSheet.getRange(row, col).getValue());
        if (!original) continue;
        const before = qasGetBrandIssuesForAsset_(original, "headline", h, adGroup, rules);
        const k = qasNormalizeTextKey_(original);
        const isDup = !!seen[k];
        if (isDup) before.push("duplicate_headline_exact");
        if (!before.length) {
          seen[k] = true;
          continue;
        }
        let next = qasGetSuggestedFix_(original, "headline", h, adGroup, rules);
        if (!next || next === original) next = qasIssueAwareRewrite_(original, "headline", QAS_HEADLINE_MAX, adGroup, keywords, before, landingContext);
        if (!next || next === original) {
          unresolved++;
          seen[k] = true;
          continue;
        }
        if (seen[qasNormalizeTextKey_(next)]) {
          unresolved++;
          seen[k] = true;
          continue;
        }
        if (!qasIsAcceptableRewrite_(next, original, "headline", QAS_HEADLINE_MAX)) {
          unresolved++;
          seen[k] = true;
          continue;
        }
        const after = qasGetBrandIssuesForAsset_(next, "headline", h, adGroup, rules);
        if (!overrideMode && !qasShouldApplyBrandRewrite_(before, after)) {
          unresolved++;
          seen[k] = true;
          continue;
        }
        adSheet.getRange(row, col).setValue(next).setBackground("#d9ead3");
        logSheet.appendRow([new Date(), runId, row, adSheet.getName() + " | " + adGroup, "Headline", meta.slot, qasColToLetter_(col), original, next, original.length, next.length, "brand_compliance_improve", 100]);
        updates++;
        seen[qasNormalizeTextKey_(next)] = true;
      }

      for (let d = 0; d < schema.descriptionCols.length; d++) {
        const meta = schema.descriptionCols[d];
        const col = meta.col + 1;
        const original = qasToText_(adSheet.getRange(row, col).getValue());
        if (!original) continue;
        const before = qasGetBrandIssuesForAsset_(original, "description", d, adGroup, rules);
        if (!before.length) continue;
        let next = qasGetSuggestedFix_(original, "description", d, adGroup, rules);
        if (!next || next === original) next = qasIssueAwareRewrite_(original, "description", QAS_DESCRIPTION_MAX, adGroup, keywords, before, landingContext);
        if (!next || next === original) {
          unresolved++;
          continue;
        }
        if (!qasIsAcceptableRewrite_(next, original, "description", QAS_DESCRIPTION_MAX)) {
          unresolved++;
          continue;
        }
        const after = qasGetBrandIssuesForAsset_(next, "description", d, adGroup, rules);
        if (!overrideMode && !qasShouldApplyBrandRewrite_(before, after)) {
          unresolved++;
          continue;
        }
        adSheet.getRange(row, col).setValue(next).setBackground("#d9ead3");
        logSheet.appendRow([new Date(), runId, row, adSheet.getName() + " | " + adGroup, "Description", meta.slot, qasColToLetter_(col), original, next, original.length, next.length, "brand_compliance_improve", 100]);
        updates++;
      }
    }
    if (stopped) break;
  }

  SpreadsheetApp.getUi().alert(
    "Brand compliance fix complete.\n" +
    "Sheets processed: " + processedSheets + "\n" +
    "Updated assets: " + updates + "\n" +
    "Unresolved assets: " + unresolved + "\n" +
    "Override mode: " + overrideMode +
    (stopped ? "\nStopped early by user request." : "")
  );
}

function qasStopRequested_() {
  return String(PropertiesService.getScriptProperties().getProperty("STOP_RUN_REQUESTED") || "false").toLowerCase() === "true";
}

function qaStandaloneUndoLastRun() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(QAS_LOG_SHEET);
  if (!logSheet) return SpreadsheetApp.getUi().alert("Nothing to undo.");
  const values = logSheet.getDataRange().getValues();
  if (values.length < 2) return SpreadsheetApp.getUi().alert("Nothing to undo.");
  const lastRunId = qasFindLastRunId_(values);
  if (!lastRunId) return SpreadsheetApp.getUi().alert("Nothing to undo.");
  const adSheets = qasGetTargetAdSheets_(ss);
  const byName = {};
  for (let i = 0; i < adSheets.length; i++) byName[adSheets[i].getName()] = adSheets[i];
  let reverted = 0;
  for (let i = values.length - 1; i >= 1; i--) {
    const row = values[i];
    if (row[1] !== lastRunId) continue;
    const sheetAndGroup = qasToText_(row[3]);
    const sheetName = sheetAndGroup.split("|")[0].trim();
    const sheet = byName[sheetName] || ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const sheetRow = Number(row[2]);
    const col = qasLetterToCol_(row[6]);
    const oldText = qasToText_(row[7]);
    if (sheetRow > 1 && col > 0) {
      sheet.getRange(sheetRow, col).setValue(oldText).setBackground(null);
      reverted++;
    }
  }
  SpreadsheetApp.getUi().alert("Undo complete. Reverted " + reverted + " cells from run " + lastRunId + ".");
}

function qaStandaloneClearColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adSheets = qasGetTargetAdSheets_(ss);
  for (let i = 0; i < adSheets.length; i++) {
    const sheet = adSheets[i];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).setBackground(null);
  }
  SpreadsheetApp.getUi().alert("Cleared QA highlight colors.");
}

function qasGetBrandRules_(ss) {
  return {
    specificHeadlineCount: qasGetNumberSetting_("SPECIFIC_HEADLINE_COUNT", 5, 3, 5),
    specificDescriptionCount: qasGetNumberSetting_("SPECIFIC_DESCRIPTION_COUNT", 2, 1, 5),
    lockSharedHeadlines: qasGetBoolSetting_("LOCK_SHARED_HEADLINES", false),
    lockSharedDescriptions: qasGetBoolSetting_("LOCK_SHARED_DESCRIPTIONS", false),
    fixedHeadlineByPosition: {},
    fixedDescriptionByPosition: {}
  };
}

function qasGetBrandIssuesForAsset_(text, type, index, adGroup, rules) {
  const maxLen = type === "headline" ? QAS_HEADLINE_MAX : QAS_DESCRIPTION_MAX;
  const out = [];
  const t = qasToText_(text);
  if (!t) return out;
  if (t.length > maxLen) out.push("too_long");
  if (/[a-z][A-Z]/.test(t) || /[A-Za-z]\/[A-Za-z]/.test(t)) out.push("joined_words");
  if (/\b(in|at|to|for|with|from|of|on|by|via)\s*$/i.test(t)) out.push("dangling_ending");
  if (/\bw\/\b/i.test(t) || /\bsvc\b/i.test(t)) out.push("bad_abbrev");
  if (qasContainsAnyTerm_(t, qasGetBlockedTerms_("BANNED_WORDS"))) out.push("banned_word");
  if (qasContainsAnyTerm_(t, qasGetBlockedTerms_("OFF_SHELF_WORDS"))) out.push("off_shelf_phrase");
  return qasUnique_(out);
}

function qasGetSuggestedFix_(_text, _type, _index, _adGroup, _rules) {
  return "";
}

function qasIsBrandIssueFixable_(_issue) {
  return true;
}

function qasTryFixAsset_(original, type, maxLen, adGroup, keywords, issues, landingContext) {
  const text = qasToText_(original);
  if (!text) return { status: "skip", text: text, method: "none", score: 0 };
  let best = qasShortenToLimit_(text, maxLen, type);
  let method = "deterministic";
  if (!qasIsAcceptableRewrite_(best, text, type, maxLen) || best === text) {
    const ai = qasAiRewriteToLimit_(text, type, maxLen, adGroup, keywords, issues, landingContext);
    if (ai) {
      best = ai;
      method = "ai";
    }
  }
  if (qasIsAcceptableRewrite_(best, text, type, maxLen)) {
    return { status: "ok", text: best, method: method, score: qasHumanScore_(best, type, adGroup, text, maxLen) };
  }
  const emergency = qasEmergencyLengthFallback_(text, type, maxLen);
  if (emergency && qasIsAcceptableRewrite_(emergency, text, type, maxLen)) {
    return { status: "ok", text: emergency, method: "emergency_length_fallback", score: qasHumanScore_(emergency, type, adGroup, text, maxLen) };
  }
  return { status: "error", text: text, method: "failed", score: 0, message: "Could not create acceptable " + type + " within " + maxLen + " chars." };
}

function qasIssueAwareRewrite_(original, type, maxLen, adGroup, keywords, issues, landingContext) {
  const list = qasUnique_(issues || []);
  if (list.indexOf("not_specific_to_ad_group") >= 0) {
    const det = qasInjectAdGroupContext_(original, adGroup, type, maxLen);
    if (det && det !== qasToText_(original)) return det;
  }
  const attempt = qasTryFixAsset_(original, type, maxLen, adGroup, keywords, list, landingContext);
  return attempt.status === "ok" ? attempt.text : "";
}

function qasInjectAdGroupContext_(text, adGroup, type, maxLen) {
  const base = qasToText_(text);
  const tokens = qasToText_(adGroup).split(/\s+/).filter(function(t) { return t && t.length > 2; });
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (base.toLowerCase().indexOf(token.toLowerCase()) >= 0) continue;
    let c = type === "headline" ? (token + " " + base) : (base + " for " + token);
    c = qasShortenToLimit_(c, maxLen, type);
    c = qasStripDanglingEnding_(c);
    if (c && c.length <= maxLen) return c;
  }
  return "";
}

function qasIsAcceptableRewrite_(candidate, original, _type, maxLen) {
  const c = qasToText_(candidate);
  const o = qasToText_(original);
  if (!c) return false;
  if (c.length > maxLen) return false;
  if (/\b(in|at|to|for|with|from|of|on|by|via)\s*$/i.test(c)) return false;
  if (/\bw\/\b/i.test(c) || /\bsvc\b/i.test(c)) return false;
  if (qasContainsAnyTerm_(c, qasGetBlockedTerms_("BANNED_WORDS"))) return false;
  if (qasContainsAnyTerm_(c, qasGetBlockedTerms_("OFF_SHELF_WORDS"))) return false;
  if (c === o && o.length > maxLen) return false;
  return true;
}

function qasShouldApplyBrandRewrite_(beforeIssues, afterIssues) {
  const before = qasUnique_(beforeIssues || []);
  const after = qasUnique_(afterIssues || []);
  if (!after.length) return true;
  const score = function(list) {
    return list.reduce(function(sum, k) {
      const w = { too_long: 5, banned_word: 5, off_shelf_phrase: 4, bad_abbrev: 3, dangling_ending: 3, duplicate_headline_exact: 4, joined_words: 2, not_specific_to_ad_group: 1 };
      return sum + (w[k] || 1);
    }, 0);
  };
  const b = score(before);
  const a = score(after);
  return a < b || (a === b && after.length < before.length);
}

function qasShortenToLimit_(text, maxLen, type) {
  const original = qasToText_(text);
  let t = qasNormalizeSpacing_(original);
  if (!t) return t;
  if (t.length <= maxLen) return qasRestoreCaseStyle_(t, original, type);
  t = t.replace(/\b(the|a|an|very|really|that|which|just|simply|completely|fully|carefully)\b/gi, "");
  t = t.replace(/\s+/g, " ").trim();
  if (t.length <= maxLen) return qasRestoreCaseStyle_(t, original, type);
  const sliced = t.slice(0, maxLen + 1);
  const lastSpace = sliced.lastIndexOf(" ");
  if (lastSpace > Math.floor(maxLen * 0.6)) return qasRestoreCaseStyle_(sliced.slice(0, lastSpace).trim(), original, type);
  return qasRestoreCaseStyle_(t.slice(0, maxLen).trim(), original, type);
}

function qasNormalizeSpacing_(text) {
  return qasToText_(text)
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/([A-Za-z])\/([A-Za-z])/g, "$1 $2")
    .replace(/\s+/g, " ")
    .trim();
}

function qasEmergencyLengthFallback_(original, type, maxLen) {
  let t = qasToText_(original);
  if (!t) return "";
  t = qasShortenToLimit_(t, maxLen, type);
  t = qasStripDanglingEnding_(t);
  if (t.length > maxLen) t = t.slice(0, maxLen).trim();
  t = t.replace(/\bw\/\b/gi, "with").replace(/\bsvc\b/gi, "service").replace(/\s+/g, " ").trim();
  if (t.length > maxLen) t = t.slice(0, maxLen).trim();
  return qasStripDanglingEnding_(t);
}

function qasAiRewriteToLimit_(original, type, maxLen, adGroup, keywords, issues, landingContext) {
  const apiKey = qasGetTextSetting_("OPENAI_API_KEY", "");
  if (!apiKey) return "";
  const model = qasGetTextSetting_("OPENAI_MODEL", "gpt-4o-mini");
  const brand = qasGetBrandContext_();
  const banned = qasGetBlockedTerms_("BANNED_WORDS");
  const offShelf = qasGetBlockedTerms_("OFF_SHELF_WORDS");
  const prompt = [
    "Rewrite this " + type + " for Google Ads brand compliance.",
    "Original: " + original,
    "Ad group label (internal only, do not copy into ad text): " + adGroup,
    "Keywords: " + (keywords || []).join(", "),
    "Landing page context: " + (qasToText_(landingContext) || "(none provided)"),
    "Detected issues: " + qasUnique_(issues || []).join(", "),
    "Brand: " + brand.brandName,
    "Website: " + brand.brandWebsite,
    "Positioning: " + brand.positioning1 + " | " + brand.positioning2,
    "Customer promise: " + brand.customerPromise,
    "CTA: " + brand.cta,
    "Rules:",
    "- Keep core meaning and fix the issue list",
    "- Do NOT use ad-group/internal labels in the final ad copy",
    "- Natural human phrasing; no robotic or spammy style",
    "- No abbreviations like w/, svc",
    "- Avoid banned words: " + banned.join(", "),
    "- Avoid off-shelf phrases: " + offShelf.join(", "),
    "- Max " + maxLen + " chars including spaces",
    "- Return 3 alternatives",
    "Return ONLY JSON:",
    '{"alternatives":["alt1","alt2","alt3"]}'
  ].join("\n");
  try {
    const resp = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + apiKey },
      payload: JSON.stringify({ model: model, messages: [{ role: "user", content: prompt }], temperature: 0.2 }),
      muteHttpExceptions: true
    });
    const json = JSON.parse(resp.getContentText());
    const out = json && json.choices && json.choices[0] && json.choices[0].message ? String(json.choices[0].message.content || "") : "";
    const candidates = qasExtractAiCandidates_(out);
    if (!candidates.length) return "";
    let best = "";
    let bestScore = -1;
    for (let i = 0; i < candidates.length; i++) {
      let c = qasShortenToLimit_(candidates[i], maxLen, type);
      c = qasNormalizeSpacing_(c);
      if (!qasIsAcceptableRewrite_(c, original, type, maxLen)) continue;
      const score = qasHumanScore_(c, type, adGroup, original, maxLen);
      if (score > bestScore) {
        bestScore = score;
        best = c;
      }
    }
    return best;
  } catch (_e) {
    return "";
  }
}

function qasExtractAiCandidates_(raw) {
  const text = String(raw || "").replace(/```json/gi, "").replace(/```/g, "").trim();
  const out = [];
  try {
    const parsed = JSON.parse(text);
    if (parsed && Array.isArray(parsed.alternatives)) {
      for (let i = 0; i < parsed.alternatives.length; i++) {
        const v = qasToText_(parsed.alternatives[i]);
        if (v) out.push(v);
      }
    }
  } catch (_e) {}
  if (out.length) return out;
  return text.split(/\r?\n/).map(qasToText_).filter(Boolean).slice(0, 3);
}

function qasGetBrandContext_() {
  return {
    brandName: qasGetTextSetting_("BRAND_NAME", ""),
    brandWebsite: qasGetTextSetting_("BRAND_WEBSITE", ""),
    positioning1: qasGetTextSetting_("POSITIONING_LINE_1", ""),
    positioning2: qasGetTextSetting_("POSITIONING_LINE_2", ""),
    customerPromise: qasGetTextSetting_("CUSTOMER_PROMISE", ""),
    cta: qasGetTextSetting_("CTA", "")
  };
}

function qasGetLandingPageContextForAdGroup_(adGroup, sheetValues, cache) {
  if (!sheetValues || !sheetValues.length) return "";
  const key = qasToText_(adGroup).toLowerCase().trim();
  const memo = cache || { byAdGroup: {}, byUrl: {} };
  if (memo.byAdGroup[key] != null) return memo.byAdGroup[key];
  const url = qasGetLandingPageUrlForAdGroup_(adGroup, sheetValues);
  if (!url) {
    memo.byAdGroup[key] = "";
    return "";
  }
  if (memo.byUrl[url] == null) memo.byUrl[url] = qasFetchWebsiteTextSafe_(url).slice(0, 1200);
  memo.byAdGroup[key] = memo.byUrl[url] || "";
  return memo.byAdGroup[key];
}

function qasGetLandingPageUrlForAdGroup_(adGroup, sheetValues) {
  if (!sheetValues || !sheetValues.length) return "";
  const headers = (sheetValues[0] || []).map(function(v) { return qasToText_(v).toLowerCase().trim(); });
  const adGroupKey = qasToText_(adGroup).toLowerCase().trim();
  let adGroupCol = headers.indexOf("ad group");
  if (adGroupCol < 0) adGroupCol = headers.indexOf("asset group");
  let urlCol = headers.indexOf("landing page");
  if (urlCol < 0) urlCol = headers.indexOf("landing page url");
  if (urlCol < 0) urlCol = headers.indexOf("url");
  if (adGroupCol >= 0 && urlCol >= 0) {
    for (let r = 1; r < sheetValues.length; r++) {
      const rowAdGroup = qasToText_(sheetValues[r][adGroupCol]).toLowerCase();
      if (rowAdGroup !== adGroupKey) continue;
      const url = qasToText_(sheetValues[r][urlCol]);
      if (url) return url;
    }
  }
  const colByHeader = headers.indexOf(adGroupKey);
  if (colByHeader >= 0) {
    for (let r = 1; r < sheetValues.length; r++) {
      const val = qasToText_(sheetValues[r][colByHeader]);
      if (val) return val;
    }
  }
  return "";
}

function qasFetchWebsiteTextSafe_(url) {
  const target = qasToText_(url);
  if (!target) return "";
  try {
    const res = UrlFetchApp.fetch(target, { muteHttpExceptions: true, followRedirects: true });
    const html = String(res.getContentText() || "");
    return html
      .replace(/<script[\s\S]*?<\/script>/gi, " ")
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<[^>]+>/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  } catch (_e) {
    return "";
  }
}

function qasHumanScore_(text, type, adGroup, original, maxLen) {
  const t = qasToText_(text);
  if (!t) return 0;
  let score = 100;
  if (t.length > maxLen) score -= 60;
  if (t === qasToText_(original)) score -= 20;
  if (type === "headline" && t.split(/\s+/).length < 2) score -= 20;
  if (qasContainsAnyTerm_(t, qasGetBlockedTerms_("BANNED_WORDS"))) score -= 35;
  if (qasContainsAnyTerm_(t, qasGetBlockedTerms_("OFF_SHELF_WORDS"))) score -= 30;
  return Math.max(score, 0);
}

function qasRestoreCaseStyle_(candidate, original, type) {
  let t = qasToText_(candidate);
  if (!t) return t;
  t = qasStripDanglingEnding_(t);
  const src = qasToText_(original);
  if (type === "headline" && src && /^[A-Z]/.test(src)) t = t.charAt(0).toUpperCase() + t.slice(1);
  return t;
}

function qasStripDanglingEnding_(text) {
  let t = qasToText_(text);
  const trailing = /\b(in|at|to|for|with|from|of|on|by|via)\s*$/i;
  let guard = 0;
  while (trailing.test(t) && guard < 3) {
    t = t.replace(trailing, "").trim();
    guard++;
  }
  return t;
}

function qasGetBlockedTerms_(propertyKey) {
  const fromSettings = qasGetTextSetting_(propertyKey, "")
    .split(/[\n,]/)
    .map(function(v) { return String(v || "").toLowerCase().trim(); })
    .filter(Boolean);
  const merged = fromSettings.concat(QAS_HARD_BLOCKED_TERMS);
  return qasUnique_(merged);
}

function qasContainsAnyTerm_(text, terms) {
  const lower = qasToText_(text).toLowerCase();
  for (let i = 0; i < (terms || []).length; i++) {
    if (terms[i] && lower.indexOf(terms[i]) >= 0) return true;
  }
  return false;
}

function qasGetKeywordsForAdGroup_(adGroup, headers, data) {
  const normalized = qasToText_(adGroup).toLowerCase();
  const col = (headers || []).findIndex(function(h) { return qasToText_(h).toLowerCase() === normalized; });
  const fromSheet = col === -1 ? [] : (data || []).slice(1).map(function(r) { return qasToText_(r[col]); }).filter(Boolean);
  return fromSheet;
}

function qasGetTargetAdSheets_(ss) {
  const tabs = qasGetTextSetting_("AD_TYPE_TABS", QAS_DEFAULT_AD_TYPE_TABS).split(/[\n,]/).map(qasToText_).filter(Boolean);
  const map = {};
  qasGetTextSetting_("AD_TYPE_SHEET_MAP", QAS_DEFAULT_AD_TYPE_SHEET_MAP).split(/[\n,]/).forEach(function(entry) {
    const m = qasToText_(entry).match(/^([^:]+)\s*:\s*(.+)$/);
    if (!m) return;
    map[String(m[1]).toLowerCase().replace(/[^a-z0-9]+/g, "")] = qasToText_(m[2]);
  });
  const out = [];
  const seen = {};
  for (let i = 0; i < tabs.length; i++) {
    const key = String(tabs[i]).toLowerCase().replace(/[^a-z0-9]+/g, "");
    const sheetName = map[key] || (key === "search" ? qasGetTextSetting_("AD_SHEET_NAME", QAS_DEFAULT_AD_SHEET) : tabs[i]);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || seen[sheet.getName()]) continue;
    seen[sheet.getName()] = true;
    out.push(sheet);
  }
  if (!out.length) {
    const fallback = ss.getSheetByName(qasGetTextSetting_("AD_SHEET_NAME", QAS_DEFAULT_AD_SHEET));
    if (fallback) out.push(fallback);
  }
  return out;
}

function qasDetectAssetColumns_(headerRow) {
  const headers = (headerRow || []).map(qasToText_);
  const normalize = function(v) { return qasToText_(v).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim(); };
  const adGroupAliases = { "ad group": true, "asset group": true, "adgroup": true, "assetgroup": true };
  const headlineCols = [];
  const descriptionCols = [];
  let adGroupCol = -1;
  for (let i = 0; i < headers.length; i++) {
    const norm = normalize(headers[i]);
    if (adGroupCol === -1 && adGroupAliases[norm]) adGroupCol = i;
    if (/^long headline\s+\d+$/.test(norm)) continue;
    let m = norm.match(/^headline\s+(\d+)$/);
    if (m) headlineCols.push({ col: i, slot: Number(m[1]) || headlineCols.length + 1 });
    m = norm.match(/^description\s+(\d+)$/);
    if (m) descriptionCols.push({ col: i, slot: Number(m[1]) || descriptionCols.length + 1 });
  }
  headlineCols.sort(function(a, b) { return a.slot - b.slot; });
  descriptionCols.sort(function(a, b) { return a.slot - b.slot; });
  return { adGroupCol: adGroupCol, headlineCols: headlineCols, descriptionCols: descriptionCols };
}

function qasEnsureErrorColumns_(sheet, headerRow) {
  let headers = (headerRow || []).slice();
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
    sheet.appendRow(["Time", "RunId", "Row", "AdGroup", "AssetType", "AssetIndex", "Column", "Old", "New", "OldLength", "NewLength", "Method", "HumanScore"]);
  }
  return sheet;
}

function qasGetOrCreateBrandReportSheet_(ss) {
  let sheet = ss.getSheetByName(QAS_BRAND_REPORT_SHEET);
  if (!sheet) sheet = ss.insertSheet(QAS_BRAND_REPORT_SHEET);
  return sheet;
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
  return String(qasGetTextSetting_(key, fallback ? "true" : "false")).toLowerCase() === "true";
}

function qasGetNumberSetting_(key, fallback, min, max) {
  let n = Number(qasGetTextSetting_(key, fallback));
  if (!isFinite(n)) n = fallback;
  if (typeof min === "number" && n < min) n = min;
  if (typeof max === "number" && n > max) n = max;
  return n;
}

function qasLooksSpecificToAdGroup_(text, adGroup) {
  const t = qasToText_(text).toLowerCase();
  const words = qasToText_(adGroup).toLowerCase().split(/\s+/).filter(function(w) { return w.length > 2; });
  if (!words.length) return true;
  return words.some(function(w) { return t.indexOf(w) >= 0; });
}

function qasNormalizeTextKey_(text) {
  return qasToText_(text).toLowerCase().replace(/\s+/g, " ").trim();
}

function qasUnique_(arr) {
  const seen = {};
  const out = [];
  for (let i = 0; i < (arr || []).length; i++) {
    const k = String(arr[i] || "");
    if (!k || seen[k]) continue;
    seen[k] = true;
    out.push(k);
  }
  return out;
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
/* ===============================
QA LENGTH AUDIT (STANDALONE)
=============================== */

(function() {

const TBX_DEFAULT_AD_SHEET = "Search";
const TBX_REPORT_SHEET = "QA Length Audit";
const TBX_FIX_LOG_SHEET = "QA Length Fix Log";
const TBX_HEADLINE_MAX = 30;
const TBX_DESCRIPTION_MAX = 90;
const TBX_LONG_HEADLINE_MAX = 90;
const TBX_MODE_KEY = "QAL_LENGTH_FIX_MODE";
const TBX_MODE_DETERMINISTIC = "deterministic";
const TBX_MODE_AI_FALLBACK = "ai_fallback";
const TBX_MODE_COMPARE = "compare_best";

function qalDetectAssetColumns_(headerRow) {
  const headers = (headerRow || []).map(function(v) { return qalToText_(v); });
  const normalize = function(v) { return qalToText_(v).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim(); };
  const adGroupAliases = { "ad group": true, "asset group": true, "adgroup": true, "assetgroup": true };
  const headlineCols = [];
  const longHeadlineCols = [];
  const descriptionCols = [];
  let adGroupCol = -1;
  for (let i = 0; i < headers.length; i++) {
    const norm = normalize(headers[i]);
    if (adGroupCol === -1 && adGroupAliases[norm]) adGroupCol = i;
    let m = norm.match(/^long headline\s+(\d+)$/);
    if (m) {
      longHeadlineCols.push({ col: i, slot: Number(m[1]) || longHeadlineCols.length + 1 });
      continue;
    }
    m = norm.match(/^headline\s+(\d+)$/);
    if (m) {
      headlineCols.push({ col: i, slot: Number(m[1]) || headlineCols.length + 1 });
      continue;
    }
    m = norm.match(/^description\s+(\d+)$/);
    if (m) {
      descriptionCols.push({ col: i, slot: Number(m[1]) || descriptionCols.length + 1 });
      continue;
    }
  }
  headlineCols.sort(function(a, b) { return a.slot - b.slot; });
  descriptionCols.sort(function(a, b) { return a.slot - b.slot; });
  longHeadlineCols.sort(function(a, b) { return a.slot - b.slot; });
  return {
    adGroupCol: adGroupCol,
    headlineCols: headlineCols,
    descriptionCols: descriptionCols,
    longHeadlineCols: longHeadlineCols
  };
}

function qaLengthAuditRun() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adSheetName = qalGetTextSetting_("AD_SHEET_NAME", TBX_DEFAULT_AD_SHEET);
  const adSheet = ss.getSheetByName(adSheetName);
  if (!adSheet) {
    SpreadsheetApp.getUi().alert("Ad sheet not found: " + adSheetName);
    return;
  }

  const data = adSheet.getDataRange().getValues();
  if (!data.length) return;
  const schema = qalDetectAssetColumns_(data[0] || []);
  if (schema.adGroupCol < 0 || !schema.headlineCols.length || !schema.descriptionCols.length) {
    SpreadsheetApp.getUi().alert("Could not detect Ad Group/Asset Group + Headline + Description columns in: " + adSheetName);
    return;
  }
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
    const adGroup = qalToText_(data[i][schema.adGroupCol]);
    if (!adGroup) continue;

    for (let h = 0; h < schema.headlineCols.length; h++) {
      const meta = schema.headlineCols[h];
      const text = qalToText_(data[i][meta.col]);
      if (!text) continue;
      if (text.length > TBX_HEADLINE_MAX) {
        report.appendRow([
          now, row, adGroup, "Headline", meta.slot,
          qalColToLetter_(meta.col + 1), text.length, TBX_HEADLINE_MAX, text
        ]);
        findings++;
      }
    }

    for (let d = 0; d < schema.descriptionCols.length; d++) {
      const meta = schema.descriptionCols[d];
      const text = qalToText_(data[i][meta.col]);
      if (!text) continue;
      if (text.length > TBX_DESCRIPTION_MAX) {
        report.appendRow([
          now, row, adGroup, "Description", meta.slot,
          qalColToLetter_(meta.col + 1), text.length, TBX_DESCRIPTION_MAX, text
        ]);
        findings++;
      }
    }

    for (let l = 0; l < schema.longHeadlineCols.length; l++) {
      const meta = schema.longHeadlineCols[l];
      const text = qalToText_(data[i][meta.col]);
      if (!text) continue;
      if (text.length > TBX_LONG_HEADLINE_MAX) {
        report.appendRow([
          now, row, adGroup, "LongHeadline", meta.slot,
          qalColToLetter_(meta.col + 1), text.length, TBX_LONG_HEADLINE_MAX, text
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
  if (!data.length) return;
  const schema = qalDetectAssetColumns_(data[0] || []);
  if (schema.adGroupCol < 0 || !schema.headlineCols.length || !schema.descriptionCols.length) {
    SpreadsheetApp.getUi().alert("Could not detect Ad Group/Asset Group + Headline + Description columns in: " + adSheetName);
    return;
  }
  const keywordData = keywordSheet ? keywordSheet.getDataRange().getValues() : [];
  const keywordHeaders = keywordData[0] || [];
  const logSheet = qalGetOrCreateFixLogSheet_(ss);
  const runId = "QALFIX_" + Date.now();
  let fixes = 0;

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const adGroup = qalToText_(data[i][schema.adGroupCol]);
    if (!adGroup) continue;

    for (let h = 0; h < schema.headlineCols.length; h++) {
      const meta = schema.headlineCols[h];
      const col = meta.col + 1;
      const original = qalToText_(data[i][meta.col]);
      if (!original || original.length <= TBX_HEADLINE_MAX) continue;
      const keywords = qalGetKeywordsForAdGroup_(adGroup, keywordHeaders, keywordData);
      const selected = qalSelectBestFix_(original, "headline", TBX_HEADLINE_MAX, adGroup, keywords, mode);
      const fixed = selected.text;
      if (!fixed || fixed === original) continue;
      adSheet.getRange(row, col).setValue(fixed).setBackground("#d9ead3");
      logSheet.appendRow([new Date(), runId, row, adGroup, "Headline", meta.slot, qalColToLetter_(col), original, fixed, original.length, fixed.length, mode, selected.method, selected.score]);
      fixes++;
    }

    for (let d = 0; d < schema.descriptionCols.length; d++) {
      const meta = schema.descriptionCols[d];
      const col = meta.col + 1;
      const original = qalToText_(data[i][meta.col]);
      if (!original || original.length <= TBX_DESCRIPTION_MAX) continue;
      const keywords = qalGetKeywordsForAdGroup_(adGroup, keywordHeaders, keywordData);
      const selected = qalSelectBestFix_(original, "description", TBX_DESCRIPTION_MAX, adGroup, keywords, mode);
      const fixed = selected.text;
      if (!fixed || fixed === original) continue;
      adSheet.getRange(row, col).setValue(fixed).setBackground("#d9ead3");
      logSheet.appendRow([new Date(), runId, row, adGroup, "Description", meta.slot, qalColToLetter_(col), original, fixed, original.length, fixed.length, mode, selected.method, selected.score]);
      fixes++;
    }

    for (let l = 0; l < schema.longHeadlineCols.length; l++) {
      const meta = schema.longHeadlineCols[l];
      const col = meta.col + 1;
      const original = qalToText_(data[i][meta.col]);
      if (!original || original.length <= TBX_LONG_HEADLINE_MAX) continue;
      const keywords = qalGetKeywordsForAdGroup_(adGroup, keywordHeaders, keywordData);
      const selected = qalSelectBestFix_(original, "description", TBX_LONG_HEADLINE_MAX, adGroup, keywords, mode);
      const fixed = selected.text;
      if (!fixed || fixed === original) continue;
      adSheet.getRange(row, col).setValue(fixed).setBackground("#d9ead3");
      logSheet.appendRow([new Date(), runId, row, adGroup, "LongHeadline", meta.slot, qalColToLetter_(col), original, fixed, original.length, fixed.length, mode, selected.method, selected.score]);
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

})();
