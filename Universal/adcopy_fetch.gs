function fetchAdcopy(){
  if(!getAccessToken()){
    showOAuthPopup();
  }
  fetchPmaxAssets();
  fetchSearchAds();
}

function fetchPmaxAssets() {
  var ACCESS_TOKEN = getAccessToken();
  var CUSTOMER_ID = PropertiesService.getScriptProperties().getProperty('customer_id');
  var MCC_ID = PropertiesService.getScriptProperties().getProperty('mcc_id');
  var SHEET_NAME = "PMax1";

  var url = "https://googleads.googleapis.com/v23/customers/" + CUSTOMER_ID + "/googleAds:searchStream";

  var query = `
    SELECT
      campaign.name,
      asset_group.name,
      asset_group.final_urls,
      asset_group_asset.field_type,
      asset.text_asset.text
    FROM asset_group_asset
    WHERE campaign.advertising_channel_type = 'PERFORMANCE_MAX'
      AND campaign.status = 'ENABLED'
      AND asset_group_asset.field_type IN ('HEADLINE', 'LONG_HEADLINE', 'DESCRIPTION')
  `;

  var payload = JSON.stringify({ query: query });
  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + ACCESS_TOKEN,
      "developer-token": PropertiesService.getScriptProperties().getProperty('GADS_DEVELOPER_TOKEN')
    },
    payload: payload
  };

  if (MCC_ID !== '-') {
    options.headers["login-customer-id"] = MCC_ID;
  }

  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());

  if (!data || data.length === 0 || (data.length === 1 && !data[0].results)) {
    Logger.log("⚠️ No PMax Assets found for this account.");
    return;
  }

  var results = {};

  data.forEach(function(chunk) {
    if (!chunk.results) return;
    chunk.results.forEach(function(row) {
      var campaign = row.campaign.name;
      var assetGroup = row.assetGroup.name;
      var finalUrl = (row.assetGroup.finalUrls && row.assetGroup.finalUrls.length > 0) ? row.assetGroup.finalUrls[0] : "";
      var fieldType = row.assetGroupAsset.fieldType;
      var text = row.asset?.textAsset?.text || "";

      var key = campaign + "||" + assetGroup;

      if (!results[key]) {
        results[key] = {
          campaign: campaign,
          assetGroup: assetGroup,
          finalUrl: finalUrl,
          headlines: [],
          longHeadlines: [],
          descriptions: []
        };
      }

      if (fieldType === "HEADLINE") {
        results[key].headlines.push(text);
      } else if (fieldType === "LONG_HEADLINE") {
        results[key].longHeadlines.push(text);
      } else if (fieldType === "DESCRIPTION") {
        results[key].descriptions.push(text);
      }
    });
  });

  if (Object.keys(results).length === 0) {
    Logger.log("⚠️ No PMax results to display.");
    return;
  }

  var header = ["Campaign", "Asset Group"];
  for (var i = 1; i <= 15; i++) header.push("Headline " + i);
  for (var i = 1; i <= 5; i++) header.push("Long Headline " + i);
  for (var i = 1; i <= 5; i++) header.push("Description " + i);
  header.push("Landing Page");

  var rows = [header];

  Object.values(results).forEach(function(r) {
    var row = [r.campaign, r.assetGroup];
    for (var i = 0; i < 15; i++) row.push(r.headlines[i] || "");
    for (var i = 0; i < 5; i++) row.push(r.longHeadlines[i] || "");
    for (var i = 0; i < 5; i++) row.push(r.descriptions[i] || "");
    row.push(r.finalUrl);
    rows.push(row);
  });

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
  } else {
    sheet.clear();
  }

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log("✅ Done: PMax data exported → " + SHEET_NAME);
}

function fetchSearchAds() {
  var ACCESS_TOKEN = getAccessToken();
  var CUSTOMER_ID = PropertiesService.getScriptProperties().getProperty('customer_id');
  var MCC_ID = PropertiesService.getScriptProperties().getProperty('mcc_id');
  var SHEET_NAME = "Search1";

  var url = "https://googleads.googleapis.com/v23/customers/" + CUSTOMER_ID + "/googleAds:searchStream";

  var query = `
    SELECT
      campaign.name,
      ad_group.name,
      ad_group_ad.ad.id,
      ad_group_ad.ad.final_urls,
      ad_group_ad.ad.responsive_search_ad.headlines,
      ad_group_ad.ad.responsive_search_ad.descriptions
    FROM ad_group_ad
    WHERE campaign.advertising_channel_type = 'SEARCH'
      AND campaign.status = 'ENABLED'
      AND ad_group.status = 'ENABLED'
      AND ad_group_ad.status = 'ENABLED'
  `;

  var payload = JSON.stringify({ query: query });
  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + ACCESS_TOKEN,
      "developer-token": PropertiesService.getScriptProperties().getProperty('GADS_DEVELOPER_TOKEN')
    },
    payload: payload
  };

  if (MCC_ID !== '-') {
    options.headers["login-customer-id"] = MCC_ID;
  }

  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());

  if (!data || data.length === 0 || (data.length === 1 && !data[0].results)) {
    Logger.log("⚠️ No Search Ads found for this account.");
    return;
  }

  var tempRows = [];
  var maxHeadlines = 0;
  var maxDescriptions = 0;

  data.forEach(function(chunk) {
    if (!chunk.results) return;
    chunk.results.forEach(function(row) {
      var campaign = row.campaign.name;
      var adGroup = row.adGroup.name;
      var adId = row.adGroupAd.ad.id;
      var finalUrl = (row.adGroupAd.ad.finalUrls && row.adGroupAd.ad.finalUrls.length > 0) ? row.adGroupAd.ad.finalUrls[0] : "";
      var headlines = [];
      var descriptions = [];

      if (row.adGroupAd.ad.responsiveSearchAd) {
        var rsa = row.adGroupAd.ad.responsiveSearchAd;
        if (rsa.headlines) rsa.headlines.forEach(function(h) { headlines.push(h.text); });
        if (rsa.descriptions) rsa.descriptions.forEach(function(d) { descriptions.push(d.text); });
      }

      maxHeadlines = Math.max(maxHeadlines, headlines.length);
      maxDescriptions = Math.max(maxDescriptions, descriptions.length);

      tempRows.push({
        campaign: campaign,
        adGroup: adGroup,
        adId: adId,
        finalUrl: finalUrl,
        headlines: headlines,
        descriptions: descriptions
      });
    });
  });

  if (tempRows.length === 0) {
    Logger.log("⚠️ No active Search Ads processed.");
    return;
  }

  var header = ["Campaign", "Ad Group", "Ad ID"];
  for (var i = 1; i <= maxHeadlines; i++) header.push("Headline " + i);
  for (var i = 1; i <= maxDescriptions; i++) header.push("Description " + i);
  header.push("Landing Page");

  var rows = [header];
  tempRows.forEach(function(r) {
    var row = [r.campaign, r.adGroup, r.adId];
    for (var i = 0; i < maxHeadlines; i++) row.push(r.headlines[i] || "");
    for (var i = 0; i < maxDescriptions; i++) row.push(r.descriptions[i] || "");
    row.push(r.finalUrl);
    rows.push(row);
  });

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
  } else {
    sheet.clear();
  }

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log("✅ Done: Search Ads exported to sheet → " + SHEET_NAME);
}
