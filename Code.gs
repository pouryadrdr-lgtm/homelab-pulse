// Paste this into the Homelab spreadsheet's Apps Script project (Extensions > Apps Script), replacing
// the earlier version of this file (keep the separate onEdit sync script untouched, if present).
// Deploy as a Web App (Execute as: Me, Who has access: Anyone). Read-only — never writes to the sheet.
// Serves raw per-row data from Data AND Targets so the dashboard can filter/aggregate/chart, and compute
// pace/attainment/variance against target, entirely client-side.

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = {
    generatedAt: new Date().toISOString(),
    raw: getRawData_(ss),
    targets: getTargets_(ss),
    channelTypes: getChannelTypes_(ss)
  };
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// "Free" was renamed from "No-insurance" — campaign rows still tagged with the old text should still
// match, same alias map the sheet's own costF formula uses (build_line.js / patch_cost_campaign_match.js).
var CAMPAIGN_ALIASES_ = { 'Free': ['Free', 'No-insurance'] };

function getChannelTypes_(ss){
  var sh = ss.getSheetByName('Settings');
  if (!sh) return [];
  var vals = sh.getRange(3, 5, 30, 1).getValues(); // Settings!E3:E32 — the Channel Type list
  return vals.map(function(r){ return r[0]; }).filter(function(v){ return v; });
}

// ---------- Data tab (per-row campaign entries) ----------

var RAW_FIELDS_ = [
  'date', 'source', 'medium', 'channelType', 'campaign', 'landingPage', 'copy',
  'cost', 'impressions', 'clicks', 'sessions', 'leads', 'inzone', 'mql', 'orders', 'sms', 'upsell',
  'leadPremium', 'leadPrescription', 'leadFree',
  'inzonePremium', 'inzonePrescription', 'inzoneFree',
  'mqlPremium', 'mqlPrescription', 'mqlFree',
  'orderPremium', 'orderPrescription', 'orderFree',
  'costPremium', 'costPrescription', 'costFree',
  'upsellPremium', 'upsellPrescription', 'upsellFree'
];

var RAW_HEADER_NAMES_ = {
  source: 'Source', medium: 'Medium', channelType: 'Channel Type', campaign: 'Campaign',
  landingPage: 'Landing Page', copy: 'Copy',
  cost: 'Cost (IRR)', impressions: 'Impressions', clicks: 'Clicks', sessions: 'Sessions',
  leads: 'Leads', inzone: 'Inzone', mql: 'MQL', orders: 'Orders', sms: 'SMS Sent', upsell: 'Upsell',
  leadPremium: 'Lead Premium', leadPrescription: 'Lead Prescription', leadFree: 'Lead Free',
  inzonePremium: 'Inzone Premium', inzonePrescription: 'Inzone Prescription', inzoneFree: 'Inzone Free',
  mqlPremium: 'MQL Premium', mqlPrescription: 'MQL Prescription', mqlFree: 'MQL Free',
  orderPremium: 'Order Premium', orderPrescription: 'Order Prescription', orderFree: 'Order Free',
  costPremium: 'Cost Premium', costPrescription: 'Cost Prescription', costFree: 'Cost Free',
  upsellPremium: 'Upsell Premium', upsellPrescription: 'Upsell Prescription', upsellFree: 'Upsell Free'
};

function getRawData_(ss) {
  var sh = ss.getSheetByName('Data');
  if (!sh) return { fields: RAW_FIELDS_, rows: [] };

  var lastCol = sh.getLastColumn();
  var hdr = sh.getRange(3, 1, 1, lastCol).getValues()[0];
  var idx = {};
  hdr.forEach(function (h, i) { if (h) idx[h] = i; });

  var dateIdx = idx['Date (Gregorian)'];
  var colIdx = {};
  for (var key in RAW_HEADER_NAMES_) colIdx[key] = idx[RAW_HEADER_NAMES_[key]];

  var lastRow = sh.getLastRow();
  var numRows = lastRow - 3;
  if (numRows <= 0 || dateIdx === undefined) return { fields: RAW_FIELDS_, rows: [] };

  var data = sh.getRange(4, 1, numRows, lastCol).getValues();
  var tz = Session.getScriptTimeZone();
  var rows = [];
  var costTypeIdx = { Premium: colIdx.costPremium, Prescription: colIdx.costPrescription, Free: colIdx.costFree };

  for (var r = 0; r < data.length; r++) {
    var d = data[r][dateIdx];
    if (!(d instanceof Date)) continue;
    var totalCost = data[r][colIdx.cost];
    totalCost = typeof totalCost === 'number' ? totalCost : 0;
    var campaign = String(data[r][colIdx.campaign] || '');
    var rec = [Utilities.formatDate(d, tz, 'yyyy-MM-dd')];
    for (var f = 1; f < RAW_FIELDS_.length; f++) {
      var name = RAW_FIELDS_[f];
      if (name === 'costPremium' || name === 'costPrescription' || name === 'costFree') {
        var type = name === 'costPremium' ? 'Premium' : (name === 'costPrescription' ? 'Prescription' : 'Free');
        var entry = data[r][costTypeIdx[type]];
        if (typeof entry === 'number') { rec.push(entry); continue; }
        var aliases = CAMPAIGN_ALIASES_[type] || [type];
        var matches = aliases.some(function (a) { return campaign.toLowerCase() === a.toLowerCase(); });
        rec.push(matches ? totalCost : 0);
        continue;
      }
      var v = data[r][colIdx[name]];
      rec.push(typeof v === 'number' ? v : (v || 0));
    }
    rows.push(rec);
  }

  return { fields: RAW_FIELDS_, rows: rows };
}

// ---------- Targets tab (one row per Jalali Year + Month) ----------

var TARGET_FIELDS_ = [
  'year', 'month',
  'cost', 'impressions', 'clicks', 'sessions', 'leads', 'inzone', 'mql', 'orders', 'sms',
  'leadPremium', 'leadPrescription', 'leadFree',
  'inzonePremium', 'inzonePrescription', 'inzoneFree',
  'mqlPremium', 'mqlPrescription', 'mqlFree',
  'orderPremium', 'orderPrescription', 'orderFree',
  'costPremium', 'costPrescription', 'costFree',
  'upsellPremium', 'upsellPrescription', 'upsellFree'
];

var TARGET_HEADER_NAMES_ = {
  cost: 'Target Cost (IRR)', impressions: 'Target Impressions', clicks: 'Target Clicks', sessions: 'Target Sessions',
  leads: 'Target Leads', inzone: 'Target Inzone', mql: 'Target MQL', orders: 'Target Orders', sms: 'Target SMS Sent',
  leadPremium: 'Target Lead Premium', leadPrescription: 'Target Lead Prescription', leadFree: 'Target Lead Free',
  inzonePremium: 'Target Inzone Premium', inzonePrescription: 'Target Inzone Prescription', inzoneFree: 'Target Inzone Free',
  mqlPremium: 'Target MQL Premium', mqlPrescription: 'Target MQL Prescription', mqlFree: 'Target MQL Free',
  orderPremium: 'Target Order Premium', orderPrescription: 'Target Order Prescription', orderFree: 'Target Order Free',
  costPremium: 'Target Cost Premium', costPrescription: 'Target Cost Prescription', costFree: 'Target Cost Free',
  upsellPremium: 'Target Upsell Premium', upsellPrescription: 'Target Upsell Prescription', upsellFree: 'Target Upsell Free'
};

function getTargets_(ss) {
  var sh = ss.getSheetByName('Targets');
  if (!sh) return { fields: TARGET_FIELDS_, rows: [] };

  var lastCol = sh.getLastColumn();
  var hdr = sh.getRange(3, 1, 1, lastCol).getValues()[0];
  var idx = {};
  hdr.forEach(function (h, i) { if (h) idx[h] = i; });

  var yearIdx = idx['Year'], monthIdx = idx['Month'];
  var colIdx = {};
  for (var key in TARGET_HEADER_NAMES_) colIdx[key] = idx[TARGET_HEADER_NAMES_[key]];

  var lastRow = sh.getLastRow();
  var numRows = lastRow - 3;
  if (numRows <= 0 || yearIdx === undefined) return { fields: TARGET_FIELDS_, rows: [] };

  var data = sh.getRange(4, 1, numRows, lastCol).getValues();
  var rows = [];

  for (var r = 0; r < data.length; r++) {
    var yr = data[r][yearIdx], mo = data[r][monthIdx];
    if (!yr || !mo) continue;
    var rec = [String(yr), String(mo)];
    for (var f = 2; f < TARGET_FIELDS_.length; f++) {
      var v = data[r][colIdx[TARGET_FIELDS_[f]]];
      rec.push(typeof v === 'number' ? v : (v || 0));
    }
    rows.push(rec);
  }

  return { fields: TARGET_FIELDS_, rows: rows };
}
