// Paste this into the Homelab spreadsheet's Apps Script project (Extensions > Apps Script), replacing
// the earlier version of this file (keep the separate onEdit sync script untouched, if present).
// Deploy as a Web App (Execute as: Me, Who has access: Anyone). Read-only — never writes to the sheet.
// Serves raw per-row data from the Data tab so the dashboard can filter/aggregate/chart client-side,
// instead of mirroring pre-computed cells.

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = {
    generatedAt: new Date().toISOString(),
    raw: getRawData_(ss)
  };
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// Field order returned per row. Index 0 is always the ISO date (from the "Date (Gregorian)" helper
// column) — everything else is looked up by header name so this survives future column inserts.
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

  for (var r = 0; r < data.length; r++) {
    var d = data[r][dateIdx];
    if (!(d instanceof Date)) continue;
    var rec = [Utilities.formatDate(d, tz, 'yyyy-MM-dd')];
    for (var f = 1; f < RAW_FIELDS_.length; f++) {
      var v = data[r][colIdx[RAW_FIELDS_[f]]];
      rec.push(typeof v === 'number' ? v : (v || 0));
    }
    rows.push(rec);
  }

  return { fields: RAW_FIELDS_, rows: rows };
}
