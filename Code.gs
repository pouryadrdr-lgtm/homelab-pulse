// Paste this into the Homelab spreadsheet's Apps Script project (Extensions > Apps Script),
// as a NEW file or appended to the existing Code.gs that already has the Type-filter onEdit sync.
// Deploy as a Web App (Execute as: Me, Who has access: Anyone) to get a JSON endpoint the PWA reads.
// Read-only — does not modify the sheet.

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = {
    generatedAt: new Date().toISOString(),
    glance: getGlance_(ss),
    typePerformance: getTypePerformance_(ss),
    charts: getCharts_(ss)
  };
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function getGlance_(ss) {
  var sh = ss.getSheetByName('Dashboard');
  if (!sh) return null;
  var finder = sh.createTextFinder('AT A GLANCE').matchEntireCell(false).findNext();
  if (!finder) return null;
  var titleRow = finder.getRow();
  var title = sh.getRange(titleRow, 1).getDisplayValue();
  var startRow = titleRow + 2; // skip the "Metric | Period A" header row
  var rows = [];
  for (var r = startRow; r < startRow + 30; r++) {
    var metric = sh.getRange(r, 1).getDisplayValue();
    if (!metric) break;
    var value = sh.getRange(r, 2).getDisplayValue();
    rows.push({ metric: metric, value: value });
  }
  return { title: title, rows: rows };
}

function getTypePerformance_(ss) {
  var sh = ss.getSheetByName('Type Performance');
  if (!sh) return null;
  var values = sh.getDataRange().getDisplayValues();
  return { values: values };
}

function getCharts_(ss) {
  var sh = ss.getSheetByName('Dashboard');
  if (!sh) return [];
  var charts = sh.getCharts();
  var out = [];
  for (var i = 0; i < charts.length; i++) {
    try {
      var chart = charts[i];
      var blob = chart.getAs('image/png');
      var b64 = Utilities.base64Encode(blob.getBytes());
      var title = '';
      try { title = chart.getOptions().get('title') || ''; } catch (e1) {}
      out.push({ title: title, image: 'data:image/png;base64,' + b64 });
    } catch (e2) {
      // skip any chart that fails to export
    }
  }
  return out;
}
