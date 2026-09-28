// Directory usability test — results receiver (Google Apps Script)
// Paste into Extensions › Apps Script of a new Google Sheet, then Deploy › New deployment › Web app
// (Execute as: Me · Who has access: Anyone). Copy the Web app URL.

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    sheet_("Sessions", ["Received", "Session", "Participant", "Role", "Started", "Screen", "Device", "Final comment", "Raw JSON"])
      .appendRow([new Date(), d.sessionId, d.participant, d.role, d.date, d.screen, d.device, d.finalComment, JSON.stringify(d)]);
    const t = sheet_("Tasks", ["Session", "Participant", "Task", "Outcome", "Seconds", "Clicks", "Wrong site", "Ease", "Comment", "Events"]);
    (d.tasks || []).forEach(function (r) {
      if (!r) return;
      t.appendRow([d.sessionId, d.participant, r.id, r.outcome, r.seconds, r.clicks, r.wrongSite ? "Yes" : "No", r.ease || "", r.comment || "",
        (r.events || []).map(function (ev) { return ev.t + "s " + ev.type + ": " + ev.label; }).join(" | ")]);
    });
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}

function doGet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sessions");
  const out = [];
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 9, sh.getLastRow() - 1, 1).getValues().forEach(function (row) { try { out.push(JSON.parse(row[0])); } catch (err) {} });
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
