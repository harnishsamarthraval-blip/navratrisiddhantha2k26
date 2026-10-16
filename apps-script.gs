// Paste this into Extensions > Apps Script inside your Google Sheet.
// Then Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.

const SHEET = "Entries";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(["Time", "Name", "Phone", "Score", "Token", "StartMs", "Status", "Seconds played"]);
    sh.getRange("C:C").setNumberFormat("@");
    sh.setFrozenRows(1);
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = e.parameter || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getSheet_();

    if (p.action === "start") {
      const phone = String(p.phone || "").replace(/\D/g, "");
      const name = String(p.name || "").trim().slice(0, 40);
      if (!/^[6-9]\d{9}$/.test(phone) || name.length < 2) return out_({ ok: false, msg: "Invalid details." });
      const phones = sh.getRange("C2:C" + Math.max(2, sh.getLastRow())).getValues().flat().map(String);
      if (phones.indexOf(phone) !== -1) return out_({ ok: false, msg: "This mobile number has already played." });
      const token = Utilities.getUuid();
      const row = sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, 8).setValues([[new Date(), name, phone, 0, token, Date.now(), "started", ""]]);
      return out_({ ok: true, token: token });
    }

    if (p.action === "score") {
      const cell = sh.getRange("E:E").createTextFinder(String(p.token || "")).matchEntireCell(true).findNext();
      if (!cell) return out_({ ok: false, msg: "Unknown entry." });
      const row = cell.getRow();
      const status = sh.getRange(row, 7).getValue();
      if (status !== "started") return out_({ ok: true }); // already saved, ignore repeats
      const startMs = Number(sh.getRange(row, 6).getValue());
      const secs = (Date.now() - startMs) / 1000;
      const score = Math.max(0, Math.floor(Number(p.score) || 0));
      const maxPossible = Math.floor(secs / 0.9) + 2; // pillars can't be passed faster than this
      sh.getRange(row, 4).setValue(score);
      sh.getRange(row, 7).setValue(score > maxPossible ? "FLAGGED (score too high for time)" : "done");
      sh.getRange(row, 8).setValue(Math.round(secs));
      return out_({ ok: true });
    }

    return out_({ ok: false, msg: "Unknown action." });
  } finally {
    lock.releaseLock();
  }
}
