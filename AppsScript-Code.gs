/**
 * Poncho Service Tracker <-> Google Sheets bridge
 * =================================================
 * Paste this whole file into the Apps Script editor for your Sheet
 * (Extensions > Apps Script), replace the CONFIG values below with
 * your own, then deploy:
 *
 *   Deploy > New deployment > select type "Web app"
 *     Execute as:      Me
 *     Who has access:  Anyone
 *
 * Copy the resulting URL (ends in /exec) into sheet-sync-config.js's
 * SHEET_SYNC_URL, and copy CONFIG.SECRET below into that same file's
 * SHEET_SYNC_SECRET. They must match exactly.
 *
 * EXPECTED SHEET LAYOUT (tab named per CONFIG.SHEET_NAME, row 1 = headers,
 * headers can be in any column order — they're matched by name, not position):
 *
 *   Site Name | Rented | Number of Packs | Hardware | Firmware | Pack Colour | Last Service Date | Next Service Due | Scheduled Date
 *
 *  - "Rented"            TRUE/FALSE, or Yes/No — only rows marked rented show up in the app.
 *  - "Hardware"           free text, e.g. "HSDU" or "WIFI".
 *  - "Pack Colour"         free text, expected values "R/G" or "R/B/P".
 *  - "Next Service Due"   left untouched by this script — keep your existing formula there
 *                          (e.g. =EDATE([Last Service Date], 3)); it recalculates automatically
 *                          whenever this script updates "Last Service Date".
 *  - "Scheduled Date"      OPTIONAL, for the upcoming in-app scheduling feature. Add this column
 *                          whenever you're ready for it — everything here already supports it
 *                          (readAllSites returns it, updateSiteRow can write it via the
 *                          "scheduleService" action, and the weekly email shows it). If the
 *                          column doesn't exist yet, it's just silently ignored.
 *
 * If your actual column headers differ, just edit the strings in CONFIG.COLUMNS
 * below to match — everything else in this file references them by name.
 */

const CONFIG = {
  SHEET_NAME: "Sites", // Tab name inside the spreadsheet
  SECRET: "TERCES", // Must match SHEET_SYNC_SECRET in sheet-sync-config.js
  COLUMNS: {
    SITE_NAME: "Site Name",
    RENTED: "Rented",
    PACKS: "Number of Packs",
    HARDWARE: "Hardware",
    FIRMWARE: "Firmware",
    PACK_COLOUR: "Pack Colour",             // R/G or R/B/P
    LAST_SERVICE_DATE: "Last Service Date",
    NEXT_SERVICE_DUE: "Next Service Due",   // Read-only here — driven by your existing formula
    SCHEDULED_DATE: "Scheduled Date",       // Optional — see note above
    SCHEDULED_TIME: "Scheduled Time",       // Optional — set alongside Scheduled Date
    SCHEDULED_REP: "Scheduled Rep"          // Optional — who's assigned to the scheduled visit
  },

  // --- WEEKLY NOTIFICATION EMAIL ---
  NOTIFY_EMAILS: ["team@example.com"], // Replace with your real recipient(s), comma-separate for multiple
  DUE_WINDOW_DAYS: 30 // "Within one month" — sites due within this many days get included
};

// --- ENTRY POINTS ---

function doGet(e) {
  try {
    if (e.parameter.secret !== CONFIG.SECRET) return jsonOutput({ ok: false, error: "Unauthorized" });

    const action = e.parameter.action;
    if (action === "listSites") {
      return jsonOutput({ ok: true, sites: readAllSites() });
    }
    return jsonOutput({ ok: false, error: "Unknown action" });
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.secret !== CONFIG.SECRET) return jsonOutput({ ok: false, error: "Unauthorized" });

    if (body.action === "updateSite") {
      updateSiteRow(body.siteName, { packs: body.packs, firmware: body.firmware, hardware: body.hardware, packColour: body.packColour });
      return jsonOutput({ ok: true });
    }
    if (body.action === "completeService") {
      updateSiteRow(body.siteName, { lastServiceDate: body.serviceDate });
      return jsonOutput({ ok: true });
    }
    if (body.action === "scheduleService") {
      updateSiteRow(body.siteName, { scheduledDate: body.scheduledDate, scheduledTime: body.scheduledTime, scheduledRep: body.scheduledRep });
      return jsonOutput({ ok: true });
    }
    return jsonOutput({ ok: false, error: "Unknown action" });
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message });
  }
}

// --- SHEET HELPERS ---

function getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet tab "${CONFIG.SHEET_NAME}" not found`);
  return sheet;
}

function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[String(h).trim()] = i + 1; }); // 1-based column index
  return map;
}

// required=false lets optional columns (currently just Scheduled Date) be absent without breaking
// everything else — useful since that column doesn't exist on anyone's sheet until they add it.
function colIndex(headerMap, key, required) {
  if (required === undefined) required = true;
  const label = CONFIG.COLUMNS[key];
  const idx = headerMap[label];
  if (!idx) {
    if (required) throw new Error(`Column "${label}" not found in sheet headers`);
    return null;
  }
  return idx - 1; // 0-based, for row array access
}

function readAllSites() {
  const sheet = getSheet();
  const headerMap = getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const numCols = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  const nameCol = colIndex(headerMap, "SITE_NAME");
  const rentedCol = colIndex(headerMap, "RENTED");
  const packsCol = colIndex(headerMap, "PACKS");
  const hardwareCol = colIndex(headerMap, "HARDWARE");
  const firmwareCol = colIndex(headerMap, "FIRMWARE");
  const packColourCol = colIndex(headerMap, "PACK_COLOUR", false); // optional
  const lastServiceCol = colIndex(headerMap, "LAST_SERVICE_DATE");
  const nextServiceCol = colIndex(headerMap, "NEXT_SERVICE_DUE");
  const scheduledCol = colIndex(headerMap, "SCHEDULED_DATE", false); // optional
  const scheduledTimeCol = colIndex(headerMap, "SCHEDULED_TIME", false); // optional
  const scheduledRepCol = colIndex(headerMap, "SCHEDULED_REP", false); // optional

  return data
    .filter(row => row[nameCol])
    .map(row => ({
      siteName: String(row[nameCol]).trim(),
      rented: parseBool(row[rentedCol]),
      packs: Number(row[packsCol]) || 0,
      hardware: String(row[hardwareCol] || "").trim(),
      firmware: String(row[firmwareCol] || "").trim(),
      packColour: packColourCol !== null ? String(row[packColourCol] || "").trim() : "",
      lastServiceDate: formatDateCell(row[lastServiceCol]),
      nextServiceDue: formatDateCell(row[nextServiceCol]),
      scheduledDate: scheduledCol !== null ? formatDateCell(row[scheduledCol]) : "",
      scheduledTime: scheduledTimeCol !== null ? formatTimeCell(row[scheduledTimeCol]) : "",
      scheduledRep: scheduledRepCol !== null ? String(row[scheduledRepCol] || "").trim() : ""
    }));
}

function updateSiteRow(siteName, updates) {
  if (!siteName) throw new Error("Missing siteName");
  const sheet = getSheet();
  const headerMap = getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("Sheet has no data rows");

  const nameColIndex = headerMap[CONFIG.COLUMNS.SITE_NAME];
  const names = sheet.getRange(2, nameColIndex, lastRow - 1, 1).getValues();

  let targetRow = -1;
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === siteName.trim()) { targetRow = i + 2; break; }
  }
  if (targetRow === -1) throw new Error(`Site "${siteName}" not found in sheet`);

  if (updates.packs !== undefined && updates.packs !== null) {
    sheet.getRange(targetRow, headerMap[CONFIG.COLUMNS.PACKS]).setValue(updates.packs);
  }
  if (updates.firmware !== undefined && updates.firmware !== null) {
    sheet.getRange(targetRow, headerMap[CONFIG.COLUMNS.FIRMWARE]).setValue(updates.firmware);
  }
  if (updates.hardware) {
    sheet.getRange(targetRow, headerMap[CONFIG.COLUMNS.HARDWARE]).setValue(updates.hardware);
  }
  if (updates.packColour) {
    const packColourColLetter = headerMap[CONFIG.COLUMNS.PACK_COLOUR];
    if (packColourColLetter) sheet.getRange(targetRow, packColourColLetter).setValue(updates.packColour);
  }
  if (updates.lastServiceDate !== undefined && updates.lastServiceDate !== null) {
    sheet.getRange(targetRow, headerMap[CONFIG.COLUMNS.LAST_SERVICE_DATE]).setValue(updates.lastServiceDate);
    // Next Service Due is intentionally left alone — your existing formula recalculates it.
  }
  if (updates.scheduledDate !== undefined && updates.scheduledDate !== null) {
    const scheduledColLetter = headerMap[CONFIG.COLUMNS.SCHEDULED_DATE];
    if (scheduledColLetter) {
      sheet.getRange(targetRow, scheduledColLetter).setValue(updates.scheduledDate);
    } // else: column doesn't exist yet on this sheet — silently skipped, nothing to write to.
  }
  if (updates.scheduledTime !== undefined && updates.scheduledTime !== null) {
    const scheduledTimeColLetter = headerMap[CONFIG.COLUMNS.SCHEDULED_TIME];
    if (scheduledTimeColLetter) sheet.getRange(targetRow, scheduledTimeColLetter).setValue(updates.scheduledTime);
  }
  if (updates.scheduledRep !== undefined && updates.scheduledRep !== null) {
    const scheduledRepColLetter = headerMap[CONFIG.COLUMNS.SCHEDULED_REP];
    if (scheduledRepColLetter) sheet.getRange(targetRow, scheduledRepColLetter).setValue(updates.scheduledRep);
  }
}

function parseBool(val) {
  if (typeof val === "boolean") return val;
  const s = String(val).trim().toLowerCase();
  return s === "true" || s === "yes" || s === "y" || s === "rented";
}

function formatDateCell(val) {
  if (!val) return "";
  if (Object.prototype.toString.call(val) === "[object Date]") {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val);
}

function formatTimeCell(val) {
  if (!val) return "";
  if (Object.prototype.toString.call(val) === "[object Date]") {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
  }
  return String(val);
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// =====================================================================
// WEEKLY "SERVICES DUE SOON" EMAIL
// =====================================================================
//
// Run createWeeklyTrigger() ONCE from this editor (select it in the function
// dropdown above, click Run, and approve the permissions prompt) to schedule
// this automatically. It'll then run every Monday morning without you doing
// anything further. To change recipients later, edit CONFIG.NOTIFY_EMAILS
// above — no need to touch the trigger.

function createWeeklyTrigger() {
  deleteAllTriggers(); // avoid creating duplicates if you run this more than once
  ScriptApp.newTrigger('sendWeeklyServiceDueEmail')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
}

function deleteAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}

// One-off manual send — run this any time from the editor's function dropdown (select
// "sendServiceDueEmailNow" > Run) to test the email, or trigger it ad-hoc outside the
// weekly schedule. It's identical to what the weekly trigger runs.
function sendServiceDueEmailNow() {
  sendWeeklyServiceDueEmail();
}

// Adds a "Service Tracker" menu to the spreadsheet itself (Extensions bar, next to Help),
// so the email can be sent — or the weekly trigger installed — without opening the Apps
// Script editor at all. Runs automatically whenever the sheet is opened.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Service Tracker')
    .addItem('📧 Send Service Due Email Now', 'sendServiceDueEmailNow')
    .addItem('⏰ Set Up Weekly Email (Mondays 7am)', 'createWeeklyTrigger')
    .addToUi();
}

function sendWeeklyServiceDueEmail() {
  const today = stripTime(new Date());
  const cutoff = new Date(today.getTime() + CONFIG.DUE_WINDOW_DAYS * 24 * 60 * 60 * 1000);

  const sites = readAllSites().filter(s => s.rented);

  const scheduled = [];
  const overdue = [];
  const dueSoon = [];

  sites.forEach(s => {
    // A scheduled service takes priority — it no longer shows as overdue/due-soon, it gets its
    // own section instead, regardless of what Next Service Due says.
    if (s.scheduledDate) { scheduled.push(s); return; }

    if (!s.nextServiceDue) return; // no date on record — nothing to flag yet
    const dueDate = parseDate(s.nextServiceDue);
    if (!dueDate) return;
    if (dueDate < today) overdue.push(s);
    else if (dueDate <= cutoff) dueSoon.push(s);
  });

  scheduled.sort((a, b) => `${a.scheduledDate}T${a.scheduledTime || '00:00'}`.localeCompare(`${b.scheduledDate}T${b.scheduledTime || '00:00'}`));
  overdue.sort((a, b) => parseDate(a.nextServiceDue) - parseDate(b.nextServiceDue));
  dueSoon.sort((a, b) => parseDate(a.nextServiceDue) - parseDate(b.nextServiceDue));

  const subject = (overdue.length + dueSoon.length + scheduled.length) > 0
    ? `Poncho Service Tracker — ${overdue.length} overdue, ${dueSoon.length} due soon, ${scheduled.length} scheduled`
    : `Poncho Service Tracker — nothing due within ${CONFIG.DUE_WINDOW_DAYS} days`;

  const html = buildEmailHtml(overdue, dueSoon, scheduled);

  MailApp.sendEmail({
    to: CONFIG.NOTIFY_EMAILS.join(','),
    subject: subject,
    htmlBody: html
  });
}

function buildEmailHtml(overdue, dueSoon, scheduled) {
  const dueSection = (title, rows, emptyMsg, color) => {
    if (rows.length === 0) return `<h3 style="font-family:Arial,sans-serif;color:${color};margin-bottom:4px;">${title}</h3><p style="font-family:Arial,sans-serif;color:#666;margin-top:0;">${emptyMsg}</p>`;
    const rowsHtml = rows.map(s => `
      <tr>
        <td style="padding:6px 10px;border-bottom:1px solid #ddd;">${escapeHtml(s.siteName)}</td>
        <td style="padding:6px 10px;border-bottom:1px solid #ddd;">${escapeHtml(s.hardware || '—')}</td>
        <td style="padding:6px 10px;border-bottom:1px solid #ddd;">${escapeHtml(s.nextServiceDue)}</td>
      </tr>`).join('');
    return `
      <h3 style="font-family:Arial,sans-serif;color:${color};margin-bottom:8px;">${title}</h3>
      <table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;width:100%;max-width:640px;margin-bottom:24px;">
        <thead><tr style="background:#f3f3f3;text-align:left;">
          <th style="padding:6px 10px;">Site</th>
          <th style="padding:6px 10px;">Hardware</th>
          <th style="padding:6px 10px;">Next Service Due</th>
        </tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>`;
  };

  const scheduledSection = () => {
    if (scheduled.length === 0) return `<h3 style="font-family:Arial,sans-serif;color:#2e7d32;margin-bottom:4px;">🟢 Scheduled</h3><p style="font-family:Arial,sans-serif;color:#666;margin-top:0;">No services currently scheduled.</p>`;
    const rowsHtml = scheduled.map(s => `
      <tr>
        <td style="padding:6px 10px;border-bottom:1px solid #ddd;">${escapeHtml(s.siteName)}</td>
        <td style="padding:6px 10px;border-bottom:1px solid #ddd;">${escapeHtml(s.scheduledDate)}</td>
        <td style="padding:6px 10px;border-bottom:1px solid #ddd;">${s.scheduledTime ? escapeHtml(s.scheduledTime) : '—'}</td>
        <td style="padding:6px 10px;border-bottom:1px solid #ddd;">${s.scheduledRep ? escapeHtml(s.scheduledRep) : '<em style="color:#999;">Unassigned</em>'}</td>
      </tr>`).join('');
    return `
      <h3 style="font-family:Arial,sans-serif;color:#2e7d32;margin-bottom:8px;">🟢 Scheduled</h3>
      <table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;width:100%;max-width:640px;margin-bottom:24px;">
        <thead><tr style="background:#f3f3f3;text-align:left;">
          <th style="padding:6px 10px;">Site</th>
          <th style="padding:6px 10px;">Date</th>
          <th style="padding:6px 10px;">Time</th>
          <th style="padding:6px 10px;">Representative</th>
        </tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>`;
  };

  return `
    <div style="font-family:Arial,sans-serif;">
      ${dueSection(`🔴 Overdue`, overdue, "No overdue services. ✅", "#c0392b")}
      ${dueSection(`🟡 Due within ${CONFIG.DUE_WINDOW_DAYS} days`, dueSoon, `Nothing due within ${CONFIG.DUE_WINDOW_DAYS} days.`, "#b8860b")}
      ${scheduledSection()}
      <p style="font-family:Arial,sans-serif;color:#999;font-size:12px;">Automated weekly digest from Poncho Service Tracker.</p>
    </div>`;
}

function parseDate(yyyyMmDd) {
  if (!yyyyMmDd) return null;
  const parts = String(yyyyMmDd).split('-');
  if (parts.length !== 3) return null;
  const d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  return isNaN(d.getTime()) ? null : d;
}

function stripTime(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}