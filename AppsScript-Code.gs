/**
 * Pack Service Tracker <-> Google Sheets bridge
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
 *   Site Name | Rented | Number of Packs | Hardware | Firmware | Pack Colour | Last Service Date | Completed By | Next Service Due | Scheduled Date
 *
 *  - "Rented"            TRUE/FALSE, or Yes/No — only rows marked rented show up in the app.
 *  - "Hardware"           free text, e.g. "HSDU" or "WIFI".
 *  - "Pack Colour"         free text, expected values "R/G" or "R/B/P".
 *  - "Next Service Due"   left untouched by this script — keep your existing formula there
 *                          (e.g. =EDATE([Last Service Date], 3)); it recalculates automatically
 *                          whenever this script updates "Last Service Date".
 *  - "Completed By"        OPTIONAL. Whoever tapped "Complete Service" in the app (the
 *                          Representative field) gets written here alongside Last Service Date.
 *                          If the column doesn't exist yet, it's just silently ignored.
 *  - "Scheduled Date"      OPTIONAL, for the in-app scheduling feature. Add this column whenever
 *                          you're ready for it — everything here already supports it (readAllSites
 *                          returns it, updateSiteRow can write it via the "scheduleService" action,
 *                          the weekly email shows it, and CONFIG.SCHEDULE_NOTIFY_EMAILS below gets
 *                          an immediate email the moment a new service is scheduled). If the column
 *                          doesn't exist yet, it's just silently ignored.
 *  - "Notes"                OPTIONAL, free text, manually maintained directly in the Sheet (this
 *                          script never writes to it). Only surfaced in the weekly digest's
 *                          "Scheduled" section, and only for rows that currently have a Scheduled
 *                          Date — left out of the overdue/due-soon rows and out of the immediate
 *                          "service scheduled" email entirely. If the column doesn't exist yet,
 *                          it's just silently ignored.
 *
 * If your actual column headers differ, just edit the strings in CONFIG.COLUMNS
 * below to match — everything else in this file references them by name.
 *
 * OPTIONAL "Team" TAB — recipient lists without ever touching this script again.
 * Add a second tab named exactly "Team" (per CONFIG.TEAM_SHEET_NAME) with two columns,
 * one email address per row, columns independent of each other (different lengths are fine):
 *
 *   Weekly Digest Emails | Schedule Notification Emails
 *   alice@example.com    | alice@example.com
 *   bob@example.com      | carol@example.com
 *
 * "Weekly Digest Emails" receives the weekly overdue/due-soon/scheduled digest;
 * "Schedule Notification Emails" receives the immediate email fired when a service is newly
 * scheduled. If the "Team" tab (or a given column on it) doesn't exist yet, this script falls
 * back to CONFIG.NOTIFY_EMAILS / CONFIG.SCHEDULE_NOTIFY_EMAILS below — so nothing breaks before
 * you've set the tab up.
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
    COMPLETED_BY: "Completed By",           // Optional — who completed the service, from the app's Representative field
    NOTES: "Notes",                         // Optional — free text, only surfaced in the weekly digest's Scheduled section
    NEXT_SERVICE_DUE: "Next Service Due",   // Read-only here — driven by your existing formula
    SCHEDULED_DATE: "Scheduled Date",       // Optional — see note above
    SCHEDULED_TIME: "Scheduled Time",       // Optional — set alongside Scheduled Date
    SCHEDULED_REP: "Scheduled Rep"          // Optional — who's assigned to the scheduled visit
  },

  // --- TEAM TAB (recipient lists live here once you set it up — see below) ---
  TEAM_SHEET_NAME: "Team",
  TEAM_COLUMNS: {
    WEEKLY_DIGEST: "Weekly Digest Emails",
    SCHEDULE_NOTIFY: "Schedule Notification Emails"
  },

  // --- WEEKLY NOTIFICATION EMAIL ---
  // Used only as a fallback if the "Team" tab (or its "Weekly Digest Emails" column) doesn't
  // exist yet — once that's set up, addresses there take over automatically.
  NOTIFY_EMAILS: ["team@example.com"],
  DUE_WINDOW_DAYS: 30, // "Within one month" — sites due within this many days get included

  // --- IMMEDIATE "SERVICE SCHEDULED" EMAIL ---
  // Fires the moment someone schedules a service in the app (not just in the weekly digest).
  // Same fallback rule as above — the "Schedule Notification Emails" column on the Team tab
  // takes over once it exists. Leave this array empty to disable the email entirely (as a
  // fallback) if you haven't set up the Team tab and don't want it firing yet.
  SCHEDULE_NOTIFY_EMAILS: ["team@example.com"]
};

// Reads a recipient list from the "Team" tab (one email per row, under the given column
// header), falling back to a hardcoded CONFIG list if the tab or column doesn't exist yet —
// so nothing breaks before you've set the tab up, and you never have to touch this script
// again to add/remove someone from a distribution list afterwards.
function getTeamEmails(columnHeader, fallbackEmails) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TEAM_SHEET_NAME);
    if (!sheet) return fallbackEmails; // "Team" tab not created yet
    const headerMap = getHeaderMap(sheet);
    const colIdx = headerMap[columnHeader];
    if (!colIdx) return fallbackEmails; // column not found on the Team tab
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return fallbackEmails;
    const values = sheet.getRange(2, colIdx, lastRow - 1, 1).getValues();
    const emails = values.map(r => String(r[0] || "").trim()).filter(v => v.length > 0);
    return emails.length > 0 ? emails : fallbackEmails;
  } catch (err) {
    console.error(`Failed to read Team tab emails for "${columnHeader}", using fallback:`, err);
    return fallbackEmails;
  }
}

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
      updateSiteRow(body.siteName, { lastServiceDate: body.serviceDate, completedBy: body.completedBy });
      return jsonOutput({ ok: true });
    }
    if (body.action === "scheduleService") {
      updateSiteRow(body.siteName, { scheduledDate: body.scheduledDate, scheduledTime: body.scheduledTime, scheduledRep: body.scheduledRep });
      // Only fire the "newly scheduled" email when a date is actually being set — this same
      // action is also used to CLEAR a schedule (e.g. after starting it, or an admin clearing it),
      // and those clears shouldn't trigger a "service scheduled" notification.
      if (body.scheduledDate) {
        sendScheduleNotificationEmail(body.siteName, body.scheduledDate, body.scheduledTime, body.scheduledRep);
      }
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
  const notesCol = colIndex(headerMap, "NOTES", false); // optional

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
      scheduledRep: scheduledRepCol !== null ? String(row[scheduledRepCol] || "").trim() : "",
      notes: notesCol !== null ? String(row[notesCol] || "").trim() : ""
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
  if (updates.completedBy) {
    const completedByColLetter = headerMap[CONFIG.COLUMNS.COMPLETED_BY];
    if (completedByColLetter) sheet.getRange(targetRow, completedByColLetter).setValue(updates.completedBy);
  }
  if (updates.scheduledDate !== undefined && updates.scheduledDate !== null) {
    const scheduledColLetter = headerMap[CONFIG.COLUMNS.SCHEDULED_DATE];
    if (scheduledColLetter) {
      sheet.getRange(targetRow, scheduledColLetter).setValue(updates.scheduledDate);
    } // else: column doesn't exist yet on this sheet — silently skipped, nothing to write to.
  }
  if (updates.scheduledTime !== undefined && updates.scheduledTime !== null) {
    const scheduledTimeColLetter = headerMap[CONFIG.COLUMNS.SCHEDULED_TIME];
    if (scheduledTimeColLetter) {
      const cell = sheet.getRange(targetRow, scheduledTimeColLetter);
      // Force plain text BEFORE writing — otherwise Sheets auto-detects "20:00"-shaped strings
      // and silently converts them into a real Time value, which then displays shifted by
      // whatever gap exists between the spreadsheet's timezone and this script's project
      // timezone (that's the "20:00 in the email vs 12:00 on the Sheet" bug). Plain text means
      // exactly what was typed in the app is exactly what shows on the Sheet, no reinterpretation.
      cell.setNumberFormat('@');
      cell.setValue(updates.scheduledTime);
    }
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

// Manual diagnostic for the Team tab — run this from the Apps Script editor's function
// dropdown (select "debugTeamEmails" > Run), then check View > Logs (or View > Execution log).
// Shows exactly what tab/columns were found and which recipient list each email will actually
// use, so a silent fallback (wrong tab name, wrong column header, empty column, etc.) is
// visible instead of just "it's using the old addresses."
function debugTeamEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TEAM_SHEET_NAME);
  if (!sheet) {
    console.log(`No tab named exactly "${CONFIG.TEAM_SHEET_NAME}" found (case-sensitive). Tabs in this spreadsheet: ${ss.getSheets().map(s => s.getName()).join(', ')}`);
    return;
  }
  console.log(`Found "${CONFIG.TEAM_SHEET_NAME}" tab. Headers detected on row 1: ${JSON.stringify(getHeaderMap(sheet))}`);
  console.log(`CONFIG.TEAM_COLUMNS.WEEKLY_DIGEST = "${CONFIG.TEAM_COLUMNS.WEEKLY_DIGEST}" -> resolved emails: ${JSON.stringify(getTeamEmails(CONFIG.TEAM_COLUMNS.WEEKLY_DIGEST, CONFIG.NOTIFY_EMAILS))}`);
  console.log(`CONFIG.TEAM_COLUMNS.SCHEDULE_NOTIFY = "${CONFIG.TEAM_COLUMNS.SCHEDULE_NOTIFY}" -> resolved emails: ${JSON.stringify(getTeamEmails(CONFIG.TEAM_COLUMNS.SCHEDULE_NOTIFY, CONFIG.SCHEDULE_NOTIFY_EMAILS))}`);
  console.log("If the resolved list above matches CONFIG.NOTIFY_EMAILS/SCHEDULE_NOTIFY_EMAILS exactly, the column header text doesn't match CONFIG.TEAM_COLUMNS (check for typos/extra spaces/case) or the column is empty.");
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
    ? `Pack Service Tracker — ${overdue.length} overdue, ${dueSoon.length} due soon, ${scheduled.length} scheduled`
    : `Pack Service Tracker — nothing due within ${CONFIG.DUE_WINDOW_DAYS} days`;

  const html = buildEmailHtml(overdue, dueSoon, scheduled);
  const recipients = getTeamEmails(CONFIG.TEAM_COLUMNS.WEEKLY_DIGEST, CONFIG.NOTIFY_EMAILS);
  if (recipients.length === 0) return; // nobody to send to

  MailApp.sendEmail({
    to: recipients.join(','),
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
        <td style="padding:6px 10px;border-bottom:1px solid #ddd;">${s.notes ? escapeHtml(s.notes) : '—'}</td>
      </tr>`).join('');
    return `
      <h3 style="font-family:Arial,sans-serif;color:#2e7d32;margin-bottom:8px;">🟢 Scheduled</h3>
      <table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;width:100%;max-width:640px;margin-bottom:24px;">
        <thead><tr style="background:#f3f3f3;text-align:left;">
          <th style="padding:6px 10px;">Site</th>
          <th style="padding:6px 10px;">Date</th>
          <th style="padding:6px 10px;">Time</th>
          <th style="padding:6px 10px;">Representative</th>
          <th style="padding:6px 10px;">Notes</th>
        </tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>`;
  };

  return `
    <div style="font-family:Arial,sans-serif;">
      ${dueSection(`🔴 Overdue`, overdue, "No overdue services. ✅", "#c0392b")}
      ${dueSection(`🟡 Due within ${CONFIG.DUE_WINDOW_DAYS} days`, dueSoon, `Nothing due within ${CONFIG.DUE_WINDOW_DAYS} days.`, "#b8860b")}
      ${scheduledSection()}
      <p style="font-family:Arial,sans-serif;color:#999;font-size:12px;">Automated weekly digest from Pack Service Tracker.</p>
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

// =====================================================================
// IMMEDIATE "SERVICE SCHEDULED" EMAIL — fires right away from doPost's
// scheduleService action, separate from the weekly digest and its own
// recipient list (CONFIG.SCHEDULE_NOTIFY_EMAILS).
// =====================================================================

function sendScheduleNotificationEmail(siteName, scheduledDate, scheduledTime, scheduledRep) {
  const recipients = getTeamEmails(CONFIG.TEAM_COLUMNS.SCHEDULE_NOTIFY, CONFIG.SCHEDULE_NOTIFY_EMAILS);
  if (!recipients || recipients.length === 0) return;
  try {
    const subject = `Pack Service Scheduled — ${siteName} on ${scheduledDate}`;
    const html = `
      <div style="font-family:Arial,sans-serif;">
        <h3 style="color:#2e7d32;margin-bottom:10px;">🟢 New Service Scheduled</h3>
        <table style="border-collapse:collapse;font-size:14px;">
          <tr><td style="padding:4px 10px;color:#666;">Site</td><td style="padding:4px 10px;"><strong>${escapeHtml(siteName)}</strong></td></tr>
          <tr><td style="padding:4px 10px;color:#666;">Date</td><td style="padding:4px 10px;"><strong>${escapeHtml(scheduledDate)}</strong></td></tr>
          <tr><td style="padding:4px 10px;color:#666;">Time</td><td style="padding:4px 10px;">${scheduledTime ? escapeHtml(scheduledTime) : '—'}</td></tr>
          <tr><td style="padding:4px 10px;color:#666;">Representative</td><td style="padding:4px 10px;">${scheduledRep ? escapeHtml(scheduledRep) : '<em style="color:#999;">Unassigned</em>'}</td></tr>
        </table>
        <p style="font-family:Arial,sans-serif;color:#999;font-size:12px;margin-top:16px;">Automated notification from Pack Service Tracker.</p>
      </div>`;
    MailApp.sendEmail({ to: recipients.join(','), subject: subject, htmlBody: html });
  } catch (err) {
    console.error('Failed to send schedule notification email:', err);
  }
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}