// ═══════════════════════════════════════════════════════════════════
// SWMT – Google Apps Script Backend
// ───────────────────────────────────────────────────────────────────
// Setup instructions:
//  1. Go to https://script.google.com → create a new project
//  2. Paste this entire file into the editor
//  3. Project Settings → Script Properties → Add the following:
//       TWILIO_SID    = ACxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
//       TWILIO_TOKEN  = (new token from console.twilio.com)
//       TWILIO_FROM   = +18255331973
//       ADMIN_EMAIL   = youngjaehur.ca@gmail.com
//  4. Deploy > New deployment > Type: Web App
//       - Execute as: Me
//       - Who has access: Anyone
//  5. Copy the deployment URL → paste into CONFIG.appsScriptUrl in the HTML file
// ═══════════════════════════════════════════════════════════════════

// ── SETUP ──────────────────────────────────────────────────────────
const PROPS          = PropertiesService.getScriptProperties();
const SHEET_NAME     = "SWMT_Data";
const BACKUP_SHEET_NAME = "SWMT_Backup";
const BACKUP_MAX_ROWS   = 100;   // keep only the latest N backups
const TWILIO_SID     = PROPS.getProperty("TWILIO_SID")   || "";
const TWILIO_TOKEN   = PROPS.getProperty("TWILIO_TOKEN") || "";
const TWILIO_FROM    = PROPS.getProperty("TWILIO_FROM")  || "";
const ONESIGNAL_APP_ID = "";   // OneSignal App ID (optional – for push notifications)
const ONESIGNAL_KEY    = "";   // OneSignal REST API Key (optional)
const ADMIN_EMAIL    = PROPS.getProperty("ADMIN_EMAIL")  || "";
// ──────────────────────────────────────────────────────────────────

function doGet(e) {
  const action = (e.parameter || {}).action;
  if (action === "load") {
    const data = loadData();
    return json(data || {});
  }
  return json({ ok: true });
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === "save")            { return json(saveData(body.data)); }
    if (action === "notify")          { sendNotifications(body.post, body.players);     return json({ ok: true }); }
    if (action === "notifyAdmin")     { sendAdminNotification(body.req);                return json({ ok: true }); }
    if (action === "notifyGameAdded") { sendGameAddedNotifications(body.game, body.players); return json({ ok: true }); }

    return json({ ok: true });
  } catch (err) {
    return json({ error: err.message });
  }
}

// ── DATA SAVE / LOAD ──────────────────────────────────────────────

function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function saveData(data) {
  // Serialize concurrent saves so the compare-and-write is atomic.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // wait up to 10s for another save to finish
  } catch (e) {
    return { ok: false, error: "busy" };
  }
  try {
    const sheet = getSheet();
    // Stale-write guard: reject data older than what is already saved.
    // C1 stores the last saved client lastModified timestamp (ms since epoch),
    // separate from B1 which records the server-side save time (ISO string).
    const incomingTs = Number(data && data.lastModified) || 0;
    const savedTs    = Number(sheet.getRange("C1").getValue()) || 0;
    if (savedTs > 0 && incomingTs < savedTs) {
      return { ok: false, error: "stale", saved: savedTs, incoming: incomingTs };
    }
    const jsonStr = JSON.stringify(data);
    sheet.getRange("A1").setValue(jsonStr);
    sheet.getRange("B1").setValue(new Date().toISOString());
    sheet.getRange("C1").setValue(incomingTs);
    // Append a backup snapshot (best-effort: never block the main save)
    try { appendBackup(jsonStr); } catch (e) { Logger.log("Backup failed: " + e.message); }
    return { ok: true, saved: incomingTs };
  } finally {
    lock.releaseLock();
  }
}

// ── AUTO BACKUP ───────────────────────────────────────────────────
// Every save appends one row [timestamp, JSON] to the SWMT_Backup sheet,
// keeping only the latest BACKUP_MAX_ROWS rows.
function getBackupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BACKUP_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(BACKUP_SHEET_NAME);
    sheet.getRange("A1").setValue("timestamp");
    sheet.getRange("B1").setValue("data");
  }
  return sheet;
}

function appendBackup(jsonStr) {
  const sheet = getBackupSheet();
  sheet.appendRow([new Date().toISOString(), jsonStr]);
  // Trim oldest rows beyond the limit (row 1 is the header).
  const dataRows = sheet.getLastRow() - 1;
  if (dataRows > BACKUP_MAX_ROWS) {
    sheet.deleteRows(2, dataRows - BACKUP_MAX_ROWS);
  }
}

function loadData() {
  const val = getSheet().getRange("A1").getValue();
  if (!val) return null;
  try { return JSON.parse(val); } catch (e) { return null; }
}

// ── SEND NOTIFICATIONS ────────────────────────────────────────────

function sendNotifications(post, players) {
  const catLabel = { notice: "Notice", general: "General", game: "Game", photo: "Photo" }[post.cat] || post.cat;
  const subject  = "[SWMT] " + (post.pinned ? "📌 " : "") + post.title;

  const emailBody =
    "Hello Shoe Wedge Mafia Tour members!\n\n" +
    "A new post has been published.\n\n" +
    "Category: " + catLabel + "\n" +
    "Author:   " + post.author + "\n" +
    "Title:    " + post.title + "\n\n" +
    post.body + "\n\n" +
    "──────────────────────\n" +
    "Shoe Wedge Mafia Tour · Open the app to read and reply.";

  const smsText = "[SWMT] " + post.author + " posted: \"" + post.title + "\" — Open the app to read.";

  players.forEach(function (player) {
    if (player.email && player.email.indexOf("@") > 0) {
      try {
        MailApp.sendEmail({ to: player.email, subject: subject, body: emailBody });
      } catch (e) {
        Logger.log("Email failed [" + player.email + "]: " + e.message);
      }
    }

    if (player.phone && TWILIO_SID && TWILIO_TOKEN) {
      try {
        twilioSMS(player.phone, smsText);
      } catch (e) {
        Logger.log("SMS failed [" + player.phone + "]: " + e.message);
      }
    }
  });

  if (ONESIGNAL_APP_ID && ONESIGNAL_KEY) {
    try { oneSignalPush(subject, post.body.slice(0, 100)); } catch (e) { Logger.log("Push failed: " + e.message); }
  }
}

// ── NEW GAME NOTIFICATION ────────────────────────────────────────

function sendGameAddedNotifications(game, players) {
  if (!game) return;
  game = game || {};
  const num     = game.gameNumber != null ? game.gameNumber : "?";
  const cls     = game.class === "B" ? "B" : "A";
  const dateRaw = game.date || "";
  const dateStr = dateRaw ? new Date(dateRaw + "T12:00:00").toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric", year: "numeric" }) : "TBD";
  const teeStr  = game.teeTime ? "Tee time: " + game.teeTime : "";
  const rsvpStr = game.rsvpBy ? "RSVP by: " + new Date(game.rsvpBy + "T12:00:00").toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric" }) : "";

  const subject = "[SWMT] 📅 New Game #" + num + ": " + (game.name || "Game " + num);
  const emailBody =
    "Hello Shoe Wedge Mafia Tour members!\n\n" +
    "A new game has been scheduled.\n\n" +
    "Game #" + num + " · Class " + cls + "\n" +
    "Name:  " + (game.name || "") + "\n" +
    "Date:  " + dateStr + "\n" +
    (teeStr ? teeStr + "\n" : "") +
    (rsvpStr ? rsvpStr + "\n" : "") +
    "\n──────────────────────\n" +
    "Open the app to RSVP.\n" +
    "Shoe Wedge Mafia Tour";

  const smsText =
    "[SWMT] 📅 Game #" + num + " (" + cls + "): " + (game.name || "") +
    " · " + dateStr +
    (game.teeTime ? " @ " + game.teeTime : "") +
    (rsvpStr ? " · " + rsvpStr : "") +
    " — Open the app to RSVP.";

  (players || []).forEach(function (player) {
    if (player.email && player.email.indexOf("@") > 0) {
      try { MailApp.sendEmail({ to: player.email, subject: subject, body: emailBody }); }
      catch (e) { Logger.log("Game-added email failed [" + player.email + "]: " + e.message); }
    }
    if (player.phone && TWILIO_SID && TWILIO_TOKEN) {
      try { twilioSMS(player.phone, smsText); }
      catch (e) { Logger.log("Game-added SMS failed [" + player.phone + "]: " + e.message); }
    }
  });

  if (ONESIGNAL_APP_ID && ONESIGNAL_KEY) {
    try { oneSignalPush("📅 New Game #" + num + ": " + (game.name || ""), dateStr + (game.teeTime ? " @ " + game.teeTime : "") + " · Class " + cls); }
    catch (e) { Logger.log("Game-added push failed: " + e.message); }
  }
}

// ── ADMIN JOIN REQUEST NOTIFICATION ──────────────────────────────

function sendAdminNotification(req) {
  if (!ADMIN_EMAIL) return;
  const subject = "[SWMT] New Join Request: " + req.firstName + " " + req.lastName;
  const body =
    "A new member has requested to join Shoe Wedge Mafia Tour.\n\n" +
    "Name:     " + req.firstName + " " + req.lastName + "\n" +
    "Phone:    " + req.phone + "\n" +
    "Email:    " + req.email + "\n" +
    (req.handicap != null ? "Handicap: " + req.handicap + "\n" : "") +
    (req.birthday ? "Birthday: " + req.birthday + "\n" : "") +
    "Requested: " + req.requestedAt + "\n\n" +
    "Open the app → Profiles page → log in as Admin to approve or reject.";
  try {
    MailApp.sendEmail({ to: ADMIN_EMAIL, subject: subject, body: body });
  } catch (e) {
    Logger.log("Admin notify failed: " + e.message);
  }
}

// ── TWILIO SMS ────────────────────────────────────────────────────

function twilioSMS(toNumber, message) {
  const url = "https://api.twilio.com/2010-04-01/Accounts/" + TWILIO_SID + "/Messages.json";
  const res = UrlFetchApp.fetch(url, {
    method: "post",
    headers: { Authorization: "Basic " + Utilities.base64Encode(TWILIO_SID + ":" + TWILIO_TOKEN) },
    payload: { To: toNumber, From: TWILIO_FROM, Body: message },
    muteHttpExceptions: true
  });
  const r = JSON.parse(res.getContentText());
  if (r.error_code) throw new Error(r.message);
}

// ── ONESIGNAL PUSH ────────────────────────────────────────────────

function oneSignalPush(title, message) {
  UrlFetchApp.fetch("https://onesignal.com/api/v1/notifications", {
    method: "post",
    headers: {
      "Authorization": "Basic " + ONESIGNAL_KEY,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({
      app_id: ONESIGNAL_APP_ID,
      included_segments: ["All"],
      headings: { en: title },
      contents: { en: message }
    }),
    muteHttpExceptions: true
  });
}

// ── HELPER ───────────────────────────────────────────────────────

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
