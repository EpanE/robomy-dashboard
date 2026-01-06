/************ CONFIG ************/
const FIREBASE_URL = "https://robomy-locking-epan-default-rtdb.asia-southeast1.firebasedatabase.app";
const DEVICE_ID_DEFAULT = "robomy_raspi5_door"; 
// Make sure this Sheet ID points to your existing Google Sheet
const SHEET_ID = "1gBu4YhtXqYmAzlScUYhPZabIxKkC-6aE0ievrLyJOqY";
const SHEET_USERS = "Users";
const SHEET_SETTINGS = "Settings";
const SHEET_LOGS = "Logs";

/************ ENTRY POINTS ************/

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : "login";
  const params = e.parameter || {};

  if (page === "logout") {
    return logoutPage_(params.token);
  }

  if (page === "dashboard") {
    const token = params.token;
    const session = getSession_(token);
    if (!session) return loginPage_("Session expired or invalid.");
    return dashboardPage_(session);
  }

  return loginPage_();
}

/************ LOGIN + SESSION ************/

function attemptLogin(username, password) {
  const user = getUser_(username);
  if (!user) return { ok: false, msg: "Invalid username/password" };
  if (String(user.active).toLowerCase() !== "true") return { ok: false, msg: "User inactive" };

  const hash = hashPassword_(password);
  if (hash !== user.pass_hash) return { ok: false, msg: "Invalid username/password" };

  const sessionToken = Utilities.getUuid();
  const sessionData = JSON.stringify({
    username: user.username,
    role: user.role,
    ts: Date.now()
  });
  
  // Cache for 30 minutes
  CacheService.getUserCache().put("session_" + sessionToken, sessionData, 60 * 30);

  log_(Date.now(), "LOGIN_OK", user.username, "WEB", "");
  return { ok: true, token: sessionToken, role: user.role };
}

function getSession_(token) {
  if (!token) return null;
  const cached = CacheService.getUserCache().get("session_" + token);
  if (!cached) return null;
  return JSON.parse(cached);
}

function invalidateSession_(token) {
  if(token) CacheService.getUserCache().remove("session_" + token);
}

/************ HTML PAGES ************/

function loginPage_(errMsg) {
  const html = HtmlService.createTemplateFromFile('Login.html');
  html.error = errMsg || "";
  return html.evaluate()
    .setTitle("Robomy Login")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function dashboardPage_(session) {
  const html = HtmlService.createTemplateFromFile('Dashboard.html');
  html.username = session.username;
  // Note: You may need to adjust how the HTML file uses this token
  html.rawToken = session.token; 
  return html.evaluate()
    .setTitle("Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function logoutPage_(token) {
  if(token) invalidateSession_(token);
  return HtmlService.createHtmlOutput(
    `<script>window.top.location.href = '${ScriptApp.getService().getUrl()}';</script>`
  );
}

/************ DASHBOARD SERVER FUNCS ************/

function issueUnlockFromWeb(token) {
  const session = getSession_(token);
  if (!session) return { ok: false, msg: "Unauthorized" };

  const who = "web:" + session.username;
  const unlockSeconds = Number(getSetting_("unlockSeconds", "5"));
  const cmdId = "web_" + Utilities.getUuid();

  writeCommand_(cmdId, unlockSeconds, who, "WEB");
  log_(Date.now(), "CMD_ISSUED", who, "WEB", "cmd_id=" + cmdId);
  
  // Notify Telegram (This requires TELEGRAM_BOT_TOKEN in this project's properties too)
  telegramNotify_("🔓 WEB unlock issued by " + session.username + " (" + unlockSeconds + "s). cmd_id=" + cmdId);

  return { ok: true, msg: "Unlock command sent. cmd_id=" + cmdId };
}

function readStatus() {
  const deviceId = getSetting_("deviceId", DEVICE_ID_DEFAULT);
  const path = `/devices/${deviceId}/status.json`;
  const obj = firebaseGet_(path) || {};
  return JSON.stringify(obj);
}

// =========================================================
// COPY "COMMON CODE" SECTION BELOW HERE
// =========================================================
// (See Step 3 for the code to paste here)

/************ COMMON SHARED FUNCTIONS ************/

function getDbSecret_() {
  const props = PropertiesService.getScriptProperties();
  const sec = props.getProperty("FIREBASE_DB_SECRET");
  if (!sec) throw new Error("FIREBASE_DB_SECRET missing in Properties");
  return sec;
}

function firebaseGet_(path) {
  const url = FIREBASE_URL + path + "?auth=" + encodeURIComponent(getDbSecret_());
  const options = { method: "get", muteHttpExceptions: true };
  const res = UrlFetchApp.fetch(url, options);
  if (res.getResponseCode() >= 200 && res.getResponseCode() < 300) {
    const txt = res.getContentText();
    return txt ? JSON.parse(txt) : null;
  }
  console.error("Firebase Get Error: " + res.getResponseCode() + " " + res.getContentText());
  return null;
}

function firebasePut_(path, obj) {
  const url = FIREBASE_URL + path + "?auth=" + encodeURIComponent(getDbSecret_());
  const res = UrlFetchApp.fetch(url, {
    method: "put",
    contentType: "application/json",
    payload: JSON.stringify(obj),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    console.error("Firebase Put Error: " + res.getResponseCode());
  }
}

function firebasePost_(path, obj) {
  const url = FIREBASE_URL + path + "?auth=" + encodeURIComponent(getDbSecret_());
  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(obj),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    console.error("Firebase Post Error: " + res.getResponseCode());
  }
}

function writeCommand_(cmdId, durationS, issuedBy, via) {
  const deviceId = getSetting_("deviceId", DEVICE_ID_DEFAULT);
  const payload = {
    cmd: "UNLOCK",
    duration_s: durationS,
    cmd_id: cmdId,
    issued_by: issuedBy,
    issued_via: via,
    issued_at: Date.now()
  };
  firebasePut_(`/devices/${deviceId}/command.json`, payload);
}

function updateSetting_(key, value) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  let foundRow = -1;
  for(let i=1; i<data.length; i++) {
    if(String(data[i][0]) === key) {
      foundRow = i + 1;
      break;
    }
  }
  if(foundRow > 0) {
    sh.getRange(foundRow, 2).setValue(value);
  } else {
    sh.appendRow([key, value]);
  }
}

function getSetting_(key, fallback) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) return fallback;
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === key) return String(values[i][1]);
  }
  return fallback;
}

function log_(ts, event, who, via, details) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_LOGS);
  if (!sh) return;
  sh.appendRow([ts, event, who, via, details]);
  const deviceId = getSetting_("deviceId", DEVICE_ID_DEFAULT);
  firebasePost_(`/devices/${deviceId}/logs.json`, { ts, event, who, via, details });
}

function hashPassword_(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return bytes.map(b => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

function findUserRow_(username) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName("Users"); // Hardcoded for simplicity in shared code
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === username) return { row: i + 1, data: data[i] };
  }
  return null;
}

function getUser_(username) {
  const found = findUserRow_(username);
  if (!found) return null;
  return {
    username: String(found.data[0]),
    pass_hash: String(found.data[1]),
    role: String(found.data[2]),
    active: found.data[3]
  };
}

function isTelegramAllowed_(userId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName("TelegramAllow");
  if (!sh) return false;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]);
    const active = String(data[i][2]).toLowerCase() === "true";
    if (id === String(userId) && active) return true;
  }
  return false;
}

function telegramNotify_(text) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("TELEGRAM_BOT_TOKEN");
  const chatId = props.getProperty("TELEGRAM_DEFAULT_CHAT_ID");
  if (token && chatId) telegramSend_(token, chatId, text);
}

function telegramReply_(chatId, text) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("TELEGRAM_BOT_TOKEN");
  if (token) telegramSend_(token, chatId, text);
}

function telegramSend_(token, chatId, text) {
  const url = "https://api.telegram.org/bot" + token + "/sendMessage";
  const res = UrlFetchApp.fetch(url, {
    method: "post",
    payload: { chat_id: chatId, text: text },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    console.error("Telegram API Error: " + res.getContentText());
  }
}
