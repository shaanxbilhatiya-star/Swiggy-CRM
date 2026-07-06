// NOTE: All timers are timestamp-based and survive server restarts:
// - Break timer: uses breakStartedAt (epoch ms) in state.json
// - 7-day delivery timer: uses recruitedAt (ISO string) in state.json
// - Daily reset: uses lastReset (YYYY-MM-DD) in state.json
// Server can restart at any time without losing timer state.

const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const multer = require('multer');
const XLSX = require('xlsx');
const { v4: uuidv4 } = require('uuid');
const path = require('path');
const fs = require('fs');
const os = require('os');
const PDFDocument = require('pdfkit');

const app = express();
const server = http.createServer(app);
const io = new Server(server);

const PORT = process.env.PORT || 3000;

function ensureDir(d) { if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true }); }

// ─── Persistent Data Location ─────────────────────────────────────────────────
// CRITICAL: data/ and uploads/ live OUTSIDE the project folder by default.
// Reason: re-cloning, re-downloading, or extracting a fresh copy of this repo to
// "update" the app replaces the whole project folder. Anything stored inside it
// (old default: data/state.json, uploads/) gets wiped along with the old code.
// Storing it under the OS user's home folder means updating the code never touches it.
// Override with the AUTOLEAD_DATA_DIR environment variable (e.g. to point at a
// mounted persistent volume on a cloud host) if you don't want the home-folder default.
let DATA_ROOT = process.env.AUTOLEAD_DATA_DIR || path.join(os.homedir(), '.autolead-crm');
try {
  ensureDir(DATA_ROOT);
} catch (e) {
  console.error('\u26A0\uFE0F  Could not use external data folder "' + DATA_ROOT + '" (' + e.message + '). ' +
    'Falling back to storing data inside the project folder — your data WILL be lost next time you ' +
    'update by re-cloning/re-downloading this project. Set AUTOLEAD_DATA_DIR to a writable folder to fix this.');
  DATA_ROOT = __dirname;
}

// ─── Container/PaaS persistence sanity check ──────────────────────────────────
// "Outside the project folder" only survives an UPDATE if the home folder itself
// survives between deploys. On a real machine (LAN PC, VPS) it does. On a
// container host (Railway, Render, Heroku, Fly, etc.) it does NOT — every
// deploy/restart can hand the container a brand-new, empty filesystem,
// home folder included. AUTOLEAD_DATA_DIR must then point at a real
// persistent volume mount, or every redeploy wipes the data again.
const looksLikeContainerHost = !!(
  process.env.RAILWAY_PROJECT_ID || process.env.RAILWAY_SERVICE_ID ||
  process.env.RAILWAY_ENVIRONMENT_ID || process.env.RENDER ||
  process.env.DYNO /* Heroku */ || process.env.FLY_APP_NAME
);
if (looksLikeContainerHost && !process.env.AUTOLEAD_DATA_DIR) {
  console.error(
    '\n\uD83D\uDEA8 DATA LOSS RISK: this looks like a container host (Railway/Render/Heroku/Fly), ' +
    'and AUTOLEAD_DATA_DIR is NOT set.\n' +
    '   Right now data is sitting at "' + DATA_ROOT + '" inside the container\'s own filesystem — ' +
    'that is NOT a persistent volume and WILL be wiped on the next deploy or restart.\n' +
    '   Fix: attach a persistent Volume to this service, mount it at e.g. /data, then set the ' +
    'environment variable AUTOLEAD_DATA_DIR=/data and redeploy. See README.md → "Deploying on Railway".\n'
  );
}

const DATA_FILE         = path.join(DATA_ROOT, 'data', 'state.json');
const UPLOADS_DIR        = path.join(DATA_ROOT, 'uploads');
const LEAD_DOCS_DIR      = path.join(UPLOADS_DIR, 'lead_docs');
const AGENT_PHOTOS_DIR   = path.join(UPLOADS_DIR, 'agent_photos');
const SCRIPTS_DIR        = path.join(UPLOADS_DIR, 'scripts');
const BACKUPS_DIR        = path.join(DATA_ROOT, 'backups');
// Original numbers-sheet uploads (.xlsx/.csv etc.) are kept here forever, exactly
// as uploaded — they used to be parsed then deleted; now they're retained so the
// admin can always pull back the exact original file later.
const NUMBER_SHEETS_DIR  = path.join(UPLOADS_DIR, 'number_sheets');
// PDF report archive — generated reports (admin/client dashboards) are stored here
// FOREVER (never auto-deleted), so any party can search by date and re-download later.
const REPORTS_DIR        = path.join(UPLOADS_DIR, 'reports');

// Ensure directories exist
[path.dirname(DATA_FILE), UPLOADS_DIR, LEAD_DOCS_DIR, AGENT_PHOTOS_DIR, SCRIPTS_DIR, BACKUPS_DIR, NUMBER_SHEETS_DIR, REPORTS_DIR].forEach(ensureDir);

function copyRecursiveSync(src, dest) {
  const stat = fs.statSync(src);
  if (stat.isDirectory()) {
    ensureDir(dest);
    // Merge: recurse into every item so pre-created (but empty) destination
    // subfolders like uploads/agent_photos don't cause their contents to be skipped.
    for (const item of fs.readdirSync(src)) copyRecursiveSync(path.join(src, item), path.join(dest, item));
  } else if (!fs.existsSync(dest)) {
    fs.copyFileSync(src, dest);
  }
}

// ─── One-time migration from the OLD in-project data/uploads folders ──────────
// Older runs of this app (or your very first run on this machine) may have data
// sitting inside the project folder. Pull it into the new external location once,
// so you don't lose anything on this transition. Safe to run on every boot —
// it only ever copies files that aren't already present at the destination.
if (DATA_ROOT !== __dirname) {
  try {
    const legacyDataFile = path.join(__dirname, 'data', 'state.json');
    if (!fs.existsSync(DATA_FILE) && fs.existsSync(legacyDataFile)) {
      fs.copyFileSync(legacyDataFile, DATA_FILE);
      console.log('\uD83D\uDCE6 Migrated existing state.json -> ' + DATA_FILE);
    }
  } catch (e) { console.error('Legacy state.json migration skipped:', e.message); }

  try {
    const legacyUploads = path.join(__dirname, 'uploads');
    if (fs.existsSync(legacyUploads)) {
      for (const item of fs.readdirSync(legacyUploads)) {
        if (item === '.gitkeep') continue;
        copyRecursiveSync(path.join(legacyUploads, item), path.join(UPLOADS_DIR, item));
      }
      console.log('\uD83D\uDCE6 Checked uploads/ for anything to migrate -> ' + UPLOADS_DIR);
    }
  } catch (e) { console.error('Legacy uploads migration skipped:', e.message); }
}

console.log('\uD83D\uDCBE Data storage location: ' + DATA_ROOT);
console.log('   (Outside the project folder — updating/re-cloning the code will never touch this.)');

const BREAK_DURATION_MS = 60 * 60 * 1000; // 1 hour
const DELIVERY_DEADLINE_MS = 7 * 24 * 60 * 60 * 1000; // 7 days for delivery completion

// ─── State Management ─────────────────────────────────────────────────────────
function getTodayStr() {
  const now = new Date();
  const ist = new Date(now.getTime() + (5.5 * 60 * 60 * 1000));
  return ist.toISOString().slice(0, 10);
}

function getTomorrowStr() {
  const now = new Date();
  const ist = new Date(now.getTime() + (5.5 * 60 * 60 * 1000));
  ist.setDate(ist.getDate() + 1);
  return ist.toISOString().slice(0, 10);
}

function loadState() {
  return loadStateWithFallback();
}

function createFreshState(preserveAllowedEids) {
  // CRITICAL: Never hardcode allowedEids — always preserve existing ones (names, photos, roles).
  // If none exist yet, start with an empty object so admin can add them fresh.
  const eids = preserveAllowedEids && typeof preserveAllowedEids === 'object'
    ? preserveAllowedEids
    : {};
  return {
    numbers: [],
    agents: {},
    uploadedFiles: [],
    dialedLog: [],
    lastReset: getTodayStr(),
    allowedEids: eids,
    reports: [],
    clientLogs: []
  };
}

function saveState(state) {
  // Atomic write: write to .tmp then rename so a power-cut mid-write
  // never leaves a corrupt state.json — rename is atomic on most OS/FS.
  const tmp = DATA_FILE + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(state, null, 2));
  fs.renameSync(tmp, DATA_FILE);
}

// Extra safety net on top of the external storage location: keep one dated
// snapshot of state.json per day (last 14 days) in BACKUPS_DIR. Cheap insurance
// against an accidental Clear All / Hard Reset, a corrupted write, or anything else.
function backupStateFile() {
  try {
    if (!fs.existsSync(DATA_FILE)) return;
    const todaysBackup = path.join(BACKUPS_DIR, 'state-' + getTodayStr() + '.json');
    if (!fs.existsSync(todaysBackup)) {
      fs.copyFileSync(DATA_FILE, todaysBackup);
    }
    const cutoffMs = Date.now() - 14 * 24 * 60 * 60 * 1000;
    fs.readdirSync(BACKUPS_DIR).forEach(f => {
      const full = path.join(BACKUPS_DIR, f);
      try { if (fs.statSync(full).mtimeMs < cutoffMs) fs.unlinkSync(full); } catch {}
    });
  } catch (e) { console.error('State backup skipped:', e.message); }
}

function loadStateWithFallback() {
  // Try main file first, then .tmp backup if main is corrupt/missing
  for (const f of [DATA_FILE, DATA_FILE + '.tmp']) {
    if (fs.existsSync(f)) {
      try { return JSON.parse(fs.readFileSync(f, 'utf8')); } catch {}
    }
  }
  return createFreshState();
}

function checkDailyReset(state) {
  const today = getTodayStr();
  if (state.lastReset !== today) {
    backupStateFile(); // snapshot yesterday's final state before today's reset mutates it
    for (const id in state.agents) {
      state.agents[id].totalDialedToday = 0;
      state.agents[id].date = today;
      state.agents[id].active = false;
      state.agents[id].currentIndex = null;
      state.agents[id].onBreak = false;
      state.agents[id].breakStartedAt = null;
      state.agents[id].totalBreakMs = 0;
      state.agents[id].currentNumberId = null;
      state.agents[id].firstLoginToday = null;
      state.agents[id].firstLoginDate  = null;
      state.agents[id].onWashroom = false;
      state.agents[id].washroomStartedAt = null;
      state.agents[id].totalWashroomMs = 0;
      state.agents[id].onMeeting = false;
      state.agents[id].meetingStartedAt = null;
      state.agents[id].totalMeetingMs = 0;
      state.agents[id].onTlMode = false;
      state.agents[id].tlModeStartedAt = null;
      state.agents[id].totalTlModeMs = 0;
    }
    state.numbers.forEach(n => {
      if ((n.disposition === 'not_received' || n.disposition === 'switch_off' || n.disposition === 'dead') && n.retryAfter && today >= n.retryAfter && !n.permanent) {
        const dispoCount = (n.retryCounts && n.retryCounts[n.disposition]) || n.retryCount || 0;
        if (dispoCount < 2) {
          n.disposition = null;
          n.retryAfter = null;
          n.dialedBy = null;
          n.dialedAt = null;
          n.assignedTo = null;
        }
      }
    });
    // Trim dialedLog: remove entries older than 90 days to prevent unbounded growth
    const cutoffDate = new Date(Date.now() - 90 * 24 * 60 * 60 * 1000).toISOString();
    if (state.dialedLog && state.dialedLog.length > 0) {
      state.dialedLog = state.dialedLog.filter(entry => entry.timestamp && entry.timestamp >= cutoffDate);
    }
    // Close any client sessions that were still "open" at midnight (tab left open overnight)
    if (state.clientLogs) {
      const midnightMs = Date.now();
      state.clientLogs.forEach(l => {
        if (!l.logoutAt) {
          l.logoutAt = new Date(midnightMs).toISOString();
          l.durationMs = midnightMs - new Date(l.loginAt).getTime();
          l.closedByReset = true;
        }
      });
      // Trim: drop logs older than 90 days
      state.clientLogs = state.clientLogs.filter(l => l.loginAt >= cutoffDate);
    }
    state.lastReset = today;
    saveState(state);
  }
  return state;
}

// ─── Delivery Status Check ────────────────────────────────────────────────────
// Evaluate all recruited riders and auto-mark completed/failed based on time and deliveries
function evaluateRecruitedRiders(state) {
  const now = Date.now();
  state.numbers.forEach(n => {
    if (n.disposition !== 'interested') return;
    if (!n.recruitedAt) return;
    if (n.riderStatus === 'completed' || n.riderStatus === 'removed') return;

    const elapsedMs = now - new Date(n.recruitedAt).getTime();
    const deliveryCount = n.deliveryCount || 0;

    // If deliveryCount >= 31, mark completed
    if (deliveryCount >= 31) {
      n.riderStatus = 'completed';
      n.deliveriesCompleted = true;
      n.completedAt = new Date().toISOString();
      return;
    }

    // If 7 days passed and deliveryCount < 31, mark failed
    if (elapsedMs >= DELIVERY_DEADLINE_MS && deliveryCount < 31) {
      n.riderStatus = 'failed';
      return;
    }
  });
}

let appState = loadState();
// NOTE: We do NOT overwrite allowedEids here anymore.
// If allowedEids is missing entirely (truly fresh install), start empty and let admin add EIDs.
// Previously this block hardcoded names/strings and stomped on saved roles+photos on every deploy.
if (!appState.allowedEids) {
  appState.allowedEids = {};
}
if (!appState.dndNumbers) {
  appState.dndNumbers = [];
}
if (!appState.clientLogs) {
  appState.clientLogs = [];
}
// PDF report archive metadata — reports themselves are kept on disk forever;
// this array just indexes them so they can be searched/downloaded by date.
if (!appState.reports) {
  appState.reports = [];
}

// ─── Auto-create/update the default "Swiggy India" client-panel user ──────────
// The Client Panel needs at least one working login out of the box.
// If EID 9000 exists with an old name (e.g. "kingfisher"), update it.
(function ensureDefaultSwiggyUser() {
  const RESERVED_EID = '9000';
  const existing = appState.allowedEids[RESERVED_EID];
  if (!existing) {
    // Create fresh
    appState.allowedEids[RESERVED_EID] = { name: 'Swiggy India', photo: null, role: 'client' };
    console.log('\uD83D\uDC64 Auto-created default Client Panel login -> EID ' + RESERVED_EID + ' ("Swiggy India")');
  } else {
    // Update name if it's still "kingfisher" or "swiggy" (old defaults)
    const currentName = getEidName(existing).toLowerCase();
    if (currentName === 'kingfisher' || currentName === 'swiggy') {
      const photo = getEidPhoto(existing);
      const role = getEidRole(existing);
      appState.allowedEids[RESERVED_EID] = { name: 'Swiggy India', photo: photo || null, role: role || 'client' };
      console.log('\uD83D\uDC64 Updated default Client Panel login -> EID ' + RESERVED_EID + ' ("Swiggy India")');
    }
  }
})();
appState = checkDailyReset(appState);

// Evaluate recruited riders on startup
evaluateRecruitedRiders(appState);

for (const id in appState.agents) {
  const a = appState.agents[id];
  if (a.active && !a.onBreak) {
    a.needsAutoResume = true;
  }
  a.active = false;
}
saveState(appState);
backupStateFile(); // guarantee at least one snapshot exists per boot, even same-day restarts

setInterval(() => {
  try { saveState(appState); } catch {}
}, 1000);

// Broadcast admin stats every 5 seconds for live timer feel
setInterval(() => {
  try { broadcastAdminStats(); } catch {}
}, 5000);

// Evaluate recruited riders every 60 seconds for auto-completion/failure
setInterval(() => {
  try { evaluateRecruitedRiders(appState); saveState(appState); } catch {}
}, 60000);

// ─── Number helpers ───────────────────────────────────────────────────────────
function getNextNumber(agentId) {
  appState = checkDailyReset(appState);
  const today = getTodayStr();
  // Collect all DND phones to exclude
  const dndPhones = new Set((appState.dndNumbers || []).map(d => d.phone));
  const undialed = appState.numbers.find(n => {
    if (n.dialedBy || n.assignedTo) return false;
    if (n.disposition === 'discard') return false;
    if (n.disposition === 'not_interested') return false;
    if (n.disposition === 'dnd') return false;
    if (n.permanent) return false;
    // Skip if phone is in DND list
    if (dndPhones.has(n.phone)) return false;
    if (n.disposition === 'dead') {
      const deadCount = (n.retryCounts && n.retryCounts.dead) || n.retryCount || 0;
      if (deadCount >= 2) return false;
      if (!n.retryAfter) return false;
      if (n.retryAfter && today < n.retryAfter) return false;
    }
    if (n.disposition === 'followup' && n.followupLockedBy && n.followupLockedBy !== agentId) return false;
    if (n.disposition === 'interested') return false;
    if (n.disposition === 'not_received' || n.disposition === 'switch_off') {
      const dispoCount = (n.retryCounts && n.retryCounts[n.disposition]) || n.retryCount || 0;
      if (dispoCount >= 2) return false;
      if (n.retryAfter && today < n.retryAfter) return false;
    }
    return true;
  });
  if (!undialed) return null;
  undialed.assignedTo = agentId;
  saveState(appState);
  return undialed;
}

function markDialed(agentId, numberId) {
  appState = checkDailyReset(appState);
  const num = appState.numbers.find(n => n.id === numberId);
  if (!num) return;
  const today = getTodayStr();
  num.dialedBy = agentId;
  num.dialedAt = new Date().toISOString();
  num.assignedTo = null;

  const agent = appState.agents[agentId];
  if (agent) {
    agent.totalDialedToday = (agent.totalDialedToday || 0) + 1;
    agent.date = today;
    agent.currentNumberId = null;
  }
  appState.dialedLog.push({
    phone: num.phone, agentId,
    agentName: agent ? agent.name : agentId,
    timestamp: new Date().toISOString()
  });
  saveState(appState);
  broadcastAdminStats();
}

function releaseNumber(agentId, numberId) {
  const num = appState.numbers.find(n => n.id === numberId && n.assignedTo === agentId);
  if (num) { num.assignedTo = null; saveState(appState); }
  const agent = appState.agents[agentId];
  if (agent) agent.currentNumberId = null;
}

// ─── Disposition System ───────────────────────────────────────────────────────
const VALID_DISPOSITIONS = ['dead', 'not_received', 'not_interested', 'followup', 'switch_off', 'interested', 'discard', 'dnd'];
const VALID_RIDER_CATEGORIES = ['Bike', 'Bicycle', 'EV_Scooter', 'Scooter'];

function applyDisposition(agentId, numberId, disposition, extra) {
  appState = checkDailyReset(appState);
  const num = appState.numbers.find(n => n.id === numberId);
  if (!num) return;
  const agent = appState.agents[agentId];
  const now = new Date().toISOString();

  switch (disposition) {
    case 'dead':
      // CNC - retry counting: first time retryAfter tomorrow, second time permanent
      if (!num.retryCounts) num.retryCounts = {};
      if (!num.retryCounts.dead) num.retryCounts.dead = 0;
      num.retryCounts.dead++;
      if (!num.retryCount) num.retryCount = 0;
      num.retryCount++;
      if (num.retryCounts.dead >= 2) {
        // Permanent removal - never dial again
        num.disposition = 'dead';
        num.permanent = true;
        num.retryAfter = null;
      } else {
        num.disposition = 'dead';
        num.retryAfter = getTomorrowStr();
      }
      num.dialedBy = agentId;
      num.dialedAt = now;
      num.assignedTo = null;
      break;
    case 'not_received':
      // CNR - retry counting: first time retryAfter tomorrow, second time permanent
      if (!num.retryCounts) num.retryCounts = {};
      if (!num.retryCounts.not_received) num.retryCounts.not_received = 0;
      num.retryCounts.not_received++;
      if (!num.retryCount) num.retryCount = 0;
      num.retryCount++;
      if (num.retryCounts.not_received >= 2) {
        // Permanent removal - never dial again
        num.disposition = 'not_received';
        num.permanent = true;
        num.retryAfter = null;
      } else {
        num.disposition = 'not_received';
        num.retryAfter = getTomorrowStr();
      }
      num.dialedBy = agentId;
      num.dialedAt = now;
      num.assignedTo = null;
      break;
    case 'not_interested':
      // Permanent removal - never dial again (no 30-day window)
      num.disposition = 'not_interested';
      num.permanent = true;
      num.blockedUntil = null;
      num.dialedBy = agentId;
      num.dialedAt = now;
      num.assignedTo = null;
      break;
    case 'followup':
      // Track followup count - auto-NI after 2 followups
      if (!num.followupCount) num.followupCount = 0;
      num.followupCount++;
      if (num.followupCount > 2) {
        // Auto-convert to not_interested after 2 followups
        num.disposition = 'not_interested';
        num.permanent = true;
        num.blockedUntil = null;
        num.followupDate = null;
        num.followupTime = null;
        num.followupLockedBy = null;
        num.followupName = null;
      } else {
        num.disposition = 'followup';
        num.followupDate = extra && extra.followupDate ? extra.followupDate : null;
        num.followupTime = extra && extra.followupTime ? extra.followupTime : null;
        num.followupName = extra && extra.followupName ? extra.followupName : '';
        num.followupLockedBy = agentId;
        if (extra && extra.vehicleType && VALID_RIDER_CATEGORIES.includes(extra.vehicleType)) {
          num.vehicleType = extra.vehicleType;
        }
      }
      num.dialedBy = agentId;
      num.dialedAt = now;
      num.assignedTo = null;
      break;
    case 'switch_off':
      // Switch Off - retry counting: first time retryAfter tomorrow, second time permanent
      if (!num.retryCounts) num.retryCounts = {};
      if (!num.retryCounts.switch_off) num.retryCounts.switch_off = 0;
      num.retryCounts.switch_off++;
      if (!num.retryCount) num.retryCount = 0;
      num.retryCount++;
      if (num.retryCounts.switch_off >= 2) {
        // Permanent removal - never dial again
        num.disposition = 'switch_off';
        num.permanent = true;
        num.retryAfter = null;
      } else {
        num.disposition = 'switch_off';
        num.retryAfter = getTomorrowStr();
      }
      num.dialedBy = agentId;
      num.dialedAt = now;
      num.assignedTo = null;
      break;
    case 'interested':
      // "Interested" now means "Recruited" — rider has been recruited
      num.disposition = 'interested';
      num.interestedBy = agentId;
      num.interestedAt = now;
      num.recruitedAt = now;
      num.riderName = extra && extra.riderName ? extra.riderName : '';
      num.vehicleType = extra && extra.vehicleType && VALID_RIDER_CATEGORIES.includes(extra.vehicleType) ? extra.vehicleType : '';
      num.remarks = extra && extra.remarks ? extra.remarks : '';
      num.area = extra && extra.area ? extra.area : '';
      num.city = extra && extra.city ? extra.city : '';
      num.riderPhone = num.phone; // phone already exists on the number
      num.deliveryCount = 0;
      num.deliveryTarget = 31;
      num.daysLimit = 7;
      num.deliveryLog = [];
      num.riderStatus = 'active';
      num.deliveriesCompleted = false;
      num.completedAt = null;
      num.dialedBy = agentId;
      num.dialedAt = now;
      num.assignedTo = null;
      break;
    case 'discard':
      // Permanent removal - never ever dial again, dead forever
      num.disposition = 'discard';
      num.permanent = true;
      num.retryAfter = null;
      num.blockedUntil = null;
      num.dialedBy = agentId;
      num.dialedAt = now;
      num.assignedTo = null;
      break;
    case 'dnd':
      // DND - Do Not Disturb - permanent, never dial again
      num.disposition = 'dnd';
      num.permanent = true;
      num.retryAfter = null;
      num.blockedUntil = null;
      num.dialedBy = agentId;
      num.dialedAt = now;
      num.assignedTo = null;
      // Also add to DND list
      if (!appState.dndNumbers) appState.dndNumbers = [];
      if (!appState.dndNumbers.find(d => d.phone === num.phone)) {
        appState.dndNumbers.push({ phone: num.phone, addedAt: now, addedBy: agentId });
      }
      break;
  }

  if (agent) {
    agent.totalDialedToday = (agent.totalDialedToday || 0) + 1;
    agent.currentNumberId = null;
  }
  appState.dialedLog.push({
    phone: num.phone, agentId,
    agentName: agent ? agent.name : agentId,
    timestamp: now,
    disposition: disposition
  });
  saveState(appState);
  broadcastAdminStats();
}

// ─── Break helpers ────────────────────────────────────────────────────────────
function startBreak(agentId) {
  const agent = appState.agents[agentId];
  if (!agent || agent.onBreak) return { error: 'Already on break or agent not found' };
  agent.onBreak = true;
  agent.breakStartedAt = Date.now();
  if (!agent.totalBreakMs) agent.totalBreakMs = 0;
  saveState(appState);
  broadcastAdminStats();
  return { success: true, breakStartedAt: agent.breakStartedAt };
}

function endBreak(agentId) {
  const agent = appState.agents[agentId];
  if (!agent || !agent.onBreak) return { error: 'Not on break' };
  const elapsed = Date.now() - (agent.breakStartedAt || Date.now());
  agent.totalBreakMs = (agent.totalBreakMs || 0) + elapsed;
  agent.onBreak = false;
  agent.breakStartedAt = null;
  saveState(appState);
  broadcastAdminStats();
  return { success: true, totalBreakMs: agent.totalBreakMs };
}

function getBreakMsRemaining(agent) {
  if (!agent.onBreak) return BREAK_DURATION_MS - (agent.totalBreakMs || 0);
  const elapsed = Date.now() - (agent.breakStartedAt || Date.now());
  return BREAK_DURATION_MS - ((agent.totalBreakMs || 0) + elapsed);
}

// ─── Washroom helpers ─────────────────────────────────────────────────────────
function startWashroom(agentId) {
  const agent = appState.agents[agentId];
  if (!agent) return { error: 'Agent not found' };
  if (agent.onWashroom) return { error: 'Already in washroom' };
  if (agent.onBreak) return { error: 'Cannot use washroom while on break' };
  if (agent.onMeeting) return { error: 'Cannot use washroom while in meeting' };
  agent.onWashroom = true;
  agent.washroomStartedAt = Date.now();
  if (!agent.totalWashroomMs) agent.totalWashroomMs = 0;
  saveState(appState);
  broadcastAdminStats();
  return { success: true, washroomStartedAt: agent.washroomStartedAt };
}

function endWashroom(agentId) {
  const agent = appState.agents[agentId];
  if (!agent || !agent.onWashroom) return { error: 'Not in washroom' };
  const elapsed = Date.now() - (agent.washroomStartedAt || Date.now());
  agent.totalWashroomMs = (agent.totalWashroomMs || 0) + elapsed;
  agent.onWashroom = false;
  agent.washroomStartedAt = null;
  saveState(appState);
  broadcastAdminStats();
  return { success: true, totalWashroomMs: agent.totalWashroomMs };
}

// ─── Meeting helpers ──────────────────────────────────────────────────────────
function startMeeting(agentId) {
  const agent = appState.agents[agentId];
  if (!agent) return { error: 'Agent not found' };
  if (agent.onMeeting) return { error: 'Already in meeting' };
  if (agent.onBreak) return { error: 'Cannot start meeting while on break' };
  if (agent.onWashroom) return { error: 'Cannot start meeting while in washroom' };
  agent.onMeeting = true;
  agent.meetingStartedAt = Date.now();
  if (!agent.totalMeetingMs) agent.totalMeetingMs = 0;
  saveState(appState);
  broadcastAdminStats();
  return { success: true, meetingStartedAt: agent.meetingStartedAt };
}

function endMeeting(agentId) {
  const agent = appState.agents[agentId];
  if (!agent || !agent.onMeeting) return { error: 'Not in meeting' };
  const elapsed = Date.now() - (agent.meetingStartedAt || Date.now());
  agent.totalMeetingMs = (agent.totalMeetingMs || 0) + elapsed;
  agent.onMeeting = false;
  agent.meetingStartedAt = null;
  saveState(appState);
  broadcastAdminStats();
  return { success: true, totalMeetingMs: agent.totalMeetingMs };
}

// ─── TL Mode helpers ──────────────────────────────────────────────────────────
function startTlMode(agentId) {
  const agent = appState.agents[agentId];
  if (!agent) return { error: 'Agent not found' };
  if (agent.onTlMode) return { error: 'Already in TL mode' };
  agent.onTlMode = true;
  agent.tlModeStartedAt = Date.now();
  if (!agent.totalTlModeMs) agent.totalTlModeMs = 0;
  saveState(appState);
  broadcastAdminStats();
  return { success: true, tlModeStartedAt: agent.tlModeStartedAt };
}

function endTlMode(agentId) {
  const agent = appState.agents[agentId];
  if (!agent || !agent.onTlMode) return { error: 'Not in TL mode' };
  const elapsed = Date.now() - (agent.tlModeStartedAt || Date.now());
  agent.totalTlModeMs = (agent.totalTlModeMs || 0) + elapsed;
  agent.onTlMode = false;
  agent.tlModeStartedAt = null;
  saveState(appState);
  broadcastAdminStats();
  return { success: true, totalTlModeMs: agent.totalTlModeMs };
}

// ─── Admin broadcast ──────────────────────────────────────────────────────────
function broadcastAdminStats() {
  const stats = getAdminStats();
  io.to('admin-room').emit('stats-update', stats);
}

function getAdminStats() {
  appState = checkDailyReset(appState);
  // Evaluate recruited riders on every stats check
  evaluateRecruitedRiders(appState);

  const total = appState.numbers.length;
  const dialed = appState.numbers.filter(n => n.dialedBy).length;
  const assigned = appState.numbers.filter(n => n.assignedTo && !n.dialedBy).length;
  const remaining = total - dialed - assigned;

  const agentStats = Object.entries(appState.agents)
    .filter(([id]) => { const eid = id.startsWith('emp_') ? id.slice(4) : id; return (appState.allowedEids[eid] || {}).role !== 'client'; })
    .map(([id, a]) => {
    const liveBreakMs = a.onBreak ? (Date.now() - (a.breakStartedAt || Date.now())) : 0;
    const totalBreakMs = (a.totalBreakMs || 0) + liveBreakMs;
    const breakRemaining = Math.max(0, BREAK_DURATION_MS - totalBreakMs);

    const liveWashroomMs = a.onWashroom ? (Date.now() - (a.washroomStartedAt || Date.now())) : 0;
    const totalWashroomMs = (a.totalWashroomMs || 0) + liveWashroomMs;

    const liveMeetingMs = a.onMeeting ? (Date.now() - (a.meetingStartedAt || Date.now())) : 0;
    const totalMeetingMs = (a.totalMeetingMs || 0) + liveMeetingMs;

    const liveTlModeMs = a.onTlMode ? (Date.now() - (a.tlModeStartedAt || Date.now())) : 0;
    const totalTlModeMs = (a.totalTlModeMs || 0) + liveTlModeMs;

    const firstLogin = a.firstLoginToday || null;
    const lateLogin  = firstLogin ? (firstLogin > '10:00') : false;

    return {
      id, name: a.name, active: a.active,
      totalDialedToday: a.totalDialedToday || 0,
      date: a.date,
      onBreak: a.onBreak || false,
      totalBreakMs,
      breakRemaining,
      breakAllowedMs: BREAK_DURATION_MS,
      onWashroom: a.onWashroom || false,
      washroomStartedAt: a.washroomStartedAt || null,
      totalWashroomMs,
      onMeeting: a.onMeeting || false,
      meetingStartedAt: a.meetingStartedAt || null,
      totalMeetingMs,
      onTlMode: a.onTlMode || false,
      tlModeStartedAt: a.tlModeStartedAt || null,
      totalTlModeMs,
      firstLoginToday: firstLogin,
      lateLogin
    };
  });

  const fileStats = appState.uploadedFiles.map(f => {
    const { sheetPath, ...publicFields } = f;
    const fileNums = appState.numbers.filter(n => n.file === f.id);
    return {
      ...publicFields,
      total: fileNums.length,
      dialed: fileNums.filter(n => n.dialedBy).length,
      remaining: fileNums.filter(n => !n.dialedBy).length,
      hasOriginal: !!(sheetPath && fs.existsSync(sheetPath))
    };
  });

  return {
    total, dialed, assigned, remaining, agentStats, fileStats,
    today: getTodayStr(),
    interestedCount: appState.numbers.filter(n => n.disposition === 'interested').length,
    recruitedCount: appState.numbers.filter(n => n.disposition === 'interested' && n.riderStatus === 'active').length,
    deliveryCompletedCount: appState.numbers.filter(n => n.disposition === 'interested' && n.riderStatus === 'completed').length,
    deliveryFailedCount: appState.numbers.filter(n => n.disposition === 'interested' && n.riderStatus === 'failed').length,
    followupCount: appState.numbers.filter(n => n.disposition === 'followup').length,
    discardCount: appState.numbers.filter(n => n.disposition === 'discard').length,
    notInterestedCount: appState.numbers.filter(n => n.disposition === 'not_interested').length,
    dndCount: (appState.dndNumbers || []).length,
    comingBackTomorrow: appState.numbers.filter(n => (n.disposition === 'not_received' || n.disposition === 'switch_off' || n.disposition === 'dead') && n.retryAfter && !n.permanent && (n.retryCount || 0) < 2 && getTodayStr() < n.retryAfter).length,
    overdueRecruitedCount: appState.numbers.filter(n => n.disposition === 'interested' && !n.deliveriesCompleted && n.riderStatus === 'active' && (Date.now() - new Date(n.recruitedAt || n.interestedAt).getTime()) >= DELIVERY_DEADLINE_MS).length
  };
}

// ─── Express Setup ─────────────────────────────────────────────────────────────
app.use(express.json());

// ─── CORS — allow other CRMs to call the cross-sync endpoints ─────────────────
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

// ─── Password Protection Middleware (Admin & Client) ──────────────────────────
const PANEL_PASSWORD = process.env.PANEL_PASSWORD || 'Jaimatadi02"';

function getPanelAuthCookie(req) {
  const cookies = (req.headers.cookie || '').split(';').reduce((acc, c) => {
    const [k, ...v] = c.trim().split('=');
    if (k) acc[k.trim()] = v.join('=');
    return acc;
  }, {});
  return cookies['panel_auth'] || '';
}

function isAuthenticated(req) {
  const token = Buffer.from(PANEL_PASSWORD).toString('base64');
  return getPanelAuthCookie(req) === token;
}

function sendPasswordPage(res) {
  res.send(`<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>Access Protected</title>
<style>*{box-sizing:border-box;margin:0;padding:0}body{min-height:100vh;display:flex;align-items:center;justify-content:center;background:#f8f9fa;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif}
.card{background:#fff;border-radius:16px;padding:40px;box-shadow:0 4px 24px rgba(0,0,0,.08);width:90%;max-width:380px;text-align:center}
.card h2{font-size:20px;margin-bottom:6px;color:#1a1a2e}.card p{font-size:13px;color:#6b7280;margin-bottom:24px}
input{width:100%;padding:14px 18px;border-radius:10px;border:1.5px solid #e5e7eb;font-size:16px;outline:none;margin-bottom:16px;transition:border-color .2s}
input:focus{border-color:#FC8019}
button{width:100%;padding:14px;border-radius:10px;border:none;background:#FC8019;color:#fff;font-size:16px;font-weight:700;cursor:pointer;transition:all .2s}
button:hover{background:#e8720f;transform:translateY(-1px)}
.err{color:#ef4444;font-size:13px;margin-top:12px;display:none}</style></head>
<body><div class="card"><h2>🔒 Access Protected</h2><p>Enter the panel password to continue</p>
<form method="POST" action="/auth/panel-login"><input type="password" name="password" id="pwd" placeholder="Enter password" autofocus>
<button type="submit">Enter →</button></form><div class="err" id="err"></div></div>
<script>const u=new URLSearchParams(window.location.search);if(u.get('err')==='1'){document.getElementById('err').style.display='block';document.getElementById('err').textContent='Incorrect password. Try again.'}</script></body></html>`);
}

// Block static access to /admin/ folder without auth
app.use('/admin', (req, res, next) => {
  if (isAuthenticated(req)) return next();
  sendPasswordPage(res);
});

// Password login endpoint
app.use(require('express').urlencoded({ extended: false }));
app.post('/auth/panel-login', (req, res) => {
  const { password } = req.body;
  if (password === PANEL_PASSWORD) {
    const token = Buffer.from(PANEL_PASSWORD).toString('base64');
    res.setHeader('Set-Cookie', 'panel_auth=' + token + '; Path=/; HttpOnly; Max-Age=86400; SameSite=Lax');
    // Redirect back to referrer or admin
    const ref = req.headers.referer || '/admin';
    const target = ref.includes('/client') ? '/client' : '/admin';
    return res.redirect(302, target);
  }
  // Wrong password — redirect back with error
  const ref = req.headers.referer || '/admin';
  const target = ref.includes('/client') ? '/client' : '/admin';
  res.redirect(302, target + '?err=1');
});

app.use(express.static(path.join(__dirname, 'public')));

// Multer for number file uploads — keep the ORIGINAL file permanently (no longer
// deleted after parsing) so the admin can retrieve the exact sheet they uploaded.
const numberSheetStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, NUMBER_SHEETS_DIR),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname) || '.xlsx';
    cb(null, uuidv4() + ext);
  }
});
const numberUpload = multer({ storage: numberSheetStorage });

// Multer for lead document ZIP uploads — store with original name under lead_docs
const docStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, LEAD_DOCS_DIR),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname) || '.zip';
    cb(null, uuidv4() + ext);
  }
});
const docUpload = multer({
  storage: docStorage,
  limits: { fileSize: 100 * 1024 * 1024 }, // 100MB
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/zip' || file.mimetype === 'application/x-zip-compressed' ||
        file.originalname.toLowerCase().endsWith('.zip')) {
      cb(null, true);
    } else {
      cb(new Error('Only ZIP files are allowed'));
    }
  }
});

// Multer for agent photo uploads
const agentPhotoStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, AGENT_PHOTOS_DIR),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname) || '.jpg';
    cb(null, (req.params.eid || 'photo') + '_' + Date.now() + ext);
  }
});
const agentPhotoUpload = multer({
  storage: agentPhotoStorage,
  limits: { fileSize: 5 * 1024 * 1024 }, // 5MB
  fileFilter: (req, file, cb) => {
    if (file.mimetype.startsWith('image/')) {
      cb(null, true);
    } else {
      cb(new Error('Only image files are allowed'));
    }
  }
});

app.post('/api/admin/upload', numberUpload.single('file'), (req, res) => {
  try {
    const wb = XLSX.readFile(req.file.path);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const fileId = uuidv4();
    const phones = [];
    const existingPhones = new Set(appState.numbers.map(n => n.phone));
    let skipped = 0;
    rows.forEach((row, i) => {
      if (i === 0 && isNaN(row[0])) return;
      const phone = String(row[0] || '').trim().replace(/\s+/g, '');
      if (!phone || phone.length < 7) return;
      if (existingPhones.has(phone)) { skipped++; return; }
      existingPhones.add(phone);
      const name = row[1] ? String(row[1]).trim() : '';
      phones.push({ id: uuidv4(), phone, name, file: fileId, assignedTo: null, dialedBy: null, dialedAt: null });
    });
    appState.numbers.push(...phones);
    appState.uploadedFiles.push({ id: fileId, name: req.file.originalname, uploadedAt: new Date().toISOString(), total: phones.length, sheetPath: req.file.path });
    saveState(appState);
    broadcastAdminStats();
    res.json({ success: true, count: phones.length, skipped, fileId });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── Lead Document ZIP Upload (Agent) ─────────────────────────────────────────
app.post('/api/agent/upload-doc-zip/:numberId', docUpload.single('docZip'), (req, res) => {
  try {
    const { numberId } = req.params;
    const { agentId } = req.body;
    if (!agentId || !numberId) {
      if (req.file) fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: 'agentId and numberId are required' });
    }
    const num = appState.numbers.find(n => n.id === numberId);
    if (!num) {
      if (req.file) fs.unlinkSync(req.file.path);
      return res.status(404).json({ error: 'Rider not found' });
    }
    if (num.disposition !== 'interested') {
      if (req.file) fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: 'Rider is not marked as recruited' });
    }
    if (num.interestedBy !== agentId) {
      if (req.file) fs.unlinkSync(req.file.path);
      return res.status(403).json({ error: 'This rider is not assigned to you' });
    }
    if (!req.file) {
      return res.status(400).json({ error: 'No ZIP file uploaded' });
    }

    // Delete old zip if exists
    if (num.docZipPath && fs.existsSync(num.docZipPath)) {
      try { fs.unlinkSync(num.docZipPath); } catch {}
    }

    // Store ZIP (kept for documentation purposes)
    num.docZipPath = req.file.path;
    num.docZipName = req.file.originalname || 'documents.zip';
    num.docZipUploadedAt = new Date().toISOString();

    saveState(appState);
    broadcastAdminStats();
    res.json({ success: true, docZipName: num.docZipName, docZipUploadedAt: num.docZipUploadedAt });
  } catch (e) {
    if (req.file) { try { fs.unlinkSync(req.file.path); } catch {} }
    res.status(500).json({ error: e.message });
  }
});

// ─── Admin: Download lead doc ZIP ─────────────────────────────────────────────
app.get('/api/admin/download-doc-zip/:numberId', (req, res) => {
  const { numberId } = req.params;
  const num = appState.numbers.find(n => n.id === numberId);
  if (!num) return res.status(404).json({ error: 'Rider not found' });
  if (!num.docZipPath || !fs.existsSync(num.docZipPath)) {
    return res.status(404).json({ error: 'No document ZIP found for this rider' });
  }
  const downloadName = num.docZipName || 'documents.zip';
  res.download(num.docZipPath, downloadName);
});

app.get('/api/admin/stats', (req, res) => res.json(getAdminStats()));

// ─── Disposition API Endpoints ────────────────────────────────────────────────
app.post('/api/agent/disposition', (req, res) => {
  const { agentId, numberId, disposition, followupDate, followupTime, followupName, riderName, vehicleType, remarks, area, city } = req.body;
  if (!agentId || !numberId || !disposition) {
    return res.status(400).json({ error: 'agentId, numberId, and disposition are required' });
  }
  if (!VALID_DISPOSITIONS.includes(disposition)) {
    return res.status(400).json({ error: 'Invalid disposition. Must be one of: ' + VALID_DISPOSITIONS.join(', ') });
  }
  applyDisposition(agentId, numberId, disposition, { followupDate, followupTime, followupName, riderName, vehicleType, remarks, area, city });
  const nextNum = getNextNumber(agentId);
  const agent = appState.agents[agentId];
  if (nextNum && agent) {
    agent.currentNumberId = nextNum.id;
    saveState(appState);
  }
  res.json({ success: true, nextNumber: nextNum ? { numberId: nextNum.id, phone: nextNum.phone, name: nextNum.name || '' } : null });
});

// ─── Delivery Tracking Endpoints ──────────────────────────────────────────────

// POST /api/agent/update-delivery - Agent updates a rider's delivery count for today
app.post('/api/agent/update-delivery', (req, res) => {
  const { agentId, numberId, todayDeliveries } = req.body;
  if (!agentId || !numberId) {
    return res.status(400).json({ error: 'agentId and numberId are required' });
  }
  if (todayDeliveries === undefined || todayDeliveries === null || isNaN(Number(todayDeliveries))) {
    return res.status(400).json({ error: 'todayDeliveries must be a valid number' });
  }
  const deliveries = Number(todayDeliveries);
  if (deliveries < 0) {
    return res.status(400).json({ error: 'todayDeliveries cannot be negative' });
  }
  const num = appState.numbers.find(n => n.id === numberId);
  if (!num) return res.status(404).json({ error: 'Rider not found' });
  if (num.disposition !== 'interested') return res.status(400).json({ error: 'Rider is not marked as recruited' });
  if (num.riderStatus === 'completed') return res.status(400).json({ error: 'Rider already completed deliveries' });
  if (num.riderStatus === 'removed') return res.status(400).json({ error: 'Rider has been removed' });

  const today = getTodayStr();

  // Initialize deliveryLog if needed
  if (!num.deliveryLog) num.deliveryLog = [];

  // Check if there's already an entry for today — update it instead of adding duplicate
  const existingEntry = num.deliveryLog.find(entry => entry.date === today);
  if (existingEntry) {
    // Update existing entry
    const oldCount = existingEntry.count;
    existingEntry.count = deliveries;
    existingEntry.updatedBy = agentId;
    // Adjust deliveryCount: subtract old, add new
    num.deliveryCount = (num.deliveryCount || 0) - oldCount + deliveries;
  } else {
    // Add new entry
    num.deliveryLog.push({ date: today, count: deliveries, updatedBy: agentId });
    num.deliveryCount = (num.deliveryCount || 0) + deliveries;
  }

  // Check if delivery target reached
  if (num.deliveryCount >= 31) {
    num.riderStatus = 'completed';
    num.deliveriesCompleted = true;
    num.completedAt = new Date().toISOString();
  }

  saveState(appState);
  broadcastAdminStats();
  res.json({
    success: true,
    numberId,
    deliveryCount: num.deliveryCount,
    deliveryTarget: num.deliveryTarget || 31,
    riderStatus: num.riderStatus,
    deliveriesCompleted: num.deliveriesCompleted || false
  });
});

// GET /api/admin/recruited - Returns all recruited riders (active, delivery in progress)
app.get('/api/admin/recruited', (req, res) => {
  const now = Date.now();
  evaluateRecruitedRiders(appState);
  const recruited = appState.numbers.filter(n => n.disposition === 'interested' && n.riderStatus === 'active').map(n => {
    const agent = appState.agents[n.interestedBy];
    const elapsedMs = now - new Date(n.recruitedAt || n.interestedAt).getTime();
    const hoursElapsed = elapsedMs / (1000 * 60 * 60);
    const hoursRemaining = Math.max(0, (7 * 24) - hoursElapsed);
    const overdue = hoursRemaining <= 0;
    return {
      id: n.id, phone: n.phone, name: n.riderName || n.name || '',
      riderName: n.riderName || '',
      vehicleType: n.vehicleType || '',
      remarks: n.remarks || '',
      area: n.area || '',
      city: n.city || '',
      riderPhone: n.riderPhone || n.phone,
      interestedBy: agent ? agent.name : n.interestedBy,
      interestedByAgentId: n.interestedBy,
      recruitedAt: n.recruitedAt || n.interestedAt,
      interestedAt: n.interestedAt,
      deliveryCount: n.deliveryCount || 0,
      deliveryTarget: n.deliveryTarget || 31,
      daysLimit: n.daysLimit || 7,
      deliveryLog: n.deliveryLog || [],
      riderStatus: n.riderStatus || 'active',
      deliveriesCompleted: n.deliveriesCompleted || false,
      completedAt: n.completedAt || null,
      hoursRemaining: Math.round(hoursRemaining * 100) / 100,
      overdue
    };
  });
  res.json(recruited);
});

// GET /api/admin/delivery-completed - Riders who completed 31 deliveries
app.get('/api/admin/delivery-completed', (req, res) => {
  evaluateRecruitedRiders(appState);
  const completed = appState.numbers.filter(n => n.disposition === 'interested' && n.riderStatus === 'completed').map(n => {
    const agent = appState.agents[n.interestedBy];
    return {
      id: n.id, phone: n.phone, name: n.riderName || n.name || '',
      riderName: n.riderName || '',
      vehicleType: n.vehicleType || '',
      remarks: n.remarks || '',
      area: n.area || '',
      city: n.city || '',
      riderPhone: n.riderPhone || n.phone,
      interestedBy: agent ? agent.name : n.interestedBy,
      interestedByAgentId: n.interestedBy,
      recruitedAt: n.recruitedAt || n.interestedAt,
      deliveryCount: n.deliveryCount || 0,
      deliveryTarget: n.deliveryTarget || 31,
      deliveryLog: n.deliveryLog || [],
      riderStatus: 'completed',
      deliveriesCompleted: true,
      completedAt: n.completedAt || null,
      adminStatus: n.adminStatus || ''
    };
  });
  res.json(completed);
});

// GET /api/admin/delivery-failed - Riders who failed (7 days passed, <31 deliveries)
app.get('/api/admin/delivery-failed', (req, res) => {
  evaluateRecruitedRiders(appState);
  const failed = appState.numbers.filter(n => n.disposition === 'interested' && n.riderStatus === 'failed').map(n => {
    const agent = appState.agents[n.interestedBy];
    return {
      id: n.id, phone: n.phone, name: n.riderName || n.name || '',
      riderName: n.riderName || '',
      vehicleType: n.vehicleType || '',
      remarks: n.remarks || '',
      area: n.area || '',
      city: n.city || '',
      riderPhone: n.riderPhone || n.phone,
      interestedBy: agent ? agent.name : n.interestedBy,
      interestedByAgentId: n.interestedBy,
      recruitedAt: n.recruitedAt || n.interestedAt,
      deliveryCount: n.deliveryCount || 0,
      deliveryTarget: n.deliveryTarget || 31,
      deliveryLog: n.deliveryLog || [],
      riderStatus: 'failed',
      deliveriesCompleted: false,
      adminStatus: n.adminStatus || ''
    };
  });
  res.json(failed);
});

// Legacy endpoint - still returns all interested (recruited) riders pending delivery
app.get('/api/admin/interested', (req, res) => {
  const now = Date.now();
  evaluateRecruitedRiders(appState);
  const interested = appState.numbers.filter(n => n.disposition === 'interested' && !n.deliveriesCompleted).map(n => {
    const agent = appState.agents[n.interestedBy];
    const elapsedMs = now - new Date(n.recruitedAt || n.interestedAt).getTime();
    const hoursElapsed = elapsedMs / (1000 * 60 * 60);
    const hoursRemaining = Math.max(0, (7 * 24) - hoursElapsed);
    const overdue = hoursRemaining <= 0;
    return {
      id: n.id, phone: n.phone, name: n.riderName || n.name || '',
      riderName: n.riderName || '',
      vehicleType: n.vehicleType || '',
      remarks: n.remarks || '',
      area: n.area || '',
      city: n.city || '',
      riderPhone: n.riderPhone || n.phone,
      interestedBy: agent ? agent.name : n.interestedBy,
      interestedByAgentId: n.interestedBy,
      interestedAt: n.interestedAt,
      recruitedAt: n.recruitedAt || n.interestedAt,
      deliveryCount: n.deliveryCount || 0,
      deliveryTarget: n.deliveryTarget || 31,
      daysLimit: n.daysLimit || 7,
      riderStatus: n.riderStatus || 'active',
      deliveriesCompleted: n.deliveriesCompleted || false,
      completedAt: n.completedAt || null,
      hoursRemaining: Math.round(hoursRemaining * 100) / 100,
      overdue
    };
  });
  res.json(interested);
});

app.get('/api/admin/followups', (req, res) => {
  const now = new Date();
  const istNow = new Date(now.getTime() + 5.5 * 60 * 60 * 1000);
  const followups = appState.numbers.filter(n => n.disposition === 'followup').map(n => {
    const agent = appState.agents[n.followupLockedBy];
    let overdue = false;
    if (n.followupDate) {
      const fDateStr = n.followupDate + 'T' + (n.followupTime || '23:59') + ':00';
      const fDate = new Date(fDateStr);
      overdue = istNow > fDate;
    }
    return {
      id: n.id, phone: n.phone, name: n.name || '',
      followupLockedBy: agent ? agent.name : n.followupLockedBy,
      followupLockedByAgentId: n.followupLockedBy,
      followupDate: n.followupDate,
      followupTime: n.followupTime,
      followupName: n.followupName || '',
      followupCount: n.followupCount || 0,
      vehicleType: n.vehicleType || '',
      overdue
    };
  });
  // Sort by nearest date and time
  followups.sort((a, b) => {
    const dateA = (a.followupDate || '9999-12-31') + ' ' + (a.followupTime || '23:59');
    const dateB = (b.followupDate || '9999-12-31') + ' ' + (b.followupTime || '23:59');
    return dateA.localeCompare(dateB);
  });
  res.json(followups);
});

app.get('/api/agent/interested/:agentId', (req, res) => {
  const agentId = req.params.agentId;
  const now = Date.now();
  evaluateRecruitedRiders(appState);
  const interested = appState.numbers.filter(n => n.disposition === 'interested' && n.interestedBy === agentId && !n.deliveriesCompleted).map(n => {
    const elapsedMs = now - new Date(n.recruitedAt || n.interestedAt).getTime();
    const hoursElapsed = elapsedMs / (1000 * 60 * 60);
    const hoursRemaining = Math.max(0, (7 * 24) - hoursElapsed);
    return {
      id: n.id, phone: n.phone, name: n.riderName || n.name || '',
      riderName: n.riderName || '',
      vehicleType: n.vehicleType || '',
      remarks: n.remarks || '',
      area: n.area || '',
      city: n.city || '',
      riderPhone: n.riderPhone || n.phone,
      recruitedAt: n.recruitedAt || n.interestedAt,
      interestedAt: n.interestedAt,
      deliveryCount: n.deliveryCount || 0,
      deliveryTarget: n.deliveryTarget || 31,
      daysLimit: n.daysLimit || 7,
      deliveryLog: n.deliveryLog || [],
      riderStatus: n.riderStatus || 'active',
      deliveriesCompleted: n.deliveriesCompleted || false,
      completedAt: n.completedAt || null,
      hoursRemaining: Math.round(hoursRemaining * 100) / 100
    };
  });
  res.json(interested);
});

app.get('/api/agent/followups/:agentId', (req, res) => {
  const agentId = req.params.agentId;
  const now = new Date();
  const istNow = new Date(now.getTime() + 5.5 * 60 * 60 * 1000);
  const followups = appState.numbers.filter(n => n.disposition === 'followup' && n.followupLockedBy === agentId).map(n => {
    let overdue = false;
    if (n.followupDate) {
      const fDateStr = n.followupDate + 'T' + (n.followupTime || '23:59') + ':00';
      const fDate = new Date(fDateStr);
      overdue = istNow > fDate;
    }
    return {
      id: n.id, phone: n.phone, name: n.name || '',
      followupDate: n.followupDate,
      followupTime: n.followupTime,
      followupName: n.followupName || '',
      followupCount: n.followupCount || 0,
      vehicleType: n.vehicleType || '',
      overdue
    };
  });
  // Sort by nearest date and time
  followups.sort((a, b) => {
    const dateA = (a.followupDate || '9999-12-31') + ' ' + (a.followupTime || '23:59');
    const dateB = (b.followupDate || '9999-12-31') + ' ' + (b.followupTime || '23:59');
    return dateA.localeCompare(dateB);
  });
  res.json(followups);
});

// Agent marks deliveries completed (manual override if needed)
app.post('/api/agent/mark-documentation-complete', (req, res) => {
  const { agentId, numberId } = req.body;
  if (!agentId || !numberId) {
    return res.status(400).json({ error: 'agentId and numberId are required' });
  }
  const num = appState.numbers.find(n => n.id === numberId);
  if (!num) return res.status(404).json({ error: 'Number not found' });
  if (num.disposition !== 'interested') return res.status(400).json({ error: 'Rider is not marked as recruited' });
  if (num.interestedBy !== agentId) return res.status(403).json({ error: 'This rider is not assigned to you' });
  num.deliveriesCompleted = true;
  num.completedAt = new Date().toISOString();
  num.riderStatus = 'completed';
  num.adminStatus = num.adminStatus || 'Completed';
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true, numberId, deliveriesCompleted: true, completedAt: num.completedAt });
});

app.post('/api/admin/transfer-interested', (req, res) => {
  const { numberId, newAgentId } = req.body;
  if (!numberId || !newAgentId) {
    return res.status(400).json({ error: 'numberId and newAgentId are required' });
  }
  if (!appState.agents[newAgentId]) {
    const eidMatch = newAgentId.match(/^emp_(\d+)$/);
    if (!eidMatch || !appState.allowedEids[eidMatch[1]]) {
      return res.status(404).json({ error: 'Target agent not found' });
    }
    const eid = eidMatch[1];
    appState.agents[newAgentId] = {
      name: getEidName(appState.allowedEids[eid]),
      employeeId: eid,
      active: false,
      totalDialedToday: 0,
      date: getTodayStr(),
      currentIndex: null,
      onBreak: false,
      breakStartedAt: null,
      totalBreakMs: 0,
      currentNumberId: null,
      firstLoginToday: null,
      firstLoginDate: null
    };
  }
  const num = appState.numbers.find(n => n.id === numberId);
  if (!num) return res.status(404).json({ error: 'Number not found' });
  if (num.disposition !== 'interested') return res.status(400).json({ error: 'Rider is not marked as recruited' });
  num.interestedBy = newAgentId;
  num.interestedAt = new Date().toISOString();
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true, numberId, newAgentId, interestedAt: num.interestedAt });
});

app.post('/api/agent/add-interested', (req, res) => {
  const { agentId, phone, riderName, vehicleType, remarks, area, city } = req.body;
  if (!agentId || !phone) {
    return res.status(400).json({ error: 'agentId and phone are required' });
  }
  if (!appState.agents[agentId]) {
    return res.status(404).json({ error: 'Agent not found' });
  }
  if (vehicleType && !VALID_RIDER_CATEGORIES.includes(vehicleType)) {
    return res.status(400).json({ error: 'Invalid vehicle type. Must be one of: ' + VALID_RIDER_CATEGORIES.join(', ') });
  }
  const existingNumber = appState.numbers.find(n => n.phone === phone);
  if (existingNumber) {
    // If existing number was marked as not_interested, switch_off, dead (CNC), or discard
    // allow overriding to interested (recruited)
    const overridableDispositions = ['not_interested', 'switch_off', 'dead', 'discard', 'not_received'];
    if (overridableDispositions.includes(existingNumber.disposition) || existingNumber.permanent) {
      // Override: convert to interested (recruited)
      const now = new Date().toISOString();
      existingNumber.disposition = 'interested';
      existingNumber.permanent = false;
      existingNumber.retryAfter = null;
      existingNumber.blockedUntil = null;
      existingNumber.interestedBy = agentId;
      existingNumber.interestedAt = now;
      existingNumber.recruitedAt = now;
      existingNumber.riderName = riderName || '';
      existingNumber.name = riderName || existingNumber.name || '';
      existingNumber.vehicleType = vehicleType || '';
      existingNumber.remarks = remarks || '';
      existingNumber.area = area || '';
      existingNumber.city = city || '';
      existingNumber.riderPhone = existingNumber.phone;
      existingNumber.deliveryCount = 0;
      existingNumber.deliveryTarget = 31;
      existingNumber.daysLimit = 7;
      existingNumber.deliveryLog = [];
      existingNumber.riderStatus = 'active';
      existingNumber.deliveriesCompleted = false;
      existingNumber.completedAt = null;
      existingNumber.docZipPath = null;
      existingNumber.docZipName = null;
      existingNumber.dialedBy = agentId;
      existingNumber.dialedAt = now;
      existingNumber.assignedTo = null;
      appState.dialedLog.push({
        phone, agentId,
        agentName: appState.agents[agentId] ? appState.agents[agentId].name : agentId,
        timestamp: now,
        disposition: 'interested'
      });
      saveState(appState);
      broadcastAdminStats();
      return res.json({ success: true, entry: existingNumber });
    }
    // If it's already interested or followup, don't allow duplicate
    return res.status(409).json({ error: 'This phone number already exists in the system as ' + (existingNumber.disposition || 'active') });
  }
  // Check DND list
  if (appState.dndNumbers && appState.dndNumbers.find(d => d.phone === phone)) {
    return res.status(409).json({ error: 'This number is in the DND list and cannot be added' });
  }
  const now = new Date().toISOString();
  const newEntry = {
    id: uuidv4(),
    phone,
    name: riderName || '',
    file: null,
    assignedTo: null,
    dialedBy: agentId,
    dialedAt: now,
    disposition: 'interested',
    interestedBy: agentId,
    interestedAt: now,
    recruitedAt: now,
    riderName: riderName || '',
    vehicleType: vehicleType || '',
    remarks: remarks || '',
    area: area || '',
    city: city || '',
    riderPhone: phone,
    deliveryCount: 0,
    deliveryTarget: 31,
    daysLimit: 7,
    deliveryLog: [],
    riderStatus: 'active',
    deliveriesCompleted: false,
    completedAt: null,
    docZipPath: null,
    docZipName: null
  };
  appState.numbers.push(newEntry);
  appState.dialedLog.push({
    phone, agentId,
    agentName: appState.agents[agentId] ? appState.agents[agentId].name : agentId,
    timestamp: now,
    disposition: 'interested'
  });
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true, entry: newEntry });
});

app.get('/api/admin/agents-list', (req, res) => {
  const agentMap = {};
  for (const [id, a] of Object.entries(appState.agents)) {
    const eidMatch = id.match(/^emp_(\d+)$/);
    if (eidMatch && !appState.allowedEids[eidMatch[1]]) continue; // agent removed
    // Skip client-role users from agent list
    if (eidMatch && appState.allowedEids[eidMatch[1]]) {
      const role = getEidRole(appState.allowedEids[eidMatch[1]]);
      if (role === 'client') continue;
    }
    agentMap[id] = { id, name: a.name };
  }
  for (const [eid, val] of Object.entries(appState.allowedEids)) {
    const role = getEidRole(val);
    if (role === 'client') continue; // Don't show client-role users in agent dropdown
    const virtualId = 'emp_' + eid;
    if (!agentMap[virtualId]) {
      agentMap[virtualId] = { id: virtualId, name: getEidName(val) };
    }
  }
  res.json(Object.values(agentMap));
});

// ─── Remove recruited/completed riders COMPLETELY from system ──────────────────
app.post('/api/agent/remove-interested', (req, res) => {
  const { agentId, numberId } = req.body;
  if (!agentId || !numberId) {
    return res.status(400).json({ error: 'agentId and numberId are required' });
  }
  const idx = appState.numbers.findIndex(n => n.id === numberId);
  if (idx === -1) return res.status(404).json({ error: 'Number not found' });
  const num = appState.numbers[idx];
  if (num.disposition !== 'interested') return res.status(400).json({ error: 'Rider is not marked as recruited' });
  if (num.interestedBy !== agentId) return res.status(403).json({ error: 'This rider is not assigned to you' });
  // Completely remove from system
  appState.numbers.splice(idx, 1);
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true });
});

app.post('/api/admin/remove-interested', (req, res) => {
  const { numberId } = req.body;
  if (!numberId) {
    return res.status(400).json({ error: 'numberId is required' });
  }
  const idx = appState.numbers.findIndex(n => n.id === numberId);
  if (idx === -1) return res.status(404).json({ error: 'Number not found' });
  // Completely remove from system (works for recruited AND delivery completed)
  appState.numbers.splice(idx, 1);
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true });
});

app.post('/api/admin/update-interested', (req, res) => {
  const { numberId, vehicleType, remarks, status, area, city } = req.body;
  if (!numberId) {
    return res.status(400).json({ error: 'numberId is required' });
  }
  const num = appState.numbers.find(n => n.id === numberId);
  if (!num) return res.status(404).json({ error: 'Number not found' });
  if (vehicleType !== undefined) {
    if (vehicleType && !VALID_RIDER_CATEGORIES.includes(vehicleType)) {
      return res.status(400).json({ error: 'Invalid vehicle type' });
    }
    num.vehicleType = vehicleType;
  }
  if (remarks !== undefined) num.remarks = remarks;
  if (status !== undefined) num.adminStatus = status;
  if (area !== undefined) num.area = area;
  if (city !== undefined) num.city = city;
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true });
});

app.post('/api/admin/update-lead-status', (req, res) => {
  const { numberId, adminStatus } = req.body;
  if (!numberId || !adminStatus) {
    return res.status(400).json({ error: 'numberId and adminStatus are required' });
  }
  const validStatuses = ['Completed', 'In Process', 'Rejected', 'Approved', 'On Hold'];
  if (!validStatuses.includes(adminStatus)) {
    return res.status(400).json({ error: 'Invalid status. Must be one of: ' + validStatuses.join(', ') });
  }
  const num = appState.numbers.find(n => n.id === numberId);
  if (!num) return res.status(404).json({ error: 'Number not found' });
  num.adminStatus = adminStatus;
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true });
});

app.get('/api/admin/completed', (req, res) => {
  evaluateRecruitedRiders(appState);
  const completed = appState.numbers.filter(n => n.disposition === 'interested' && n.deliveriesCompleted).map(n => {
    const agent = appState.agents[n.interestedBy];
    return {
      id: n.id, phone: n.phone, name: n.riderName || n.name || '',
      riderName: n.riderName || '',
      vehicleType: n.vehicleType || '',
      remarks: n.remarks || '',
      area: n.area || '',
      city: n.city || '',
      riderPhone: n.riderPhone || n.phone,
      interestedBy: agent ? agent.name : n.interestedBy,
      interestedByAgentId: n.interestedBy,
      recruitedAt: n.recruitedAt || n.interestedAt,
      deliveryCount: n.deliveryCount || 0,
      deliveryTarget: n.deliveryTarget || 31,
      completedAt: n.completedAt || null,
      adminStatus: n.adminStatus || '',
      hasDocZip: !!(n.docZipPath && fs.existsSync(n.docZipPath)),
      docZipName: n.docZipName || null
    };
  });
  res.json(completed);
});

app.get('/api/agent/completed/:agentId', (req, res) => {
  const agentId = req.params.agentId;
  evaluateRecruitedRiders(appState);
  const completed = appState.numbers.filter(n => n.disposition === 'interested' && n.deliveriesCompleted && n.interestedBy === agentId).map(n => ({
    id: n.id, phone: n.phone, name: n.riderName || n.name || '',
    riderName: n.riderName || '',
    vehicleType: n.vehicleType || '',
    remarks: n.remarks || '',
    area: n.area || '',
    city: n.city || '',
    riderPhone: n.riderPhone || n.phone,
    recruitedAt: n.recruitedAt || n.interestedAt,
    deliveryCount: n.deliveryCount || 0,
    deliveryTarget: n.deliveryTarget || 31,
    completedAt: n.completedAt || null,
    adminStatus: n.adminStatus || '',
    hasDocZip: !!(n.docZipPath && fs.existsSync(n.docZipPath)),
    docZipName: n.docZipName || null
  }));
  res.json(completed);
});

// ─── Agent Daily Report Data ──────────────────────────────────────────────────
app.get('/api/agent/dialed-today/:agentId', (req, res) => {
  const agentId = req.params.agentId;
  const today = getTodayStr();
  const todayLog = (appState.dialedLog || []).filter(entry => {
    if (entry.agentId !== agentId) return false;
    if (!entry.timestamp) return false;
    return entry.timestamp.slice(0, 10) === today;
  }).map(entry => {
    const t = new Date(entry.timestamp);
    const ist = new Date(t.getTime() + 5.5 * 60 * 60 * 1000);
    let h = ist.getHours(); const m = String(ist.getMinutes()).padStart(2, '0');
    const ampm = h >= 12 ? 'PM' : 'AM';
    if (h === 0) h = 12; else if (h > 12) h -= 12;
    const DISPO_MAP = { dead: 'CNC (Dead)', not_received: 'CNR', not_interested: 'Not Interested', followup: 'Followup', switch_off: 'Switch Off', interested: 'Recruited', discard: 'Discard', dnd: 'DND' };
    return {
      phone: entry.phone || '',
      disposition: DISPO_MAP[entry.disposition] || entry.disposition || 'Dialed',
      time: h + ':' + m + ' ' + ampm
    };
  });
  res.json(todayLog);
});

app.get('/api/agent/stats/:agentId', (req, res) => {
  const agentId = req.params.agentId;
  const agent = appState.agents[agentId];
  const today = getTodayStr();
  const todayLog = (appState.dialedLog || []).filter(entry => entry.agentId === agentId && entry.timestamp && entry.timestamp.slice(0, 10) === today);
  const counts = { totalDialedToday: agent ? (agent.totalDialedToday || 0) : 0 };
  const dispoCount = (d) => todayLog.filter(e => e.disposition === d).length;
  counts.interestedToday = dispoCount('interested');
  counts.followupToday = dispoCount('followup');
  counts.notInterestedToday = dispoCount('not_interested');
  counts.deadToday = dispoCount('dead');
  counts.notReceivedToday = dispoCount('not_received');
  counts.switchOffToday = dispoCount('switch_off');
  counts.dndToday = dispoCount('dnd');
  counts.discardToday = dispoCount('discard');
  res.json(counts);
});

app.delete('/api/admin/file/:fileId', (req, res) => {
  const fid = req.params.fileId;
  // SAFE DELETE: never remove recruited riders when deleting a file batch —
  // they are real business data (delivery tracking in progress or completed).
  // Only wipe undisposed / non-converting numbers from that batch.
  const fileInfo = appState.uploadedFiles.find(f => f.id === fid);
  if (fileInfo && fileInfo.sheetPath) {
    try { fs.unlinkSync(fileInfo.sheetPath); } catch {}
  }
  const protectedCount = appState.numbers.filter(
    n => n.file === fid && n.disposition === 'interested'
  ).length;
  appState.numbers = appState.numbers.filter(
    n => n.file !== fid || n.disposition === 'interested'
  );
  appState.uploadedFiles = appState.uploadedFiles.filter(f => f.id !== fid);
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true, protected_interested: protectedCount });
});

app.post('/api/admin/reset-today', (req, res) => {
  for (const id in appState.agents) {
    appState.agents[id].totalDialedToday = 0;
    appState.agents[id].active = false;
    appState.agents[id].currentIndex = null;
    appState.agents[id].onBreak = false;
    appState.agents[id].breakStartedAt = null;
    appState.agents[id].totalBreakMs = 0;
    appState.agents[id].currentNumberId = null;
    appState.agents[id].firstLoginToday = null;
    appState.agents[id].firstLoginDate  = null;
    appState.agents[id].onWashroom = false;
    appState.agents[id].washroomStartedAt = null;
    appState.agents[id].totalWashroomMs = 0;
    appState.agents[id].onMeeting = false;
    appState.agents[id].meetingStartedAt = null;
    appState.agents[id].totalMeetingMs = 0;
    appState.agents[id].onTlMode = false;
    appState.agents[id].tlModeStartedAt = null;
    appState.agents[id].totalTlModeMs = 0;
  }
  appState.lastReset = getTodayStr();
  saveState(appState);
  broadcastAdminStats();
  io.emit('force-stop');
  res.json({ success: true });
});

app.post('/api/admin/clear-all', (req, res) => {
  // SAFE CLEAR: keep all recruited riders (active + completed) — they are permanent business data.
  // Only wipe undisposed numbers, file registry, and dialed log.
  (appState.uploadedFiles || []).forEach(f => {
    if (f.sheetPath) { try { fs.unlinkSync(f.sheetPath); } catch {} }
  });
  appState.numbers = appState.numbers.filter(n => n.disposition === 'interested');
  appState.uploadedFiles = [];
  appState.dialedLog = [];
  for (const id in appState.agents) {
    appState.agents[id].totalDialedToday = 0;
    appState.agents[id].active = false;
    appState.agents[id].onBreak = false;
    appState.agents[id].breakStartedAt = null;
    appState.agents[id].totalBreakMs = 0;
    appState.agents[id].currentNumberId = null;
    appState.agents[id].onWashroom = false;
    appState.agents[id].washroomStartedAt = null;
    appState.agents[id].totalWashroomMs = 0;
    appState.agents[id].onMeeting = false;
    appState.agents[id].meetingStartedAt = null;
    appState.agents[id].totalMeetingMs = 0;
  }
  saveState(appState);
  broadcastAdminStats();
  io.emit('force-stop');
  res.json({ success: true });
});

app.post('/api/admin/hard-reset', (req, res) => {
  // ─── NUCLEAR HARD RESET ─────────────────────────────────────────────────────
  // Requires admin password. Wipes EVERYTHING: all numbers, all riders,
  // all agents, all EIDs, all reports, all uploaded files, all logs.
  // System becomes completely fresh — like a brand new install.
  const { password } = req.body;
  if (password !== PANEL_PASSWORD) {
    return res.status(403).json({ error: 'Incorrect admin password' });
  }

  // Delete ALL uploaded files from disk
  (appState.uploadedFiles || []).forEach(f => {
    if (f.sheetPath) { try { fs.unlinkSync(f.sheetPath); } catch {} }
  });
  // Delete ALL lead doc ZIPs
  try {
    const leadFiles = fs.readdirSync(LEAD_DOCS_DIR);
    leadFiles.forEach(f => { if (f === '.gitkeep') return; try { fs.unlinkSync(path.join(LEAD_DOCS_DIR, f)); } catch {} });
  } catch {}
  // Delete ALL agent photos
  try {
    const agentPhotos = fs.readdirSync(AGENT_PHOTOS_DIR);
    agentPhotos.forEach(f => { if (f === '.gitkeep') return; try { fs.unlinkSync(path.join(AGENT_PHOTOS_DIR, f)); } catch {} });
  } catch {}
  // Delete ALL reports from disk
  try {
    const reportFiles = fs.readdirSync(REPORTS_DIR);
    reportFiles.forEach(f => { if (f === '.gitkeep') return; try { fs.unlinkSync(path.join(REPORTS_DIR, f)); } catch {} });
  } catch {}
  // Delete ALL number sheets
  try {
    const sheetFiles = fs.readdirSync(NUMBER_SHEETS_DIR);
    sheetFiles.forEach(f => { if (f === '.gitkeep') return; try { fs.unlinkSync(path.join(NUMBER_SHEETS_DIR, f)); } catch {} });
  } catch {}
  // Delete scripts
  try {
    const scriptFiles = fs.readdirSync(SCRIPTS_DIR);
    scriptFiles.forEach(f => { if (f === '.gitkeep') return; try { fs.unlinkSync(path.join(SCRIPTS_DIR, f)); } catch {} });
  } catch {}

  // Create completely fresh state — NO preserved data at all
  appState = createFreshState({});
  appState.dndNumbers = [];
  appState.reports = [];
  saveState(appState);
  io.emit('force-stop');
  broadcastAdminStats();
  res.json({ success: true, message: 'NUCLEAR RESET COMPLETE — system is completely fresh' });
});

app.post('/api/agent/register', (req, res) => {
  let { name, employeeId } = req.body;
  if (!employeeId || !/^\d+$/.test(employeeId)) return res.status(400).json({ error: 'Valid numeric Employee ID required' });

  if (!appState.allowedEids[employeeId]) {
    return res.status(403).json({ error: 'Employee ID not recognised. Please contact your admin.' });
  }
  // Auto-fill name from allowedEids if not provided
  if (!name || !name.trim()) { name = getEidName(appState.allowedEids[employeeId]); }
  appState = checkDailyReset(appState);
  const agentId = 'emp_' + employeeId;
  const today   = getTodayStr();

  function getISTTimeStr() {
    const now = new Date();
    const ist = new Date(now.getTime() + (5.5 * 60 * 60 * 1000));
    return ist.toISOString().slice(11, 16);
  }

  if (!appState.agents[agentId]) {
    appState.agents[agentId] = {
      name, employeeId, active: false,
      totalDialedToday: 0, date: today,
      currentIndex: null, onBreak: false,
      breakStartedAt: null, totalBreakMs: 0,
      currentNumberId: null,
      firstLoginToday: getISTTimeStr(),
      firstLoginDate:  today
    };
  } else {
    appState.agents[agentId].name   = name;
    appState.agents[agentId].active = false;
    if (appState.agents[agentId].firstLoginDate !== today) {
      appState.agents[agentId].firstLoginToday = getISTTimeStr();
      appState.agents[agentId].firstLoginDate  = today;
    }
  }
  saveState(appState);
  broadcastAdminStats();

  const agent = appState.agents[agentId];
  let resumeNumber = null;
  if (agent.currentNumberId) {
    const num = appState.numbers.find(n => n.id === agent.currentNumberId);
    if (num && num.assignedTo === agentId && !num.dialedBy) {
      resumeNumber = { numberId: num.id, phone: num.phone, name: num.name || '' };
    }
  }

  const needsAutoResume = agent.needsAutoResume || false;
  if (agent.needsAutoResume) { delete agent.needsAutoResume; saveState(appState); }

  res.json({
    agentId, name, employeeId,
    role: getEidRole(appState.allowedEids[employeeId]),
    resumeNumber,
    needsAutoResume,
    totalDialedToday: agent.totalDialedToday || 0,
    onBreak: agent.onBreak || false,
    breakStartedAt: agent.breakStartedAt || null,
    totalBreakMs: agent.totalBreakMs || 0,
    breakAllowedMs: BREAK_DURATION_MS,
    onWashroom: agent.onWashroom || false,
    washroomStartedAt: agent.washroomStartedAt || null,
    totalWashroomMs: agent.totalWashroomMs || 0,
    onMeeting: agent.onMeeting || false,
    meetingStartedAt: agent.meetingStartedAt || null,
    totalMeetingMs: agent.totalMeetingMs || 0,
    onTlMode: agent.onTlMode || false,
    tlModeStartedAt: agent.tlModeStartedAt || null,
    totalTlModeMs: agent.totalTlModeMs || 0,
    lateLogin: (agent.firstLoginToday && agent.firstLoginToday > '10:00') || false
  });
});

// Break / washroom / meeting endpoints
app.post('/api/agent/break/start',    (req, res) => { const { agentId } = req.body; if (!agentId) return res.status(400).json({ error: 'agentId required' }); res.json(startBreak(agentId)); });
app.post('/api/agent/break/end',      (req, res) => { const { agentId } = req.body; if (!agentId) return res.status(400).json({ error: 'agentId required' }); res.json(endBreak(agentId)); });
app.post('/api/agent/washroom/start', (req, res) => { const { agentId } = req.body; if (!agentId) return res.status(400).json({ error: 'agentId required' }); res.json(startWashroom(agentId)); });
app.post('/api/agent/washroom/end',   (req, res) => { const { agentId } = req.body; if (!agentId) return res.status(400).json({ error: 'agentId required' }); res.json(endWashroom(agentId)); });
app.post('/api/agent/meeting/start',  (req, res) => { const { agentId } = req.body; if (!agentId) return res.status(400).json({ error: 'agentId required' }); res.json(startMeeting(agentId)); });
app.post('/api/agent/meeting/end',    (req, res) => { const { agentId } = req.body; if (!agentId) return res.status(400).json({ error: 'agentId required' }); res.json(endMeeting(agentId)); });

// ─── Cross-CRM Timer Sync Endpoints ──────────────────────────────────────────
// GET  /api/config/crm-urls  — returns peer CRM URLs from env vars (for frontend)
// Set env vars: PEER_CRM_1=https://dsa-crm.up.railway.app
//               PEER_CRM_2=https://kingfisher-crm.up.railway.app
app.get('/api/config/crm-urls', (req, res) => {
  const peers = [process.env.PEER_CRM_1, process.env.PEER_CRM_2]
    .filter(Boolean)
    .map(u => u.replace(/\/$/, ''));
  res.json({ peers });
});

// GET  /api/sync/timer-state/:empId  — read current timer state for an employee
app.get('/api/sync/timer-state/:empId', (req, res) => {
  const agentId = 'emp_' + req.params.empId;
  const agent = appState.agents[agentId];
  if (!agent) return res.json({ found: false });
  res.json({
    found: true,
    onBreak:          agent.onBreak         || false,
    breakStartedAt:   agent.breakStartedAt  || null,
    totalBreakMs:     agent.totalBreakMs    || 0,
    breakAllowedMs:   BREAK_DURATION_MS,
    onWashroom:        agent.onWashroom       || false,
    washroomStartedAt: agent.washroomStartedAt || null,
    totalWashroomMs:   agent.totalWashroomMs   || 0,
    onMeeting:         agent.onMeeting         || false,
    meetingStartedAt:  agent.meetingStartedAt  || null,
    totalMeetingMs:    agent.totalMeetingMs    || 0
  });
});

// POST /api/sync/timer-action  — trigger a timer start/end by empId (called by other CRMs)
app.post('/api/sync/timer-action', (req, res) => {
  const { empId, type, action } = req.body;
  if (!empId || !type || !action) return res.status(400).json({ error: 'empId, type, action required' });
  const agentId = 'emp_' + empId;
  if (!appState.agents[agentId]) return res.json({ found: false });
  let result;
  if      (type === 'break'    && action === 'start') result = startBreak(agentId);
  else if (type === 'break'    && action === 'end')   result = endBreak(agentId);
  else if (type === 'washroom' && action === 'start') result = startWashroom(agentId);
  else if (type === 'washroom' && action === 'end')   result = endWashroom(agentId);
  else if (type === 'meeting'  && action === 'start') result = startMeeting(agentId);
  else if (type === 'meeting'  && action === 'end')   result = endMeeting(agentId);
  else return res.status(400).json({ error: 'Invalid type/action' });
  // Broadcast to all sockets on this server so the agent tab here also updates
  io.emit('timer-update', { agentId, type, action, ...result });
  broadcastAdminStats();
  res.json({ found: true, ...result });
});

app.get('/api/agent/state/:agentId', (req, res) => {
  const agent = appState.agents[req.params.agentId];
  if (!agent) return res.status(404).json({ error: 'Agent not found' });
  let resumeNumber = null;
  if (agent.currentNumberId) {
    const num = appState.numbers.find(n => n.id === agent.currentNumberId);
    if (num && num.assignedTo === req.params.agentId && !num.dialedBy) {
      resumeNumber = { numberId: num.id, phone: num.phone, name: num.name || '' };
    }
  }
  const needsAutoResume = agent.needsAutoResume || false;
  if (agent.needsAutoResume) { delete agent.needsAutoResume; saveState(appState); }
  res.json({
    resumeNumber, needsAutoResume,
    totalDialedToday: agent.totalDialedToday || 0,
    onBreak: agent.onBreak || false,
    breakStartedAt: agent.breakStartedAt || null,
    totalBreakMs: agent.totalBreakMs || 0,
    breakAllowedMs: BREAK_DURATION_MS,
    onWashroom: agent.onWashroom || false,
    washroomStartedAt: agent.washroomStartedAt || null,
    totalWashroomMs: agent.totalWashroomMs || 0,
    onMeeting: agent.onMeeting || false,
    meetingStartedAt: agent.meetingStartedAt || null,
    totalMeetingMs: agent.totalMeetingMs || 0,
    onTlMode: agent.onTlMode || false,
    tlModeStartedAt: agent.tlModeStartedAt || null,
    totalTlModeMs: agent.totalTlModeMs || 0
  });
});

// ─── Socket.IO ─────────────────────────────────────────────────────────────────
io.on('connection', (socket) => {
  let socketAgentId = null;
  let socketCurrentNumber = null;

  socket.on('join-admin', () => {
    socket.join('admin-room');
    socket.emit('stats-update', getAdminStats());
  });

  socket.on('disconnect', () => {
    if (socketAgentId) {
      const agent = appState.agents[socketAgentId];
      if (agent) { agent.active = false; saveState(appState); }
      broadcastAdminStats();
    }
  });

  socket.on('agent-start', ({ agentId }) => {
    socketAgentId = agentId;
    socket.join('agent-room'); // join agent-room so auto-report broadcasts reach clients
    appState = checkDailyReset(appState);
    const agent = appState.agents[agentId];
    if (!agent) return socket.emit('error', 'Agent not found');
    agent.active = true;
    saveState(appState);
    broadcastAdminStats();

    if (agent.currentNumberId) {
      const num = appState.numbers.find(n => n.id === agent.currentNumberId);
      if (num && num.assignedTo === agentId && !num.dialedBy) {
        socketCurrentNumber = num.id;
        return socket.emit('show-number', {
          numberId: num.id, phone: num.phone, name: num.name || '',
          totalDialedToday: agent.totalDialedToday || 0,
          resumed: true
        });
      }
    }

    const num = getNextNumber(agentId);
    if (!num) {
      socket.emit('no-numbers');
    } else {
      socketCurrentNumber = num.id;
      agent.currentNumberId = num.id;
      saveState(appState);
      socket.emit('show-number', {
        numberId: num.id, phone: num.phone, name: num.name || '',
        totalDialedToday: agent.totalDialedToday || 0
      });
    }
  });

  socket.on('agent-next', ({ agentId, prevNumberId }) => {
    appState = checkDailyReset(appState);
    const agent = appState.agents[agentId];
    if (!agent) return socket.emit('error', 'Agent not found');

    if (prevNumberId) markDialed(agentId, prevNumberId);

    const num = getNextNumber(agentId);
    if (!num) {
      socketCurrentNumber = null;
      if (agent) agent.currentNumberId = null;
      saveState(appState);
      socket.emit('no-numbers', { totalDialedToday: agent.totalDialedToday || 0 });
    } else {
      socketCurrentNumber = num.id;
      agent.currentNumberId = num.id;
      saveState(appState);
      socket.emit('show-number', {
        numberId: num.id, phone: num.phone, name: num.name || '',
        totalDialedToday: agent.totalDialedToday || 0
      });
    }
    broadcastAdminStats();
  });

  socket.on('agent-stop', ({ agentId, currentNumberId }) => {
    const agent = appState.agents[agentId];
    if (agent) {
      agent.active = false;
      agent.currentNumberId = null;
    }
    if (currentNumberId) releaseNumber(agentId, currentNumberId);
    saveState(appState);
    broadcastAdminStats();
  });

  socket.on('agent-disposition', ({ agentId, numberId, disposition, followupDate, followupTime, followupName, riderName, vehicleType, remarks, area, city }) => {
    appState = checkDailyReset(appState);
    const agent = appState.agents[agentId];
    if (!agent) return socket.emit('error', 'Agent not found');
    if (!VALID_DISPOSITIONS.includes(disposition)) return socket.emit('error', 'Invalid disposition');

    applyDisposition(agentId, numberId, disposition, { followupDate, followupTime, followupName, riderName, vehicleType, remarks, area, city });

    const num = getNextNumber(agentId);
    if (!num) {
      socketCurrentNumber = null;
      if (agent) agent.currentNumberId = null;
      saveState(appState);
      socket.emit('no-numbers', { totalDialedToday: agent.totalDialedToday || 0 });
    } else {
      socketCurrentNumber = num.id;
      agent.currentNumberId = num.id;
      saveState(appState);
      socket.emit('show-number', {
        numberId: num.id, phone: num.phone, name: num.name || '',
        totalDialedToday: agent.totalDialedToday || 0
      });
    }
    broadcastAdminStats();
  });

  socket.on('agent-break-start',    ({ agentId }) => { const r = startBreak(agentId);    socket.emit('break-started', r);    io.emit('timer-update', { agentId, type: 'break', action: 'start', ...r }); broadcastAdminStats(); });
  socket.on('agent-break-end',      ({ agentId }) => { const r = endBreak(agentId);      socket.emit('break-ended', r);      io.emit('timer-update', { agentId, type: 'break', action: 'end', ...r }); broadcastAdminStats(); });
  socket.on('agent-washroom-start', ({ agentId }) => { const r = startWashroom(agentId); socket.emit('washroom-started', r); io.emit('timer-update', { agentId, type: 'washroom', action: 'start', ...r }); broadcastAdminStats(); });
  socket.on('agent-washroom-end',   ({ agentId }) => { const r = endWashroom(agentId);   socket.emit('washroom-ended', r);   io.emit('timer-update', { agentId, type: 'washroom', action: 'end', ...r }); broadcastAdminStats(); });
  socket.on('agent-meeting-start',  ({ agentId }) => { const r = startMeeting(agentId);  socket.emit('meeting-started', r);  io.emit('timer-update', { agentId, type: 'meeting', action: 'start', ...r }); broadcastAdminStats(); });
  socket.on('agent-meeting-end',    ({ agentId }) => { const r = endMeeting(agentId);    socket.emit('meeting-ended', r);    io.emit('timer-update', { agentId, type: 'meeting', action: 'end', ...r }); broadcastAdminStats(); });
  socket.on('agent-tlmode-start',   ({ agentId }) => { const r = startTlMode(agentId);   socket.emit('tlmode-started', r);   io.emit('timer-update', { agentId, type: 'tlmode', action: 'start', ...r }); broadcastAdminStats(); });
  socket.on('agent-tlmode-end',     ({ agentId }) => { const r = endTlMode(agentId);     socket.emit('tlmode-ended', r);     io.emit('timer-update', { agentId, type: 'tlmode', action: 'end', ...r }); broadcastAdminStats(); });

  socket.on('ping-alive', ({ agentId }) => {
    const agent = appState.agents[agentId];
    if (agent) appState = checkDailyReset(appState);
  });
});

// ─── Disposition Stats (shared logic — used by the HTTP route AND the automatic
// daily PDF report generator below, so both stay perfectly in sync) ───────────
function computeDispositionStats(period, agentId) {
  const now = new Date();
  const istNow = new Date(now.getTime() + (5.5 * 60 * 60 * 1000));
  const istTodayStr = istNow.toISOString().slice(0, 10);

  let daysBack = 0;
  switch (period) {
    case 'daily': daysBack = 0; break;
    case 'weekly': daysBack = 7; break;
    case 'monthly': daysBack = 30; break;
    case 'yearly': daysBack = 365; break;
    default: daysBack = 0;
  }

  let cutoffDate;
  if (daysBack === 0) {
    cutoffDate = new Date(istTodayStr + 'T00:00:00.000+05:30');
  } else {
    const cutoffIST = new Date(istNow);
    cutoffIST.setDate(cutoffIST.getDate() - daysBack);
    const cutoffStr = cutoffIST.toISOString().slice(0, 10);
    cutoffDate = new Date(cutoffStr + 'T00:00:00.000+05:30');
  }

  const filteredLogs = appState.dialedLog.filter(entry => {
    if (!entry.timestamp) return false;
    const entryDate = new Date(entry.timestamp);
    if (entryDate < cutoffDate) return false;
    if (agentId && entry.agentId !== agentId) return false;
    return true;
  });

  const stats = {
    period,
    totalCalls: filteredLogs.length,
    dead: 0, not_received: 0, not_interested: 0, followup: 0, switch_off: 0, interested: 0, discard: 0, dnd: 0
  };

  filteredLogs.forEach(entry => {
    const d = entry.disposition;
    if (d && stats.hasOwnProperty(d)) stats[d]++;
  });

  return stats;
}

app.get('/api/stats/dispositions', (req, res) => {
  const period = req.query.period || 'daily';
  const agentId = req.query.agentId || null;
  res.json(computeDispositionStats(period, agentId));
});

// ─── Admin EID Management ──────────────────────────────────────────────────────
// Helper: get agent name from allowedEids (supports both string and object format)
function getEidName(eidVal) {
  if (!eidVal) return '';
  if (typeof eidVal === 'string') return eidVal;
  if (typeof eidVal === 'object' && eidVal.name) return eidVal.name;
  return '';
}
function getEidPhoto(eidVal) {
  if (!eidVal) return null;
  if (typeof eidVal === 'object' && eidVal.photo) return eidVal.photo;
  return null;
}
function getEidRole(eidVal) {
  if (!eidVal) return 'agent';
  if (typeof eidVal === 'object' && eidVal.role) return eidVal.role;
  return 'agent';
}

// Helper: resolve who added a DND number to a display name + role (agent / client / admin)
// so admin.html and client.html can show a note like "Added by Rohan (Client)".
function resolveDndAddedBy(rawId) {
  if (!rawId || rawId === 'admin') {
    return { id: 'admin', name: 'Admin', role: 'admin' };
  }
  const agent = appState.agents && appState.agents[rawId];
  if (agent) {
    const eid = agent.employeeId;
    const role = (eid && appState.allowedEids[eid]) ? getEidRole(appState.allowedEids[eid]) : 'agent';
    return { id: rawId, name: agent.name || rawId, role: role };
  }
  return { id: rawId, name: rawId, role: 'agent' };
}

app.get('/api/admin/eids', (req, res) => {
  const list = Object.entries(appState.allowedEids).map(([eid, val]) => ({
    eid,
    name: getEidName(val),
    photo: getEidPhoto(val),
    role: (typeof val === 'object' && val.role) ? val.role : 'agent'
  }));
  res.json({ eids: list });
});

const VALID_ROLES = ['agent', 'tl', 'client', 'admin'];

app.post('/api/admin/eids', (req, res) => {
  const { eid, name, role } = req.body;
  if (!eid || !/^\d+$/.test(eid)) return res.status(400).json({ error: 'Valid numeric EID required' });
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name required' });
  const existing = appState.allowedEids[eid];
  const existingPhoto = getEidPhoto(existing);
  const existingRole = getEidRole(existing); // fallback: preserve existing role if none supplied
  let finalRole = existingRole;
  if (role !== undefined && role !== null && String(role).trim() !== '') {
    if (!VALID_ROLES.includes(role)) return res.status(400).json({ error: 'Invalid role. Must be one of: ' + VALID_ROLES.join(', ') });
    finalRole = role;
  }
  appState.allowedEids[eid] = { name: name.trim(), photo: existingPhoto || null, role: finalRole };
  saveState(appState);
  res.json({ success: true, eid, name: name.trim(), role: finalRole });
});

// Assign / change a user's role
app.put('/api/admin/eids/:eid/role', (req, res) => {
  const eid = req.params.eid;
  const { role } = req.body;
  if (!appState.allowedEids[eid]) return res.status(404).json({ error: 'EID not found' });
  if (!role || !VALID_ROLES.includes(role)) return res.status(400).json({ error: 'Invalid role. Must be one of: ' + VALID_ROLES.join(', ') });
  const existing = appState.allowedEids[eid];
  const name = getEidName(existing);
  const photo = getEidPhoto(existing);
  appState.allowedEids[eid] = { name, photo: photo || null, role };
  saveState(appState);
  res.json({ success: true, eid, name, role });
});

app.delete('/api/admin/eids/:eid', (req, res) => {
  const eid = req.params.eid;
  if (!appState.allowedEids[eid]) return res.status(404).json({ error: 'EID not found' });
  delete appState.allowedEids[eid];
  saveState(appState);
  res.json({ success: true });
});

// ─── Client Panel Auth — check if EID has client/admin role ──────────────────
function handleClientAuth(req, res) {
  let { employeeId, name } = req.body;
  if (!employeeId) return res.status(400).json({ error: 'employeeId required' });
  const eidData = appState.allowedEids[employeeId];
  if (!eidData) return res.status(403).json({ error: 'Employee ID not recognised. Please contact your admin.' });
  if (!name || !name.trim()) { name = getEidName(eidData); }
  const role = getEidRole(eidData);
  if (role !== 'client' && role !== 'admin' && role !== 'tl') {
    return res.status(403).json({ error: 'You do not have Client Panel access. Contact your admin.' });
  }
  const agentId = 'emp_' + employeeId;
  // Register/update agent if not exists
  appState = checkDailyReset(appState);
  const today = getTodayStr();
  function getISTTimeStr() {
    const now = new Date();
    const ist = new Date(now.getTime() + (5.5 * 60 * 60 * 1000));
    return ist.toISOString().slice(11, 16);
  }
  if (!appState.agents[agentId]) {
    appState.agents[agentId] = {
      name, employeeId, active: false,
      totalDialedToday: 0, date: today,
      currentIndex: null, onBreak: false,
      breakStartedAt: null, totalBreakMs: 0,
      currentNumberId: null,
      firstLoginToday: getISTTimeStr(),
      firstLoginDate: today
    };
  } else {
    appState.agents[agentId].name = name;
    if (appState.agents[agentId].firstLoginDate !== today) {
      appState.agents[agentId].firstLoginToday = getISTTimeStr();
      appState.agents[agentId].firstLoginDate = today;
    }
  }
  saveState(appState);
  broadcastAdminStats();
  const agent = appState.agents[agentId];
  const lateLogin = (agent.firstLoginToday && agent.firstLoginToday > '10:00') || false;
  res.json({ success: true, agentId, name, employeeId, role, lateLogin });
}
app.post('/api/client/auth', handleClientAuth);
// Backward-compatible alias — anything still pointing at the old TL panel endpoint.
app.post('/api/tl/auth', handleClientAuth);

// ─── Agent Photo Upload ─────────────────────────────────────────────────────────
app.post('/api/admin/agent-photo/:eid', agentPhotoUpload.single('photo'), (req, res) => {
  try {
    const eid = req.params.eid;
    if (!req.file) return res.status(400).json({ error: 'No image file uploaded' });
    const photoPath = '/api/admin/agent-photo/' + eid + '?t=' + Date.now();
    const existing = appState.allowedEids[eid];
    if (existing) {
      const name = getEidName(existing);
      const role = getEidRole(existing); // PRESERVE existing TL role
      appState.allowedEids[eid] = { name, photo: req.file.path, role };
    } else {
      appState.allowedEids[eid] = { name: '', photo: req.file.path, role: 'agent' };
    }
    saveState(appState);
    res.json({ success: true, photoUrl: photoPath });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/admin/agent-photo/:eid', (req, res) => {
  const eid = req.params.eid;
  const val = appState.allowedEids[eid];
  const photoPath = getEidPhoto(val);
  if (!photoPath || !fs.existsSync(photoPath)) {
    return res.status(404).json({ error: 'No photo found' });
  }
  res.sendFile(path.resolve(photoPath));
});

// ─── Rankings/Leaderboard API ─────────────────────────────────────────────────
app.get('/api/rankings', (req, res) => {
  const period = req.query.period || 'daily';

  const now = new Date();
  const istNow = new Date(now.getTime() + (5.5 * 60 * 60 * 1000));
  const istTodayStr = istNow.toISOString().slice(0, 10);

  let daysBack = 0;
  switch (period) {
    case 'daily': daysBack = 0; break;
    case 'weekly': daysBack = 7; break;
    case 'monthly': daysBack = 30; break;
    default: daysBack = 0;
  }

  let cutoffDate;
  if (daysBack === 0) {
    cutoffDate = new Date(istTodayStr + 'T00:00:00.000+05:30');
  } else {
    const cutoffIST = new Date(istNow);
    cutoffIST.setDate(cutoffIST.getDate() - daysBack);
    const cutoffStr = cutoffIST.toISOString().slice(0, 10);
    cutoffDate = new Date(cutoffStr + 'T00:00:00.000+05:30');
  }

  // Collect all agent IDs from dialedLog and agents registry
  const agentScores = {};

  // Initialize from known agents — skip anyone whose EID has since been removed
  for (const [id, a] of Object.entries(appState.agents)) {
    const eidMatch = id.match(/^emp_(\d+)$/);
    if (eidMatch && !appState.allowedEids[eidMatch[1]]) continue;
    agentScores[id] = { agentId: id, name: a.name, interested: 0, followups: 0, totalCalls: 0, notInterested: 0, discard: 0, dead: 0, switchOff: 0 };
  }

  // Also ensure agents from allowedEids appear
  for (const [eid, val] of Object.entries(appState.allowedEids)) {
    const agId = 'emp_' + eid;
    if (!agentScores[agId]) {
      agentScores[agId] = { agentId: agId, name: getEidName(val), interested: 0, followups: 0, totalCalls: 0, notInterested: 0, discard: 0, dead: 0, switchOff: 0 };
    }
  }

  // Filter dialedLog by period and tally — skip entries for agents who have since been removed
  appState.dialedLog.forEach(entry => {
    if (!entry.timestamp) return;
    const entryDate = new Date(entry.timestamp);
    if (entryDate < cutoffDate) return;

    const aid = entry.agentId;
    const eidMatch = aid && aid.match(/^emp_(\d+)$/);
    // If this agentId maps to an EID that's no longer in allowedEids, the agent was removed — skip
    if (eidMatch && !appState.allowedEids[eidMatch[1]]) return;
    if (!agentScores[aid]) {
      agentScores[aid] = { agentId: aid, name: entry.agentName || aid, interested: 0, followups: 0, totalCalls: 0, notInterested: 0, discard: 0, dead: 0, switchOff: 0 };
    }

    agentScores[aid].totalCalls++;
    if (entry.disposition === 'interested') agentScores[aid].interested++;
    else if (entry.disposition === 'followup') agentScores[aid].followups++;
    else if (entry.disposition === 'not_interested') agentScores[aid].notInterested++;
    else if (entry.disposition === 'discard') agentScores[aid].discard++;
    else if (entry.disposition === 'dead') agentScores[aid].dead++;
    else if (entry.disposition === 'switch_off') agentScores[aid].switchOff++;
  });

  // Score formula: MAX(0, MIN(100, (((100*Recruited) + (25*FollowUp) - (10*NotInterested) - (15*Discard) - (2*CNC) - (2*SwitchOff)) / (TotalCalls*100)) * 100))
  const rankings = Object.values(agentScores).map(a => {
    let score = 0;
    if (a.totalCalls > 0) {
      const rawScore = ((100 * a.interested) + (25 * a.followups) - (10 * a.notInterested) - (15 * a.discard) - (2 * a.dead) - (2 * a.switchOff)) / (a.totalCalls * 100) * 100;
      score = Math.max(0, Math.min(100, rawScore));
    }
    score = Math.round(score * 100) / 100;
    // Get profile photo
    const eidMatch = a.agentId.match(/^emp_(\d+)$/);
    let profilePhoto = null;
    if (eidMatch) {
      const eidVal = appState.allowedEids[eidMatch[1]];
      const photoPath = getEidPhoto(eidVal);
      if (photoPath && fs.existsSync(photoPath)) {
        profilePhoto = '/api/admin/agent-photo/' + eidMatch[1];
      }
    }
    return { ...a, score, profilePhoto };
  });

  rankings.sort((a, b) => b.score - a.score || b.interested - a.interested);

  // Add rank and remarks
  rankings.forEach((r, i) => {
    r.rank = i + 1;
    if (i === 0 && r.score > 0) {
      r.remarks = `Top performer! Score: ${r.score}/100 with ${r.interested} recruited riders and ${r.followups} followups`;
    } else if (r.score === 0 && r.totalCalls === 0) {
      r.remarks = 'No calls made yet in this period';
    } else if (r.score < 20) {
      r.remarks = 'Focus on quality calls - increase recruited and followup conversions';
    } else {
      r.remarks = `Score: ${r.score}/100. ${r.interested} recruited, ${r.followups} followups. Keep improving!`;
    }
  });

  const formulaDescription = 'Score (0-100) = ((100 x Recruited) + (25 x FollowUp) - (10 x NotInterested) - (15 x Not-Eligible) - (2 x CNC) - (2 x SwitchOff)) / (TotalCalls x 100) x 100. Higher recruited and followup calls improve your score. Not Interested, Not-Eligible, CNC and SwitchOff reduce it.';

  res.json({ rankings, formulaDescription });
});

// ─── Followup Management Endpoints ─────────────────────────────────────────────
// PUT /api/admin/followup/:numberId - Edit followup date, time, and name
app.put('/api/admin/followup/:numberId', (req, res) => {
  const { numberId } = req.params;
  const { followupDate, followupTime, followupName } = req.body;
  const num = appState.numbers.find(n => n.id === numberId);
  if (!num) return res.status(404).json({ error: 'Number not found' });
  if (num.disposition !== 'followup') return res.status(400).json({ error: 'Number is not a followup' });
  if (followupDate !== undefined) num.followupDate = followupDate;
  if (followupTime !== undefined) num.followupTime = followupTime;
  if (followupName !== undefined) num.followupName = followupName;
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true, id: num.id, followupDate: num.followupDate, followupTime: num.followupTime, followupName: num.followupName || '' });
});

// DELETE /api/admin/followup/:numberId - Remove followup (lead goes to NI)
app.delete('/api/admin/followup/:numberId', (req, res) => {
  const { numberId } = req.params;
  const num = appState.numbers.find(n => n.id === numberId);
  if (!num) return res.status(404).json({ error: 'Number not found' });
  if (num.disposition !== 'followup') return res.status(400).json({ error: 'Number is not a followup' });
  num.disposition = 'not_interested';
  num.permanent = true;
  num.blockedUntil = null;
  num.followupDate = null;
  num.followupTime = null;
  num.followupLockedBy = null;
  num.followupName = null;
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true });
});

// DELETE /api/agent/followup/:numberId - Agent removes followup (lead goes to NI)
app.delete('/api/agent/followup/:numberId', (req, res) => {
  const { numberId } = req.params;
  const { agentId } = req.body || {};
  const num = appState.numbers.find(n => n.id === numberId);
  if (!num) return res.status(404).json({ error: 'Number not found' });
  if (num.disposition !== 'followup') return res.status(400).json({ error: 'Number is not a followup' });
  if (agentId && num.followupLockedBy && num.followupLockedBy !== agentId) {
    return res.status(403).json({ error: 'Not authorized to remove this followup' });
  }
  num.disposition = 'not_interested';
  num.permanent = true;
  num.blockedUntil = null;
  num.followupDate = null;
  num.followupTime = null;
  num.followupLockedBy = null;
  num.followupName = null;
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true });
});

// GET /api/agent/due-followups/:agentId - Followups whose date+time has arrived (popup trigger)
app.get('/api/agent/due-followups/:agentId', (req, res) => {
  const agentId = req.params.agentId;
  const now = new Date();
  const ist = new Date(now.getTime() + (5.5 * 60 * 60 * 1000));
  const todayStr = ist.toISOString().slice(0, 10);
  const currentTime = ist.toISOString().slice(11, 16); // HH:MM

  const dueFollowups = appState.numbers.filter(n => {
    if (n.disposition !== 'followup') return false;
    if (n.followupLockedBy !== agentId) return false;
    if (!n.followupDate) return false;
    // Due if date is today and time has passed, or date is in the past
    if (n.followupDate < todayStr) return true;
    if (n.followupDate === todayStr) {
      const fTime = n.followupTime || '00:00';
      if (fTime <= currentTime) return true;
    }
    return false;
  }).map(n => ({
    id: n.id, phone: n.phone, name: n.name || '',
    followupDate: n.followupDate,
    followupTime: n.followupTime,
    followupName: n.followupName || '',
    followupCount: n.followupCount || 0
  }));

  // Sort by nearest first
  dueFollowups.sort((a, b) => {
    const dateA = (a.followupDate || '9999-12-31') + ' ' + (a.followupTime || '23:59');
    const dateB = (b.followupDate || '9999-12-31') + ' ' + (b.followupTime || '23:59');
    return dateA.localeCompare(dateB);
  });

  res.json(dueFollowups);
});

// POST /api/admin/upload-followups - Upload custom Excel with followups
const followupUpload = multer({ dest: UPLOADS_DIR });
app.post('/api/admin/upload-followups', followupUpload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    const wb = XLSX.readFile(req.file.path);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    let added = 0;
    let skipped = 0;
    const existingPhones = new Set(appState.numbers.map(n => n.phone));

    rows.forEach((row, i) => {
      if (i === 0) return; // Skip header row
      const phone = String(row[0] || '').trim().replace(/\s+/g, '');
      if (!phone || phone.length < 7) return;
      const name = row[1] ? String(row[1]).trim() : '';
      const followupDate = row[2] ? String(row[2]).trim() : null;
      const followupTime = row[3] ? String(row[3]).trim() : null;
      const agentId = row[4] ? String(row[4]).trim() : null;
      const followupName = row[5] ? String(row[5]).trim() : '';

      if (existingPhones.has(phone)) {
        // If number exists, update its followup if not already permanently removed
        const existing = appState.numbers.find(n => n.phone === phone);
        if (existing && existing.disposition !== 'discard' && existing.disposition !== 'interested' && !existing.permanent) {
          if (!existing.followupCount) existing.followupCount = 0;
          if (existing.followupCount >= 2) {
            // Enforce 2-max cap: auto-convert to NI
            existing.disposition = 'not_interested';
            existing.permanent = true;
            existing.blockedUntil = null;
            existing.followupDate = null;
            existing.followupTime = null;
            existing.followupLockedBy = null;
            existing.followupName = null;
            skipped++;
          } else {
            existing.disposition = 'followup';
            existing.followupDate = followupDate;
            existing.followupTime = followupTime;
            existing.followupLockedBy = agentId || existing.followupLockedBy;
            existing.followupName = followupName || existing.followupName || '';
            existing.followupCount++;
            added++;
          }
        } else {
          skipped++;
        }
        return;
      }

      existingPhones.add(phone);
      const newEntry = {
        id: uuidv4(),
        phone,
        name,
        file: null,
        assignedTo: null,
        dialedBy: null,
        dialedAt: null,
        disposition: 'followup',
        followupDate,
        followupTime,
        followupLockedBy: agentId || null,
        followupName: followupName,
        followupCount: 1
      };
      appState.numbers.push(newEntry);
      added++;
    });

    saveState(appState);
    fs.unlinkSync(req.file.path);
    broadcastAdminStats();
    res.json({ success: true, added, skipped });
  } catch (e) {
    if (req.file) { try { fs.unlinkSync(req.file.path); } catch {} }
    res.status(500).json({ error: e.message });
  }
});

// GET /api/admin/download-followup-sample - Download sample Excel for followup upload
app.get('/api/admin/download-followup-sample', (req, res) => {
  const sampleData = [
    ['Phone', 'Name', 'FollowupDate (YYYY-MM-DD)', 'FollowupTime (HH:MM)', 'AgentId (emp_XXX)', 'FollowupName'],
    ['9876543210', 'John Doe', '2025-01-20', '10:30', 'emp_101', 'Rider recruitment call'],
    ['9876543211', 'Jane Smith', '2025-01-21', '14:00', 'emp_102', 'Vehicle verification']
  ];
  const ws = XLSX.utils.aoa_to_sheet(sampleData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Followups');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', 'attachment; filename=followup-sample.xlsx');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// ─── Admin Timer Control (fix broadcast to all) ───────────────────────────────
app.post('/api/admin/agent/break/end', (req, res) => {
  const { agentId } = req.body;
  if (!agentId) return res.status(400).json({ error: 'agentId required' });
  const result = endBreak(agentId);
  // Broadcast to ALL connected clients so agent sees the update
  io.emit('timer-update', { agentId, type: 'break', action: 'end', ...result });
  // Force a full page reload on the agent's browser so dialing resumes immediately
  io.emit('force-page-reload', { agentId, reason: 'break-removed-by-admin' });
  broadcastAdminStats();
  res.json(result);
});

app.post('/api/admin/agent/washroom/end', (req, res) => {
  const { agentId } = req.body;
  if (!agentId) return res.status(400).json({ error: 'agentId required' });
  const result = endWashroom(agentId);
  // Broadcast to ALL connected clients so agent sees the update
  io.emit('timer-update', { agentId, type: 'washroom', action: 'end', ...result });
  broadcastAdminStats();
  res.json(result);
});

app.post('/api/admin/agent/meeting/end', (req, res) => {
  const { agentId } = req.body;
  if (!agentId) return res.status(400).json({ error: 'agentId required' });
  const result = endMeeting(agentId);
  // Broadcast to ALL connected clients so agent sees the update
  io.emit('timer-update', { agentId, type: 'meeting', action: 'end', ...result });
  broadcastAdminStats();
  res.json(result);
});

// ─── DND (Do Not Disturb) Management ──────────────────────────────────────────
app.get('/api/admin/dnd', (req, res) => {
  const list = (appState.dndNumbers || []).map(d => {
    const info = resolveDndAddedBy(d.addedBy);
    return Object.assign({}, d, { addedByName: info.name, addedByRole: info.role });
  });
  res.json({ dndNumbers: list });
});

app.post('/api/admin/dnd', (req, res) => {
  const { phone, addedBy } = req.body;
  if (!phone || !/^\d{7,15}$/.test(phone.replace(/\s+/g,''))) {
    return res.status(400).json({ error: 'Valid phone number required' });
  }
  const cleanPhone = phone.replace(/\s+/g,'');
  if (!appState.dndNumbers) appState.dndNumbers = [];
  if (appState.dndNumbers.find(d => d.phone === cleanPhone)) {
    return res.status(409).json({ error: 'Number already in DND list' });
  }
  // Respect who actually submitted this (agent / TL / admin); default to 'admin' only when nobody is identified.
  const info = resolveDndAddedBy(addedBy);
  appState.dndNumbers.push({ phone: cleanPhone, addedAt: new Date().toISOString(), addedBy: info.id });
  // Also mark any existing number with this phone as dnd
  const existing = appState.numbers.find(n => n.phone === cleanPhone);
  if (existing && existing.disposition !== 'interested') {
    existing.disposition = 'dnd';
    existing.permanent = true;
    existing.assignedTo = null;
  }
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true, phone: cleanPhone, addedByName: info.name, addedByRole: info.role });
});

app.delete('/api/admin/dnd/:phone', (req, res) => {
  const phone = req.params.phone;
  if (!appState.dndNumbers) appState.dndNumbers = [];
  const idx = appState.dndNumbers.findIndex(d => d.phone === phone);
  if (idx === -1) return res.status(404).json({ error: 'Number not in DND list' });
  appState.dndNumbers.splice(idx, 1);
  saveState(appState);
  res.json({ success: true });
});

// ─── Script Upload & Management ──────────────────────────────────────────────
const scriptUpload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, SCRIPTS_DIR),
    filename: (req, file, cb) => cb(null, 'call_script.txt')
  }),
  limits: { fileSize: 2 * 1024 * 1024 }, // 2MB max
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'text/plain' || file.originalname.toLowerCase().endsWith('.txt')) {
      cb(null, true);
    } else {
      cb(new Error('Only TXT files are allowed'));
    }
  }
});

app.post('/api/admin/upload-script', scriptUpload.single('script'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No TXT file uploaded' });
    res.json({ success: true, filename: req.file.originalname });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/script', (req, res) => {
  const scriptPath = path.join(SCRIPTS_DIR, 'call_script.txt');
  if (!fs.existsSync(scriptPath)) {
    return res.json({ script: null });
  }
  try {
    const content = fs.readFileSync(scriptPath, 'utf8');
    res.json({ script: content });
  } catch (e) {
    res.json({ script: null });
  }
});

// ─── Disposition Stats Copy ───────────────────────────────────────────────────
app.get('/api/stats/daily-numbers', (req, res) => {
  // Returns numbers dialed today between 10:00 AM and 5:43 PM IST, grouped by disposition
  const now = new Date();
  const istNow = new Date(now.getTime() + (5.5 * 60 * 60 * 1000));
  const istTodayStr = istNow.toISOString().slice(0, 10);
  
  // Today start at 10:00 AM IST and end at 5:43 PM IST
  const startIST = new Date(istTodayStr + 'T10:00:00.000+05:30');
  const endIST = new Date(istTodayStr + 'T17:43:00.000+05:30');
  
  // Filter dialed log for today between 10:00 AM and 5:43 PM
  const filteredLogs = appState.dialedLog.filter(entry => {
    if (!entry.timestamp) return false;
    const entryDate = new Date(entry.timestamp);
    return entryDate >= startIST && entryDate <= endIST;
  });
  
  // Group by disposition
  const groups = {};
  filteredLogs.forEach(entry => {
    const dispo = entry.disposition || 'unknown';
    if (!groups[dispo]) groups[dispo] = [];
    // Only add phone if not already in this disposition group
    if (!groups[dispo].includes(entry.phone)) {
      groups[dispo].push(entry.phone);
    }
  });
  
  // If no data, provide dummy data for testing
  const hasSomeData = Object.keys(groups).length > 0;
  if (!hasSomeData) {
    groups.dead = ['9876543210', '9876543211'];
    groups.not_received = ['9876543212', '9876543213'];
    groups.not_interested = ['9876543214'];
    groups.followup = ['9876543215'];
    groups.switch_off = ['9876543216'];
    groups.interested = ['9876543217'];
    groups._isDummy = true;
  }
  
  res.json({ 
    date: istTodayStr, 
    timeRange: '10:00 AM - 5:43 PM',
    groups,
    isDummy: !hasSomeData
  });
});

// ─── Admin: Download uploaded numbers sheet back as Excel ──────────────────────
app.get('/api/admin/download-numbers/:fileId', (req, res) => {
  const fid = req.params.fileId;
  const fileInfo = appState.uploadedFiles.find(f => f.id === fid);
  if (!fileInfo) return res.status(404).json({ error: 'File not found' });
  const fileNumbers = appState.numbers.filter(n => n.file === fid);
  if (fileNumbers.length === 0) return res.status(404).json({ error: 'No numbers found for this file' });
  const rows = [['Phone', 'Name', 'Disposition', 'Dialed By', 'Dialed At']];
  fileNumbers.forEach(n => {
    const agentName = n.dialedBy && appState.agents[n.dialedBy] ? appState.agents[n.dialedBy].name : (n.dialedBy || '');
    rows.push([n.phone || '', n.name || '', n.disposition || 'Pending', agentName, n.dialedAt || '']);
  });
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Numbers');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const downloadName = (fileInfo.name || 'numbers').replace(/\.[^.]+$/, '') + '_export.xlsx';
  res.setHeader('Content-Disposition', 'attachment; filename=' + encodeURIComponent(downloadName));
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// ─── Admin: Download the ORIGINAL uploaded sheet exactly as it was uploaded ────
app.get('/api/admin/original-file/:fileId', (req, res) => {
  const fid = req.params.fileId;
  const fileInfo = appState.uploadedFiles.find(f => f.id === fid);
  if (!fileInfo || !fileInfo.sheetPath || !fs.existsSync(fileInfo.sheetPath)) {
    return res.status(404).json({ error: 'Original file is not available for this upload (uploaded before this feature, or already removed).' });
  }
  res.download(fileInfo.sheetPath, fileInfo.name);
});

// ─── Manual Lead Addition (Agent / TL-Agent) ──────────────────────────────────
app.post('/api/agent/add-manual-number', (req, res) => {
  const { agentId, phone, name } = req.body;
  const clean = String(phone || '').trim().replace(/\s+/g, '');
  if (!/^\d{10}$/.test(clean)) return res.status(400).json({ error: 'Valid 10-digit phone required' });
  if (appState.numbers.find(n => n.phone === clean)) return res.status(400).json({ error: 'Number already exists in system' });
  let manualFile = appState.uploadedFiles.find(f => f.id === 'manual');
  if (!manualFile) {
    manualFile = { id: 'manual', name: 'Manual Entries', uploadedAt: new Date().toISOString(), total: 0, hasOriginal: false };
    appState.uploadedFiles.push(manualFile);
  }
  const newNum = { id: uuidv4(), phone: clean, name: String(name || '').trim(), file: 'manual', assignedTo: null, dialedBy: null, dialedAt: null };
  appState.numbers.push(newNum);
  manualFile.total = appState.numbers.filter(n => n.file === 'manual').length;
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true, numberId: newNum.id });
});

// Remove a number from pool — blocked if rider is recruited (real business data)
app.delete('/api/agent/number/:numberId', (req, res) => {
  const num = appState.numbers.find(n => n.id === req.params.numberId);
  if (!num) return res.status(404).json({ error: 'Number not found' });
  if (num.disposition === 'interested') return res.status(400).json({ error: 'Cannot remove a recruited rider' });
  appState.numbers = appState.numbers.filter(n => n.id !== req.params.numberId);
  const manualFile = appState.uploadedFiles.find(f => f.id === 'manual');
  if (manualFile) manualFile.total = appState.numbers.filter(n => n.file === 'manual').length;
  saveState(appState);
  broadcastAdminStats();
  res.json({ success: true });
});

// ─── PDF Report Archive (permanent, searchable by date, generated automatically) ─
// Reports are built server-side with pdfkit and written straight to REPORTS_DIR —
// no manual upload/generate step required. They are indexed in appState.reports
// (which survives Hard Reset) so admin.html and client.html can search by date
// and re-download at any time.
const DISPO_LABELS = {
  dead: 'CNC (Dead)', not_received: 'CNR (Not Received)', not_interested: 'Not Interested',
  followup: 'Followup', switch_off: 'Switch Off', interested: 'Recruited',
  discard: 'Not-Eligible (Discard)', dnd: 'DND'
};

function registerReport(entry) {
  if (!appState.reports) appState.reports = [];
  appState.reports.push(entry);
  saveState(appState);
  return entry;
}

// Builds one PDF for the given IST date string (YYYY-MM-DD) summarizing daily
// disposition stats + admin/lead stats, saves it to disk, and registers it in
// the permanent report archive. Used by the automatic 6:30 PM IST scheduler.
function generateDailyReport(dateStr) {
  return new Promise((resolve, reject) => {
    try {
      const dispo = computeDispositionStats('daily', null);
      const stats = getAdminStats();
      const fileName = uuidv4() + '.pdf';
      const filePath = path.join(REPORTS_DIR, fileName);
      const doc = new PDFDocument({ margin: 50 });
      const writeStream = fs.createWriteStream(filePath);
      doc.pipe(writeStream);

      doc.fontSize(18).fillColor('#fc8019').text('Swiggy Rider Recruitment CRM — Daily Report', { align: 'left' });
      doc.moveDown(0.3);
      doc.fontSize(11).fillColor('#12293f').text('Date: ' + dateStr + '  (Auto-generated at 6:30 PM IST)');
      doc.moveDown(1);

      doc.fontSize(13).fillColor('#fc8019').text('Rider Recruitment Overview');
      doc.moveDown(0.3);
      doc.fontSize(11).fillColor('#12293f');
      doc.text('Total Numbers Uploaded: ' + (stats.total || 0));
      doc.text('Remaining To Dial: ' + (stats.remaining || 0));
      doc.text('Recruited Riders (Active): ' + (stats.recruitedCount || 0));
      doc.text('Delivery Completed: ' + (stats.deliveryCompletedCount || 0));
      doc.text('Delivery Failed (7d expired): ' + (stats.deliveryFailedCount || 0));
      doc.text('Overdue Riders: ' + (stats.overdueRecruitedCount || 0));
      doc.text('Followups Pending: ' + (stats.followupCount || 0));
      doc.text('Redialing Tomorrow: ' + (stats.comingBackTomorrow || 0));
      doc.text('DND Numbers: ' + (stats.dndCount || 0));
      doc.moveDown(1);

      doc.fontSize(13).fillColor('#fc8019').text('Disposition Breakdown (Today)');
      doc.moveDown(0.3);
      doc.fontSize(11).fillColor('#12293f');
      doc.text('Total Calls: ' + (dispo.totalCalls || 0));
      Object.keys(DISPO_LABELS).forEach(key => {
        doc.text(DISPO_LABELS[key] + ': ' + (dispo[key] || 0));
      });

      doc.end();
      writeStream.on('finish', () => {
        const sizeBytes = fs.statSync(filePath).size;
        const entry = registerReport({
          id: uuidv4(),
          fileName,
          originalName: dateStr + '_Swiggy_Rider_Recruitment_Daily_Report.pdf',
          title: 'Daily Report',
          reportDate: dateStr,
          generatedBy: 'system-auto-6:30pm-IST',
          scope: 'general',
          uploadedAt: new Date().toISOString(),
          sizeBytes
        });
        resolve(entry);
      });
      writeStream.on('error', reject);
    } catch (e) {
      reject(e);
    }
  });
}

// ─── Automatic Daily Report Scheduler (6:30 PM IST, every day) ────────────────
// Computes ms until the next 6:30 PM IST, generates+saves a report at that
// moment, then reschedules itself for the following day. Also guards against
// duplicate reports if the server restarts around the trigger time on the same day.
function msUntilNext630PM_IST() {
  const now = new Date();
  const istNow = new Date(now.getTime() + 5.5 * 60 * 60 * 1000);
  const target = new Date(istNow);
  target.setUTCHours(18, 30, 0, 0); // 6:30 PM on the IST-shifted clock (stored as UTC fields)
  if (target <= istNow) target.setUTCDate(target.getUTCDate() + 1);
  return target.getTime() - istNow.getTime();
}

function scheduleAutoDailyReport() {
  const delay = msUntilNext630PM_IST();
  setTimeout(async () => {
    try {
      const todayStr = getTodayStr();
      const already = (appState.reports || []).some(r => r.reportDate === todayStr && r.generatedBy === 'system-auto-6:30pm-IST');
      if (!already) {
        await generateDailyReport(todayStr);
        console.log('\uD83D\uDCC4 Auto-generated daily report for ' + todayStr + ' at 6:30 PM IST');
      }
    } catch (e) {
      console.error('Auto daily report generation failed:', e.message);
    }
    scheduleAutoDailyReport(); // reschedule for the next day
  }, delay);
}
scheduleAutoDailyReport();

// List / search reports — optional query params: date=YYYY-MM-DD, from=YYYY-MM-DD, to=YYYY-MM-DD, scope=
app.get('/api/reports', (req, res) => {
  const { date, from, to, scope } = req.query;
  let list = (appState.reports || []).slice();
  if (date) list = list.filter(r => r.reportDate === date);
  if (from) list = list.filter(r => r.reportDate >= from);
  if (to) list = list.filter(r => r.reportDate <= to);
  if (scope) list = list.filter(r => r.scope === scope);
  // Newest first
  list.sort((a, b) => (b.reportDate + b.uploadedAt).localeCompare(a.reportDate + a.uploadedAt));
  res.json(list);
});

app.get('/api/reports/:id/download', (req, res) => {
  const entry = (appState.reports || []).find(r => r.id === req.params.id);
  if (!entry) return res.status(404).json({ error: 'Report not found' });
  const filePath = path.join(REPORTS_DIR, entry.fileName);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'Report file missing on disk' });
  res.download(filePath, entry.originalName || (entry.title + '.pdf'));
});

// ─── Client Panel Session Tracking ────────────────────────────────────────────
// Records each client panel login/logout so admin can see daily usage logs.

app.post('/api/client/session-start', (req, res) => {
  const { eid, name } = req.body || {};
  if (!eid) return res.status(400).json({ error: 'eid required' });
  const sessionId = Date.now() + '-' + Math.random().toString(36).slice(2, 7);
  const entry = {
    sessionId,
    eid: String(eid),
    name: name || 'Client',
    loginAt: new Date().toISOString(),
    logoutAt: null,
    durationMs: null,
    date: getTodayStr()
  };
  if (!appState.clientLogs) appState.clientLogs = [];
  appState.clientLogs.push(entry);
  saveState(appState);
  res.json({ ok: true, sessionId });
});

app.post('/api/client/session-end', (req, res) => {
  const { sessionId } = req.body || {};
  if (!sessionId) return res.json({ ok: true });
  if (!appState.clientLogs) return res.json({ ok: true });
  const entry = appState.clientLogs.find(l => l.sessionId === sessionId);
  if (entry && !entry.logoutAt) {
    entry.logoutAt = new Date().toISOString();
    entry.durationMs = Date.now() - new Date(entry.loginAt).getTime();
    saveState(appState);
  }
  res.json({ ok: true });
});

app.get('/api/admin/client-logs', (req, res) => {
  const { date } = req.query;
  const target = date || getTodayStr();
  const logs = (appState.clientLogs || []).filter(l => l.date === target);
  res.json({ logs, date: target });
});

// ─── Page Routes ──────────────────────────────────────────────────────────────
app.get('/admin', (req, res) => {
  if (!isAuthenticated(req)) return sendPasswordPage(res);
  res.sendFile(path.join(__dirname, 'public/admin/index.html'));
});
app.get('/agent', (req, res) => res.sendFile(path.join(__dirname, 'public/agent/index.html')));
app.get('/client', (req, res) => res.sendFile(path.join(__dirname, 'public/client/index.html')));
// Backward-compatible alias for old bookmarks/links pointing at the previous TL panel URL.
app.get('/tl', (req, res) => res.redirect(301, '/client'));

// ─── Start ─────────────────────────────────────────────────────────────────────
server.listen(PORT, '0.0.0.0', () => {
  console.log(`\n✅  Swiggy Rider Recruitment CRM running on http://0.0.0.0:${PORT}`);
  console.log(`   Admin Panel  : http://YOUR-LAN-IP:${PORT}/admin`);
  console.log(`   Agent Panel  : http://YOUR-LAN-IP:${PORT}/agent`);
  console.log(`   Client Panel : http://YOUR-LAN-IP:${PORT}/client\n`);
});

