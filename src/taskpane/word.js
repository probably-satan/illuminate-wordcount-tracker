/* global document, Office, Word */

const STORAGE_KEY = "wordSessionTrackerStatsV1";
const EMPTY_STATS = {
  totalAdded: 0,
  totalDeleted: 0,
  totalSeconds: 0,
  totalSessions: 0,
  lastSavedAt: null,
  byDate: {},
};

const state = {
  isTracking: false,
  timerIntervalId: null,
  scanIntervalId: null,
  persistIntervalId: null,
  scanInProgress: false,
  persistInProgress: false,
  wordsAdded: 0,
  wordsDeleted: 0,
  elapsedSeconds: 0,
  previousWordMap: new Map(),
  totalsVisible: false,
  lifetimeVisible: false,
  hasUnsavedChanges: false,
  confirmResolver: null,
  documentStats: cloneEmptyStats(),
  pendingPersist: {
    added: 0,
    deleted: 0,
    seconds: 0,
  },
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initializeWordTracker().catch((error) => {
      setStatus(`Initialization failed: ${toErrorMessage(error)}`);
    });
  }
});

async function initializeWordTracker() {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  document.getElementById("start-session").addEventListener("click", () => {
    startSession(false).catch((error) => {
      setStatus(`Unable to start: ${toErrorMessage(error)}`);
    });
  });
  document.getElementById("stop-session").addEventListener("click", stopSession);
  document.getElementById("reset-session").addEventListener("click", () => {
    resetSession().catch((error) => {
      setStatus(`Unable to reset: ${toErrorMessage(error)}`);
    });
  });
  document.getElementById("save-now").addEventListener("click", () => {
    saveNow().catch((error) => {
      setStatus(`Unable to save: ${toErrorMessage(error)}`);
    });
  });
  document.getElementById("confirm-cancel").addEventListener("click", () => respondToConfirmation(false));
  document.getElementById("confirm-ok").addEventListener("click", () => respondToConfirmation(true));
  document.getElementById("reset-lifetime").addEventListener("click", () => {
    resetLifetime().catch((error) => {
      setStatus(`Unable to reset lifetime stats: ${toErrorMessage(error)}`);
    });
  });
  document.getElementById("toggle-stats").addEventListener("click", toggleProductivityPanel);
  document.getElementById("toggle-lifetime").addEventListener("click", toggleLifetimePanel);

  loadDocumentStatsFromSettings();
  render();

  await startSession(true);
}

async function startSession(isAutoStart) {
  if (state.isTracking) {
    return;
  }

  try {
    state.previousWordMap = await getDocumentWordMap();
    state.isTracking = true;
    recordNewSessionStart();

    state.timerIntervalId = window.setInterval(onTimerTick, 1000);
    state.scanIntervalId = window.setInterval(scanForWordChanges, 2000);
    state.persistIntervalId = window.setInterval(() => {
      persistPendingStats().catch(() => {
        // Best effort persistence; status remains focused on tracking.
      });
    }, 12000);

    if (isAutoStart) {
      setStatus("Auto-tracking started for this document.");
    } else {
      setStatus("Tracking in progress.");
    }

    render();
  } catch (error) {
    setStatus(`Unable to start: ${toErrorMessage(error)}`);
  }
}

function stopSession() {
  if (!state.isTracking) {
    return;
  }

  state.isTracking = false;
  clearIntervals();

  persistPendingStats().catch(() => {
    // Preserve stop workflow even if persistence fails.
  });

  render();
  setStatus("Session stopped.");
}

async function resetSession() {
  const confirmed = await requestConfirmation(
    "Reset current writing session counters? This clears session time, words added, and words deleted for this session only."
  );

  if (!confirmed) {
    setStatus("Session reset canceled.");
    return;
  }

  state.wordsAdded = 0;
  state.wordsDeleted = 0;
  state.elapsedSeconds = 0;

  if (state.isTracking) {
    state.previousWordMap = await getDocumentWordMap();
  }

  render();
  setStatus("Current writing session counters reset.");
}

async function resetLifetime() {
  const confirmed = await requestConfirmation(
    "Reset all lifetime stats for this document? This cannot be undone."
  );

  if (!confirmed) {
    setStatus("Lifetime reset canceled.");
    return;
  }

  state.documentStats = cloneEmptyStats();
  state.pendingPersist = {
    added: 0,
    deleted: 0,
    seconds: 0,
  };
  state.hasUnsavedChanges = true;

  await persistStatsToSettings();
  renderProductivityStats();
  renderLastSaved();
  setStatus("Lifetime stats reset for this document.");
}

async function saveNow() {
  await persistStatsToSettings();
  state.pendingPersist = {
    added: 0,
    deleted: 0,
    seconds: 0,
  };
  renderLastSaved();
  renderSaveState();
  setStatus("Stats saved to document.");
}

async function scanForWordChanges() {
  if (!state.isTracking || state.scanInProgress) {
    return;
  }

  state.scanInProgress = true;

  try {
    const nextWordMap = await getDocumentWordMap();
    const delta = calculateWordDelta(state.previousWordMap, nextWordMap);

    state.wordsAdded += delta.added;
    state.wordsDeleted += delta.deleted;
    addPendingDelta(delta.added, delta.deleted, 0);
    state.previousWordMap = nextWordMap;

    applyDeltaToDocumentStats(delta.added, delta.deleted, 0);
    renderWordCounters();
    renderProductivityStats();
  } catch (error) {
    setStatus(`Tracking paused briefly: ${toErrorMessage(error)}`);
  } finally {
    state.scanInProgress = false;
  }
}

function calculateWordDelta(previousMap, nextMap) {
  let added = 0;
  let deleted = 0;

  const vocabulary = new Set([...previousMap.keys(), ...nextMap.keys()]);

  vocabulary.forEach((word) => {
    const previousCount = previousMap.get(word) || 0;
    const nextCount = nextMap.get(word) || 0;

    if (nextCount > previousCount) {
      added += nextCount - previousCount;
    }

    if (previousCount > nextCount) {
      deleted += previousCount - nextCount;
    }
  });

  return { added, deleted };
}

async function getDocumentWordMap() {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    return buildWordMap(body.text || "");
  });
}

function buildWordMap(text) {
  const normalizedWords = text.toLowerCase().match(/[a-z0-9']+/gi);
  const wordMap = new Map();

  if (!normalizedWords) {
    return wordMap;
  }

  normalizedWords.forEach((word) => {
    wordMap.set(word, (wordMap.get(word) || 0) + 1);
  });

  return wordMap;
}

function onTimerTick() {
  if (!state.isTracking) {
    return;
  }

  state.elapsedSeconds += 1;
  addPendingDelta(0, 0, 1);
  applyDeltaToDocumentStats(0, 0, 1);

  renderElapsedTime();
  renderProductivityStats();
}

function formatDuration(totalSeconds) {
  const safeTotal = Math.max(0, totalSeconds);
  const hours = String(Math.floor(safeTotal / 3600)).padStart(2, "0");
  const minutes = String(Math.floor((safeTotal % 3600) / 60)).padStart(2, "0");
  const seconds = String(safeTotal % 60).padStart(2, "0");
  return `${hours}:${minutes}:${seconds}`;
}

function clearIntervals() {
  if (state.timerIntervalId) {
    clearInterval(state.timerIntervalId);
    state.timerIntervalId = null;
  }

  if (state.scanIntervalId) {
    clearInterval(state.scanIntervalId);
    state.scanIntervalId = null;
  }

  if (state.persistIntervalId) {
    clearInterval(state.persistIntervalId);
    state.persistIntervalId = null;
  }
}

function render() {
  renderElapsedTime();
  renderWordCounters();
  renderProductivityStats();
  renderLastSaved();
  renderSaveState();
  updateButtons();
}

function renderSaveState() {
  const el = document.getElementById("save-state");
  el.textContent = state.hasUnsavedChanges ? "Changes pending save" : "All changes saved";
}

function renderLastSaved() {
  const label = document.getElementById("last-saved");
  const stamp = state.documentStats.lastSavedAt;

  if (!stamp) {
    label.textContent = "Last saved: Never";
    return;
  }

  const parsed = new Date(stamp);
  if (Number.isNaN(parsed.getTime())) {
    label.textContent = "Last saved: Unknown";
    return;
  }

  label.textContent = `Last saved: ${parsed.toLocaleString()}`;
}

function renderElapsedTime() {
  document.getElementById("elapsed-time").textContent = formatDuration(state.elapsedSeconds);
}

function renderWordCounters() {
  document.getElementById("words-added").textContent = String(state.wordsAdded);
  document.getElementById("words-deleted").textContent = String(state.wordsDeleted);
}

function renderProductivityStats() {
  const byDateKeys = Object.keys(state.documentStats.byDate).sort();
  const weekly = calculateWeeklyTotals(state.documentStats.byDate);
  const best = calculateBestDay(state.documentStats.byDate);
  const activeDays = byDateKeys.length;
  const lifetimeNet = state.documentStats.totalAdded - state.documentStats.totalDeleted;
  const sessionNet = state.wordsAdded - state.wordsDeleted;
  const sessionMinutes = Math.max(state.elapsedSeconds / 60, 1 / 60);
  const sessionWpm = state.wordsAdded / sessionMinutes;
  const avgDailyNet = activeDays > 0 ? lifetimeNet / activeDays : 0;

  document.getElementById("session-net-words").textContent = String(sessionNet);
  document.getElementById("session-wpm").textContent = sessionWpm.toFixed(1);
  document.getElementById("weekly-net-words").textContent = String(weekly.net);
  document.getElementById("weekly-added").textContent = String(weekly.added);
  document.getElementById("weekly-deleted").textContent = String(weekly.deleted);
  document.getElementById("lifetime-net-words").textContent = String(lifetimeNet);
  document.getElementById("active-days").textContent = String(activeDays);
  document.getElementById("avg-daily-net").textContent = avgDailyNet.toFixed(1);
  document.getElementById("best-day").textContent = best ? `${best.date}: ${best.net}` : "-";

  renderDailyTableRows(byDateKeys);
  renderLifetimeStats(byDateKeys);
}

function renderLifetimeStats(sortedDateKeys) {
  const lifetimeAdded = state.documentStats.totalAdded;
  const lifetimeDeleted = state.documentStats.totalDeleted;
  const lifetimeNet = lifetimeAdded - lifetimeDeleted;
  const totalSessions = state.documentStats.totalSessions;
  const lifetimeMinutes = state.documentStats.totalSeconds / 60;
  const avgNetSession = totalSessions > 0 ? lifetimeNet / totalSessions : 0;
  const streaks = calculateStreakStats(state.documentStats.byDate);
  const weeklyRollups = calculateWeeklyRollups(state.documentStats.byDate);
  const bestWeek = weeklyRollups.length > 0 ? weeklyRollups.reduce((best, item) => (item.net > best.net ? item : best), weeklyRollups[0]) : null;

  document.getElementById("lifetime-added").textContent = String(lifetimeAdded);
  document.getElementById("lifetime-deleted").textContent = String(lifetimeDeleted);
  document.getElementById("lifetime-net").textContent = String(lifetimeNet);
  document.getElementById("lifetime-minutes").textContent = lifetimeMinutes.toFixed(1);
  document.getElementById("lifetime-sessions").textContent = String(totalSessions);
  document.getElementById("avg-net-session").textContent = avgNetSession.toFixed(1);
  document.getElementById("current-streak").textContent = String(streaks.current);
  document.getElementById("longest-streak").textContent = String(streaks.longest);
  document.getElementById("best-week").textContent = bestWeek ? `${bestWeek.weekOf}: ${bestWeek.net}` : "-";

  renderWeeklyTableRows(weeklyRollups, sortedDateKeys.length === 0);
}

function renderDailyTableRows(sortedDateKeys) {
  const tbody = document.getElementById("daily-stats-rows");
  tbody.innerHTML = "";

  if (sortedDateKeys.length === 0) {
    const emptyRow = document.createElement("tr");
    emptyRow.innerHTML = "<td colspan=\"5\">No document totals yet.</td>";
    tbody.appendChild(emptyRow);
    return;
  }

  sortedDateKeys.slice().reverse().forEach((dateKey) => {
    const day = state.documentStats.byDate[dateKey];
    const net = day.added - day.deleted;
    const minutes = (day.seconds / 60).toFixed(1);

    const row = document.createElement("tr");
    row.innerHTML = `<td>${dateKey}</td><td>${day.added}</td><td>${day.deleted}</td><td>${net}</td><td>${minutes}</td>`;
    tbody.appendChild(row);
  });
}

function calculateWeeklyTotals(byDate) {
  const keys = Object.keys(byDate);
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 6);
  cutoff.setHours(0, 0, 0, 0);

  let added = 0;
  let deleted = 0;

  keys.forEach((key) => {
    const dayDate = parseLocalDateKey(key);
    if (dayDate >= cutoff) {
      added += byDate[key].added;
      deleted += byDate[key].deleted;
    }
  });

  return {
    added,
    deleted,
    net: added - deleted,
  };
}

function calculateBestDay(byDate) {
  let best = null;

  Object.keys(byDate).forEach((key) => {
    const day = byDate[key];
    const net = day.added - day.deleted;

    if (!best || net > best.net) {
      best = {
        date: key,
        net,
      };
    }
  });

  return best;
}

function calculateStreakStats(byDate) {
  const positiveDays = Object.keys(byDate)
    .filter((key) => {
      const day = byDate[key];
      return (day.added - day.deleted) > 0;
    })
    .sort();

  if (positiveDays.length === 0) {
    return { current: 0, longest: 0 };
  }

  let longest = 1;
  let run = 1;

  for (let i = 1; i < positiveDays.length; i += 1) {
    const prev = parseLocalDateKey(positiveDays[i - 1]);
    const curr = parseLocalDateKey(positiveDays[i]);
    const diffDays = Math.round((curr - prev) / 86400000);

    if (diffDays === 1) {
      run += 1;
      if (run > longest) {
        longest = run;
      }
    } else {
      run = 1;
    }
  }

  let current = 1;
  const today = getLocalDateKey(new Date());
  const yesterdayDate = new Date();
  yesterdayDate.setDate(yesterdayDate.getDate() - 1);
  const yesterday = getLocalDateKey(yesterdayDate);
  const last = positiveDays[positiveDays.length - 1];

  if (last !== today && last !== yesterday) {
    current = 0;
  } else {
    for (let i = positiveDays.length - 1; i > 0; i -= 1) {
      const prev = parseLocalDateKey(positiveDays[i - 1]);
      const curr = parseLocalDateKey(positiveDays[i]);
      const diffDays = Math.round((curr - prev) / 86400000);
      if (diffDays === 1) {
        current += 1;
      } else {
        break;
      }
    }
  }

  return { current, longest };
}

function getWeekStartDateKey(date) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = d.getDay();
  const diffToMonday = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + diffToMonday);
  return getLocalDateKey(d);
}

function calculateWeeklyRollups(byDate) {
  const weeks = {};

  Object.keys(byDate).forEach((dateKey) => {
    const weekOf = getWeekStartDateKey(parseLocalDateKey(dateKey));
    const day = byDate[dateKey];

    if (!weeks[weekOf]) {
      weeks[weekOf] = {
        weekOf,
        added: 0,
        deleted: 0,
        seconds: 0,
        net: 0,
      };
    }

    weeks[weekOf].added += day.added;
    weeks[weekOf].deleted += day.deleted;
    weeks[weekOf].seconds += day.seconds;
    weeks[weekOf].net = weeks[weekOf].added - weeks[weekOf].deleted;
  });

  return Object.values(weeks).sort((a, b) => a.weekOf.localeCompare(b.weekOf));
}

function renderWeeklyTableRows(weeklyRollups, hasNoDailyData) {
  const tbody = document.getElementById("weekly-stats-rows");
  tbody.innerHTML = "";

  if (hasNoDailyData || weeklyRollups.length === 0) {
    const emptyRow = document.createElement("tr");
    emptyRow.innerHTML = "<td colspan=\"5\">No weekly totals yet.</td>";
    tbody.appendChild(emptyRow);
    return;
  }

  weeklyRollups.slice().reverse().forEach((week) => {
    const minutes = (week.seconds / 60).toFixed(1);
    const row = document.createElement("tr");
    row.innerHTML = `<td>${week.weekOf}</td><td>${week.added}</td><td>${week.deleted}</td><td>${week.net}</td><td>${minutes}</td>`;
    tbody.appendChild(row);
  });
}

function parseLocalDateKey(dateKey) {
  const [year, month, day] = dateKey.split("-").map((value) => parseInt(value, 10));
  return new Date(year, month - 1, day);
}

function toggleProductivityPanel() {
  state.totalsVisible = !state.totalsVisible;

  const panel = document.getElementById("productivity-panel");
  const toggleButton = document.getElementById("toggle-stats");
  panel.style.display = state.totalsVisible ? "flex" : "none";
  toggleButton.textContent = state.totalsVisible ? "Hide Productivity Stats" : "Show Productivity Stats";

  if (state.totalsVisible) {
    renderProductivityStats();
  }
}

function toggleLifetimePanel() {
  state.lifetimeVisible = !state.lifetimeVisible;

  const panel = document.getElementById("lifetime-panel");
  const toggleButton = document.getElementById("toggle-lifetime");
  panel.style.display = state.lifetimeVisible ? "flex" : "none";
  toggleButton.textContent = state.lifetimeVisible ? "Hide Lifetime Stats" : "Show Lifetime Stats";

  if (state.lifetimeVisible) {
    renderProductivityStats();
  }
}

function cloneEmptyStats() {
  return {
    totalAdded: EMPTY_STATS.totalAdded,
    totalDeleted: EMPTY_STATS.totalDeleted,
    totalSeconds: EMPTY_STATS.totalSeconds,
    totalSessions: EMPTY_STATS.totalSessions,
    lastSavedAt: EMPTY_STATS.lastSavedAt,
    byDate: {},
  };
}

function loadDocumentStatsFromSettings() {
  const saved = Office.context.document.settings.get(STORAGE_KEY);

  if (!saved || typeof saved !== "object") {
    state.documentStats = cloneEmptyStats();
    return;
  }

  state.documentStats = {
    totalAdded: toNonNegativeInt(saved.totalAdded),
    totalDeleted: toNonNegativeInt(saved.totalDeleted),
    totalSeconds: toNonNegativeInt(saved.totalSeconds),
    totalSessions: toNonNegativeInt(saved.totalSessions),
    lastSavedAt: typeof saved.lastSavedAt === "string" ? saved.lastSavedAt : null,
    byDate: sanitizeByDate(saved.byDate),
  };
}

function sanitizeByDate(byDate) {
  const clean = {};

  if (!byDate || typeof byDate !== "object") {
    return clean;
  }

  Object.keys(byDate).forEach((dateKey) => {
    const day = byDate[dateKey] || {};
    clean[dateKey] = {
      added: toNonNegativeInt(day.added),
      deleted: toNonNegativeInt(day.deleted),
      seconds: toNonNegativeInt(day.seconds),
      sessions: toNonNegativeInt(day.sessions),
    };
  });

  return clean;
}

function toNonNegativeInt(value) {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) {
    return 0;
  }

  return Math.floor(n);
}

function recordNewSessionStart() {
  state.documentStats.totalSessions += 1;

  const dateKey = getLocalDateKey(new Date());
  ensureDateBucket(dateKey);
  state.documentStats.byDate[dateKey].sessions += 1;
  state.hasUnsavedChanges = true;
  renderSaveState();

  persistStatsToSettings().catch(() => {
    // Best effort persistence on start.
  });
}

function getLocalDateKey(date) {
  const year = String(date.getFullYear());
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function ensureDateBucket(dateKey) {
  if (!state.documentStats.byDate[dateKey]) {
    state.documentStats.byDate[dateKey] = {
      added: 0,
      deleted: 0,
      seconds: 0,
      sessions: 0,
    };
  }
}

function addPendingDelta(added, deleted, seconds) {
  state.pendingPersist.added += added;
  state.pendingPersist.deleted += deleted;
  state.pendingPersist.seconds += seconds;
  if (added > 0 || deleted > 0 || seconds > 0) {
    state.hasUnsavedChanges = true;
    renderSaveState();
  }
}

function applyDeltaToDocumentStats(added, deleted, seconds) {
  state.documentStats.totalAdded += added;
  state.documentStats.totalDeleted += deleted;
  state.documentStats.totalSeconds += seconds;

  const dateKey = getLocalDateKey(new Date());
  ensureDateBucket(dateKey);
  state.documentStats.byDate[dateKey].added += added;
  state.documentStats.byDate[dateKey].deleted += deleted;
  state.documentStats.byDate[dateKey].seconds += seconds;
  if (added > 0 || deleted > 0 || seconds > 0) {
    state.hasUnsavedChanges = true;
  }
}

async function persistPendingStats() {
  if (state.persistInProgress) {
    return;
  }

  const hasPending = state.pendingPersist.added > 0 || state.pendingPersist.deleted > 0 || state.pendingPersist.seconds > 0;

  if (!hasPending) {
    return;
  }

  state.persistInProgress = true;

  try {
    await persistStatsToSettings();
    state.pendingPersist = {
      added: 0,
      deleted: 0,
      seconds: 0,
    };
  } finally {
    state.persistInProgress = false;
  }
}

async function persistStatsToSettings() {
  state.documentStats.lastSavedAt = new Date().toISOString();
  Office.context.document.settings.set(STORAGE_KEY, state.documentStats);

  await new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        state.hasUnsavedChanges = false;
        renderSaveState();
        resolve();
      } else {
        state.hasUnsavedChanges = true;
        renderSaveState();
        reject(new Error(result.error && result.error.message ? result.error.message : "save failed"));
      }
    });
  });
}

function requestConfirmation(message) {
  return new Promise((resolve) => {
    if (state.confirmResolver) {
      resolve(false);
      return;
    }

    state.confirmResolver = resolve;
    document.getElementById("confirm-message").textContent = message;
    setResetButtonsDisabled(true);
    document.getElementById("confirm-overlay").style.display = "flex";
  });
}

function respondToConfirmation(confirmed) {
  if (!state.confirmResolver) {
    return;
  }

  const resolver = state.confirmResolver;
  state.confirmResolver = null;
  document.getElementById("confirm-overlay").style.display = "none";
  setResetButtonsDisabled(false);
  resolver(confirmed);
}

function setResetButtonsDisabled(disabled) {
  document.getElementById("reset-session").disabled = disabled;
  document.getElementById("reset-lifetime").disabled = disabled;
}

function updateButtons() {
  document.getElementById("start-session").disabled = state.isTracking;
  document.getElementById("stop-session").disabled = !state.isTracking;
}

function setStatus(message) {
  document.getElementById("status").textContent = message;
}

function toErrorMessage(error) {
  if (error && typeof error.message === "string") {
    return error.message;
  }

  return "unknown error";
}
