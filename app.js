// ============ UTILS ============
function escapeHtml(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

function pad(n) { return String(n).padStart(2, '0'); }
function now24() { const d = new Date(); return pad(d.getHours()) + ':' + pad(d.getMinutes()); }
function todayStr() { const d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }

function to12(t) {
  if (!t) return '';
  const p = t.split(':').map(Number);
  const h = p[0] % 12 || 12;
  return h + ':' + pad(p[1]) + (p[0] >= 12 ? ' PM' : ' AM');
}

function to24(t) {
  if (!t) return '';
  if (/^\d{2}:\d{2}$/.test(t.trim())) return t.trim();
  const match = t.trim().match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (!match) return '';
  let h = parseInt(match[1]);
  const m = match[2];
  const ampm = match[3].toUpperCase();
  if (ampm === 'PM' && h !== 12) h += 12;
  if (ampm === 'AM' && h === 12) h = 0;
  return pad(h) + ':' + m;
}

function normTime(t) {
  if (!t) return '';
  if (/^\d{1,2}:\d{2}\s*(AM|PM)$/i.test(t.trim())) return t.trim();
  if (/^\d{1,2}:\d{2}$/.test(t.trim())) return to12(t.trim());
  const d = new Date(t);
  if (!isNaN(d)) { const h = d.getHours(), m = d.getMinutes(); return (h % 12 || 12) + ':' + pad(m) + (h >= 12 ? ' PM' : ' AM'); }
  return t;
}

function normDate(ds) {
  if (!ds) return '';
  let d;
  if (ds.includes('T')) d = new Date(ds);
  else if (ds.match(/^\d{4}-\d{2}-\d{2}$/)) d = new Date(ds + 'T12:00:00');
  else d = new Date(ds);
  if (isNaN(d)) return ds;
  return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
}

function formatDateNice(ds) {
  if (!ds) return '';
  let d;
  if (ds.includes('T')) d = new Date(ds);
  else if (ds.match(/^\d{4}-\d{2}-\d{2}$/)) d = new Date(ds + 'T12:00:00');
  else d = new Date(ds);
  if (isNaN(d)) return ds;
  return d.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });
}

function toast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.style.display = 'block';
  setTimeout(() => t.style.display = 'none', 2500);
}

// ============ ELAPSED TIME (Fix #2) ============
function parseTime12(timeStr) {
  if (!timeStr) return null;
  const match = timeStr.trim().match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (!match) return null;
  let h = parseInt(match[1]);
  const m = parseInt(match[2]);
  const ampm = match[3].toUpperCase();
  if (ampm === 'PM' && h !== 12) h += 12;
  if (ampm === 'AM' && h === 12) h = 0;
  return { hours: h, minutes: m };
}

function getElapsedStr(timeStr, dateStr) {
  const parsed = parseTime12(timeStr);
  if (!parsed) return '-';
  const now = new Date();
  const then = new Date();
  then.setHours(parsed.hours, parsed.minutes, 0, 0);

  // If the entry is from a different day, use that date
  if (dateStr && dateStr !== todayStr()) {
    const dp = dateStr.split('-').map(Number);
    then.setFullYear(dp[0], dp[1] - 1, dp[2]);
  }

  // If calculated time is in the future, it was yesterday
  if (then > now) then.setDate(then.getDate() - 1);

  const diffMs = now - then;
  const diffMin = Math.floor(diffMs / 60000);
  if (diffMin < 1) return 'just now';
  if (diffMin < 60) return diffMin + 'm ago';
  const h = Math.floor(diffMin / 60);
  const m = diffMin % 60;
  if (h >= 24) {
    const days = Math.floor(h / 24);
    return days + 'd ago';
  }
  if (m === 0) return h + 'h ago';
  return h + 'h ' + m + 'm ago';
}

// ============ STATE ============
let allEntries = [];
let currentFilter = 'all';
let scriptUrl = localStorage.getItem('elliot_script_url') || '';
let picks = { fSide: '', fShield: '' };
let editingIdx = null; // Fix #5: track which entry is being edited
let diaperPick = '';
let pendingDeletes = JSON.parse(localStorage.getItem('elliot_pending_deletes') || '[]'); // Fix #3

const EMOJIS = { Feed: '\u{1F37C}', Diaper: '\u{1F9F7}', Pump: '\u{1F95B}', Note: '\u{1F4DD}' };
const SIDE_MAP = { 'L': 'L', 'R': 'R', '\u2192': '', 'L\u2192R': 'LR', 'R\u2192L': 'RL', 'LRL': 'LRL', 'RLR': 'RLR' };
const BASE_URL = window.location.origin + window.location.pathname;

// ============ LOCAL STORAGE ============
function saveLocal() { localStorage.setItem('elliot_entries', JSON.stringify(allEntries)); }
function loadLocal() {
  const r = localStorage.getItem('elliot_entries');
  if (r) {
    allEntries = JSON.parse(r);
    allEntries.forEach(e => { if (e._synced === undefined) e._synced = true; });
  }
}
function savePendingDeletes() { localStorage.setItem('elliot_pending_deletes', JSON.stringify(pendingDeletes)); }

// ============ CLOCK ============
function updateClock() {
  const d = new Date();
  document.getElementById('headerTime').textContent =
    d.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' }) +
    ' \u2022 ' + to12(now24());
  renderStats(); // Update elapsed time every tick
}

function setSyncState(s) { document.getElementById('syncDot').className = 'sync-dot ' + s; }

// ============ TABS & FILTER ============
function switchTab(el) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById(el.dataset.view + 'View').classList.add('active');
}

function setFilter(f, el) {
  currentFilter = f;
  document.querySelectorAll('.filter-pill').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  renderToday();
}

// ============ MODALS ============
function showModal(id) { document.getElementById(id).classList.add('show'); }
function closeModal(id) {
  document.getElementById(id).classList.remove('show');
  // Reset editing state when closing any modal
  if (['feedModal', 'pumpModal', 'noteModal', 'diaperModal'].includes(id)) {
    editingIdx = null;
  }
}

function showSetup() {
  document.getElementById('scriptUrl').value = scriptUrl;
  document.getElementById('setupStatus').innerHTML = '';
  document.getElementById('appBaseUrl').textContent = BASE_URL;
  showModal('setupModal');
}

function showFeedForm() {
  document.getElementById('fTime').value = now24();
  document.getElementById('fAmount').value = '';
  document.getElementById('fNotes').value = '';
  picks.fSide = ''; picks.fShield = '';
  document.querySelectorAll('#feedModal .pill').forEach(p => p.classList.remove('selected'));
  document.getElementById('feedSubmitBtn').textContent = 'Log Feed';
  showModal('feedModal');
}

function showPumpForm() {
  document.getElementById('pTime').value = now24();
  document.getElementById('pAmount').value = '';
  document.getElementById('pNotes').value = '';
  document.getElementById('pumpSubmitBtn').textContent = 'Log Pump';
  showModal('pumpModal');
}

function showNoteForm() {
  document.getElementById('nTime').value = now24();
  document.getElementById('nText').value = '';
  document.getElementById('noteSubmitBtn').textContent = 'Log Note';
  showModal('noteModal');
}

function showMoreMenu() { showModal('moreModal'); }

function showDiaperEditForm(entry) {
  document.getElementById('dTime').value = to24(entry.time);
  diaperPick = '';
  document.querySelectorAll('#diaperModal .pill').forEach(p => p.classList.remove('selected'));

  // Pre-select the diaper type pill
  const details = (entry.details || '').toLowerCase();
  let preselect = '';
  if (details.includes('pee & poo') || details.includes('pee and poo')) preselect = 'Pee & poo diaper';
  else if (details.includes('poo')) preselect = 'Poo diaper';
  else if (details.includes('pee')) preselect = 'Pee diaper';

  if (preselect) {
    diaperPick = preselect;
    document.querySelectorAll('#diaperModal .pill.diaper-type').forEach(p => {
      if (p.dataset.value === preselect) p.classList.add('selected');
    });
  }

  // Extract notes from parentheses
  const parenMatch = (entry.details || '').match(/\(([^)]+)\)/);
  let notes = '';
  if (parenMatch && parenMatch[1] !== 'live') notes = parenMatch[1];
  document.getElementById('dNotes').value = notes;

  document.getElementById('diaperSubmitBtn').textContent = 'Update Diaper';
  showModal('diaperModal');
}

function pickPill(btn, key) {
  btn.parentElement.querySelectorAll('.pill').forEach(p => p.classList.remove('selected'));
  btn.classList.add('selected');
  picks[key] = btn.textContent.trim();
}

// ============ API ============
async function apiCall(params) {
  if (!scriptUrl) return null;
  setSyncState('syncing');
  try {
    const url = new URL(scriptUrl);
    Object.keys(params).forEach(k => {
      if (k !== '_synced') url.searchParams.append(k, params[k]);
    });
    const res = await fetch(url.toString(), { redirect: 'follow' });
    const data = await res.json();
    setSyncState('connected');
    return data;
  } catch (e) { setSyncState('offline'); return null; }
}

// Fix #3: Add entry with sync tracking
async function addEntry(entry) {
  entry._synced = false;
  allEntries.push(entry);
  saveLocal(); renderAll();

  const result = await apiCall({ action: 'add', ...entry });
  if (result && result.success) {
    entry._synced = true;
    saveLocal();
    toast('\u2705 Saved to sheet!');
  } else {
    toast('\u{1F4F1} Saved locally (will sync later)');
  }
}

// Fix #4: Merge instead of overwrite on refresh
async function refreshData() {
  toast('\u{1F504} Loading...');
  const data = await apiCall({ action: 'read' });
  if (data && data.success && data.entries) {
    const sheetEntries = data.entries.map(e => ({ ...e, _synced: true, date: normDate(e.date), time: normTime(e.time) }));

    // Keep unsynced local entries (Fix #4)
    const unsynced = allEntries.filter(e => e._synced === false);

    allEntries = [...sheetEntries, ...unsynced];
    saveLocal(); renderAll();

    // Try to sync unsynced entries in the background
    let syncedCount = 0;
    for (const entry of unsynced) {
      const result = await apiCall({ action: 'add', ...entry });
      if (result && result.success) {
        entry._synced = true;
        syncedCount++;
      }
    }

    // Try to replay pending deletes
    const remainingDeletes = [];
    for (const del of pendingDeletes) {
      const result = await apiCall({ action: 'delete', ...del });
      if (!result || !result.success) remainingDeletes.push(del);
    }
    pendingDeletes = remainingDeletes;
    savePendingDeletes();

    if (syncedCount > 0) saveLocal();

    const stillUnsynced = allEntries.filter(e => e._synced === false).length;
    if (stillUnsynced > 0) {
      toast('\u2705 Loaded ' + allEntries.length + ' entries (' + stillUnsynced + ' pending sync)');
    } else {
      toast('\u2705 Loaded ' + allEntries.length + ' entries');
    }
  } else {
    toast('Using local cache');
  }
}

// ============ DELETE (Fix #7: match on all fields) ============
function confirmDelete(idx) {
  const entry = allEntries[idx];
  if (!entry) return;
  const emoji = EMOJIS[entry.type] || '\u{1F4DD}';
  const dc = document.getElementById('deleteConfirm');
  dc.style.display = 'block';
  dc.innerHTML = '<div class="delete-confirm"><div class="delete-confirm-box">' +
    '<h3>Delete entry?</h3>' +
    '<p>' + emoji + ' ' + escapeHtml(entry.time || '') + ' \u2014 ' + escapeHtml(entry.details || '') + '</p>' +
    '<div class="dc-btns">' +
    '<button class="dc-btn cancel" data-action="cancel-delete">Keep</button>' +
    '<button class="dc-btn confirm" data-action="do-delete" data-idx="' + idx + '">Delete</button>' +
    '</div></div></div>';
}

function cancelDelete() { document.getElementById('deleteConfirm').style.display = 'none'; }

async function doDelete(idx) {
  document.getElementById('deleteConfirm').style.display = 'none';
  const entry = allEntries[idx];

  allEntries.splice(idx, 1);
  saveLocal(); renderAll();

  if (scriptUrl && entry && entry._synced) {
    // Fix #7: send all fields for precise matching
    const result = await apiCall({
      action: 'delete',
      date: entry.date || '',
      time: entry.time || '',
      type: entry.type || '',
      side: entry.side || '',
      shield: entry.shield || '',
      amount: entry.amount || '',
      details: entry.details || ''
    });
    if (result && result.success) {
      toast('\u2705 Deleted from sheet');
    } else {
      // Fix #3: queue delete for later sync
      pendingDeletes.push({
        date: entry.date || '',
        time: entry.time || '',
        type: entry.type || '',
        side: entry.side || '',
        shield: entry.shield || '',
        amount: entry.amount || '',
        details: entry.details || ''
      });
      savePendingDeletes();
      toast('\u{1F4F1} Deleted locally (will sync later)');
    }
  } else {
    toast('\u{1F4F1} Deleted locally');
  }
}

// ============ EDIT (Fix #5) ============
function extractFeedNotes(entry) {
  let d = entry.details || '';
  d = d.replace(/^Feed/, '').trim();
  d = d.replace(/^on\s+\S+/, '').trim();
  d = d.replace(/\(shield:\s*\S+\)/, '').replace(/\(no shield\)/, '').trim();
  if (entry.amount) d = d.replace(entry.amount, '').trim();
  d = d.replace(/^[;\s,]+/, '').replace(/[;\s,]+$/, '').trim();
  return d;
}

function extractPumpNotes(entry) {
  let d = entry.details || '';
  d = d.replace(/^Pump/, '').trim();
  if (entry.amount) d = d.replace(entry.amount, '').trim();
  d = d.replace(/^[;\s,]+/, '').replace(/[;\s,]+$/, '').trim();
  return d;
}

function editEntry(idx) {
  const entry = allEntries[idx];
  if (!entry) return;
  editingIdx = idx;

  switch (entry.type) {
    case 'Feed':
      showFeedForm();
      document.getElementById('fTime').value = to24(entry.time);
      // Pre-select side pill
      if (entry.side) {
        const sideDisplay = { L: 'L', R: 'R', LR: 'L\u2192R', RL: 'R\u2192L', LRL: 'LRL', RLR: 'RLR' };
        const sideText = sideDisplay[entry.side] || entry.side;
        picks.fSide = sideText;
        document.querySelectorAll('#feedModal .pill-group').forEach((pg, pgIdx) => {
          if (pgIdx === 0) { // side group
            pg.querySelectorAll('.pill').forEach(p => {
              if (p.textContent.trim() === sideText) p.classList.add('selected');
            });
          }
        });
      }
      // Pre-select shield pill
      if (entry.shield) {
        picks.fShield = entry.shield;
        document.querySelectorAll('#feedModal .pill.shield').forEach(p => {
          if (p.textContent.trim() === entry.shield) p.classList.add('selected');
        });
      }
      document.getElementById('fAmount').value = entry.amount || '';
      document.getElementById('fNotes').value = extractFeedNotes(entry);
      document.getElementById('feedSubmitBtn').textContent = 'Update Feed';
      break;

    case 'Pump':
      showPumpForm();
      document.getElementById('pTime').value = to24(entry.time);
      document.getElementById('pAmount').value = entry.amount || '';
      document.getElementById('pNotes').value = extractPumpNotes(entry);
      document.getElementById('pumpSubmitBtn').textContent = 'Update Pump';
      break;

    case 'Note':
      showNoteForm();
      document.getElementById('nTime').value = to24(entry.time);
      document.getElementById('nText').value = entry.details || '';
      document.getElementById('noteSubmitBtn').textContent = 'Update Note';
      break;

    case 'Diaper':
      showDiaperEditForm(entry);
      break;
  }
}

// ============ SUBMIT HANDLERS ============
function submitFeed() {
  const time = to12(document.getElementById('fTime').value);
  const side = SIDE_MAP[picks.fSide] || picks.fSide || '';
  const shield = picks.fShield || '';
  const amount = document.getElementById('fAmount').value.trim();
  const notes = document.getElementById('fNotes').value.trim();
  let details = 'Feed';
  if (side) details += ' on ' + side;
  if (shield && shield !== 'None') details += ' (shield: ' + shield + ')';
  else if (shield === 'None') details += ' (no shield)';
  if (amount) details += '; ' + amount;
  if (notes) details += '; ' + notes;

  const entry = { date: todayStr(), time, type: 'Feed', side, shield, amount, details };

  if (editingIdx !== null) {
    replaceEntry(editingIdx, entry);
    editingIdx = null;
  } else {
    addEntry(entry);
  }
  closeModal('feedModal');
}

function submitPump() {
  const time = to12(document.getElementById('pTime').value);
  const amount = document.getElementById('pAmount').value.trim();
  const notes = document.getElementById('pNotes').value.trim();
  let details = 'Pump';
  if (amount) details += ' ' + amount;
  if (notes) details += '; ' + notes;

  const entry = { date: todayStr(), time, type: 'Pump', side: '', shield: '', amount, details };

  if (editingIdx !== null) {
    replaceEntry(editingIdx, entry);
    editingIdx = null;
  } else {
    addEntry(entry);
  }
  closeModal('pumpModal');
}

function submitNote() {
  const time = to12(document.getElementById('nTime').value);
  const text = document.getElementById('nText').value.trim();
  if (!text) { toast('Enter a note'); return; }

  const entry = { date: todayStr(), time, type: 'Note', side: '', shield: '', amount: '', details: text };

  if (editingIdx !== null) {
    replaceEntry(editingIdx, entry);
    editingIdx = null;
  } else {
    addEntry(entry);
  }
  closeModal('noteModal');
}

function submitDiaper() {
  const time = to12(document.getElementById('dTime').value);
  const notes = document.getElementById('dNotes').value.trim();
  const dType = diaperPick || 'Diaper change';
  let details = dType;
  if (notes) details += ' (' + notes + ')';

  const entry = { date: todayStr(), time, type: 'Diaper', side: '', shield: '', amount: '', details };

  if (editingIdx !== null) {
    replaceEntry(editingIdx, entry);
    editingIdx = null;
  } else {
    addEntry(entry);
  }
  closeModal('diaperModal');
}

// Fix #5: Replace an entry (delete old + add new)
async function replaceEntry(idx, newEntry) {
  const oldEntry = allEntries[idx];

  // Remove old entry locally
  allEntries.splice(idx, 1);

  // Delete old entry from sheet if it was synced
  if (scriptUrl && oldEntry && oldEntry._synced) {
    await apiCall({
      action: 'delete',
      date: oldEntry.date || '',
      time: oldEntry.time || '',
      type: oldEntry.type || '',
      side: oldEntry.side || '',
      shield: oldEntry.shield || '',
      amount: oldEntry.amount || '',
      details: oldEntry.details || ''
    });
  }

  // Add new entry
  await addEntry(newEntry);
  toast('\u2705 Entry updated!');
}

function quickDiaper(details) {
  addEntry({ date: todayStr(), time: to12(now24()), type: 'Diaper', side: '', shield: '', amount: '', details });
}

// ============ VOICE INPUT ============
let recognition = null;
let voiceState = 'idle';
let parsedEntry = null;

function hasVoiceSupport() {
  return !!(window.SpeechRecognition || window.webkitSpeechRecognition);
}

function startVoice() {
  if (!hasVoiceSupport()) {
    toast('Voice not supported in this browser. Use Chrome or Safari.');
    return;
  }

  voiceState = 'listening';
  parsedEntry = null;

  const overlay = document.getElementById('voiceOverlay');
  const mic = document.getElementById('voiceMic');
  const status = document.getElementById('voiceStatus');
  const transcript = document.getElementById('voiceTranscript');
  const parsed = document.getElementById('voiceParsed');
  const btns = document.getElementById('voiceBtns');

  overlay.classList.add('show');
  mic.classList.add('listening');
  status.textContent = 'Listening...';
  transcript.textContent = '';
  parsed.style.display = 'none';
  btns.innerHTML = '<button class="voice-btn cancel" data-action="cancel-voice">Cancel</button>';

  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  recognition = new SpeechRecognition();
  recognition.continuous = false;
  recognition.interimResults = true;
  recognition.lang = 'en-US';

  recognition.onresult = function (event) {
    let interim = '';
    let final = '';
    for (let i = event.resultIndex; i < event.results.length; i++) {
      if (event.results[i].isFinal) final += event.results[i][0].transcript;
      else interim += event.results[i][0].transcript;
    }
    transcript.textContent = final || interim;
    if (final) handleVoiceResult(final.trim());
  };

  recognition.onerror = function (event) {
    mic.classList.remove('listening');
    if (event.error === 'no-speech') {
      status.textContent = 'No speech detected. Try again?';
      btns.innerHTML =
        '<button class="voice-btn cancel" data-action="cancel-voice">Cancel</button>' +
        '<button class="voice-btn retry" data-action="retry-voice">Retry</button>';
    } else if (event.error === 'not-allowed') {
      status.textContent = 'Microphone access denied.';
      btns.innerHTML = '<button class="voice-btn cancel" data-action="cancel-voice">Close</button>';
    } else {
      status.textContent = 'Error: ' + event.error;
      btns.innerHTML =
        '<button class="voice-btn cancel" data-action="cancel-voice">Cancel</button>' +
        '<button class="voice-btn retry" data-action="retry-voice">Retry</button>';
    }
  };

  recognition.onend = function () {
    mic.classList.remove('listening');
    if (voiceState === 'listening' && !parsedEntry) {
      if (!transcript.textContent) {
        status.textContent = 'No speech detected. Try again?';
        btns.innerHTML =
          '<button class="voice-btn cancel" data-action="cancel-voice">Cancel</button>' +
          '<button class="voice-btn retry" data-action="retry-voice">Retry</button>';
      }
    }
  };

  recognition.start();
}

function cancelVoice() {
  voiceState = 'idle';
  if (recognition) { try { recognition.stop(); } catch (e) { } }
  recognition = null;
  document.getElementById('voiceOverlay').classList.remove('show');
  document.getElementById('voiceMic').classList.remove('listening');
}

function handleVoiceResult(text) {
  voiceState = 'parsed';
  const mic = document.getElementById('voiceMic');
  const status = document.getElementById('voiceStatus');
  const parsed = document.getElementById('voiceParsed');
  const btns = document.getElementById('voiceBtns');

  mic.classList.remove('listening');
  parsedEntry = parseVoice(text);

  if (!parsedEntry) {
    status.textContent = "Couldn't parse that. Try again?";
    btns.innerHTML =
      '<button class="voice-btn cancel" data-action="cancel-voice">Cancel</button>' +
      '<button class="voice-btn retry" data-action="retry-voice">Retry</button>';
    return;
  }

  const emoji = EMOJIS[parsedEntry.type] || '\u{1F4DD}';
  status.textContent = emoji + ' ' + parsedEntry.type + ' detected';

  document.getElementById('vpType').textContent = emoji + ' ' + parsedEntry.summary;
  document.getElementById('vpDetails').textContent = parsedEntry.details;
  parsed.style.display = 'block';

  btns.innerHTML =
    '<button class="voice-btn cancel" data-action="cancel-voice">Cancel</button>' +
    '<button class="voice-btn retry" data-action="retry-voice">Redo</button>' +
    '<button class="voice-btn confirm" data-action="confirm-voice">Log It</button>';
}

function confirmVoice() {
  if (!parsedEntry) return;
  addEntry(parsedEntry.entry);
  cancelVoice();
  toast(EMOJIS[parsedEntry.entry.type] + ' Logged via voice!');
}

// ============ VOICE PARSER ============
function parseVoice(text) {
  const t = text.toLowerCase().trim();

  // ---- DIAPER ----
  if (matchesAny(t, ['diaper', 'pee', 'poo', 'poop', 'wet', 'dirty', 'meconium', 'shart', 'poopy', 'pooped', 'peed'])) {
    const hasPee = matchesAny(t, ['pee', 'wet', 'peed']);
    const hasPoo = matchesAny(t, ['poo', 'poop', 'dirty', 'meconium', 'shart', 'poopy', 'pooped']);
    const isLive = matchesAny(t, ['live']);
    let diaperType = '';

    if ((hasPee && hasPoo) || matchesAny(t, ['both'])) diaperType = 'Pee & poo diaper';
    else if (hasPoo) diaperType = 'Poo diaper';
    else if (hasPee) diaperType = 'Pee diaper';
    else diaperType = 'Diaper change';

    let extraNotes = [];
    ['runny', 'green', 'yellow', 'brown', 'black', 'watery', 'seedy', 'mucus', 'sticky', 'dark', 'light', 'big', 'small', 'little', 'huge', 'massive']
      .forEach(d => { if (t.includes(d)) extraNotes.push(d); });

    if (isLive) diaperType += ' (live)';
    let details = diaperType;
    if (extraNotes.length) details += ' (' + extraNotes.join(', ') + ')';

    return {
      type: 'Diaper', summary: diaperType, details,
      entry: { date: todayStr(), time: to12(now24()), type: 'Diaper', side: '', shield: '', amount: '', details }
    };
  }

  // ---- PUMP ----
  if (matchesAny(t, ['pump', 'pumped', 'pumping', 'expressed'])) {
    let amount = '';
    let notes = [];
    const mlMatch = t.match(/(\d+)\s*(?:ml|milliliters?)/i);
    const ozMatch = t.match(/(\d+)\s*(?:oz|ounces?)/i);
    if (mlMatch) amount = mlMatch[1] + 'ml';
    else if (ozMatch) amount = ozMatch[1] + 'oz';

    const durMatch = t.match(/(\d+)\s*(?:min|minutes?)/i);
    if (durMatch) notes.push(durMatch[1] + ' min');

    if (t.includes('mostly left')) notes.push('mostly left');
    else if (t.includes('mostly right')) notes.push('mostly right');
    else if (t.includes('left side') || (t.includes('left') && !t.includes('right'))) notes.push('left');
    else if (t.includes('right side') || (t.includes('right') && !t.includes('left'))) notes.push('right');

    let details = 'Pump';
    if (amount) details += ' ' + amount;
    if (notes.length) details += '; ' + notes.join(', ');
    let summary = 'Pump';
    if (amount) summary += ' \u2014 ' + amount;
    if (notes.length) summary += ', ' + notes.join(', ');

    return {
      type: 'Pump', summary, details,
      entry: { date: todayStr(), time: to12(now24()), type: 'Pump', side: '', shield: '', amount, details }
    };
  }

  // ---- FEED ----
  if (matchesAny(t, ['fed', 'feed', 'feeding', 'nursed', 'nursing', 'breastfed', 'ate', 'eating', 'breast', 'bottle', 'latch', 'latched'])) {
    let side = '', shield = '', amount = '';
    let notes = [];

    if (matchesAny(t, ['left right left', 'l r l', 'lrl'])) side = 'LRL';
    else if (matchesAny(t, ['right left right', 'r l r', 'rlr'])) side = 'RLR';
    else if (matchesAny(t, ['left then right', 'left to right', 'left right', 'l then r', 'left and right'])) side = 'LR';
    else if (matchesAny(t, ['right then left', 'right to left', 'right left', 'r then l', 'right and left'])) side = 'RL';
    else if (t.includes('both sides')) side = 'LR';
    else if (matchesAny(t, ['just left', 'only left', 'left side', 'on the left', 'on left'])) side = 'L';
    else if (matchesAny(t, ['just right', 'only right', 'right side', 'on the right', 'on right'])) side = 'R';
    else {
      const hasL = /\bleft\b/.test(t), hasR = /\bright\b/.test(t);
      if (hasL && hasR) { side = t.indexOf('left') < t.indexOf('right') ? 'LR' : 'RL'; }
      else if (hasL) side = 'L';
      else if (hasR) side = 'R';
    }

    if (matchesAny(t, ['no shield', 'without shield', 'shield off', 'no nipple shield'])) shield = 'None';
    else if (matchesAny(t, ['shield on both', 'shield both', 'both shields'])) shield = 'Both';
    else if (matchesAny(t, ['mixed shield', 'some shield', 'sometimes shield'])) shield = 'Mixed';
    else if (matchesAny(t, ['shield on left', 'shield left', 'left shield'])) shield = 'L';
    else if (matchesAny(t, ['shield on right', 'shield right', 'right shield'])) shield = 'R';
    else if (matchesAny(t, ['with shield', 'shield on', 'using shield', 'nipple shield'])) shield = 'Both';

    const mlMatch = t.match(/(\d+)\s*(?:ml|milliliters?)/i);
    const ozMatch = t.match(/(\d+)\s*(?:oz|ounces?)/i);
    if (mlMatch) amount = mlMatch[1] + 'ml';
    else if (ozMatch) amount = ozMatch[1] + 'oz';

    const noteWords = {
      'lots of milk': 'lots of milk', 'tons of milk': 'tons of milk',
      'not much milk': 'not much milk', 'medium milk': 'medium milk',
      'good latch': 'good latch', 'bad latch': 'bad latch', 'poor latch': 'poor latch',
      'very sleepy': 'very sleepy', 'super sleepy': 'super sleepy', 'sleepy': 'sleepy',
      'fussy': 'fussy', 'cluster': 'cluster feeding', 'quick snack': 'quick snack',
      'fell asleep': 'fell asleep', 'drowsy': 'drowsy',
      'formula': 'formula', 'bottle': 'bottle'
    };
    Object.keys(noteWords).forEach(key => { if (t.includes(key)) notes.push(noteWords[key]); });

    let details = 'Feed';
    if (side) details += ' on ' + side;
    if (shield && shield !== 'None') details += ' (shield: ' + shield + ')';
    else if (shield === 'None') details += ' (no shield)';
    if (amount) details += '; ' + amount;
    if (notes.length) details += '; ' + notes.join(', ');

    const sideDisplay = { L: 'Left', R: 'Right', LR: 'L\u2192R', RL: 'R\u2192L', LRL: 'L\u2192R\u2192L', RLR: 'R\u2192L\u2192R' };
    let summary = 'Feed';
    if (side) summary += ' \u2014 ' + (sideDisplay[side] || side);
    if (shield && shield !== 'None') summary += ', shield: ' + shield;
    else if (shield === 'None') summary += ', no shield';
    if (notes.length) summary += ', ' + notes.join(', ');

    return {
      type: 'Feed', summary, details,
      entry: { date: todayStr(), time: to12(now24()), type: 'Feed', side, shield, amount, details }
    };
  }

  // ---- NOTE ----
  if (t.length > 2) {
    return {
      type: 'Note', summary: 'Note', details: text,
      entry: { date: todayStr(), time: to12(now24()), type: 'Note', side: '', shield: '', amount: '', details: text }
    };
  }

  return null;
}

function matchesAny(text, phrases) {
  return phrases.some(p => text.includes(p));
}

// ============ RENDERING (Fix #1: escapeHtml everywhere) ============
function renderAll() { renderStats(); renderToday(); renderHistory(); }

function renderStats() {
  const today = todayStr();
  const te = allEntries.filter(e => normDate(e.date) === today);
  document.getElementById('statFeeds').textContent = te.filter(e => e.type === 'Feed').length;
  document.getElementById('statDiapers').textContent = te.filter(e => e.type === 'Diaper').length;
  document.getElementById('statPumps').textContent = te.filter(e => e.type === 'Pump').length;

  // Fix #2: Show elapsed time since last feed
  const feeds = allEntries.filter(e => e.type === 'Feed');
  if (feeds.length) {
    const lastFeed = feeds[feeds.length - 1];
    document.getElementById('statLastFeed').textContent = getElapsedStr(lastFeed.time, lastFeed.date);
  } else {
    document.getElementById('statLastFeed').textContent = '-';
  }
}

// Fix #1: All user content is escaped
function renderEntry(e, idx) {
  const emoji = EMOJIS[e.type] || '\u{1F4DD}';
  const syncIcon = e._synced === false ? ' <span class="pending-dot" title="Pending sync">\u25CF</span>' : '';
  return '<div class="entry-wrapper">' +
    '<div class="entry ' + escapeHtml(e.type || '') + '" data-action="edit-entry" data-idx="' + idx + '">' +
    '<span class="e-emoji">' + emoji + '</span>' +
    '<span class="e-time">' + escapeHtml(e.time || '') + '</span>' +
    '<span class="e-detail">' + escapeHtml(e.details || '') + syncIcon + '</span>' +
    '<button class="e-delete-btn" data-action="delete-entry" data-idx="' + idx + '" title="Delete">\u00d7</button>' +
    '</div></div>';
}

function renderToday() {
  const today = todayStr();
  let entries = allEntries.map((e, i) => ({ ...e, _idx: i })).filter(e => normDate(e.date) === today);
  if (currentFilter !== 'all') entries = entries.filter(e => e.type === currentFilter);
  const c = document.getElementById('todayEntries');
  if (!entries.length) { c.innerHTML = '<div class="empty">No entries yet today.<br>Use the buttons below to start logging!</div>'; return; }
  c.innerHTML = entries.slice().reverse().map(e => renderEntry(e, e._idx)).join('');
}

function renderHistory() {
  const grouped = {};
  allEntries.forEach((e, i) => { const d = normDate(e.date) || 'Unknown'; if (!grouped[d]) grouped[d] = []; grouped[d].push({ ...e, _idx: i }); });
  const dates = Object.keys(grouped).sort().reverse();
  const c = document.getElementById('historyEntries');
  if (!dates.length) { c.innerHTML = '<div class="empty">No history yet.</div>'; return; }
  c.innerHTML = dates.map(date => {
    const entries = grouped[date];
    const feeds = entries.filter(e => e.type === 'Feed').length;
    const diapers = entries.filter(e => e.type === 'Diaper').length;
    const wet = entries.filter(e => e.type === 'Diaper' && /pee|wet/i.test(e.details || '')).length;
    const dirty = entries.filter(e => e.type === 'Diaper' && /poo|dirty|meconium/i.test(e.details || '')).length;
    const pumps = entries.filter(e => e.type === 'Pump').length;
    const isToday = date === todayStr();
    return '<div class="day-group">' +
      '<div class="day-header" data-action="toggle-day">' +
      '<h3>' + escapeHtml(formatDateNice(date)) + '</h3>' +
      '<span class="day-meta">' + entries.length + ' entries \u25BE</span></div>' +
      '<div class="day-entries" style="' + (isToday ? '' : 'display:none') + '">' +
      '<div class="day-summary">' +
      '<span class="day-badge feed">\u{1F37C} ' + feeds + ' feeds</span>' +
      '<span class="day-badge diaper">\u{1F9F7} ' + diapers + ' (' + wet + 'W/' + dirty + 'D)</span>' +
      (pumps ? '<span class="day-badge pump">\u{1F95B} ' + pumps + ' pumps</span>' : '') +
      '</div>' +
      entries.map(e => renderEntry(e, e._idx)).join('') +
      '</div></div>';
  }).join('');
}

// ============ SETUP ============
async function saveSetup() {
  scriptUrl = document.getElementById('scriptUrl').value.trim();
  localStorage.setItem('elliot_script_url', scriptUrl);
  const status = document.getElementById('setupStatus');
  status.innerHTML = '<div class="status-msg loading">Testing...</div>';
  const data = await apiCall({ action: 'read' });
  if (data && data.success) {
    const sheetEntries = (data.entries || []).map(e => ({ ...e, _synced: true, date: normDate(e.date), time: normTime(e.time) }));
    const unsynced = allEntries.filter(e => e._synced === false);
    allEntries = [...sheetEntries, ...unsynced];
    saveLocal(); renderAll();
    status.innerHTML = '<div class="status-msg ok">\u2705 Connected! Loaded ' + allEntries.length + ' entries.</div>';
  } else {
    status.innerHTML = '<div class="status-msg err">\u274c Could not connect. Check URL and deployment.</div>';
  }
}

// ============ URL PARAMETER AUTO-ADD ============
function handleUrlParams() {
  const params = new URLSearchParams(window.location.search);
  const type = params.get('type');
  if (!type) return;
  const time = params.get('time') || to12(now24());
  const side = params.get('side') || '';
  const shield = params.get('shield') || '';
  const amount = params.get('amount') || '';
  const notes = params.get('notes') || '';
  const date = params.get('date') || todayStr();
  let details = type;
  if (type === 'Feed') {
    details = 'Feed';
    if (side) details += ' on ' + side;
    if (shield && shield !== 'None') details += ' (shield: ' + shield + ')';
    else if (shield === 'None') details += ' (no shield)';
    if (amount) details += '; ' + amount;
    if (notes) details += '; ' + notes;
  } else if (type === 'Pump') {
    details = 'Pump';
    if (amount) details += ' ' + amount;
    if (notes) details += '; ' + notes;
  } else if (type === 'Diaper') {
    details = notes || 'Diaper change';
  } else if (type === 'Note') {
    details = notes || 'Note';
  }
  const entry = { date, time, type, side, shield, amount, details, _synced: false };
  allEntries.push(entry);
  saveLocal(); renderAll();
  const banner = document.getElementById('autoAddBanner');
  const emoji = EMOJIS[type] || '\u{1F4DD}';
  banner.innerHTML = '<div class="auto-add-banner"><h3>' + emoji + ' Entry Added!</h3><p>' + escapeHtml(time) + ' \u2014 ' + escapeHtml(details) + '</p></div>';
  setTimeout(() => { banner.innerHTML = ''; }, 5000);
  apiCall({ action: 'add', ...entry }).then(r => {
    if (r && r.success) { entry._synced = true; saveLocal(); toast('\u2705 Saved to sheet!'); }
    else toast('\u{1F4F1} Saved locally');
  });
  window.history.replaceState({}, '', window.location.pathname);
}

// ============ EVENT DELEGATION (Fix #8) ============
document.addEventListener('click', function (e) {
  const target = e.target.closest('[data-action]');
  if (!target) return;
  const action = target.dataset.action;

  switch (action) {
    // Bottom bar
    case 'quick-diaper':
      quickDiaper(target.dataset.diaper);
      break;
    case 'show-feed-form':
      showFeedForm();
      break;
    case 'show-pump-form':
      showPumpForm();
      break;
    case 'start-voice':
      startVoice();
      break;
    case 'show-more-menu':
      showMoreMenu();
      break;

    // Tabs
    case 'switch-tab':
      switchTab(target);
      break;

    // Filters
    case 'set-filter':
      setFilter(target.dataset.filter, target);
      break;

    // Setup
    case 'show-setup':
      showSetup();
      break;
    case 'save-setup':
      saveSetup();
      break;

    // Modal close
    case 'close-modal':
      closeModal(target.dataset.modal);
      break;

    // Pill selection
    case 'pick-pill':
      pickPill(target, target.dataset.group);
      break;
    case 'pick-diaper-pill':
      target.parentElement.querySelectorAll('.pill').forEach(p => p.classList.remove('selected'));
      target.classList.add('selected');
      diaperPick = target.dataset.value;
      break;

    // Submit buttons
    case 'submit-feed':
      submitFeed();
      break;
    case 'submit-pump':
      submitPump();
      break;
    case 'submit-note':
      submitNote();
      break;
    case 'submit-diaper':
      submitDiaper();
      break;

    // Entries
    case 'edit-entry':
      e.preventDefault();
      editEntry(parseInt(target.dataset.idx));
      break;
    case 'delete-entry':
      e.stopPropagation();
      confirmDelete(parseInt(target.dataset.idx));
      break;

    // Delete confirm
    case 'cancel-delete':
      cancelDelete();
      break;
    case 'do-delete':
      doDelete(parseInt(target.dataset.idx));
      break;

    // More modal actions
    case 'quick-diaper-close':
      quickDiaper(target.dataset.diaper);
      closeModal('moreModal');
      break;
    case 'show-note-form':
      closeModal('moreModal');
      showNoteForm();
      break;
    case 'refresh-data':
      refreshData();
      break;

    // Voice actions
    case 'cancel-voice':
      cancelVoice();
      break;
    case 'retry-voice':
      cancelVoice();
      startVoice();
      break;
    case 'confirm-voice':
      confirmVoice();
      break;

    // Day group toggle (history)
    case 'toggle-day':
      const dayEntries = target.closest('.day-group').querySelector('.day-entries');
      if (dayEntries) dayEntries.style.display = dayEntries.style.display === 'none' ? 'block' : 'none';
      break;
  }
});

// ============ INIT ============
loadLocal();
renderAll();
updateClock();
setInterval(updateClock, 30000);
if (scriptUrl) { setSyncState('connected'); setTimeout(refreshData, 1000); } else setSyncState('offline');
setInterval(() => { if (scriptUrl) refreshData(); }, 120000);
handleUrlParams();

// Disable voice button if not supported
if (!hasVoiceSupport()) {
  const vb = document.getElementById('voiceBtn');
  if (vb) { vb.style.opacity = '0.4'; vb.dataset.action = ''; vb.addEventListener('click', () => toast('Voice not supported. Use Chrome or Safari.')); }
}

// Fix #6: Register service worker for PWA/offline support
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('./sw.js').catch(() => {});
}
