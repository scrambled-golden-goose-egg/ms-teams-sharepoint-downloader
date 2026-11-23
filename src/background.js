// Colin transcript collector: the timer and the tabs. Once an hour, while the browser is open,
// it opens the user's OneDrive and each wanted recording in a minimised window, lets the page
// script do the reading as the signed-in user, and closes the window again.
//
// It never touches cookies or tokens: every SharePoint call happens inside a SharePoint page, the
// same way the downloader's own buttons work (see CLAUDE.md, Don'ts).
importScripts('colin-core.js');
const core = self.ColinCore;

const ALARM = 'colin-collector';
const roles = new Map();       // tabId -> { role: 'home' } | { role: 'fetch', id, subject }
const waiters = new Map();     // tabId -> resolve(message)

function schedule() {
  chrome.alarms.get(ALARM, (a) => { if (!a) chrome.alarms.create(ALARM, { delayInMinutes: 2, periodInMinutes: 60 }); });
}
chrome.runtime.onInstalled.addListener(schedule);
chrome.runtime.onStartup.addListener(schedule);
chrome.alarms.onAlarm.addListener((a) => { if (a.name === ALARM) run('timer'); });

// A page may ask for its role a moment before the worker has written it down, so wait up to 3
// seconds before saying it has none. Every other SharePoint page the user browses gets 'none'.
function roleFor(tabId) {
  return new Promise((resolve) => {
    const start = Date.now();
    const look = () => {
      if (roles.has(tabId)) return resolve(roles.get(tabId));
      if (!running || Date.now() - start > 3000) return resolve({ role: 'none' });
      setTimeout(look, 100);
    };
    look();
  });
}

// And once a collector tab has loaded, tell it its role as well, in case its own ask went astray.
chrome.tabs.onUpdated.addListener((tabId, info) => {
  if (info.status === 'complete' && roles.has(tabId)) chrome.tabs.sendMessage(tabId, Object.assign({ type: 'colin-start' }, roles.get(tabId)), () => void chrome.runtime.lastError);
});

chrome.runtime.onMessage.addListener((msg, sender, reply) => {
  if (!msg || typeof msg.type !== 'string' || !msg.type.startsWith('colin-')) return false;
  const tabId = sender.tab && sender.tab.id;
  if (msg.type === 'colin-role') { roleFor(tabId).then(reply); return true; }
  if (msg.type === 'colin-run-now') { run('button').then(reply); return true; }
  if (tabId != null && waiters.has(tabId)) { waiters.get(tabId)(msg); waiters.delete(tabId); }
  return false;
});

const get = (keys) => new Promise((r) => chrome.storage.local.get(keys, r));
const set = (obj) => new Promise((r) => chrome.storage.local.set(obj, r));

// Wait for one message from a tab (its page script reports in), or give up.
function waitFor(tabId, ms) {
  return new Promise((resolve) => {
    const t = setTimeout(() => { waiters.delete(tabId); resolve({ type: 'colin-timeout' }); }, ms);
    waiters.set(tabId, (m) => { clearTimeout(t); resolve(m); });
  });
}
function ask(tabId, msg, ms) {
  return new Promise((resolve) => {
    const t = setTimeout(() => resolve({ ok: false, error: 'No answer from the OneDrive page' }), ms);
    chrome.tabs.sendMessage(tabId, msg, (r) => { clearTimeout(t); resolve(r || { ok: false, error: (chrome.runtime.lastError || {}).message || 'No answer' }); });
  });
}

// Chrome stops an idle service worker after about 30 seconds; a run waits longer than that for a
// recording page, so it keeps itself awake while it works.
function keepAwake() {
  const t = setInterval(() => chrome.runtime.getPlatformInfo(() => {}), 20000);
  return () => clearInterval(t);
}

let running = false;
async function run(why) {
  const awake = keepAwake();
  try { return await runOnce(why); } finally { awake(); }
}

async function runOnce(why) {
  const { colinEnabled, colinHome, colinAttempts = {} } = await get(['colinEnabled', 'colinHome', 'colinAttempts']);
  const home = core.homeFrom(colinHome || '');
  if (!colinEnabled || !home) return { ok: false, summary: 'Off, or no OneDrive address set' };
  if (running) return { ok: false, summary: 'Already running' };
  running = true;
  const results = {};
  let win = null;
  try {
    win = await chrome.windows.create({ url: home + '/_layouts/15/onedrive.aspx', state: 'minimized', focused: false });
    const homeTab = win.tabs[0].id;
    const homeReady = waitFor(homeTab, 60000);
    roles.set(homeTab, { role: 'home' });
    const ready = await homeReady;
    if (ready.type !== 'colin-home-ready') throw new Error('OneDrive did not open (signed out?)');

    const state = await ask(homeTab, { type: 'colin-home-read' }, 30000);
    if (!state.ok) throw new Error(state.error || 'Could not read Colin/wanted.json');
    const wanted = core.parseWanted(state.wantedText);
    if (!wanted.ok) throw new Error(wanted.error);
    const plan = core.plan(wanted.items, state.saved, colinAttempts);

    for (const w of plan.recycle) {
      const r = await ask(homeTab, { type: 'colin-home-recycle', name: core.fileName(w.id) }, 30000);
      results[w.id] = r.ok ? { status: 'recycled' } : { status: 'error', error: r.error };
    }

    for (const w of plan.fetch) {
      const tab = await chrome.tabs.create({ windowId: win.id, url: w.url, active: false });
      const pending = waitFor(tab.id, core.FETCH_TIMEOUT_MS);
      roles.set(tab.id, { role: 'fetch', id: w.id, subject: w.subject });
      const got = await pending;
      chrome.tabs.remove(tab.id).catch(() => {});
      roles.delete(tab.id);
      if (got.type === 'colin-transcript' && got.text) {
        const up = await ask(homeTab, { type: 'colin-home-save', name: core.fileName(w.id), text: got.text }, 30000);
        results[w.id] = up.ok ? { status: 'saved' } : { status: 'error', error: up.error };
        if (up.ok) delete colinAttempts[w.id];
      } else {
        const prev = colinAttempts[w.id];
        colinAttempts[w.id] = { n: (prev && prev.n >= core.MAX_ATTEMPTS ? 0 : (prev ? prev.n : 0)) + 1, at: Date.now() };
        results[w.id] = { status: got.type === 'colin-timeout' ? 'timeout' : (got.status || 'error'), error: got.error || 'The recording page did not give a transcript (signed out of that organisation?)' };
      }
    }

    await ask(homeTab, { type: 'colin-home-status', status: { version: 1, ran_at: new Date().toISOString(), why, results } }, 30000);
    roles.delete(homeTab);
    const n = (s) => Object.values(results).filter((r) => r.status === s).length;
    const failed = Object.keys(results).length - n('saved') - n('recycled') - n('none');
    const summary = `Saved ${n('saved')}, recycled ${n('recycled')}, no transcript yet ${n('none')}, failed ${failed}`;
    await set({ colinAttempts, colinLast: { at: new Date().toISOString(), summary } });
    return { ok: true, summary };
  } catch (e) {
    const summary = 'Not run: ' + (e && e.message ? e.message : String(e));
    await set({ colinLast: { at: new Date().toISOString(), summary } });
    return { ok: false, summary };
  } finally {
    running = false;
    if (win) chrome.windows.remove(win.id).catch(() => {});
  }
}
