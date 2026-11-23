// Colin transcript collector: plain helpers shared by the background worker, the page script and
// the tests. No browser APIs here, so every rule can be checked in Node.
//
// The collector reads Colin/wanted.json from the user's own OneDrive (Colin writes it), saves each
// wanted recording's transcript to Colin/Transcripts/<id>.txt, and moves a transcript to the
// OneDrive recycle bin once Colin marks it done. It never reads, copies or exports cookies.
(function (root) {
  'use strict';
  const core = {};

  core.MAX_PER_RUN = 5;
  core.MAX_ATTEMPTS = 3;
  core.FETCH_TIMEOUT_MS = 90000;

  const ID = /^[a-z0-9][a-z0-9-]{5,63}$/;

  // Only SharePoint recordings (a Teams recording lives in the organiser's OneDrive or a site).
  core.allowedUrl = function (url) {
    try {
      const u = new URL(url);
      return u.protocol === 'https:' && (/\.sharepoint\.com$/i.test(u.hostname) || /\.sharepoint\.com\.mcas\.ms$/i.test(u.hostname));
    } catch (_) { return false; }
  };

  // The user's OneDrive address, e.g. https://contoso-my.sharepoint.com/personal/ann_contoso_com
  core.homeFrom = function (input) {
    try {
      const u = new URL(String(input).trim());
      const m = u.pathname.match(/^\/personal\/[^/]+/);
      if (u.protocol !== 'https:' || !/-my\.sharepoint\.com$/i.test(u.hostname) || !m) return null;
      return u.origin + m[0];
    } catch (_) { return null; }
  };

  // wanted.json, as Colin writes it. Anything malformed is dropped, never guessed at.
  core.parseWanted = function (text) {
    let data;
    try { data = JSON.parse(text); } catch (_) { return { ok: false, error: 'wanted.json is not valid JSON', items: [] }; }
    if (!data || data.version !== 1 || !Array.isArray(data.wanted)) return { ok: false, error: 'wanted.json has no version 1 list', items: [] };
    const seen = new Set();
    const items = [];
    for (const w of data.wanted) {
      if (!w || typeof w.id !== 'string' || !ID.test(w.id) || seen.has(w.id)) continue;
      if (typeof w.url !== 'string' || !core.allowedUrl(w.url)) continue;
      seen.add(w.id);
      items.push({ id: w.id, url: w.url, subject: typeof w.subject === 'string' ? w.subject.slice(0, 200) : '', done: w.done === true });
    }
    return { ok: true, items };
  };

  core.fileName = (id) => id + '.txt';

  // What this run does: transcripts still wanted that are not saved yet (resting any that failed
  // MAX_ATTEMPTS times until a day after the last try, so signing in later still brings them), and
  // saved transcripts Colin has finished with. `attempts`: { id: { n, at } }.
  core.plan = function (items, savedNames, attempts, nowMs) {
    const saved = new Set(savedNames);
    const now = nowMs || Date.now();
    const resting = (id) => { const a = attempts[id]; return a && a.n >= core.MAX_ATTEMPTS && now - a.at < 86400000; };
    const fetch = items
      .filter((w) => !w.done && !saved.has(core.fileName(w.id)) && !resting(w.id))
      .slice(0, core.MAX_PER_RUN);
    const recycle = items.filter((w) => w.done && saved.has(core.fileName(w.id)));
    return { fetch, recycle };
  };

  function seconds(t) {
    const parts = String(t || '0:0:0').split(':');
    const h = parseInt(parts[0], 10) || 0, m = parseInt(parts[1], 10) || 0, s = parseFloat(parts[2]) || 0;
    return h * 3600 + m * 60 + s;
  }
  function clock(sec) {
    const f = Math.floor(sec);
    return [Math.floor(f / 3600), Math.floor((f % 3600) / 60), f % 60].map((n) => String(n).padStart(2, '0')).join(':');
  }

  // The Teams transcript JSON as plain text Colin reads: one paragraph per turn, with the time the
  // speaker started. "[00:04:12] Ann Lee: ..."
  core.toText = function (jsonText, header) {
    const entries = (JSON.parse(jsonText).entries || []).filter((e) => e && e.text);
    const out = [];
    let who = null, buf = '', at = 0;
    for (const e of entries) {
      const speaker = e.speakerDisplayName || 'Unknown';
      if (speaker !== who) {
        if (buf) out.push(`[${clock(at)}] ${who}: ${buf.trim()}`);
        who = speaker; buf = e.text; at = seconds(e.startOffset);
      } else buf += ' ' + e.text;
    }
    if (buf) out.push(`[${clock(at)}] ${who}: ${buf.trim()}`);
    return (header ? header + '\n\n' : '') + out.join('\n\n') + '\n';
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = core;
  else root.ColinCore = core;
})(typeof self !== 'undefined' ? self : this);
