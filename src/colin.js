// Colin transcript collector: the page script. It does nothing unless the background worker opened
// this tab for the collector. Two roles:
//  - home: a page on the user's own OneDrive. Reads Colin/wanted.json, lists Colin/Transcripts,
//    saves transcripts there, moves finished ones to the recycle bin, writes Colin/collector.json.
//  - fetch: a recording's page. Reads its transcript the way the downloader's button does
//    (drive and item from the page, SharePoint's own transcript API, the signed-in session).
// Everything runs as the signed-in user, inside SharePoint's own pages; no cookie or token leaves.
(function () {
  'use strict';
  if (window.top !== window) return;
  const core = self.ColinCore;

  chrome.runtime.sendMessage({ type: 'colin-role' }, (r) => {
    if (chrome.runtime.lastError || !r) return;
    if (r.role === 'home') home();
    if (r.role === 'fetch') fetchTranscript(r);
  });

  // ---------- home: the user's own OneDrive ----------
  function home() {
    const base = location.origin + (location.pathname.match(/^\/personal\/[^/]+/) || [''])[0];
    const folder = (base.replace(location.origin, '') + '/Documents/Colin');
    const api = (p) => base + '/_api/web/' + p;
    const json = { Accept: 'application/json;odata=nometadata' };
    let digest = null;

    async function getDigest() {
      if (digest) return digest;
      const r = await fetch(base + '/_api/contextinfo', { method: 'POST', credentials: 'include', headers: json });
      if (!r.ok) throw new Error('SharePoint refused a write token (' + r.status + ')');
      digest = (await r.json()).FormDigestValue;
      return digest;
    }
    async function call(url, init) {
      const r = await fetch(url, Object.assign({ credentials: 'include' }, init));
      if (!r.ok) throw new Error(`${r.status} ${r.statusText}`.trim());
      return r;
    }
    const fileApi = (path) => api(`GetFileByServerRelativePath(decodedurl='${encodeURIComponent(path)}')`);
    const folderApi = (path) => api(`GetFolderByServerRelativePath(decodedurl='${encodeURIComponent(path)}')`);
    async function save(dir, name, text) {
      await call(`${folderApi(dir)}/Files/AddUsingPath(decodedurl='${encodeURIComponent(name)}',overwrite=true)`, {
        method: 'POST', headers: Object.assign({ 'X-RequestDigest': await getDigest() }, json), body: text,
      });
    }

    chrome.runtime.onMessage.addListener((msg, _s, reply) => {
      const done = (p) => p.then((v) => reply(Object.assign({ ok: true }, v || {}))).catch((e) => reply({ ok: false, error: String(e.message || e) }));
      if (msg.type === 'colin-home-read') {
        done((async () => {
          const wantedText = await (await call(`${fileApi(folder + '/wanted.json')}/$value`, {})).text();
          const files = await (await call(`${folderApi(folder + '/Transcripts')}/Files?$select=Name`, { headers: json })).json();
          return { wantedText, saved: (files.value || []).map((f) => f.Name) };
        })());
        return true;
      }
      if (msg.type === 'colin-home-save') { done(save(folder + '/Transcripts', msg.name, msg.text)); return true; }
      if (msg.type === 'colin-home-recycle') {
        // The recycle bin, not a delete: the file can be restored for as long as OneDrive keeps it.
        done((async () => { await call(`${fileApi(folder + '/Transcripts/' + msg.name)}/recycle()`, { method: 'POST', headers: Object.assign({ 'X-RequestDigest': await getDigest() }, json) }); })());
        return true;
      }
      if (msg.type === 'colin-home-status') { done(save(folder, 'collector.json', JSON.stringify(msg.status, null, 1))); return true; }
      return false;
    });
    chrome.runtime.sendMessage({ type: 'colin-home-ready' });
  }

  // ---------- fetch: one recording's page ----------
  function fetchTranscript(job) {
    let ctx = null, bearer = null, sent = false;
    const report = (m) => { if (!sent) { sent = true; chrome.runtime.sendMessage(Object.assign({ type: 'colin-transcript' }, m)); } };

    // intercept.js (the page's own world) relays the recording's drive and item, and the API bearer.
    window.addEventListener('message', (e) => {
      if (e.source !== window || !e.data) return;
      if (e.data.type === 'TRANSCRIPT_CONTEXT' && e.data.driveId && e.data.itemId) ctx = e.data;
      if (e.data.type === 'SP_API_BEARER' && e.data.authorization) bearer = e.data.authorization;
    });

    const started = Date.now();
    const tick = setInterval(async () => {
      if (sent) return clearInterval(tick);
      if (!ctx) {
        window.postMessage({ type: 'TTD_REQUEST_CONTEXT' }, '*');
        if (Date.now() - started > 60000) { clearInterval(tick); report({ status: 'no-context', error: 'The recording page never loaded its details (signed out, or no access)' }); }
        return;
      }
      clearInterval(tick);
      if (ctx.hasTranscripts === false) return report({ status: 'none', error: 'This recording has no transcript' });
      try {
        const base = `${location.origin}${ctx.sitePath}/_api/v2.1/drives/${ctx.driveId}/items/${ctx.itemId}`;
        const headers = { Accept: 'application/json' };
        if (bearer) headers.Authorization = bearer;
        let r = await fetch(`${base}?select=media%2Ftranscripts&%24expand=media%2Ftranscripts`, { credentials: 'include', headers });
        if (!r.ok) r = await fetch(`${base}/media/transcripts`, { credentials: 'include', headers });
        if (!r.ok) return report({ status: 'error', error: 'Transcript details refused (' + r.status + ')' });
        const data = await r.json();
        const list = (data.media && data.media.transcripts) || data.value || [];
        const t = list.find((x) => x.isDefault) || list[0];
        if (!t || !t.temporaryDownloadUrl) return report({ status: 'none', error: 'This recording has no transcript' });
        const u = new URL(t.temporaryDownloadUrl);
        u.searchParams.set('format', 'json');
        if (location.hostname.endsWith('.sharepoint.com.mcas.ms') && u.hostname.endsWith('.sharepoint.com')) u.hostname += '.mcas.ms';
        const body = await fetch(u.href);
        if (!body.ok) return report({ status: 'error', error: 'Transcript download refused (' + body.status + ')' });
        const header = `Meeting: ${job.subject || ctx.fileName || ''}\nRecording: ${location.href.split('?')[0]}`;
        report({ text: core.toText(await body.text(), header) });
      } catch (e) {
        report({ status: 'error', error: String(e.message || e) });
      }
    }, 1000);
  }
})();
