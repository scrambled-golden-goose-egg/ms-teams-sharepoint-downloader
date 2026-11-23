// End-to-end: the real extension in Chromium against a fake SharePoint (all hosts mapped to a local
// HTTPS server). Checks: wanted.json read, two transcripts saved (one with none skipped), a done one
// recycled, collector.json written, the minimised window closed.
//
// Run: PW=<path to the playwright package> node tests/colin-e2e.mjs src
// Needs openssl (a throwaway certificate for *.sharepoint.com is made in a temp folder).
import https from 'node:https';
import { readFileSync, mkdtempSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const { chromium } = require(process.env.PW || 'playwright');
const EXT = resolve(process.argv[2] || 'src');
const here = mkdtempSync(join(tmpdir(), 'colin-e2e-')) + '/';
execFileSync('openssl', ['req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-keyout', here + 'key.pem', '-out', here + 'cert.pem', '-days', '1', '-subj', '/CN=sharepoint.com', '-addext', 'subjectAltName=DNS:*.sharepoint.com'], { stdio: 'ignore' });

const HOME = '/personal/ann_contoso_com';
const saved = {}, recycled = [], seen = [];
const wanted = { version: 1, wanted: [
  { id: 'r-1000001', url: 'https://contoso-my.sharepoint.com/personal/bob_contoso_com/rec1', subject: 'Site walk with Galaxy' },
  { id: 'r-1000002', url: 'https://partner.sharepoint.com/sites/x/rec2', subject: 'Partner hosted' },
  { id: 'r-1000003', url: 'https://contoso-my.sharepoint.com/personal/bob_contoso_com/rec3', subject: 'No transcript' },
  { id: 'r-done001', url: 'https://contoso-my.sharepoint.com/personal/bob_contoso_com/rec9', subject: 'Old', done: true },
] };
const transcript = JSON.stringify({ entries: [
  { speakerDisplayName: 'Ann Lee', startOffset: '00:00:01', endOffset: '00:00:02', text: 'Morning.' },
  { speakerDisplayName: 'Bo Kim', startOffset: '00:01:00', endOffset: '00:01:02', text: 'Quote by Friday.' },
] });
const recPage = (spItemUrl, has) => `<!doctype html><html><head><script>window.g_fileInfo={'.spItemUrl':${JSON.stringify(spItemUrl)},hasTranscripts:${has},name:'Rec.mp4'};</script></head><body>player</body></html>`;

const server = https.createServer({ key: readFileSync(here + 'key.pem'), cert: readFileSync(here + 'cert.pem') }, (req, res) => {
  const host = req.headers.host.split(':')[0];
  const url = decodeURIComponent(req.url);
  seen.push(`${req.method} ${host}${url}`.slice(0, 160));
  let body = '';
  req.on('data', (c) => (body += c));
  req.on('end', () => {
    const send = (code, data, type = 'application/json') => { res.writeHead(code, { 'content-type': type }); res.end(typeof data === 'string' ? data : JSON.stringify(data)); };
    if (url.startsWith(HOME + '/_layouts/15/onedrive.aspx')) return send(200, '<!doctype html><title>OneDrive</title><p>files</p>', 'text/html');
    if (url === HOME + '/_api/contextinfo' && req.method === 'POST') return send(200, { FormDigestValue: 'digest-1' });
    if (url.includes("GetFileByServerRelativePath(decodedurl='/personal/ann_contoso_com/Documents/Colin/wanted.json')/$value")) return send(200, JSON.stringify(wanted));
    if (url.includes("GetFolderByServerRelativePath(decodedurl='/personal/ann_contoso_com/Documents/Colin/Transcripts')/Files?$select=Name")) return send(200, { value: [{ Name: 'r-done001.txt' }] });
    const add = url.match(/GetFolderByServerRelativePath\(decodedurl='([^']+)'\)\/Files\/AddUsingPath\(decodedurl='([^']+)',overwrite=true\)/);
    if (add && req.method === 'POST') { if (req.headers['x-requestdigest'] !== 'digest-1') return send(403, {}); saved[add[1] + '/' + add[2]] = body; return send(200, {}); }
    const rc = url.match(/GetFileByServerRelativePath\(decodedurl='([^']+)'\)\/recycle\(\)/);
    if (rc && req.method === 'POST') { recycled.push(rc[1]); return send(200, {}); }
    if (url === '/personal/bob_contoso_com/rec1') return send(200, recPage('https://contoso-my.sharepoint.com/personal/bob_contoso_com/_api/v2.0/drives/D1/items/I1', true), 'text/html');
    if (url === '/personal/bob_contoso_com/rec3') return send(200, recPage('https://contoso-my.sharepoint.com/personal/bob_contoso_com/_api/v2.0/drives/D1/items/I3', false), 'text/html');
    if (url === '/sites/x/rec2' && host === 'partner.sharepoint.com') return send(200, recPage('https://partner.sharepoint.com/sites/x/_api/v2.0/drives/D2/items/I2', true), 'text/html');
    if (/\/_api\/v2\.1\/drives\/D[12]\/items\/I[12]\?/.test(url)) return send(200, { media: { transcripts: [{ isDefault: true, temporaryDownloadUrl: `https://${host}/tx/1?sig=abc` }] } });
    if (url.startsWith('/tx/1?sig=abc') && url.includes('format=json')) return send(200, transcript);
    send(404, { error: 'not mocked' });
  });
});
await new Promise((r) => server.listen(8443, '127.0.0.1', r));

const ctx = await chromium.launchPersistentContext('', {
  channel: 'chromium', headless: true, ignoreHTTPSErrors: true,
  args: [`--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`, '--host-resolver-rules=MAP *.sharepoint.com 127.0.0.1:8443', '--ignore-certificate-errors', '--no-proxy-server'],
});
const logs = [];
ctx.on('page', (p) => { p.on('console', (m) => logs.push('page: ' + m.text().slice(0, 200))); p.on('pageerror', (e) => logs.push('pageerror: ' + e.message)); });
let [sw] = ctx.serviceWorkers();
if (!sw) sw = await ctx.waitForEvent('serviceworker');
await sw.evaluate(() => new Promise((r) => chrome.storage.local.set({ colinEnabled: true, colinHome: 'https://contoso-my.sharepoint.com/personal/ann_contoso_com' }, r)));
const t0 = Date.now();
const result = await sw.evaluate(() => run('test'));
const out = {
  result, seconds: Math.round((Date.now() - t0) / 1000),
  saved: Object.keys(saved),
  rec1: saved['/personal/ann_contoso_com/Documents/Colin/Transcripts/r-1000001.txt'],
  collector: saved['/personal/ann_contoso_com/Documents/Colin/collector.json'] && JSON.parse(saved['/personal/ann_contoso_com/Documents/Colin/collector.json']).results,
  recycled,
  pagesLeftOpen: ctx.pages().map((p) => p.url()),
  seen: seen.slice(0, 30),
  logs: logs.slice(0, 30),
  attempts: await sw.evaluate(() => new Promise((r) => chrome.storage.local.get(['colinAttempts', 'colinLast'], r))),
};
console.log(JSON.stringify(out, null, 1));
await ctx.close();
server.close();
