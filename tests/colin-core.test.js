// node --test tests/
const { test } = require('node:test');
const assert = require('node:assert/strict');
const core = require('../src/colin-core.js');

const rec = (id, extra = {}) => ({ id, url: `https://contoso-my.sharepoint.com/:v:/g/personal/ann/${id}`, subject: 'Site walk', ...extra });

test('only SharePoint recordings are fetched', () => {
  assert.equal(core.allowedUrl('https://contoso-my.sharepoint.com/:v:/g/x'), true);
  assert.equal(core.allowedUrl('https://contoso.sharepoint.com.mcas.ms/x'), true);
  assert.equal(core.allowedUrl('http://contoso.sharepoint.com/x'), false);
  assert.equal(core.allowedUrl('https://evil.com/?sharepoint.com'), false);
  assert.equal(core.allowedUrl('https://sharepoint.com.evil.com/x'), false);
});

test('the OneDrive address is cut to the personal site', () => {
  assert.equal(core.homeFrom('https://contoso-my.sharepoint.com/personal/ann_contoso_com/Documents/Forms/All.aspx'), 'https://contoso-my.sharepoint.com/personal/ann_contoso_com');
  assert.equal(core.homeFrom('https://contoso.sharepoint.com/sites/team'), null);
  assert.equal(core.homeFrom('not a url'), null);
});

test('wanted.json: malformed entries are dropped, never guessed', () => {
  const ok = core.parseWanted(JSON.stringify({ version: 1, wanted: [rec('r-1795001'), rec('r-1795001'), { id: 'BAD ID', url: 'https://x.sharepoint.com/a' }, { id: 'r-1795002', url: 'https://evil.com/a' }, rec('r-1795003', { done: true })] }));
  assert.equal(ok.ok, true);
  assert.deepEqual(ok.items.map((i) => [i.id, i.done]), [['r-1795001', false], ['r-1795003', true]]);
  assert.equal(core.parseWanted('{').ok, false);
  assert.equal(core.parseWanted('{"version":2,"wanted":[]}').ok, false);
});

test('each run fetches what is missing, up to five, and recycles what is done', () => {
  const items = [1, 2, 3, 4, 5, 6, 7].map((n) => rec('r-179500' + n)).concat([rec('r-1795099', { done: true }), rec('r-1795098', { done: true })]);
  const now = Date.parse('2026-11-24T06:00:00Z');
  const plan = core.plan(items, ['r-1795001.txt', 'r-1795099.txt'], { 'r-1795002': { n: 3, at: now - 3600000 } }, now);
  assert.deepEqual(plan.fetch.map((w) => w.id), ['r-1795003', 'r-1795004', 'r-1795005', 'r-1795006', 'r-1795007']);
  assert.deepEqual(plan.recycle.map((w) => w.id), ['r-1795099'], 'a done one that was never saved has nothing to recycle');
  const nextDay = core.plan([rec('r-1795002')], [], { 'r-1795002': { n: 3, at: now - 3600000 } }, now + 86400000);
  assert.deepEqual(nextDay.fetch.map((w) => w.id), ['r-1795002'], 'three failures rest a day, then it tries again');
});

test('the transcript becomes one paragraph per turn, with times', () => {
  const json = JSON.stringify({ entries: [
    { speakerDisplayName: 'Ann Lee', startOffset: '00:00:01.200', endOffset: '00:00:03', text: 'Morning all.' },
    { speakerDisplayName: 'Ann Lee', startOffset: '00:00:03.100', endOffset: '00:00:05', text: 'Quick one today.' },
    { speakerDisplayName: 'Bo Kim', startOffset: '01:02:03.900', endOffset: '01:02:05', text: 'I will send the quote Friday.' },
    { speakerDisplayName: 'Bo Kim', startOffset: '01:02:06', endOffset: '01:02:07', text: '' },
  ] });
  assert.equal(core.toText(json, 'Meeting: Site walk'), 'Meeting: Site walk\n\n[00:00:01] Ann Lee: Morning all. Quick one today.\n\n[01:02:03] Bo Kim: I will send the quote Friday.\n');
});
