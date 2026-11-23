// Toolbar popup: read the version from the manifest so it never goes stale
// (MV3 extension pages can't run inline scripts, hence this file)
document.getElementById('version').textContent = 'v' + chrome.runtime.getManifest().version;

// Colin transcript collector settings.
(function () {
  const home = document.getElementById('colin-home');
  const on = document.getElementById('colin-on');
  const runBtn = document.getElementById('colin-run');
  const last = document.getElementById('colin-last');
  const show = (l) => { last.textContent = l ? `Last run ${new Date(l.at).toLocaleString()}: ${l.summary}` : 'Not run yet'; };

  chrome.storage.local.get(['colinHome', 'colinEnabled', 'colinLast'], (s) => {
    home.value = s.colinHome || '';
    on.checked = !!s.colinEnabled;
    show(s.colinLast);
  });
  home.addEventListener('change', () => {
    const ok = self.ColinCore ? self.ColinCore.homeFrom(home.value) : home.value.trim();
    home.setCustomValidity(ok ? '' : 'Your OneDrive address, starting https:// and ending /personal/<you>');
    if (ok) { home.value = ok; chrome.storage.local.set({ colinHome: ok }); }
    home.reportValidity();
  });
  on.addEventListener('change', () => chrome.storage.local.set({ colinEnabled: on.checked }));
  runBtn.addEventListener('click', () => {
    runBtn.disabled = true;
    last.textContent = 'Running…';
    chrome.runtime.sendMessage({ type: 'colin-run-now' }, (r) => {
      runBtn.disabled = false;
      chrome.storage.local.get(['colinLast'], (s) => show(s.colinLast || (r && { at: Date.now(), summary: r.summary })));
    });
  });
})();
