# CLAUDE.md

Guidance for Claude when working in this repo.

## Project layout

- `src/` — extension source. `content.js`, `intercept.js`, `mux-worker.js`, `modal.css`, `manifest.json`, `icons/`.
- `src/key.pem` — local signing key. **Must not** ship in any release zip.
- `scripts/` — repo utilities.
- `demo-website/` — separate marketing site (Cloudflare Workers). Unrelated to the extension build.
- `releases/` — Chrome Web Store upload artifacts, one zip per release. Gitignored.

## Building a release zip

Always use `scripts/package-extension.ps1` — do not zip `src/` manually. The script reads `version` from `src/manifest.json`, refuses to overwrite an existing zip (bump the version first), excludes `key.pem` + `.DS_Store` + `Thumbs.db`, and writes `ms-teams-downloader-v<version>.zip` to the repo root.

```pwsh
# 1. Bump src/manifest.json "version"
# 2. Run:
pwsh scripts/package-extension.ps1
# Use -Force only if you intentionally want to overwrite an existing zip.
```

Verify the result with `unzip -l releases/ms-teams-downloader-v<version>.zip` — expect 8 entries (`icons/icon128.png`, `content.js`, `intercept.js`, `manifest.json`, `modal.css`, `mux-worker.js`, `popup.html`, `popup.js`) and no `key.pem`.

## Colin transcript collector (this fork)

`src/background.js`, `src/colin.js` and `src/colin-core.js` add an hourly collector for Jayden's chief-of-staff stack (sicko-mode). It reads `Colin/wanted.json` from his own OneDrive, opens each listed recording in a minimised window, saves the transcript to `Colin/Transcripts/<id>.txt`, and moves a transcript to the recycle bin once `wanted.json` marks it done. Every SharePoint call is made inside a SharePoint page as the signed-in user; the worker only runs the timer and opens and closes tabs. It never reads, copies or exports cookies or tokens, which keeps it inside the Don'ts below. Tests: `node --test tests/colin-core.test.js`, and `PW=<playwright> node tests/colin-e2e.mjs src` (the real extension against a fake SharePoint). A release zip now also holds `background.js`, `colin.js` and `colin-core.js`.

## Don'ts

- Never add `chrome.cookies` permission or a background service worker to extract SharePoint session cookies, even to "improve" CLI tool support. Tested: segment URLs require `FedAuth` + `rtFa` (HttpOnly) and baking those into a copy-paste script exposes the user's full session. The in-browser path is the only one we support.
- Don't bring back the ffmpeg / yt-dlp tabs in the video modal. Microsoft now DASH-SEA AES-128-CBC encrypts Stream segments; yt-dlp refuses all `<ContentProtection>` streams as DRM, and ffmpeg's DASH demuxer doesn't implement DASH-SEA. The modal note states this — don't soften it.

## Agent skills

### Issue tracker

GitHub Issues at `brendangooden/ms-teams-sharepoint-downloader` via `gh` CLI. See `docs/agents/issue-tracker.md`.

### Triage labels

Canonical defaults (`needs-triage`, `needs-info`, `ready-for-agent`, `ready-for-human`, `wontfix`). See `docs/agents/triage-labels.md`.

### Domain docs

Single-context — one `CONTEXT.md` + `docs/adr/` at repo root. See `docs/agents/domain.md`.
