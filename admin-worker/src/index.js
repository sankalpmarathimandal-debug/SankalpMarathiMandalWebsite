/**
 * Sankalp Marathi Mandal — Website Admin Panel
 * ------------------------------------------------------------
 * A single Cloudflare Worker that:
 *   1. Serves a password-protected admin dashboard, organized by which
 *      site page you want to change.
 *   2. Lets you view and edit the current contents of the site's Excel
 *      workbooks in a spreadsheet-like table (or upload a replacement
 *      file wholesale), and see thumbnails of images already in each
 *      folder before uploading or deleting.
 *   3. Commits every change directly to the GitHub repo via the
 *      GitHub REST API, so GitHub Pages redeploys automatically —
 *      nobody needs GitHub access.
 *
 * See SETUP.md for how to configure and deploy this.
 */

// ---------------------------------------------------------------------------
// Config: every file/folder the site depends on (see README.md in the repo)
// ---------------------------------------------------------------------------

const XLSX_FILES = [
  { path: "data/home-events.xlsx", label: "Homepage Events", desc: "Homepage event cards" },
  { path: "data/timeline.xlsx", label: "Event Timeline", desc: "Events page timeline" },
  { path: "data/testimonials.xlsx", label: "Testimonials", desc: "Homepage “Community Voices”" },
  { path: "data/partners.xlsx", label: "Partners", desc: "Homepage partner logos list" },
  { path: "data/team.xlsx", label: "Team", desc: "Our Team page" },
  { path: "data/shala-team.xlsx", label: "Shala Team", desc: "Marathi Shala team section" },
  { path: "data/faq.xlsx", label: "FAQ", desc: "FAQs page" },
  { path: "data/shala-faq.xlsx", label: "Shala FAQ", desc: "Marathi Shala FAQs" },
  { path: "data/shala-calendar.xlsx", label: "Shala Calendar", desc: "Shala Calendar page + download button" },
  { path: "data/forms.xlsx", label: "Forms & Sign-ups", desc: "Forms & Sign-ups page" },
  { path: "data/showcase.xlsx", label: "Showcase", desc: "Showcase page entries" },
];

const SINGLE_DOCS = [
  { path: "docs/constitution.pdf", label: "Constitution PDF", desc: "Constitution page", accept: "application/pdf" },
];

const FOLDERS = [
  { path: "assets/images/events", label: "Event Images", accept: "image/*", yearFolders: true, note: "Organized by year — pick/type a year below", xlsxRef: "data/home-events.xlsx / data/timeline.xlsx" },
  { path: "assets/images/team", label: "Team Photos", accept: "image/*", xlsxRef: "data/team.xlsx" },
  { path: "assets/images/shala", label: "Shala Images", accept: "image/*" },
  { path: "assets/images/showcase", label: "Showcase Photos", accept: "image/*", xlsxRef: "data/showcase.xlsx" },
  { path: "assets/images/partners", label: "Partner Logos", accept: "image/*", xlsxRef: "data/partners.xlsx" },
  {
    path: "assets/images/branding/logo-variants",
    label: "Community Pride Wall Logos",
    accept: "image/*",
    note: "Auto-updates homepage — no xlsx edit needed",
    autoManifest: true,
  },
  {
    path: "assets/images/highlights",
    label: "Homepage Highlight Photos",
    accept: "image/*",
    note: "Auto-updates homepage slider — no xlsx edit needed",
    autoManifest: true,
    namingHint: "Filename becomes the caption shown on the homepage slider. Use the format “YYYY - Title”, e.g. “2026 - Summer Picnic” — not a camera/screenshot default name.",
  },
  { path: "docs/showcase", label: "Showcase Documents (PDFs)", accept: "application/pdf", xlsxRef: "data/showcase.xlsx" },
];

// Groups everything above by the actual page on the live site it affects —
// this drives the "which page do you want to change?" selector.
const PAGES = [
  { key: "home", label: "Home Page", xlsx: ["data/home-events.xlsx", "data/testimonials.xlsx", "data/partners.xlsx"], folders: ["assets/images/highlights", "assets/images/branding/logo-variants", "assets/images/partners"] },
  { key: "events", label: "Events Page", xlsx: ["data/timeline.xlsx"], folders: ["assets/images/events"] },
  { key: "team", label: "Team Page", xlsx: ["data/team.xlsx"], folders: ["assets/images/team"] },
  { key: "shala", label: "Shala Page", xlsx: ["data/shala-team.xlsx", "data/shala-faq.xlsx"], folders: ["assets/images/shala"] },
  { key: "calendar", label: "Shala Calendar", xlsx: ["data/shala-calendar.xlsx"], folders: [] },
  { key: "faq", label: "FAQ Page", xlsx: ["data/faq.xlsx"], folders: [] },
  { key: "forms", label: "Forms & Sign-ups", xlsx: ["data/forms.xlsx"], folders: [] },
  { key: "showcase", label: "Showcase Page", xlsx: ["data/showcase.xlsx"], folders: ["assets/images/showcase", "docs/showcase"] },
  { key: "constitution", label: "Constitution Page", xlsx: [], folders: [], docs: ["docs/constitution.pdf"] },
];

// ---------------------------------------------------------------------------
// GitHub Contents API helpers
// ---------------------------------------------------------------------------

function ghHeaders(env) {
  return {
    Authorization: `Bearer ${env.GITHUB_TOKEN}`,
    Accept: "application/vnd.github+json",
    "User-Agent": "sankalp-admin-worker",
    "X-GitHub-Api-Version": "2022-11-28",
  };
}

function ghUrl(env, path) {
  return `https://api.github.com/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}/contents/${path
    .split("/")
    .map(encodeURIComponent)
    .join("/")}`;
}

async function ghGetSha(env, path) {
  const res = await fetch(`${ghUrl(env, path)}?ref=${env.GITHUB_BRANCH}`, { headers: ghHeaders(env) });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`GitHub GET ${path} failed: ${res.status} ${await res.text()}`);
  const json = await res.json();
  return json.sha;
}

async function ghGetFile(env, path) {
  const res = await fetch(`${ghUrl(env, path)}?ref=${env.GITHUB_BRANCH}`, { headers: ghHeaders(env) });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`GitHub GET ${path} failed: ${res.status} ${await res.text()}`);
  const json = await res.json();
  return { sha: json.sha, contentBase64: json.content };
}

async function ghListFolder(env, path, recurseDirs = false) {
  const res = await fetch(`${ghUrl(env, path)}?ref=${env.GITHUB_BRANCH}`, { headers: ghHeaders(env) });
  if (res.status === 404) return [];
  if (!res.ok) throw new Error(`GitHub list ${path} failed: ${res.status} ${await res.text()}`);
  const json = await res.json();
  if (!Array.isArray(json)) return [];
  const files = json.filter((f) => f.type === "file");
  if (!recurseDirs) return files;
  const dirs = json.filter((f) => f.type === "dir");
  const nested = await Promise.all(dirs.map((d) => ghListFolder(env, d.path, false)));
  return files.concat(...nested);
}

async function ghCommitFile(env, path, base64Content, message) {
  const sha = await ghGetSha(env, path);
  const res = await fetch(ghUrl(env, path), {
    method: "PUT",
    headers: { ...ghHeaders(env), "Content-Type": "application/json" },
    body: JSON.stringify({
      message,
      content: base64Content,
      branch: env.GITHUB_BRANCH,
      ...(sha ? { sha } : {}),
    }),
  });
  if (!res.ok) throw new Error(`GitHub commit ${path} failed: ${res.status} ${await res.text()}`);
  return res.json();
}

async function ghListCommits(env, perPage = 40) {
  const url = `https://api.github.com/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}/commits?sha=${env.GITHUB_BRANCH}&per_page=${perPage}`;
  const res = await fetch(url, { headers: ghHeaders(env) });
  if (!res.ok) throw new Error(`GitHub list commits failed: ${res.status} ${await res.text()}`);
  const list = await res.json();
  return list.map((c) => ({
    sha: c.sha,
    shortSha: c.sha.slice(0, 7),
    message: c.commit.message,
    date: c.commit.author?.date || c.commit.committer?.date || null,
    authorName: c.commit.author?.name || "unknown",
    url: c.html_url,
  }));
}

// Contents API has no native "rename" — recreate at the new path with the
// same bytes, then delete the old path. Two commits, same effect.
async function ghRenameFile(env, oldPath, newPath, message) {
  const existing = await ghGetFile(env, oldPath);
  if (!existing) throw new Error(`File not found: ${oldPath}`);
  const clash = await ghGetSha(env, newPath);
  if (clash) throw new Error(`A file already exists at ${newPath}`);
  await ghCommitFile(env, newPath, existing.contentBase64, message);
  await ghDeleteFile(env, oldPath, message);
}

async function ghDeleteFile(env, path, message) {
  const sha = await ghGetSha(env, path);
  if (!sha) throw new Error(`File not found: ${path}`);
  const res = await fetch(ghUrl(env, path), {
    method: "DELETE",
    headers: { ...ghHeaders(env), "Content-Type": "application/json" },
    body: JSON.stringify({ message, sha, branch: env.GITHUB_BRANCH }),
  });
  if (!res.ok) throw new Error(`GitHub delete ${path} failed: ${res.status} ${await res.text()}`);
  return res.json();
}

// ---------------------------------------------------------------------------
// Session (signed cookie, no server-side storage needed)
// ---------------------------------------------------------------------------

const SESSION_HOURS = 12;

async function hmac(secret, message) {
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const sig = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(message));
  return [...new Uint8Array(sig)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function makeSession(env) {
  const expiry = Date.now() + SESSION_HOURS * 3600 * 1000;
  const sig = await hmac(env.SESSION_SECRET, String(expiry));
  return `${expiry}.${sig}`;
}

async function verifySession(env, cookieHeader) {
  if (!cookieHeader) return false;
  const match = cookieHeader.match(/(?:^|;\s*)session=([^;]+)/);
  if (!match) return false;
  const [expiryStr, sig] = decodeURIComponent(match[1]).split(".");
  if (!expiryStr || !sig) return false;
  if (Date.now() > Number(expiryStr)) return false;
  const expected = await hmac(env.SESSION_SECRET, expiryStr);
  return expected === sig;
}

function setCookieHeader(token) {
  return `session=${encodeURIComponent(token)}; Path=/; HttpOnly; Secure; SameSite=Strict; Max-Age=${SESSION_HOURS * 3600}`;
}

// ---------------------------------------------------------------------------
// HTML
// ---------------------------------------------------------------------------

const BASE_STYLE = `
  :root { color-scheme: light; }
  * { box-sizing: border-box; }
  body { font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif; background:#f5f4f0; margin:0; color:#242018; }
  header { background:#8c2f39; color:#fff; padding:18px 24px; display:flex; justify-content:space-between; align-items:center; }
  header h1 { font-size:18px; margin:0; }
  header a { color:#fff; text-decoration:none; font-size:13px; opacity:0.85; }
  main { max-width:1040px; margin:0 auto; padding:24px; }
  .card { background:#fff; border:1px solid #e4e0d8; border-radius:10px; padding:18px 20px; margin-bottom:16px; }
  .card h3 { margin:0 0 4px; font-size:15px; }
  .card p.desc { margin:0 0 12px; font-size:13px; color:#6b6558; }
  .row { display:flex; gap:10px; align-items:center; flex-wrap:wrap; }
  input[type=file] { font-size:13px; }
  button { background:#8c2f39; color:#fff; border:none; padding:8px 16px; border-radius:6px; cursor:pointer; font-size:13px; }
  button:hover { background:#732530; }
  button.secondary { background:#e4e0d8; color:#242018; }
  button.danger { background:#a33; }
  button:disabled { opacity:0.5; cursor:default; }
  .status { font-size:12px; margin-top:8px; min-height:16px; }
  .status.ok { color:#2a7a34; }
  .status.err { color:#a33; }
  .thumbs { display:flex; flex-wrap:wrap; gap:10px; margin-top:12px; }
  .thumb { width:84px; text-align:center; font-size:10px; }
  .thumb img { width:84px; height:84px; object-fit:cover; border-radius:6px; border:1px solid #e4e0d8; }
  .thumb .name { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .thumb button { margin-top:4px; padding:2px 6px; font-size:10px; }
  .note { display:inline-block; background:#f1e6c8; color:#6b5a10; font-size:11px; padding:2px 8px; border-radius:10px; margin-left:8px; }
  .naming-hint { font-size:12px; color:#8c2f39; background:#fbeef0; border:1px solid #f0d5d8; border-radius:6px; padding:8px 10px; margin:0 0 12px; }
  .pending-list { margin-top:10px; }
  .pending-row { display:flex; align-items:center; gap:8px; padding:6px 0; border-bottom:1px solid #f0ede6; font-size:13px; }
  .pending-row:last-child { border-bottom:none; }
  .pending-row .orig-name { color:#999; font-size:11px; flex-shrink:0; max-width:140px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .pending-row input[type=text] { flex:1; padding:6px 8px; border:1px solid #ccc; border-radius:6px; font-size:13px; }
  .pending-row .ext { color:#6b6558; font-size:12px; flex-shrink:0; }
  .thumb .rename-btn { margin-top:4px; padding:2px 6px; font-size:10px; }
  h2.section { margin:28px 0 6px; font-size:16px; color:#8c2f39; }
  .login-box { max-width:340px; margin:80px auto; background:#fff; border-radius:10px; padding:28px; border:1px solid #e4e0d8; }
  .login-box input { width:100%; padding:10px; margin:10px 0; border:1px solid #ccc; border-radius:6px; font-size:14px; }
  .login-box button { width:100%; padding:10px; font-size:14px; }
  .page-tabs { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:20px; }
  .page-tab { background:#fff; color:#242018; border:1px solid #d9d3c6; padding:8px 14px; border-radius:20px; cursor:pointer; font-size:13px; }
  .page-tab.active { background:#8c2f39; color:#fff; border-color:#8c2f39; }
  .page-section { display:none; }
  .page-section.active { display:block; }
  .editor-toggle { margin-left:auto; }
  .xlsx-editor { display:none; margin-top:14px; border-top:1px solid #eee; padding-top:14px; }
  .xlsx-editor .table-wrap { overflow-x:auto; max-height:420px; overflow-y:auto; border:1px solid #e4e0d8; border-radius:6px; }
  table.xlsx-table { border-collapse:collapse; font-size:12px; width:100%; }
  table.xlsx-table th, table.xlsx-table td { border:1px solid #e4e0d8; padding:6px 8px; min-width:100px; text-align:left; }
  table.xlsx-table th { background:#f5f4f0; position:sticky; top:0; font-weight:600; }
  table.xlsx-table td[contenteditable="true"]:focus, table.xlsx-table th[contenteditable="true"]:focus { outline:2px solid #8c2f39; outline-offset:-2px; background:#fff8ee; }
  .editor-controls { display:flex; gap:8px; margin-top:10px; }
`;

function loginPage(error) {
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Sankalp Admin — Login</title><style>${BASE_STYLE}</style></head>
  <body><div class="login-box">
    <h2 style="margin-top:0">Sankalp Website Admin</h2>
    <form method="POST" action="/login">
      <input type="password" name="password" placeholder="Admin password" required autofocus>
      <button type="submit">Log in</button>
    </form>
    ${error ? `<p class="status err">${error}</p>` : ""}
  </div></body></html>`;
}

function cssId(path) {
  return path.replace(/[^a-zA-Z0-9]/g, "-");
}

function fileCardHtml(f) {
  const isXlsx = f.path.endsWith(".xlsx");
  return `<div class="card" data-path="${f.path}">
    <div class="row">
      <div>
        <h3 style="margin-bottom:2px">${f.label}</h3>
        <p class="desc" style="margin-bottom:6px">${f.desc || ""} &middot; <code>${f.path}</code></p>
      </div>
      ${isXlsx ? `<button class="secondary editor-toggle" onclick="toggleEditor(this, '${f.path}')">View / Edit Data</button>` : ""}
    </div>

    ${isXlsx ? `<div class="xlsx-editor" id="editor-${cssId(f.path)}"></div>` : ""}

    <div class="row" style="margin-top:14px;padding-top:14px;border-top:1px solid #eee;">
      <input type="file" accept="${f.accept || ".xlsx"}">
      <button onclick="uploadSingle(this, '${f.path}')">${isXlsx ? "Or upload a replacement file" : "Upload / Replace"}</button>
    </div>
    <div class="status"></div>
  </div>`;
}

function folderCardHtml(f) {
  const yearInput = f.yearFolders
    ? `<input type="text" class="year-input" placeholder="Year (e.g. ${new Date().getFullYear()})" style="width:130px;padding:8px;border:1px solid #ccc;border-radius:6px;" value="${new Date().getFullYear()}">`
    : "";
  return `<div class="card" data-folder="${f.path}" data-auto="${f.autoManifest ? "1" : "0"}" data-xlsx-ref="${f.xlsxRef || ""}">
    <h3>${f.label} ${f.note ? `<span class="note">${f.note}</span>` : ""}</h3>
    <p class="desc"><code>${f.path}/</code></p>
    ${f.namingHint ? `<p class="naming-hint">${f.namingHint}</p>` : ""}
    ${f.xlsxRef ? `<p class="naming-hint">Filenames here are referenced by <code>${f.xlsxRef}</code> — if you rename an existing file, update that row's filename too, or the site will show a broken image.</p>` : ""}
    <div class="row">
      ${yearInput}
      <input type="file" accept="${f.accept}" multiple onchange="onFilesSelected(this, '${f.path}', ${!!f.yearFolders})">
    </div>
    <div class="pending-list" id="pending-${cssId(f.path)}"></div>
    <div class="status"></div>
    <div class="thumbs" id="thumbs-${cssId(f.path)}">Loading current files…</div>
  </div>`;
}

function byPath(list, path) {
  return list.find((x) => x.path === path);
}

function pageSectionHtml(page, idx) {
  const xlsxCards = (page.xlsx || []).map((p) => fileCardHtml(byPath(XLSX_FILES, p))).join("\n");
  const docCards = (page.docs || []).map((p) => fileCardHtml(byPath(SINGLE_DOCS, p))).join("\n");
  const folderCards = (page.folders || []).map((p) => folderCardHtml(byPath(FOLDERS, p))).join("\n");
  return `<section class="page-section${idx === 0 ? " active" : ""}" id="page-${page.key}">
    ${xlsxCards || docCards ? `<h2 class="section">Content</h2>${xlsxCards}${docCards}` : ""}
    ${folderCards ? `<h2 class="section">Images &amp; Documents</h2>${folderCards}` : ""}
  </section>`;
}

function logSectionHtml() {
  return `<section class="page-section" id="page-log">
    <h2 class="section">Activity Log</h2>
    <p style="font-size:13px;color:#6b6558;margin:0 0 14px;">Every change made here (or by anyone else pushing to the repo) is a real GitHub commit — this list is pulled live from GitHub, not a separate log file, so it can never fall out of sync. Click a row to see the exact before/after diff on GitHub.</p>
    <div class="card" style="padding:0;overflow:hidden;">
      <div class="table-wrap" style="max-height:520px;">
        <table class="xlsx-table" style="width:100%;">
          <thead><tr><th>When</th><th>Change</th><th></th></tr></thead>
          <tbody id="log-tbody"><tr><td colspan="3" style="padding:14px;color:#999;">Loading…</td></tr></tbody>
        </table>
      </div>
    </div>
    <div class="row" style="margin-top:12px;"><button class="secondary" onclick="loadLogs()">Refresh</button></div>
  </section>`;
}

function adminPage() {
  const tabs = PAGES.map(
    (p, i) => `<button class="page-tab${i === 0 ? " active" : ""}" data-page="${p.key}" onclick="showPage('${p.key}')">${p.label}</button>`
  ).join("\n") + `\n<button class="page-tab" data-page="log" onclick="showPage('log')">Activity Log</button>`;
  const sections = PAGES.map(pageSectionHtml).join("\n") + "\n" + logSectionHtml();
  const allFolderPaths = JSON.stringify(FOLDERS.map((f) => f.path));

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Sankalp Website Admin</title><style>${BASE_STYLE}</style>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  </head>
  <body>
  <header>
    <h1>Sankalp Website Admin</h1>
    <a href="/logout">Log out</a>
  </header>
  <main>
    <p style="font-size:13px;color:#6b6558">Pick the page you want to change below. Uploads and edits commit directly to the live site's GitHub repo — changes go live within a minute or two.</p>

    <div class="page-tabs">
      ${tabs}
    </div>

    ${sections}
  </main>

  <script>
  // ---- helpers ----------------------------------------------------------

  function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result.split(',')[1]);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  function base64ToArrayBuffer(base64) {
    const clean = base64.replace(/\\s/g, '');
    const binary = atob(clean);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes.buffer;
  }

  function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
  }

  function setStatus(card, msg, ok) {
    const el = card.querySelector('.status');
    el.textContent = msg;
    el.className = 'status ' + (ok ? 'ok' : 'err');
  }

  // ---- page tabs ----------------------------------------------------------

  function showPage(key) {
    document.querySelectorAll('.page-section').forEach(function(s) { s.classList.toggle('active', s.id === 'page-' + key); });
    document.querySelectorAll('.page-tab').forEach(function(b) { b.classList.toggle('active', b.dataset.page === key); });
  }

  // ---- xlsx upload / replace ----------------------------------------------

  async function uploadSingle(btn, path) {
    const card = btn.closest('.card');
    const input = card.querySelector('input[type=file]');
    const file = input.files[0];
    if (!file) { setStatus(card, 'Choose a file first.', false); return; }
    btn.disabled = true;
    setStatus(card, 'Uploading…', true);
    try {
      const contentBase64 = await fileToBase64(file);
      const res = await fetch('/api/commit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ path, contentBase64, message: 'Admin: update ' + path })
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || 'Upload failed');
      setStatus(card, 'Updated. Live in a minute or two.', true);
      input.value = '';
      const editor = document.getElementById('editor-' + path.replace(/[^a-zA-Z0-9]/g, '-'));
      if (editor) { editor.dataset.loaded = ''; editor.innerHTML = ''; editor.style.display = 'none'; }
    } catch (e) {
      setStatus(card, e.message, false);
    } finally {
      btn.disabled = false;
    }
  }

  // ---- xlsx in-browser table editor ----------------------------------------

  async function toggleEditor(btn, path) {
    const card = btn.closest('.card');
    const editorId = 'editor-' + path.replace(/[^a-zA-Z0-9]/g, '-');
    const editor = document.getElementById(editorId);
    const isOpen = editor.style.display === 'block';
    if (isOpen) { editor.style.display = 'none'; btn.textContent = 'View / Edit Data'; return; }
    editor.style.display = 'block';
    btn.textContent = 'Hide Data';
    if (editor.dataset.loaded === '1') return;

    setStatus(card, 'Loading current data…', true);
    try {
      const res = await fetch('/api/file?path=' + encodeURIComponent(path));
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Failed to load file');
      const buf = base64ToArrayBuffer(data.contentBase64);
      const wb = XLSX.read(buf, { type: 'array' });
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      renderEditor(editor, path, sheetName, aoa, card);
      editor.dataset.loaded = '1';
      setStatus(card, '', true);
    } catch (e) {
      setStatus(card, e.message, false);
    }
  }

  function renderEditor(editor, path, sheetName, aoa, card) {
    editor.innerHTML = '';
    const wrap = document.createElement('div');
    wrap.className = 'table-wrap';
    const table = document.createElement('table');
    table.className = 'xlsx-table';
    if (!aoa.length) aoa = [['']];
    const colCount = aoa.reduce(function(m, r) { return Math.max(m, r.length); }, 1);

    aoa.forEach(function(row, ri) {
      const tr = document.createElement('tr');
      for (let ci = 0; ci < colCount; ci++) {
        const cellEl = document.createElement(ri === 0 ? 'th' : 'td');
        cellEl.contentEditable = 'true';
        cellEl.textContent = row[ci] !== undefined ? row[ci] : '';
        tr.appendChild(cellEl);
      }
      table.appendChild(tr);
    });
    wrap.appendChild(table);
    editor.appendChild(wrap);

    const controls = document.createElement('div');
    controls.className = 'editor-controls';

    const addRowBtn = document.createElement('button');
    addRowBtn.className = 'secondary';
    addRowBtn.textContent = '+ Add row';
    addRowBtn.onclick = function() {
      const cols = table.rows[0].cells.length;
      const tr = document.createElement('tr');
      for (let i = 0; i < cols; i++) {
        const td = document.createElement('td');
        td.contentEditable = 'true';
        tr.appendChild(td);
      }
      table.appendChild(tr);
      tr.cells[0].focus();
    };

    const delRowBtn = document.createElement('button');
    delRowBtn.className = 'secondary';
    delRowBtn.textContent = 'Delete last row';
    delRowBtn.onclick = function() {
      if (table.rows.length > 1) table.deleteRow(table.rows.length - 1);
    };

    const saveBtn = document.createElement('button');
    saveBtn.textContent = 'Save changes';
    saveBtn.onclick = function() { saveEditor(saveBtn, path, sheetName, table, card); };

    controls.appendChild(addRowBtn);
    controls.appendChild(delRowBtn);
    controls.appendChild(saveBtn);
    editor.appendChild(controls);
  }

  async function saveEditor(btn, path, sheetName, table, card) {
    btn.disabled = true;
    setStatus(card, 'Saving…', true);
    try {
      const aoa = Array.from(table.rows).map(function(tr) {
        return Array.from(tr.cells).map(function(cell) { return cell.textContent; });
      }).filter(function(row, idx) {
        return idx === 0 || row.some(function(c) { return c.trim() !== ''; });
      });
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Sheet1');
      const outBuf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const contentBase64 = arrayBufferToBase64(outBuf);
      const res = await fetch('/api/commit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ path, contentBase64, message: 'Admin: edit ' + path + ' via table editor' })
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || 'Save failed');
      setStatus(card, 'Saved. Live in a minute or two.', true);
    } catch (e) {
      setStatus(card, e.message, false);
    } finally {
      btn.disabled = false;
    }
  }

  // ---- image / doc folders --------------------------------------------------

  // Splits "IMG_4821 copy.jpeg" into a suggested clean base name and the
  // original extension, so the admin edits a readable name but can't
  // accidentally break the file type.
  function suggestName(fileName) {
    const dot = fileName.lastIndexOf('.');
    const base = dot > 0 ? fileName.slice(0, dot) : fileName;
    const ext = dot > 0 ? fileName.slice(dot) : '';
    const cleaned = base.replace(/_+/g, ' ').replace(/\\s+/g, ' ').trim();
    return { base: cleaned, ext: ext };
  }

  // Selecting files doesn't upload immediately — it shows one editable name
  // field per file so the admin can standardize naming before anything is
  // committed. This is the main defense against random screenshot/camera
  // filenames ending up as page content (see the naming-hint text on
  // folders like Homepage Highlights, where the filename becomes a caption).
  function onFilesSelected(input, folder, hasYear) {
    const card = input.closest('.card');
    const container = document.getElementById('pending-' + folder.replace(/[^a-zA-Z0-9]/g, '-'));
    container.innerHTML = '';
    setStatus(card, '', true);
    const files = Array.from(input.files);
    if (!files.length) return;

    const rows = [];
    files.forEach(function(file) {
      const suggestion = suggestName(file.name);
      const row = document.createElement('div');
      row.className = 'pending-row';

      const orig = document.createElement('span');
      orig.className = 'orig-name';
      orig.title = 'Original filename: ' + file.name;
      orig.textContent = file.name;

      const nameInput = document.createElement('input');
      nameInput.type = 'text';
      nameInput.value = suggestion.base;

      const extSpan = document.createElement('span');
      extSpan.className = 'ext';
      extSpan.textContent = suggestion.ext;

      row.appendChild(orig);
      row.appendChild(nameInput);
      row.appendChild(extSpan);
      container.appendChild(row);
      rows.push({ file: file, nameInput: nameInput, ext: suggestion.ext });
    });

    const confirmBtn = document.createElement('button');
    confirmBtn.textContent = 'Confirm names & Upload';
    confirmBtn.style.marginTop = '8px';
    confirmBtn.onclick = function() { confirmPendingUpload(confirmBtn, folder, hasYear, rows, card, input); };
    container.appendChild(confirmBtn);
  }

  async function confirmPendingUpload(btn, folder, hasYear, rows, card, input) {
    let subfolder = '';
    if (hasYear) {
      const yearInput = card.querySelector('.year-input');
      subfolder = (yearInput.value || '').trim();
      if (!subfolder) { setStatus(card, 'Enter a year first.', false); return; }
    }
    for (const row of rows) {
      if (!row.nameInput.value.trim()) { setStatus(card, 'Give every file a name before uploading.', false); return; }
    }
    btn.disabled = true;
    for (const row of rows) {
      const finalName = row.nameInput.value.trim() + row.ext;
      setStatus(card, 'Uploading ' + finalName + '…', true);
      try {
        const contentBase64 = await fileToBase64(row.file);
        const path = folder + (subfolder ? '/' + subfolder : '') + '/' + finalName;
        const res = await fetch('/api/commit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ path, contentBase64, message: 'Admin: add ' + path })
        });
        const json = await res.json();
        if (!res.ok) throw new Error(json.error || 'Upload failed');
      } catch (e) {
        setStatus(card, finalName + ': ' + e.message, false);
        btn.disabled = false;
        return;
      }
    }
    setStatus(card, 'Uploaded. Live in a minute or two.', true);
    input.value = '';
    document.getElementById('pending-' + folder.replace(/[^a-zA-Z0-9]/g, '-')).innerHTML = '';
    btn.disabled = false;
    loadFolder(folder);
  }

  async function deleteFromFolder(path, folder) {
    if (!confirm('Delete ' + path + '?')) return;
    try {
      const res = await fetch('/api/delete', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ path, message: 'Admin: remove ' + path })
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || 'Delete failed');
      loadFolder(folder);
    } catch (e) {
      alert(e.message);
    }
  }

  async function renameInFolder(oldPath, folder, warnXlsxRef) {
    const slash = oldPath.lastIndexOf('/');
    const dir = oldPath.slice(0, slash);
    const fileName = oldPath.slice(slash + 1);
    const dot = fileName.lastIndexOf('.');
    const base = dot > 0 ? fileName.slice(0, dot) : fileName;
    const ext = dot > 0 ? fileName.slice(dot) : '';
    const promptMsg = warnXlsxRef
      ? 'Rename "' + fileName + '" to (extension stays "' + ext + '"). Remember: ' + warnXlsxRef + ' references this filename — update that row too, or the site will show a broken image.'
      : 'Rename "' + fileName + '" to (extension stays "' + ext + '"):';
    const newBase = prompt(promptMsg, base);
    if (newBase === null) return;
    const trimmed = newBase.trim();
    if (!trimmed) { alert('Name cannot be empty.'); return; }
    const newPath = dir + '/' + trimmed + ext;
    if (newPath === oldPath) return;
    try {
      const res = await fetch('/api/rename', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ oldPath: oldPath, newPath: newPath, message: 'Admin: rename ' + oldPath + ' -> ' + newPath })
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || 'Rename failed');
      loadFolder(folder);
    } catch (e) {
      alert(e.message);
    }
  }

  async function loadFolder(folder) {
    const container = document.getElementById('thumbs-' + folder.replace(/[^a-zA-Z0-9]/g, '-'));
    if (!container) return;
    const card = container.closest('.card');
    const isAuto = card && card.dataset.auto === '1';
    const xlsxRef = card ? card.dataset.xlsxRef : '';
    try {
      const res = await fetch('/api/list?folder=' + encodeURIComponent(folder));
      const items = await res.json();
      if (!res.ok) throw new Error((items && items.error) || 'Failed to load (' + res.status + ')');
      container.innerHTML = '';
      if (!items.length) {
        const span = document.createElement('span');
        span.style.cssText = 'font-size:12px;color:#999';
        span.textContent = 'No files yet.';
        container.appendChild(span);
        return;
      }
      items.forEach(function(it) {
        const isImg = /\\.(png|jpe?g|gif|webp|svg)$/i.test(it.name);
        const wrap = document.createElement('div');
        wrap.className = 'thumb';

        let media;
        if (isImg) {
          media = document.createElement('img');
          media.src = it.download_url;
          media.loading = 'lazy';
        } else {
          media = document.createElement('div');
          media.style.cssText = 'width:84px;height:84px;display:flex;align-items:center;justify-content:center;background:#eee;border-radius:6px;font-size:24px;';
          media.textContent = '📄';
        }
        wrap.appendChild(media);

        const displayName = it.path.startsWith(folder + '/') ? it.path.slice(folder.length + 1) : it.name;
        const nameEl = document.createElement('div');
        nameEl.className = 'name';
        nameEl.title = displayName;
        nameEl.textContent = displayName;
        wrap.appendChild(nameEl);

        const renameBtn = document.createElement('button');
        renameBtn.className = 'secondary rename-btn';
        renameBtn.textContent = 'Rename';
        renameBtn.onclick = function() { renameInFolder(it.path, folder, isAuto ? null : xlsxRef); };
        wrap.appendChild(renameBtn);

        const delBtn = document.createElement('button');
        delBtn.className = 'danger';
        delBtn.textContent = 'Delete';
        delBtn.onclick = function() { deleteFromFolder(it.path, folder); };
        wrap.appendChild(delBtn);

        container.appendChild(wrap);
      });
    } catch (e) {
      container.innerHTML = '';
      const span = document.createElement('span');
      span.style.cssText = 'font-size:12px;color:#a33';
      span.textContent = 'Could not load: ' + e.message;
      container.appendChild(span);
    }
  }

  // ---- activity log ---------------------------------------------------------

  async function loadLogs() {
    const tbody = document.getElementById('log-tbody');
    tbody.innerHTML = '<tr><td colspan="3" style="padding:14px;color:#999;">Loading…</td></tr>';
    try {
      const res = await fetch('/api/logs');
      const items = await res.json();
      if (!res.ok) throw new Error((items && items.error) || 'Failed to load log');
      if (!items.length) {
        tbody.innerHTML = '<tr><td colspan="3" style="padding:14px;color:#999;">No commits yet.</td></tr>';
        return;
      }
      tbody.innerHTML = '';
      items.forEach(function(c) {
        const tr = document.createElement('tr');

        const whenTd = document.createElement('td');
        whenTd.style.whiteSpace = 'nowrap';
        whenTd.textContent = c.date ? new Date(c.date).toLocaleString() : '—';

        const msgTd = document.createElement('td');
        msgTd.textContent = (c.message || '').split('\\n')[0];

        const linkTd = document.createElement('td');
        linkTd.style.whiteSpace = 'nowrap';
        const a = document.createElement('a');
        a.href = c.url;
        a.target = '_blank';
        a.rel = 'noopener';
        a.style.fontSize = '12px';
        a.textContent = 'View on GitHub';
        linkTd.appendChild(a);

        tr.appendChild(whenTd);
        tr.appendChild(msgTd);
        tr.appendChild(linkTd);
        tbody.appendChild(tr);
      });
    } catch (e) {
      tbody.innerHTML = '<tr><td colspan="3" style="padding:14px;color:#a33;">' + e.message + '</td></tr>';
    }
  }

  JSON.parse('${allFolderPaths}').forEach(loadFolder);
  loadLogs();
  </script>
  </body></html>`;
}

// ---------------------------------------------------------------------------
// Router
// ---------------------------------------------------------------------------

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const cookie = request.headers.get("Cookie") || "";

    try {
      if (url.pathname === "/login" && request.method === "POST") {
        const form = await request.formData();
        const password = form.get("password");
        if (password !== env.ADMIN_PASSWORD) {
          return new Response(loginPage("Wrong password."), {
            status: 401,
            headers: { "Content-Type": "text/html" },
          });
        }
        const token = await makeSession(env);
        return new Response(null, {
          status: 302,
          headers: { Location: "/admin", "Set-Cookie": setCookieHeader(token) },
        });
      }

      if (url.pathname === "/logout") {
        return new Response(null, {
          status: 302,
          headers: {
            Location: "/",
            "Set-Cookie": "session=; Path=/; HttpOnly; Secure; SameSite=Strict; Max-Age=0",
          },
        });
      }

      const authed = await verifySession(env, cookie);

      if (url.pathname === "/" || url.pathname === "/login") {
        if (authed) return Response.redirect(url.origin + "/admin", 302);
        return new Response(loginPage(), { headers: { "Content-Type": "text/html" } });
      }

      if (url.pathname === "/admin") {
        if (!authed) return Response.redirect(url.origin + "/", 302);
        return new Response(adminPage(), { headers: { "Content-Type": "text/html" } });
      }

      if (url.pathname.startsWith("/api/")) {
        if (!authed) return json({ error: "Not logged in" }, 401);

        if (url.pathname === "/api/list" && request.method === "GET") {
          const folder = url.searchParams.get("folder");
          const cfg = FOLDERS.find((f) => f.path === folder);
          if (!cfg) return json({ error: "Unknown folder" }, 400);
          const items = await ghListFolder(env, folder, !!cfg.yearFolders);
          return json(items.map((i) => ({ name: i.name, path: i.path, download_url: i.download_url })));
        }

        if (url.pathname === "/api/file" && request.method === "GET") {
          const path = url.searchParams.get("path");
          if (!isAllowedPath(path)) return json({ error: "Path not allowed" }, 400);
          const file = await ghGetFile(env, path);
          if (!file) return json({ error: "File not found" }, 404);
          return json(file);
        }

        if (url.pathname === "/api/commit" && request.method === "POST") {
          const body = await request.json();
          if (!isAllowedPath(body.path)) return json({ error: "Path not allowed" }, 400);
          const result = await ghCommitFile(env, body.path, body.contentBase64, body.message || `Admin update: ${body.path}`);
          return json({ ok: true, commit: result.commit?.sha });
        }

        if (url.pathname === "/api/logs" && request.method === "GET") {
          const perPage = Math.min(Number(url.searchParams.get("per_page")) || 40, 100);
          const commits = await ghListCommits(env, perPage);
          return json(commits);
        }

        if (url.pathname === "/api/rename" && request.method === "POST") {
          const body = await request.json();
          if (!isAllowedPath(body.oldPath) || !isAllowedPath(body.newPath)) {
            return json({ error: "Path not allowed" }, 400);
          }
          await ghRenameFile(env, body.oldPath, body.newPath, body.message || `Admin: rename ${body.oldPath} -> ${body.newPath}`);
          return json({ ok: true });
        }

        if (url.pathname === "/api/delete" && request.method === "POST") {
          const body = await request.json();
          if (!isAllowedPath(body.path)) return json({ error: "Path not allowed" }, 400);
          const result = await ghDeleteFile(env, body.path, body.message || `Admin delete: ${body.path}`);
          return json({ ok: true, commit: result.commit?.sha });
        }

        return json({ error: "Not found" }, 404);
      }

      return new Response("Not found", { status: 404 });
    } catch (err) {
      return json({ error: err.message || String(err) }, 500);
    }
  },
};

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), { status, headers: { "Content-Type": "application/json" } });
}

// Only allow writes to paths this admin panel is meant to manage —
// prevents the API being used to overwrite arbitrary repo files.
function isAllowedPath(path) {
  if (!path || typeof path !== "string") return false;
  if (XLSX_FILES.some((f) => f.path === path)) return true;
  if (SINGLE_DOCS.some((f) => f.path === path)) return true;
  return FOLDERS.some((f) => path === f.path || path.startsWith(f.path + "/"));
}
