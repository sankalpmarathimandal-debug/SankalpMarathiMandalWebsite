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
  { path: "data/events.xlsx", label: "Events & Performances", desc: "One sheet: homepage event cards, Event Timeline page, the homepage Book a Performance teaser, and the Shala Events section on the Marathi Shala page. Set 'Event Type' to Event or Performance; Event rows also need a Type (previous/current/future) to appear as one of the 3 featured homepage cards — leave Type blank for timeline-only events. To also show an event on the Shala page, set Audience to 'Shala' (leave blank for general Mandal events)." },
  { path: "data/testimonials.xlsx", label: "Testimonials", desc: "Homepage “Community Voices”" },
  { path: "data/sponsors.xlsx", label: "Presenting Sponsors", desc: "Homepage sponsor cards. To add a headshot photo: upload it under 'Sponsor Photos' below, then paste its path into AvatarURL (e.g. assets/images/partners/radhika.jpg) — leave blank for a generic icon. Columns: Active (Yes/No), DisplayOrder, Tier (e.g. Gold Partner — shown as a badge), Name, Role, AvatarURL (headshot, shown in the circle), AccentColor (hex, drives that card's colors), Icon (small Font Awesome solid icon name shown next to the name, e.g. chart-line or house-chimney), LogoURL (business logo image shown in the middle badge — also uploaded via 'Sponsor Photos'), BusinessName, Tagline, Category, Description, Tags (comma-separated), Website, Phone, Email, SocialLabel, SocialURL, CTAText, CTALink." },
  { path: "data/team.xlsx", label: "Team", desc: "Our Team page" },
  { path: "data/shala-team.xlsx", label: "Shala Team", desc: "Marathi Shala team section" },
  { path: "data/faq.xlsx", label: "FAQ", desc: "FAQs page" },
  { path: "data/shala-faq.xlsx", label: "Shala FAQ", desc: "Marathi Shala FAQs" },
  { path: "data/shala-guidelines-parents.xlsx", label: "Shala Guidelines — Parents", desc: "Parent handbook accordion on the Marathi Shala page. Columns: ID, Question (used as the accordion heading), Answer, Order, Active (Yes/No)." },
  { path: "data/shala-guidelines-teachers.xlsx", label: "Shala Guidelines — Teachers", desc: "Teacher handbook accordion on the Marathi Shala page. Same columns as the Parents guidelines above." },
  { path: "data/shala-calendar.xlsx", label: "Shala Calendar", desc: "Shala Calendar page + download button" },
  { path: "data/forms.xlsx", label: "Forms & Sign-ups", desc: "Forms & Sign-ups page" },
  { path: "data/showcase.xlsx", label: "Showcase", desc: "Showcase page entries" },
  { path: "data/programs.xlsx", label: "Performance Programs", desc: "Book a Performance page — the program menu" },
  { path: "data/program-participants.xlsx", label: "Performance Participants", desc: "Book a Performance page — singers/directors/etc per program (Program column must match a Title in Performance Programs exactly)" },
];

const SINGLE_DOCS = [
  { path: "docs/constitution.pdf", label: "Constitution PDF", desc: "Constitution page", accept: "application/pdf" },
];

// Small hand-edited JSON config files (not xlsx, not a folder of uploads) —
// each gets a dedicated form in the admin UI instead of the generic
// spreadsheet/upload card. Currently just the homepage marquee banner.
const SIMPLE_JSON_FILES = [
  { path: "data/marquee.json", label: "Homepage Announcement Marquee" },
];

const FOLDERS = [
  { path: "assets/images/events", label: "Event Images", accept: "image/*", yearFolders: true, note: "Organized by year — pick/type a year below", xlsxRef: "data/events.xlsx" },
  { path: "assets/images/team", label: "Team Photos", accept: "image/*", xlsxRef: "data/team.xlsx" },
  { path: "assets/images/shala", label: "Shala Images", accept: "image/*" },
  { path: "assets/images/showcase", label: "Showcase Photos", accept: "image/*", xlsxRef: "data/showcase.xlsx" },
  { path: "assets/images/partners", label: "Sponsor Photos", accept: "image/*", xlsxRef: "data/sponsors.xlsx" },
  {
    path: "assets/images/branding/logo-variants",
    label: "Community Pride Wall Logos",
    accept: "image/*",
    note: "Auto-updates homepage — no xlsx edit needed",
    autoManifest: true,
  },
  {
    path: "assets/images/culture-icons",
    label: "Homepage Culture Icon Ribbon",
    accept: "image/*",
    note: "Auto-updates homepage — no xlsx edit needed",
    autoManifest: true,
    namingHint: "No captions are shown for these — file names are just for your own organization, upload/delete freely.",
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
  { path: "assets/images/programs", label: "Performance Program Logos/Photos", accept: "image/*", xlsxRef: "data/programs.xlsx" },
];

// Groups everything above by the actual page on the live site it affects —
// this drives the "which page do you want to change?" selector.
const PAGES = [
  { key: "home", label: "Home Page", xlsx: ["data/events.xlsx", "data/testimonials.xlsx", "data/sponsors.xlsx"], folders: ["assets/images/highlights", "assets/images/branding/logo-variants", "assets/images/culture-icons", "assets/images/partners"] },
  { key: "events", label: "Events Page", xlsx: ["data/events.xlsx"], folders: ["assets/images/events"] },
  { key: "team", label: "Team Page", xlsx: ["data/team.xlsx"], folders: ["assets/images/team"] },
  { key: "shala", label: "Shala Page", xlsx: ["data/shala-team.xlsx", "data/shala-faq.xlsx", "data/shala-guidelines-parents.xlsx", "data/shala-guidelines-teachers.xlsx"], folders: ["assets/images/shala"] },
  { key: "calendar", label: "Shala Calendar", xlsx: ["data/shala-calendar.xlsx"], folders: [] },
  { key: "faq", label: "FAQ Page", xlsx: ["data/faq.xlsx"], folders: [] },
  { key: "forms", label: "Forms & Sign-ups", xlsx: ["data/forms.xlsx"], folders: [] },
  { key: "showcase", label: "Showcase Page", xlsx: ["data/showcase.xlsx"], folders: ["assets/images/showcase", "docs/showcase"] },
  { key: "performances", label: "Book a Performance", xlsx: ["data/programs.xlsx", "data/program-participants.xlsx"], folders: ["assets/images/programs"] },
  { key: "constitution", label: "Constitution Page", xlsx: [], folders: [], docs: ["docs/constitution.pdf"] },
];

// A second, restricted login (SHALA_ADMIN_PASSWORD) only sees these PAGES
// keys — Shala team/FAQ/guidelines and the Shala calendar. Deliberately
// does NOT include "home" or "events", so the Shala-only login can never
// touch data/events.xlsx (it's shared with general Mandal events) — an
// admin with the full ADMIN_PASSWORD still tags an event Audience=Shala
// to make it show up on the Shala page.
const SHALA_ROLE_PAGE_KEYS = ["shala", "calendar"];

// ---------------------------------------------------------------------------
// Event Flyer Builder — a fully self-contained, client-side flyer design
// tool (background/hero image, sponsors, QR code, PDF export). Served at
// /flyer, behind the same login as the rest of the admin panel. No GitHub
// commits happen here — it's pure browser-side editing + PDF/HTML download,
// so nothing here can touch the live site by itself.
// ---------------------------------------------------------------------------

const FLYER_BUILDER_HTML = String.raw`<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Event Flyer Builder — Sankalp Marathi Mandal</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,400;0,500;0,600;0,700;0,800;0,900&family=Tiro+Devanagari+Marathi&display=swap" rel="stylesheet">
<style>
  :root{
    --gold:#e3ab3c;
    --gold-light:#f6dfa4;
    --cream:#fdf3dd;
    --saffron:#e2711d;
    --saffron-deep:#c85a10;
    --gray:#6b6560;
    --green:#2f8a4e;
  }
  *{margin:0;padding:0;box-sizing:border-box;}
  body{
    background:#e9e4da;
    font-family:'Poppins', Arial, sans-serif;
    display:flex; flex-direction:column; align-items:center;
    padding:24px 8px 60px;
    color:#231f1f;
  }

  /* ---------- Toolbar (never printed / exported) ---------- */
  .toolbar{
    width:820px; max-width:100%;
    background:#fff;
    border:1px solid #d8d2c4;
    border-radius:10px;
    padding:16px 20px;
    margin-bottom:18px;
    box-shadow:0 4px 16px rgba(0,0,0,.06);
  }
  .toolbar h1{ font-size:17px; font-weight:700; margin-bottom:4px; }
  .toolbar p.sub{ font-size:12.5px; color:var(--gray); margin-bottom:12px; line-height:1.5; }
  .toolbar-buttons{ display:flex; flex-wrap:wrap; gap:8px; margin-bottom:10px; }
  .btn{
    font-family:inherit; font-size:13px; font-weight:600;
    padding:9px 14px; border-radius:7px; border:1px solid #d8d2c4;
    background:#fff; color:#231f1f; cursor:pointer;
    transition:background .15s, transform .05s;
  }
  .btn:hover{background:#f3efe6;}
  .btn:active{transform:scale(.97);}
  .btn-primary{background:var(--saffron); border-color:var(--saffron-deep); color:#fff;}
  .btn-primary:hover{background:var(--saffron-deep);}
  .btn-gold{background:var(--gold); border-color:var(--saffron-deep); color:#3a1a02;}
  .btn-gold:hover{background:var(--saffron);}
  .btn-danger{color:#a83232;}
  .btn:disabled{opacity:.5; cursor:not-allowed;}
  .status-msg{ font-size:12.5px; font-weight:600; min-height:16px; color:var(--green); }
  details.help{ font-size:12.5px; border-top:1px dashed #d8d2c4; padding-top:10px; margin-top:4px; }
  details.help summary{cursor:pointer; font-weight:600; color:var(--saffron-deep);}
  details.help ol{margin:8px 0 0 18px; line-height:1.7;}

  /* background controls row */
  .bg-controls{
    display:flex; align-items:center; flex-wrap:wrap; gap:8px;
    padding:10px 0; border-top:1px dashed #d8d2c4; border-bottom:1px dashed #d8d2c4;
    margin-bottom:10px;
  }
  .bg-controls-label{ font-size:12.5px; font-weight:700; margin-right:2px; }
  .swatches{ display:flex; gap:6px; }
  .swatch{
    width:26px; height:26px; border-radius:50%; border:2px solid #fff;
    box-shadow:0 0 0 1.5px #d8d2c4; cursor:pointer; padding:0;
  }
  .swatch.active{ box-shadow:0 0 0 2px var(--saffron); }
  .color-picker-wrap{ display:flex; align-items:center; gap:6px; font-size:12px; color:var(--gray); }
  #bgColorPicker{ width:34px; height:26px; border:1px solid #d8d2c4; border-radius:6px; cursor:pointer; padding:0; }

  /* ---------- Flyer ---------- */
  .flyer{
    width:820px; max-width:100%;
    border:10px solid var(--gold);
    outline:2px solid rgba(0,0,0,.4);
    box-shadow:0 24px 60px rgba(0,0,0,.3);
  }
  .garland{
    height:20px;
    background:repeating-linear-gradient(90deg, var(--saffron) 0 14px, #e2600f 14px 28px, var(--gold) 28px 42px);
  }
  .garland.bottom{ transform:rotate(180deg); }

  /* ---------- Stage: customizable background (color or image) lives here ---------- */
  .stage{ position:relative; overflow:hidden; }
  .bg-color-layer{ position:absolute; inset:0; z-index:0; }
  .bg-image{
    position:absolute; inset:0; width:100%; height:100%; object-fit:cover;
    display:none; z-index:1;
  }
  .stage.has-bg-image .bg-image{ display:block; }
  .bg-overlay{
    position:absolute; inset:0; z-index:2; pointer-events:none;
    background:linear-gradient(180deg, rgba(0,0,0,.18) 0%, rgba(0,0,0,.04) 35%, rgba(0,0,0,.22) 100%);
  }
  .bg-change-btn{
    position:absolute; top:14px; right:14px; z-index:6;
    background:rgba(0,0,0,.5); color:var(--gold-light); border:1px solid var(--gold);
    font-size:11.5px; font-weight:600; padding:7px 12px; border-radius:20px; cursor:pointer;
  }
  .bg-change-btn:hover{ background:rgba(0,0,0,.7); }

  .inner{ position:relative; z-index:5; padding:28px 36px 32px; }

  /* panel = translucent neutral card so text/logos stay legible over any background */
  .panel{
    background:rgba(12,9,7,.58);
    border:1px solid rgba(227,171,60,.5);
    border-radius:12px;
    padding:16px 20px;
    margin-bottom:18px;
  }

  /* brand header */
  .brand-header{ display:flex; align-items:center; justify-content:center; }
  .brand-logo-wrap{
    width:88px; height:88px; border-radius:50%; background:#fff; flex:0 0 auto;
    border:3px solid var(--gold); box-shadow:0 3px 10px rgba(0,0,0,.35);
    display:flex; align-items:center; justify-content:center; margin-right:18px; overflow:hidden;
  }
  .brand-logo-wrap img{ width:82%; height:82%; object-fit:contain; }
  .brand-text{ text-align:center; }
  .brand-text .mr{
    font-family:'Tiro Devanagari Marathi', serif;
    font-size:25px; color:var(--gold-light); line-height:1.2;
  }
  .brand-text .en{
    font-size:12px; letter-spacing:3px; text-transform:uppercase;
    color:var(--cream); opacity:.9; margin-top:3px;
  }
  .presents{ font-size:12.5px; font-style:italic; color:var(--gold-light); margin-top:8px; }

  /* hero image — the main event photo, full clarity, its own frame */
  .hero-label{
    text-align:center; font-size:10.5px; font-weight:700; letter-spacing:2px;
    text-transform:uppercase; color:var(--gold-light); text-shadow:0 1px 4px rgba(0,0,0,.6); margin-bottom:8px;
  }
  .hero{
    position:relative; width:100%; padding-top:60%;
    border:4px solid var(--gold); border-radius:4px;
    background:#1c1c1c;
    cursor:pointer; margin-bottom:6px; overflow:hidden;
    box-shadow:0 6px 18px rgba(0,0,0,.4);
  }
  .hero img{
    position:absolute; inset:0; width:100%; height:100%; object-fit:cover; display:none;
  }
  .hero.has-image img{ display:block; }
  .hero.has-image .hero-placeholder{ display:none; }
  .hero-placeholder{
    position:absolute; inset:0; display:flex; flex-direction:column;
    align-items:center; justify-content:center; text-align:center;
    color:var(--gold-light); padding:20px; pointer-events:none;
  }
  .hero-placeholder .icon{ font-size:40px; display:block; margin-bottom:10px; }
  .hero-placeholder .txt{ font-size:15px; font-weight:700; }
  .hero-placeholder .sub{ font-size:11.5px; margin-top:5px; color:var(--gold); opacity:.85; }
  .hero-change-btn{
    position:absolute; top:12px; right:12px; z-index:5;
    background:rgba(0,0,0,.55); color:var(--gold-light); border:1px solid var(--gold);
    font-size:11.5px; font-weight:600; padding:7px 12px; border-radius:20px; cursor:pointer;
  }
  .hero-change-btn:hover{ background:rgba(0,0,0,.75); }
  .hero-cap{ text-align:center; font-size:10.5px; color:var(--gold-light); text-shadow:0 1px 3px rgba(0,0,0,.6); margin-bottom:20px; }

  /* price ribbon */
  .badge-row{ display:flex; justify-content:center; margin-bottom:8px; }
  .badge-control{ display:inline-flex; border:1px solid #d8d2c4; border-radius:20px; overflow:hidden; }
  .badge-control label{ font-size:11.5px; font-weight:600; cursor:pointer; color:var(--gray); background:#fff; }
  .badge-control input{ display:none; }
  .badge-control label span{ display:block; padding:5px 14px; }
  .badge-control input:checked + span{ background:var(--saffron); color:#fff; }
  .price-ribbon{ display:flex; justify-content:center; margin-top:4px; }
  .price-ribbon .pill{
    display:inline-block; padding:10px 28px; border-radius:22px;
    font-size:14px; font-weight:800; letter-spacing:.6px;
    background:var(--gold); color:#2a1704;
    border:2px solid var(--gold-light); box-shadow:0 3px 10px rgba(0,0,0,.3);
    min-width:170px; text-align:center;
  }

  /* title */
  .title-block{ text-align:center; }
  .event-name{
    font-family:'Poppins', sans-serif; font-weight:800; text-transform:uppercase;
    letter-spacing:1px; font-size:32px; line-height:1.25; color:#fff;
    outline:none; text-shadow:0 2px 8px rgba(0,0,0,.5);
  }
  .event-subtitle{ font-size:14.5px; color:var(--gold-light); opacity:.95; margin-top:9px; font-style:italic; }

  /* editable placeholder styling */
  .editable{ outline-offset:3px; border-radius:4px; }
  .editable:hover{ background:rgba(227,171,60,.18); }
  .editable:focus{ background:rgba(227,171,60,.24); outline:1.5px dashed var(--gold); }
  .editable.is-empty:before{ content:attr(data-placeholder); color:rgba(253,243,221,.5); font-weight:400; font-style:normal; text-transform:none; letter-spacing:normal; }

  /* details — flex based (no gap, html2canvas-safe) */
  .details-row{ display:flex; flex-wrap:wrap; justify-content:space-between; }
  .detail{ flex:0 0 47%; margin-bottom:14px; }
  .detail.wide{ flex:0 0 100%; }
  .detail:last-child, .detail.wide:last-child{ margin-bottom:0; }
  .detail .label{
    font-size:10.5px; font-weight:700; letter-spacing:1.6px; text-transform:uppercase;
    color:var(--gold); margin-bottom:4px;
  }
  .detail .value{ font-size:15.5px; line-height:1.5; color:#fff; }

  /* sponsors — big, opaque white tiles so logos are always visible */
  .sponsors-heading{
    text-align:center; font-size:12px; font-weight:700; letter-spacing:2.5px;
    text-transform:uppercase; color:var(--gold-light); margin-bottom:16px;
  }
  .sponsors-heading .line{ display:inline-block; width:50px; height:1px; background:var(--gold); vertical-align:middle; margin:0 12px; }
  .sponsors-heading .editable{ display:inline-block; vertical-align:middle; }
  .sponsors-grid{ display:flex; flex-wrap:wrap; justify-content:center; }
  .sponsor-slot{ width:158px; text-align:center; position:relative; margin:0 12px 18px; }
  .sponsor-logo-upload{
    width:150px; height:106px; margin:0 auto 8px;
    border:1px solid #e8e2d2; border-radius:10px;
    background:#ffffff; box-shadow:0 3px 10px rgba(0,0,0,.3);
    display:flex; align-items:center; justify-content:center;
    cursor:pointer; overflow:hidden;
  }
  .sponsor-logo-upload img{ width:100%; height:100%; object-fit:contain; display:none; padding:8px; }
  .sponsor-logo-upload.has-image img{ display:block; }
  .sponsor-logo-upload.has-image .up-placeholder{ display:none; }
  .up-placeholder{ font-size:11px; font-weight:600; color:var(--saffron-deep); line-height:1.4; }
  .up-placeholder .icon{ font-size:22px; display:block; }
  .sponsor-name{ font-size:12.5px; font-weight:700; min-height:16px; color:#fff; }
  .remove-sponsor{
    position:absolute; top:-8px; right:8px; width:20px; height:20px; border-radius:50%;
    background:#a83232; color:#fff; border:2px solid #fff; font-size:12px; line-height:1;
    cursor:pointer; display:flex; align-items:center; justify-content:center;
  }
  .add-sponsor-tile{ width:158px; text-align:center; margin:0 12px 18px; display:inline-block; }
  .add-sponsor-tile button{
    width:150px; height:106px; border:2px dashed var(--gold-light); background:transparent;
    border-radius:10px; color:var(--gold-light); font-size:26px; cursor:pointer;
    display:flex; align-items:center; justify-content:center;
  }
  .add-sponsor-tile .lbl{ font-size:11px; color:var(--gold-light); opacity:.85; margin-top:7px; }
  .sponsors-panel .sponsors-grid{ margin-bottom:0; }

  /* payment / QR panel */
  .payment-panel.hidden{ display:none; }
  .payment-heading{
    text-align:center; font-size:12px; font-weight:700; letter-spacing:2.5px;
    text-transform:uppercase; color:var(--gold-light); margin-bottom:16px;
  }
  .payment-heading .line{ display:inline-block; width:50px; height:1px; background:var(--gold); vertical-align:middle; margin:0 12px; }
  .payment-heading .editable{ display:inline-block; vertical-align:middle; }
  .payment-row{ display:flex; flex-wrap:wrap; align-items:flex-start; justify-content:center; }
  .qr-upload{
    width:132px; height:132px; flex:0 0 auto; margin:0 20px 14px 0;
    border:2px solid var(--gold); border-radius:8px;
    background:#ffffff; box-shadow:0 3px 10px rgba(0,0,0,.3);
    display:flex; align-items:center; justify-content:center;
    cursor:pointer; overflow:hidden;
  }
  .qr-upload img{ width:100%; height:100%; object-fit:contain; display:none; padding:6px; }
  .qr-upload.has-image img{ display:block; }
  .qr-upload.has-image .up-placeholder{ display:none; }
  .payment-fields{ flex:1 1 220px; min-width:220px; margin-bottom:14px; }
  .payment-fields .detail{ flex:0 0 100%; margin-bottom:14px; }
  .payment-fields .detail:last-child{ margin-bottom:0; }
  .payment-toggle-row{
    display:flex; align-items:center; gap:8px; font-size:12.5px; color:var(--gray);
    padding:10px 0; border-top:1px dashed #d8d2c4; border-bottom:1px dashed #d8d2c4; margin-bottom:10px;
  }
  .payment-toggle-row input{ width:16px; height:16px; cursor:pointer; }
  .payment-toggle-row label{ cursor:pointer; font-weight:600; }

  /* footer */
  .footer{ text-align:center; }
  .footer .contact{ font-size:13px; color:#fff; margin-bottom:8px; }
  .footer .org{ font-size:11.5px; font-weight:700; letter-spacing:2.5px; color:var(--gold); }

  .file-input{ display:none; }

  @media print{
    body{ background:#fff; padding:0; }
    @page{ size:auto; margin:8mm; }
    .toolbar{ display:none; }
    .flyer{ box-shadow:none; width:100%; }
    .no-export{ display:none !important; }
  }
  body.exporting .toolbar{ display:none; }
  body.exporting .no-export{ display:none !important; }
  body.exporting .editable:hover,
  body.exporting .editable:focus{ background:transparent; outline:none; }
</style>
</head>
<body>

<div class="toolbar no-export">
  <h1>Event Flyer Builder</h1>
  <p class="sub">
    Click any text to edit it. Click the hero image box to upload the event's main photo.
    <strong>Download PDF</strong> gives a guaranteed single-page export — Print is only a backup and may
    split long flyers across pages.
  </p>

  <div class="bg-controls">
    <span class="bg-controls-label">Background color:</span>
    <div class="swatches" id="swatches">
      <button type="button" class="swatch active" data-color="#4a0d12" style="background:#4a0d12" title="Maroon &amp; Gold (classic)"></button>
      <button type="button" class="swatch" data-color="#0d2b4a" style="background:#0d2b4a" title="Deep Navy"></button>
      <button type="button" class="swatch" data-color="#1f3d1f" style="background:#1f3d1f" title="Forest Green"></button>
      <button type="button" class="swatch" data-color="#3a1246" style="background:#3a1246" title="Royal Purple"></button>
      <button type="button" class="swatch" data-color="#1c1c1c" style="background:#1c1c1c" title="Charcoal Black"></button>
      <button type="button" class="swatch" data-color="#7a3210" style="background:#7a3210" title="Terracotta"></button>
    </div>
    <div class="color-picker-wrap">
      <input type="color" id="bgColorPicker" value="#4a0d12" title="Custom background color">
      <span>Custom</span>
    </div>
    <button type="button" class="btn" id="bgImageBtn">🖼 Background Image</button>
    <button type="button" class="btn" id="clearBgImageBtn">✕ Remove Background Image</button>
  </div>

  <div class="payment-toggle-row">
    <input type="checkbox" id="paymentToggle">
    <label for="paymentToggle">Include Payment / RSVP section (optional QR code, payment link, instructions — shows above the sponsor logos)</label>
  </div>

  <div class="toolbar-buttons">
    <button type="button" class="btn btn-gold" id="addHeroBtn">📷 Add / Change Hero Image</button>
    <button type="button" class="btn btn-gold" id="addQrBtn">▦ Add / Change QR Code</button>
    <button type="button" class="btn btn-primary" id="downloadPdfBtn">⬇ Download PDF</button>
    <button type="button" class="btn" id="printBtn">🖶 Print</button>
    <button type="button" class="btn" id="saveHtmlBtn">💾 Save Editable Copy</button>
    <button type="button" class="btn" id="copyJsonBtn">🔗 Copy Details for Website</button>
    <button type="button" class="btn btn-danger" id="resetBtn">↺ Reset Template</button>
  </div>
  <div class="status-msg" id="statusMsg"></div>
  <details class="help">
    <summary>How to use this template for a new event</summary>
    <ol>
      <li>Copy this file into the event's folder and rename it, e.g. <em>Diwali-Sammelan-Flyer.html</em>.</li>
      <li>Open it in a browser (double‑click the file).</li>
      <li>Pick a background color swatch (or your own custom color), or click <strong>Background Image</strong> to use a photo behind the whole flyer instead.</li>
      <li>Click the hero image box and choose the event's main photo from this same folder — this stays sharp and undimmed regardless of the background.</li>
      <li>Click the event name, date, time, venue, and instructions to type them in directly.</li>
      <li>Choose Free or Paid, and enter the fee if it's a paid event (this also switches on the Payment / RSVP section below).</li>
      <li>Optional: tick "Include Payment / RSVP section" to add a QR code, payment link, and pay instructions — it appears above the sponsor logos.</li>
      <li>Add each sponsor's logo and name (up to 5) using "+ Add Sponsor".</li>
      <li>Click <strong>Download PDF</strong> to export a single-page flyer, and/or <strong>Save Editable Copy</strong> to keep a filled‑in HTML version in the folder.</li>
      <li>Use <strong>Copy Details for Website</strong> to grab the event info as text your web team can paste onto the website.</li>
    </ol>
  </details>
</div>

<div class="flyer" id="flyer">
  <div class="garland"></div>

  <div class="stage" id="stage">
    <div class="bg-color-layer" id="bgColorLayer"></div>
    <img class="bg-image" id="bgImage" alt="Flyer background">
    <div class="bg-overlay"></div>
    <button type="button" class="bg-change-btn no-export" id="bgChangeBtn">🖼 Change Background</button>
    <input type="file" accept="image/*" class="file-input" id="bgInput">

    <div class="inner">

      <div class="brand-header panel">
        <div class="brand-logo-wrap"><img id="brandLogo" src="" alt="Sankalp Marathi Mandal logo"></div>
        <div class="brand-text">
          <div class="mr">संकल्प मराठी मंडळ</div>
          <div class="en">Sankalp Marathi Mandal</div>
          <div class="presents">proudly presents</div>
        </div>
      </div>

    
      <div class="hero" id="heroBox">
        <img id="heroImg" alt="Event hero image">
        <div class="hero-placeholder">
          <span class="icon">🪔</span>
          <span class="txt">Click to add hero image</span>
          <span class="sub">recommended 1400 × 870px, from this event's folder</span>
        </div>
        <button type="button" class="hero-change-btn no-export" id="heroChangeBtn">📷 Change</button>
        <input type="file" accept="image/*" class="file-input" id="heroInput">
      </div>
      

      <div class="title-block panel">
        <div class="badge-row no-export">
          <div class="badge-control">
            <label><input type="radio" name="priceState" value="free" id="radioFree" checked><span>Free Event</span></label>
            <label><input type="radio" name="priceState" value="paid" id="radioPaid"><span>Paid Event</span></label>
          </div>
        </div>
        <div class="price-ribbon" style="margin-bottom:16px;">
          <span class="pill editable" id="priceText" contenteditable="true" data-placeholder="FREE ADMISSION"></span>
        </div>
        <div class="event-name editable" id="eventName" contenteditable="true" data-placeholder="CLICK TO ENTER EVENT NAME"></div>
        <div class="event-subtitle editable" id="eventSubtitle" contenteditable="true" data-placeholder="Optional subtitle or short description"></div>
      </div>

      <div class="details-row panel">
        <div class="detail">
          <div class="label">Date</div>
          <div class="value editable" id="fieldDate" contenteditable="true" data-placeholder="Saturday, Month Date, 2026"></div>
        </div>
        <div class="detail">
          <div class="label">Time</div>
          <div class="value editable" id="fieldTime" contenteditable="true" data-placeholder="5:00 PM onwards"></div>
        </div>
        <div class="detail wide">
          <div class="label">Venue</div>
          <div class="value editable" id="fieldVenue" contenteditable="true" data-placeholder="Venue name · Street address, Town, MA"></div>
        </div>
        <div class="detail wide">
          <div class="label">Special Instructions</div>
          <div class="value editable" id="fieldNote" contenteditable="true" data-placeholder="Parking, RSVP, dress code, potluck, seating, etc."></div>
        </div>
      </div>

      <div class="panel payment-panel hidden" id="paymentPanel">
        <div class="payment-heading">
          <span class="line"></span>
          <span class="editable" id="paymentHeading" contenteditable="true" data-placeholder="Scan to Pay / RSVP"></span>
          <span class="line"></span>
        </div>
        <div class="payment-row">
          <div class="qr-upload" id="qrBox">
            <img id="qrImg" alt="Payment / RSVP QR code">
            <div class="up-placeholder"><span class="icon">▦</span>QR Code<br>(optional)</div>
            <input type="file" accept="image/*" class="file-input" id="qrInput">
          </div>
          <div class="payment-fields">
            <div class="detail">
              <div class="label">Payment Link</div>
              <div class="value editable" id="fieldPayLink" contenteditable="true" data-placeholder="e.g. venmo.com/@SankalpMM or PayPal.me/..."></div>
            </div>
            <div class="detail">
              <div class="label">How to Pay</div>
              <div class="value editable" id="fieldPayInstructions" contenteditable="true" data-placeholder="Scan the QR code or use the link above. Please include your name in the notes."></div>
            </div>
          </div>
        </div>
      </div>

      <div class="panel sponsors-panel">
        <div class="sponsors-heading">
          <span class="line"></span>
          <span class="editable" id="sponsorsTitle" contenteditable="true" data-placeholder="With thanks to our sponsors"></span>
          <span class="line"></span>
        </div>
        <div class="sponsors-grid" id="sponsorsGrid"></div>
        <div style="text-align:center;" class="no-export">
          <div class="add-sponsor-tile" id="addSponsorTile">
            <button type="button" id="addSponsorBtn">+</button>
            <div class="lbl">Add Sponsor<br>(max 5)</div>
          </div>
        </div>
      </div>

      <div class="footer panel">
        <div class="contact editable" id="fieldContact" contenteditable="true" data-placeholder="https://www.sankalpmarathi.org · sankalpmarathimandal@gmail.com"></div>
        <div class="org">SANKALP MARATHI MANDAL</div>
      </div>

    </div>
  </div>

  <div class="garland bottom"></div>
</div>

<template id="sponsorTemplate">
  <div class="sponsor-slot">
    <button type="button" class="remove-sponsor no-export" title="Remove sponsor">×</button>
    <div class="sponsor-logo-upload">
      <img alt="Sponsor logo">
      <div class="up-placeholder"><span class="icon">＋</span>Logo</div>
      <input type="file" accept="image/*" class="file-input">
    </div>
    <div class="sponsor-name editable" contenteditable="true" data-placeholder="Sponsor Name"></div>
  </div>
</template>

<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script>
(function(){

  document.getElementById('brandLogo').src = "data:image/jpeg;base64,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2xwam96/TKdaNSKnB3TPnHFrRj6KKK1JCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAY6k4K8MDx/hSqwYZFOqNvkbcOh+9/jQMkooooEFFFFAH/0v34ooorQzCkYhQSfypajX523dVHT3PrQAqLj5m5Y9f8+lPoooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiikyOtABRmkJxzXlvi74l2Gh+ZY6UVu74ZU4/1cZ/2iOrf7I/HFeJnef4LKcO8Tjp8q/F+nc7cHgauKqKnRjdnc61r+maBaG81OZYoxnAPLMfRR1J9hX5IftT6da33jg+OtKtTa2usARzrxkzxLje2OMyIATz1B619W6pq+o6xdNe6lO08zdz0UeijoB7CuB8ceGovFnhi90VuJJE3wP3WVPmT8zwfY1+CPxpr185pKnHlw97O+7T0u+x93Hg2NPCScnepbTsv8z87qKkmikt5XgnQxyRsUdW42sDgj8CKjr+kITU4KUdmfnsk07M3dCvPIufIbhJuB6Bv/AK9doa8vBYEMpwQePrXoOnXYvbVJs84ww9xX4/4j5L7OrHMKS0ej9eh/QfhFxH7WjLKa71jrH06r5F6iiivy0/awooooAKKKKACiiigAooooAKKKKACiiigAoopfc0AAGaCKkUMxCICzE4CgZJPoAOpr27wR+zz8TPG3lXEen/2TYyHP2rUMx5X1WPBd89uAPeu3B5ficXP2eGg5M4MyzfA5fT9rjKigvNnh2PzFd74K+F/jj4gTrH4Y0uWeEkg3cg8u3XHrK2FP0XJ9q+/vAn7LPgDwsYbzXlfxBfxnduuRtgU/7MIJBH++Wr6UtbS1soI7WzhSCGNcIkShVUegUYA/Kv0bKPDarP38wlZdlv8Aefj2f+MVKN6eU0+Z/wA0tvkt/vPkL4f/ALI/hvShFf8Aju6bWLpcMbSIGO2U+jc75PxIHtX1vpekaXotjFp2kWsNnawjakMChEUewUAVoUtfp+WZLgsBDkw1NLz6/efi2ccQZhmdT2mNquXl0XothNtGMUtFetY8W4UUUUrCuJjtVG50vT7v/j4gRj64wfzHNX6KipShNWmrlKbWzONu/CMDZazlMZ/uvyPz61zdzoGp2pyYTIv9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tPzLlxPX5bKKueK2vwwvHUNd3scZ7qiFsH6nGa3LT4Z2EEiyyXs7MpBG0Kv9DXpuBSnmvpsLwBklCzjRu11bZ59TOsXPeZ+af7avwqi0t9M+JOjQN5MoGn6mRzhxzbyt/vDKH6LX5/1+/fj/AMHad4/8Hat4P1NR5Op2zRBj/BJ1jkHoUcBh9K/Hab9mr45QzSQ/8IleyeWxTfGYirbTjcMv0OMivr4U1CKjFWSPOcm9WeG0V7d/wzd8cv8AoT9Q/OL/AOOUf8M3fHL/AKE/UPzi/wDjlVZiPEa+lf2WPiZ/wrz4n2tpfziLSPEO3T7zd91ZCf8AR5T6bZDtJ9GNcx/wzd8cv+hP1D84v/jlKv7OHx0Vgy+ENQUjkEGLIPqD5nagD9vB0pa4H4X3vii/8A6JP40s5bDXFtVivoZsb/Njyhf5SR+827/xrvqshhRRRQIKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooqMsSdqckdT2FACs23gDJPIFCrzubk/y+lKqbfcnqT1NOoGFFFFAgooooAKKKKACiiigAooooAKKKKACiiigD/9P9+KKKK0MwooooAKKKKACiiigAooooAKKKKACiiigAooooAKjKnO5Ov6H/AOvUlFADVcMcHgjqKdTWUN17d+9N3lOJOn97/H0oGSUUZzRQIKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKTHOaWigdxMD3paKKAuFFFFAgoxiiiiw7hRRRSsFwooopiCiiihoApMUtFKw+YaV5p2BiiiiwgooopgFFFFABRgUUUAGBRgUUUAGBRgUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUcUAFITgZNNZwvHVj0FIFLHdJzjoB0/+vQMPmkHBIX17n/61PACgKowBS0UBcKKKKBBRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQB//U/fiiiitDMKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAo68UUUARbCnKdP7v+FPVw31HUHrTqayK/XgjoR1FAx1FR7mTh+R/eH9RUgORkUCCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKazgfKOT6CgB1R7mc4j4H94/0pNhY5fn/Z7f/XqWgY1UC+5PU072oooEFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAf//V/fiiiitDMKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACo9mOUOM9uxqSigBgcZw3yn3/oafSEAjBGc1GQ0Y3Kcr6Ht9DQNktFIrBhuHQ0tAgooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACkZggyTgVGZCX8tevXJ9KeqAfN1PqaB2Ggu/wDsj9f/AK1OVQvAp1FAgooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP/9k=";

  function setStatus(msg, isError){
    var el = document.getElementById('statusMsg');
    el.textContent = msg;
    el.style.color = isError ? '#a83232' : '#2f8a4e';
    if(msg){ setTimeout(function(){ if(el.textContent === msg) el.textContent = ''; }, 6000); }
  }

  function slugify(text){
    text = (text || 'Event').trim();
    if(!text) text = 'Event';
    return text.replace(/[^a-z0-9]+/gi, '-').replace(/^-+|-+$/g,'').slice(0,60) || 'Event';
  }

  /* ---------- placeholder handling for contenteditable fields ---------- */
  function refreshEmpty(el){
    var text = el.textContent.replace(/​/g,'').trim();
    el.classList.toggle('is-empty', text.length === 0);
  }
  function wireEditable(el){
    refreshEmpty(el);
    el.addEventListener('input', function(){ refreshEmpty(el); });
    el.addEventListener('blur', function(){ refreshEmpty(el); });
    el.addEventListener('paste', function(e){
      e.preventDefault();
      var text = (e.clipboardData || window.clipboardData).getData('text/plain');
      document.execCommand('insertText', false, text);
    });
  }
  document.querySelectorAll('.editable[contenteditable]').forEach(wireEditable);

  /* ---------- image upload helper (click + drag/drop) ---------- */
  function bindImageUpload(box, input, img, onLoaded){
    function handleFile(file){
      if(!file || file.type.indexOf('image/') !== 0) return;
      var reader = new FileReader();
      reader.onload = function(evt){
        img.src = evt.target.result;
        if(onLoaded) onLoaded(); else box.classList.add('has-image');
      };
      reader.readAsDataURL(file);
    }
    if(box){
      box.addEventListener('click', function(e){
        if(e.target.closest('.remove-sponsor') || e.target.closest('.hero-change-btn') || e.target.closest('.bg-change-btn')) return;
        input.click();
      });
      box.addEventListener('dragover', function(e){ e.preventDefault(); e.stopPropagation(); box.style.opacity='.75'; });
      box.addEventListener('dragleave', function(){ box.style.opacity='1'; });
      box.addEventListener('drop', function(e){
        e.preventDefault(); e.stopPropagation(); box.style.opacity='1';
        if(e.dataTransfer.files && e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
      });
    }
    input.addEventListener('change', function(e){ handleFile(e.target.files[0]); });
    return handleFile;
  }

  /* ---------- hero image ---------- */
  var heroBox = document.getElementById('heroBox');
  var heroInput = document.getElementById('heroInput');
  var heroImg = document.getElementById('heroImg');
  bindImageUpload(heroBox, heroInput, heroImg);
  document.getElementById('addHeroBtn').addEventListener('click', function(){ heroInput.click(); });
  document.getElementById('heroChangeBtn').addEventListener('click', function(){ heroInput.click(); });

  /* ---------- payment / QR section ---------- */
  var paymentPanel = document.getElementById('paymentPanel');
  var paymentToggle = document.getElementById('paymentToggle');
  function showPaymentPanel(){
    paymentPanel.classList.remove('hidden');
    paymentToggle.checked = true;
  }
  function setPaymentPanelVisible(visible){
    paymentPanel.classList.toggle('hidden', !visible);
    paymentToggle.checked = visible;
  }
  paymentToggle.addEventListener('change', function(){ setPaymentPanelVisible(this.checked); });

  var qrBox = document.getElementById('qrBox');
  var qrInput = document.getElementById('qrInput');
  var qrImg = document.getElementById('qrImg');
  bindImageUpload(qrBox, qrInput, qrImg);
  document.getElementById('addQrBtn').addEventListener('click', function(){ showPaymentPanel(); qrInput.click(); });

  /* ---------- background color + image ---------- */
  function shadeColor(hex, percent){
    hex = hex.replace('#','');
    if(hex.length === 3){ hex = hex[0]+hex[0]+hex[1]+hex[1]+hex[2]+hex[2]; }
    var num = parseInt(hex, 16);
    var r = (num >> 16) + percent, g = ((num >> 8) & 0x00FF) + percent, b = (num & 0x0000FF) + percent;
    r = Math.min(255, Math.max(0, r)); g = Math.min(255, Math.max(0, g)); b = Math.min(255, Math.max(0, b));
    return '#' + (0x1000000 + r*0x10000 + g*0x100 + b).toString(16).slice(1);
  }
  var bgColorLayer = document.getElementById('bgColorLayer');
  var bgColorPicker = document.getElementById('bgColorPicker');
  function applyBgColor(hex){
    bgColorLayer.style.background = 'linear-gradient(180deg, ' + shadeColor(hex, 35) + ' 0%, ' + hex + ' 45%, ' + shadeColor(hex, -35) + ' 100%)';
    bgColorPicker.value = hex;
    document.querySelectorAll('.swatch').forEach(function(s){
      s.classList.toggle('active', s.getAttribute('data-color').toLowerCase() === hex.toLowerCase());
    });
  }
  document.querySelectorAll('.swatch').forEach(function(btn){
    btn.addEventListener('click', function(){ applyBgColor(btn.getAttribute('data-color')); });
  });
  bgColorPicker.addEventListener('input', function(){ applyBgColor(this.value); });
  applyBgColor('#4a0d12');

  var stage = document.getElementById('stage');
  var bgImage = document.getElementById('bgImage');
  var bgInput = document.getElementById('bgInput');
  var handleBgFile = bindImageUpload(null, bgInput, bgImage, function(){ stage.classList.add('has-bg-image'); });
  document.getElementById('bgImageBtn').addEventListener('click', function(){ bgInput.click(); });
  document.getElementById('bgChangeBtn').addEventListener('click', function(){ bgInput.click(); });
  document.getElementById('clearBgImageBtn').addEventListener('click', function(){
    bgImage.src = '';
    stage.classList.remove('has-bg-image');
    setStatus('Background image removed — showing solid color ✓');
  });
  stage.addEventListener('dragover', function(e){ e.preventDefault(); });
  stage.addEventListener('drop', function(e){
    e.preventDefault();
    if(e.dataTransfer.files && e.dataTransfer.files[0]) handleBgFile(e.dataTransfer.files[0]);
  });

  /* ---------- Free / Paid toggle ---------- */
  var priceText = document.getElementById('priceText');
  var freeVal = '', paidVal = '';
  function applyPlaceholder(){
    if(document.getElementById('radioPaid').checked){
      priceText.setAttribute('data-placeholder', 'Entry: $15 per person');
    } else {
      priceText.setAttribute('data-placeholder', 'FREE ADMISSION');
    }
    refreshEmpty(priceText);
  }
  document.getElementById('radioFree').addEventListener('change', function(){
    if(this.checked){ paidVal = priceText.textContent; priceText.textContent = freeVal; applyPlaceholder(); }
  });
  document.getElementById('radioPaid').addEventListener('change', function(){
    if(this.checked){ freeVal = priceText.textContent; priceText.textContent = paidVal; applyPlaceholder(); showPaymentPanel(); }
  });
  applyPlaceholder();

  /* ---------- sponsors ---------- */
  var grid = document.getElementById('sponsorsGrid');
  var tpl = document.getElementById('sponsorTemplate');
  var addTile = document.getElementById('addSponsorTile');
  var MAX_SPONSORS = 5;

  function sponsorCount(){ return grid.querySelectorAll('.sponsor-slot').length; }
  function updateAddTileVisibility(){
    addTile.style.display = sponsorCount() >= MAX_SPONSORS ? 'none' : 'inline-block';
  }

  function addSponsor(){
    if(sponsorCount() >= MAX_SPONSORS) return;
    var node = tpl.content.cloneNode(true);
    var slot = node.querySelector('.sponsor-slot');
    var box = node.querySelector('.sponsor-logo-upload');
    var input = node.querySelector('.file-input');
    var img = node.querySelector('img');
    var nameEl = node.querySelector('.sponsor-name');
    var removeBtn = node.querySelector('.remove-sponsor');

    bindImageUpload(box, input, img);
    wireEditable(nameEl);
    removeBtn.addEventListener('click', function(){ slot.remove(); updateAddTileVisibility(); });

    grid.appendChild(slot);
    updateAddTileVisibility();
  }
  document.getElementById('addSponsorBtn').addEventListener('click', addSponsor);
  addSponsor();
  addSponsor();

  /* ---------- Download PDF ---------- */
  document.getElementById('downloadPdfBtn').addEventListener('click', function(){
    if(typeof html2canvas === 'undefined' || !window.jspdf){
      setStatus('PDF export needs an internet connection the first time it runs (loads a small helper library). Please check your connection and try again.', true);
      return;
    }
    var btn = this;
    var original = btn.textContent;
    btn.disabled = true; btn.textContent = 'Preparing PDF…';
    document.body.classList.add('exporting');

    var flyer = document.getElementById('flyer');
    html2canvas(flyer, {
      scale: 2,
      useCORS: true,
      allowTaint: false,
      backgroundColor: '#1c1c1c',
      windowWidth: flyer.scrollWidth,
      windowHeight: flyer.scrollHeight,
      height: flyer.scrollHeight,
      width: flyer.scrollWidth
    }).then(function(canvas){
      var imgData = canvas.toDataURL('image/jpeg', 0.95);
      var jsPDF = window.jspdf.jsPDF;
      var pxToPt = 72 / 96;
      var widthPt = (canvas.width / 2) * pxToPt;
      var heightPt = (canvas.height / 2) * pxToPt;
      var pdf = new jsPDF({
        orientation: widthPt > heightPt ? 'landscape' : 'portrait',
        unit: 'pt',
        format: [widthPt, heightPt]
      });
      pdf.addImage(imgData, 'JPEG', 0, 0, widthPt, heightPt);
      var name = slugify(document.getElementById('eventName').textContent) + '-Flyer.pdf';
      pdf.save(name);
      setStatus('PDF downloaded ✓ (single page)');
    }).catch(function(err){
      console.error(err);
      setStatus('Could not generate the PDF (' + (err && err.message ? err.message : 'unknown error') + '). Try "Print" instead, or check your connection.', true);
    }).finally(function(){
      document.body.classList.remove('exporting');
      btn.disabled = false; btn.textContent = original;
    });
  });

  /* ---------- Print ---------- */
  document.getElementById('printBtn').addEventListener('click', function(){ window.print(); });

  /* ---------- Save editable copy as standalone HTML ---------- */
  document.getElementById('saveHtmlBtn').addEventListener('click', function(){
    var html = '<!DOCTYPE html>\n' + document.documentElement.outerHTML;
    var blob = new Blob([html], { type: 'text/html' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = slugify(document.getElementById('eventName').textContent) + '-Flyer.html';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    setStatus('Editable copy saved ✓');
  });

  /* ---------- Copy details for website ---------- */
  document.getElementById('copyJsonBtn').addEventListener('click', function(){
    var sponsors = [];
    grid.querySelectorAll('.sponsor-slot').forEach(function(slot){
      var n = slot.querySelector('.sponsor-name').textContent.trim();
      if(n) sponsors.push(n);
    });
    var data = {
      eventName: document.getElementById('eventName').textContent.trim(),
      subtitle: document.getElementById('eventSubtitle').textContent.trim(),
      date: document.getElementById('fieldDate').textContent.trim(),
      time: document.getElementById('fieldTime').textContent.trim(),
      venue: document.getElementById('fieldVenue').textContent.trim(),
      specialInstructions: document.getElementById('fieldNote').textContent.trim(),
      priceType: document.getElementById('radioPaid').checked ? 'Paid' : 'Free',
      priceDetails: priceText.textContent.trim(),
      paymentSectionShown: !paymentPanel.classList.contains('hidden'),
      hasQrCode: qrBox.classList.contains('has-image'),
      paymentLink: document.getElementById('fieldPayLink').textContent.trim(),
      paymentInstructions: document.getElementById('fieldPayInstructions').textContent.trim(),
      contact: document.getElementById('fieldContact').textContent.trim(),
      sponsors: sponsors
    };
    var json = JSON.stringify(data, null, 2);
    var done = function(){ setStatus('Event details copied — paste into your website form ✓'); };
    var fail = function(){ window.prompt('Copy the event details below:', json); };
    if(navigator.clipboard && navigator.clipboard.writeText){
      navigator.clipboard.writeText(json).then(done).catch(fail);
    } else {
      fail();
    }
  });

  /* ---------- Reset ---------- */
  document.getElementById('resetBtn').addEventListener('click', function(){
    if(confirm('Clear all edits and start over with a blank template?')) location.reload();
  });

})();
</script>
</body>
</html>
`;

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

// Session token carries a role ("full" or "shala") alongside the expiry,
// both covered by the HMAC signature so a shala-role cookie can't be
// hand-edited into a full-role one without knowing SESSION_SECRET.
async function makeSession(env, role) {
  const expiry = Date.now() + SESSION_HOURS * 3600 * 1000;
  const sig = await hmac(env.SESSION_SECRET, `${expiry}.${role}`);
  return `${expiry}.${role}.${sig}`;
}

// Returns the role string ("full"/"shala") if the session cookie is valid,
// or null otherwise. Callers should treat any falsy return as "not logged in".
async function verifySession(env, cookieHeader) {
  if (!cookieHeader) return null;
  const match = cookieHeader.match(/(?:^|;\s*)session=([^;]+)/);
  if (!match) return null;
  const parts = decodeURIComponent(match[1]).split(".");
  if (parts.length !== 3) return null;
  const [expiryStr, role, sig] = parts;
  if (!expiryStr || !role || !sig) return null;
  if (Date.now() > Number(expiryStr)) return null;
  const expected = await hmac(env.SESSION_SECRET, `${expiryStr}.${role}`);
  return expected === sig ? role : null;
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

    ${isXlsx ? `<div class="xlsx-editor"></div>` : ""}

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
    <div class="row" style="margin-top:14px;">
      <strong style="font-size:12px;color:#6b6558;">Current files</strong>
      <button class="secondary" style="margin-left:auto;font-size:11px;padding:4px 10px;" onclick="loadFolder('${f.path}')">Refresh list</button>
    </div>
    <div class="thumbs" id="thumbs-${cssId(f.path)}">Loading current files…</div>
  </div>`;
}

function byPath(list, path) {
  return list.find((x) => x.path === path);
}

function marqueeCardHtml() {
  return `<div class="card" data-path="data/marquee.json">
    <h3 style="margin-bottom:2px">Announcement Marquee <span class="note">Scrolling banner under the homepage nav</span></h3>
    <p class="desc" style="margin-bottom:12px">For occasional announcements (admissions open, an upcoming event, etc.). Turn it off between announcements — it disappears from the homepage entirely, not just blank.</p>
    <div class="row" style="margin-bottom:10px;">
      <label style="display:flex;align-items:center;gap:8px;font-size:13px;cursor:pointer;">
        <input type="checkbox" id="marquee-active"> Show this banner on the homepage
      </label>
    </div>
    <div class="row" style="margin-bottom:10px;">
      <input type="text" id="marquee-text" placeholder="Announcement text, e.g. Marathi Shala Admissions open now" style="flex:1;min-width:260px;padding:8px;border:1px solid #ccc;border-radius:6px;">
    </div>
    <div class="row" style="margin-bottom:10px;">
      <input type="text" id="marquee-link" placeholder="Page to open when clicked, e.g. shala.html (optional)" style="flex:1;min-width:260px;padding:8px;border:1px solid #ccc;border-radius:6px;">
    </div>
    <button onclick="saveMarquee(this)">Save</button>
    <div class="status" id="marquee-status"></div>
  </div>`;
}

function pageSectionHtml(page, idx) {
  const xlsxCards = (page.xlsx || []).map((p) => fileCardHtml(byPath(XLSX_FILES, p))).join("\n");
  const docCards = (page.docs || []).map((p) => fileCardHtml(byPath(SINGLE_DOCS, p))).join("\n");
  const folderCards = (page.folders || []).map((p) => folderCardHtml(byPath(FOLDERS, p))).join("\n");
  const marqueeCard = page.key === "home" ? marqueeCardHtml() : "";
  return `<section class="page-section${idx === 0 ? " active" : ""}" id="page-${page.key}">
    ${marqueeCard}
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

function adminPage(role) {
  const isShalaRole = role === "shala";
  const pages = isShalaRole ? PAGES.filter((p) => SHALA_ROLE_PAGE_KEYS.includes(p.key)) : PAGES;

  const tabs = pages.map(
    (p, i) => `<button class="page-tab${i === 0 ? " active" : ""}" data-page="${p.key}" onclick="showPage('${p.key}')">${p.label}</button>`
  ).join("\n") + (isShalaRole ? "" :
      `\n<button class="page-tab" data-page="log" onclick="showPage('log')">Activity Log</button>`
    + `\n<a href="/flyer" target="_blank" class="page-tab" style="text-decoration:none;display:inline-block;">🎨 Flyer Builder</a>`);
  const sections = pages.map(pageSectionHtml).join("\n") + (isShalaRole ? "" : "\n" + logSectionHtml());
  // Folder listing is gated server-side by isAllowedPath too, but only
  // wiring up the folders this role can actually see keeps the page's own
  // JS from even trying to load anything out of scope.
  const allFolderPaths = JSON.stringify(pages.flatMap((p) => p.folders || []));

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${isShalaRole ? "Sankalp Shala Admin" : "Sankalp Website Admin"}</title><style>${BASE_STYLE}</style>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  </head>
  <body>
  <header>
    <h1>${isShalaRole ? "Sankalp Shala Admin" : "Sankalp Website Admin"}</h1>
    <a href="/logout">Log out</a>
  </header>
  <main>
    <p style="font-size:13px;color:#6b6558">${isShalaRole
      ? "Shala team login — you can edit the Shala page's team, FAQs, guidelines, and calendar. Uploads and edits commit directly to the live site's GitHub repo — changes go live within a minute or two."
      : "Pick the page you want to change below. Uploads and edits commit directly to the live site's GitHub repo — changes go live within a minute or two."}</p>

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
    // Scoped to this card (not a global ID lookup) since the same file — e.g.
    // the unified events sheet — can be listed on more than one page tab.
    const editor = card.querySelector('.xlsx-editor');
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

    const btnRow = document.createElement('div');
    btnRow.style.cssText = 'display:flex;gap:8px;margin-top:8px;';

    const confirmBtn = document.createElement('button');
    confirmBtn.textContent = 'Confirm names & Upload';
    confirmBtn.onclick = function() { confirmPendingUpload(confirmBtn, folder, hasYear, rows, card, input); };

    const cancelBtn = document.createElement('button');
    cancelBtn.className = 'secondary';
    cancelBtn.textContent = 'Cancel';
    cancelBtn.onclick = function() {
      container.innerHTML = '';
      input.value = '';
      setStatus(card, '', true);
    };

    btnRow.appendChild(confirmBtn);
    btnRow.appendChild(cancelBtn);
    container.appendChild(btnRow);
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
    reloadFolderSoon(folder);
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
      reloadFolderSoon(folder);
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
      reloadFolderSoon(folder);
    } catch (e) {
      alert(e.message);
    }
  }

  // GitHub's list-directory API can briefly lag right after a commit, so a
  // single immediate reload sometimes still shows the old state. Reload now
  // and again shortly after — on top of the manual "Refresh list" button.
  function reloadFolderSoon(folder) {
    loadFolder(folder);
    setTimeout(function() { loadFolder(folder); }, 1800);
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

  // ---- homepage announcement marquee ---------------------------------------

  function utf8ToBase64(str) {
    return arrayBufferToBase64(new TextEncoder().encode(str).buffer);
  }

  function base64ToUtf8(b64) {
    return new TextDecoder('utf-8').decode(base64ToArrayBuffer(b64));
  }

  async function loadMarquee() {
    const statusEl = document.getElementById('marquee-status');
    try {
      const res = await fetch('/api/file?path=' + encodeURIComponent('data/marquee.json'));
      if (res.status === 404) return; // never saved yet — leave the form at its defaults
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Failed to load');
      const cfg = JSON.parse(base64ToUtf8(data.contentBase64) || '{}');
      document.getElementById('marquee-active').checked = !!cfg.active;
      document.getElementById('marquee-text').value = cfg.text || '';
      document.getElementById('marquee-link').value = cfg.link || '';
    } catch (e) {
      if (statusEl) { statusEl.textContent = 'Could not load current banner: ' + e.message; statusEl.className = 'status err'; }
    }
  }

  async function saveMarquee(btn) {
    const statusEl = document.getElementById('marquee-status');
    const active = document.getElementById('marquee-active').checked;
    const text = document.getElementById('marquee-text').value.trim();
    const link = document.getElementById('marquee-link').value.trim();
    if (active && !text) {
      statusEl.textContent = 'Add announcement text before turning the banner on.';
      statusEl.className = 'status err';
      return;
    }
    btn.disabled = true;
    statusEl.textContent = 'Saving…';
    statusEl.className = 'status ok';
    try {
      const payload = JSON.stringify({ active: active, text: text, link: link }, null, 2);
      const res = await fetch('/api/commit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ path: 'data/marquee.json', contentBase64: utf8ToBase64(payload), message: 'Admin: update homepage marquee' })
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || 'Save failed');
      statusEl.textContent = 'Saved. Live on the homepage in a minute or two.';
      statusEl.className = 'status ok';
    } catch (e) {
      statusEl.textContent = e.message;
      statusEl.className = 'status err';
    } finally {
      btn.disabled = false;
    }
  }

  JSON.parse('${allFolderPaths}').forEach(loadFolder);
  loadLogs();
  loadMarquee();
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
        // Two passwords, two roles: ADMIN_PASSWORD gets the full dashboard,
        // SHALA_ADMIN_PASSWORD (optional — only checked if it's been set)
        // gets a restricted view scoped to Shala content only.
        let role = null;
        if (password === env.ADMIN_PASSWORD) role = "full";
        else if (env.SHALA_ADMIN_PASSWORD && password === env.SHALA_ADMIN_PASSWORD) role = "shala";
        if (!role) {
          return new Response(loginPage("Wrong password."), {
            status: 401,
            headers: { "Content-Type": "text/html" },
          });
        }
        const token = await makeSession(env, role);
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

      const role = await verifySession(env, cookie);

      if (url.pathname === "/" || url.pathname === "/login") {
        if (role) return Response.redirect(url.origin + "/admin", 302);
        return new Response(loginPage(), { headers: { "Content-Type": "text/html" } });
      }

      if (url.pathname === "/admin") {
        if (!role) return Response.redirect(url.origin + "/", 302);
        return new Response(adminPage(role), { headers: { "Content-Type": "text/html" } });
      }

      if (url.pathname === "/flyer") {
        // Flyer Builder isn't part of the Shala-scoped role for now — keep
        // that login narrowly focused on Shala content only.
        if (role !== "full") return Response.redirect(url.origin + "/", 302);
        return new Response(FLYER_BUILDER_HTML, { headers: { "Content-Type": "text/html" } });
      }

      if (url.pathname.startsWith("/api/")) {
        if (!role) return json({ error: "Not logged in" }, 401);

        if (url.pathname === "/api/list" && request.method === "GET") {
          const folder = url.searchParams.get("folder");
          if (!isAllowedPath(folder, role)) return json({ error: "Path not allowed" }, 400);
          const cfg = FOLDERS.find((f) => f.path === folder);
          if (!cfg) return json({ error: "Unknown folder" }, 400);
          const items = await ghListFolder(env, folder, !!cfg.yearFolders);
          return json(items.map((i) => ({ name: i.name, path: i.path, download_url: i.download_url })));
        }

        if (url.pathname === "/api/file" && request.method === "GET") {
          const path = url.searchParams.get("path");
          if (!isAllowedPath(path, role)) return json({ error: "Path not allowed" }, 400);
          const file = await ghGetFile(env, path);
          if (!file) return json({ error: "File not found" }, 404);
          return json(file);
        }

        if (url.pathname === "/api/commit" && request.method === "POST") {
          const body = await request.json();
          if (!isAllowedPath(body.path, role)) return json({ error: "Path not allowed" }, 400);
          const result = await ghCommitFile(env, body.path, body.contentBase64, body.message || `Admin update: ${body.path}`);
          return json({ ok: true, commit: result.commit?.sha });
        }

        if (url.pathname === "/api/logs" && request.method === "GET") {
          // Activity Log is hidden from the Shala-scoped UI and blocked
          // here too, since it surfaces commits outside Shala's own files.
          if (role !== "full") return json({ error: "Not allowed" }, 403);
          const perPage = Math.min(Number(url.searchParams.get("per_page")) || 40, 100);
          const commits = await ghListCommits(env, perPage);
          return json(commits);
        }

        if (url.pathname === "/api/rename" && request.method === "POST") {
          const body = await request.json();
          if (!isAllowedPath(body.oldPath, role) || !isAllowedPath(body.newPath, role)) {
            return json({ error: "Path not allowed" }, 400);
          }
          await ghRenameFile(env, body.oldPath, body.newPath, body.message || `Admin: rename ${body.oldPath} -> ${body.newPath}`);
          return json({ ok: true });
        }

        if (url.pathname === "/api/delete" && request.method === "POST") {
          const body = await request.json();
          if (!isAllowedPath(body.path, role)) return json({ error: "Path not allowed" }, 400);
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

// Only allow writes to paths this admin panel is meant to manage — prevents
// the API being used to overwrite arbitrary repo files. The "shala" role is
// additionally restricted to only the files/folders reachable from
// SHALA_ROLE_PAGE_KEYS (enforced here server-side, not just hidden in the
// UI) — this is what actually keeps a Shala-only login off data/events.xlsx
// and everything else on the site.
function isAllowedPath(path, role) {
  if (!path || typeof path !== "string") return false;

  if (role === "shala") {
    const pages = PAGES.filter((p) => SHALA_ROLE_PAGE_KEYS.includes(p.key));
    const allowedXlsx = pages.flatMap((p) => p.xlsx || []);
    const allowedFolders = pages.flatMap((p) => p.folders || []);
    if (allowedXlsx.includes(path)) return true;
    return allowedFolders.some((f) => path === f || path.startsWith(f + "/"));
  }

  if (XLSX_FILES.some((f) => f.path === path)) return true;
  if (SINGLE_DOCS.some((f) => f.path === path)) return true;
  if (SIMPLE_JSON_FILES.some((f) => f.path === path)) return true;
  return FOLDERS.some((f) => path === f.path || path.startsWith(f.path + "/"));
}
