# संकल्प - आपलं मराठी मंडळ (Sankalp Marathi Mandal)

Official website for Sankalp Marathi Mandal, Massachusetts.
Live at: https://sankalpmarathimandal-debug.github.io/SankalpMarathiMandalWebsite/ (custom domain: [sankalpmarathi.org](https://www.sankalpmarathi.org))

**How it works:** every page reads its content from an Excel workbook in `data/`. Edit the workbook → commit → the site updates. No Google Sheets dependency.

## Handoff guide — accounts & services this site depends on

Read this section before touching anything if you're new to maintaining this site. It's not just code in one repo — a handful of outside accounts have to keep working together for the live site and the admin panel to function. If whoever holds the login to one of these disappears, that piece breaks with no easy recovery, so know who owns what.

| Service | What it does here | Where | Who needs access |
|---|---|---|---|
| **GitHub** | Hosts the code and (via GitHub Pages) the live site | github.com/sankalpmarathimandal-debug/SankalpMarathiMandalWebsite | Anyone maintaining content/code needs collaborator access (repo **Settings → Collaborators**). This is a personal GitHub account (`sankalpmarathimandal-debug`), not an organization — whoever controls that account's login controls everything. |
| **GitHub Pages** | Serves the site publicly from this repo | Repo **Settings → Pages** | Currently building from the `main` branch. |
| **Cloudflare Workers** | Hosts the private Admin panel (upload Excel/images/PDFs without using git, plus a point-and-click Event Flyer Builder) | dash.cloudflare.com → Workers & Pages → `sankalp-admin` | Whoever has the Cloudflare account login. Full setup/rotation steps: `admin-worker/SETUP.md`. |
| **GitHub fine-grained token** | Lets the Admin panel commit files into this repo on your behalf | Created at github.com/settings/tokens?type=beta, stored as the Cloudflare secret `GITHUB_TOKEN` (never in code) | Must stay scoped to `SankalpMarathiMandalWebsite` only, Contents: Read & write, and should have an expiration date set (rotate before it lapses). |
| **Web3Forms** | Delivers Join Us / Become a Sponsor submissions by email | web3forms.com — key lives in `assets/js/main.js` (`CONFIG.WEB3FORMS_ACCESS_KEY`) | No login required; it's a free access-key relay. The key being visible in the JS is expected — that's how Web3Forms works, it's not a secret. |
| **Google Forms / Tally / etc.** | Whatever individual sign-up forms are listed in `data/forms.xlsx` | Each row's `Link` column points at a form built by whoever created that specific sign-up | Not centralized — every form's owner is whoever built it in their own Google/Tally account. If that person leaves, that specific link may need rebuilding. |
| **Domain registrar for sankalpmarathi.org** | Controls DNS for the custom domain | Not identified anywhere in this repo — find out who manages billing/DNS for the domain and document it here | See the domain warning just below — this is currently the biggest gap. |
| **Instagram / Facebook / YouTube** | Linked from the site footer | Login held by whoever originally set up each account | Not part of this repo, just linked out to. |
| **cdnjs.cloudflare.com** | Delivers the SheetJS (`xlsx.js`) and Font Awesome libraries used by every page | Public CDN | No account of any kind needed — just a script tag. Not the same thing as the Cloudflare Workers account above. |

### ⚠️ The live domain still points at the old Google Sites version

As of this writing, `www.sankalpmarathi.org` serves the **old Google Sites site**, not this repo. This new site is fully built and working, but only at the raw GitHub Pages URL — https://sankalpmarathimandal-debug.github.io/SankalpMarathiMandalWebsite/. The custom domain was never cut over. Concretely:

1. This repo has no `CNAME` file at its root (the "Custom domain" section further down describes adding one — it hasn't been done).
2. Repo **Settings → Pages → Custom domain** has presumably not been set to `www.sankalpmarathi.org`.
3. At the domain registrar, the `www` DNS record still points at Google Sites' hosting, not `sankalpmarathimandal-debug.github.io`.
4. The old Google Sites page itself still has the custom domain mapped on its end, which would conflict if DNS were repointed without removing it there first.

Until all four are done, any updates made in this repo (new events, admin uploads, etc.) are invisible to anyone visiting sankalpmarathi.org — they're still seeing the old site. See "Custom domain" below for the exact steps once you're ready to cut over.

### The "automatic" GitHub Action for highlights / Community Pride Wall / culture icon ribbon

`data/logo-variants.json` (Community Pride Wall), `data/highlights.json` (homepage slider), and `data/culture-icons.json` (homepage culture icon ribbon, right below the hero) are regenerated automatically by `.github/workflows/update-manifests.yml` whenever a file is added, renamed, or removed in `assets/images/highlights/`, `assets/images/branding/logo-variants/`, or `assets/images/culture-icons/` — including uploads through the admin panel, since those are real commits to those folders. The workflow runs the same three scripts below and pushes the result back within about a minute.

**Note for context:** this workflow was missing for a while (the scripts existed and were documented as "automatic" but nothing actually ran them) — if a photo you uploaded to any of these folders still isn't showing up after a couple minutes, check the **Actions** tab on GitHub to confirm the workflow ran and didn't fail, or regenerate by hand as a fallback:

```
python3 scripts/generate_logo_variants.py
python3 scripts/generate_highlights.py
python3 scripts/generate_culture_icons.py
```

then commit the updated JSON files. (`Start Local Preview.command` already runs `generate_highlights.py` automatically for local preview — that's unrelated to the live site and unaffected by any of this.)

## Folder structure

```
SankalpMarathiMandalWebsite/
├── index.html          Home page
├── events.html         Event timeline (by year)
├── team.html           Our Team
├── shala.html          Marathi Shala
├── calendar.html       Shala Calendar (monthly, real 2026-2027 schedule)
├── join.html           Join Us form (Web3Forms — no Google Forms)
├── sponsor.html        Become a Sponsor form (Web3Forms — no Google Forms)
├── faq.html            FAQs
├── showcase.html       Showcase — videos, photos & documents, event-grouped (see below)
├── forms.html           Forms & Sign-ups (self-service — see below)
├── book-a-performance.html  Book a Performance — program menu + request form (see below)
├── constitution.html   Constitution (embedded PDF)
│
├── data/               ← ALL SITE CONTENT (edit these to update the site)
│   ├── events.xlsx         Homepage event cards + Events page timeline +
│   │                       homepage "Book a Performance" teaser — one sheet
│   │                       (Type/Event Type columns control where each row
│   │                       shows up; see the header comment in main.js)
│   ├── testimonials.xlsx   Homepage "Community Voices"
│   ├── highlights.json     Homepage slider (auto-generated — don't edit)
│   ├── culture-icons.json  Homepage culture icon ribbon (auto-generated — don't edit)
│   ├── sponsors.xlsx       Homepage Presenting Sponsors cards
│   ├── team.xlsx           Team page members
│   ├── shala-team.xlsx     Shala page team
│   ├── faq.xlsx            FAQ page questions
│   ├── shala-faq.xlsx      Shala page FAQs
│   ├── shala-guidelines-parents.xlsx   Shala page — Parent guidelines accordion
│   ├── shala-guidelines-teachers.xlsx  Shala page — Teacher guidelines accordion
│   ├── shala-admissions.xlsx           Shala page — occasional banner (Active toggle, not just for admissions)
│   ├── shala-events.xlsx               Shala page — Events section, fully separate from data/events.xlsx
│   ├── shala-calendar.xlsx Shala Calendar page (Year, Month, Day, Title, Type, Time, Notes)
│   ├── forms.xlsx          Forms & Sign-ups page (Title, Description, Link, Active, Order)
│   ├── showcase.xlsx       Showcase page (Event, Title, Description, Category, YouTubeURL, ImageURL, DocumentURL, Active, Order)
│   ├── programs.xlsx       Book a Performance program menu (Title, Description, ImageURL, InstagramURL, PriceType, PriceDetails, Active, Order)
│   └── program-participants.xlsx  Book a Performance participants (Program, Name, Role, Order)
│
├── assets/
│   ├── css/style.css          All styling — editing this auto-bumps the `?v=` cache-busting tag on every page via `.github/workflows/bump-cache-version.yml`, so visitors don't get stuck with a stale cached copy
│   ├── css/style-guide.html   Living style guide — open in a browser to see every color, font, button, and card style rendered for real (loads the actual style.css, always in sync)
│   ├── js/main.js             All logic (reads the workbooks) — same auto cache-bust as style.css applies here
│   └── images/                Site images (branding, events, team, highlights, showcase, programs…)
│
├── docs/
│   ├── constitution.pdf        The constitution document
│   └── showcase/               Showcase PDFs (Aarti sheets, event docs — see below)
│
├── source/                 Reference only — NOT used by the site
│   ├── google-links.md         Links to the original Google Sheets & Forms
│   └── reference/              Guidelines, schedules, and other workbooks
│
└── Start Local Preview.command   Double-click to preview the site locally
```

## Common updates

| To change | Edit |
|---|---|
| Homepage event cards | `data/events.xlsx` — needs a `Type` (previous/current/future) to show as one of the 3 featured cards |
| Event timeline | `data/events.xlsx` — every row with Event Type = Event shows here, grouped by Year |
| Homepage "Book a Performance" teaser | `data/events.xlsx` — rows with Event Type = Performance |
| Testimonials | `data/testimonials.xlsx` |
| Highlight photos | just add/remove photos in `assets/images/highlights/` and push — updates automatically |
| Community Pride Wall (homepage) | just add/remove image files in `assets/images/branding/logo-variants/` and push — updates automatically, see "Updating the Community Pride Wall" below |
| Culture icon ribbon (homepage, below the hero) | just add/remove image files in `assets/images/culture-icons/` and push — updates automatically, no captions shown |
| Presenting Sponsors | `data/sponsors.xlsx` (optional headshot in `assets/images/partners/`) |
| Team members | `data/team.xlsx` + photo in `assets/images/team/` |
| Shala team / org chart | `data/shala-team.xlsx` + photos in `assets/images/shala/team/` — ⚠️ any org structure change needs Sankalp Board + EC approval before publishing |
| Shala FAQs | `data/shala-faq.xlsx` |
| Shala guidelines (parents/teachers) | `data/shala-guidelines-parents.xlsx` / `data/shala-guidelines-teachers.xlsx` |
| Shala-specific events | `data/shala-events.xlsx` + photos/flyers in `assets/images/shala/events/` — standalone sheet, separate from the main Events & Performances sheet |
| Shala calendar | `data/shala-calendar.xlsx` — one row per date (see below) |
| General FAQs | `data/faq.xlsx` |
| Constitution | replace `docs/constitution.pdf` |
| Join Us / Become a Sponsor forms | `join.html` / `sponsor.html` — see "Setting up form delivery" below |
| Forms & Sign-ups (event RSVPs, surveys, etc.) | `data/forms.xlsx` — no coding, see "Updating Forms & Sign-ups" below |
| Showcase videos, photos & documents | `data/showcase.xlsx` — no coding, see "Updating the Showcase" below |
| Book a Performance — program menu | `data/programs.xlsx` + `data/program-participants.xlsx` — no coding, see "Updating Book a Performance" below |
| Announcement banner | `data/marquee.json` — or the "Announcement Marquee" card on the Home Page tab in the admin panel (no coding) |

Keep the header row of each workbook intact, and keep images web-friendly (≤1200px, JPG preferred).

### Updating the Shala Calendar

`data/shala-calendar.xlsx` has one row per date: `Year, Month, Day, Title, Type, Time, Notes`.

- **Month** is the full name (`September`), **Day** is the day of the month (`7`).
- **Type** controls the color and must be one of: `Class`, `Event`, `Holiday`, `Exam`.
- **Time** and **Notes** are optional. Notes is a good place for context like "Online · Week 5" or "In-person".
- For a weekly recurring class, add one row per week — there's no recurrence feature yet.

The calendar is loaded from the 2026-2027 detailed schedule (weekly online classes, holidays, exams, and in-person events like Diwali/Gudhi Padwa/Picnic). Multi-day breaks (Christmas/New Year, Winter Break, Spring Break) are expanded into one row per day so every day in the break shows correctly. The page opens on the current month if it has entries, otherwise the nearest month that does.

**Known thing to double-check:** the source schedule's detail sheet lists Diwali on **November 8, 2026**, but its own Summary tab lists **November 15, 2026** — the calendar currently uses November 8. Confirm which date is correct and let me know if it needs correcting.

`calendar.html` also has a **"Download Full Year Schedule (Excel)"** button right under the page title, linking directly to `data/shala-calendar.xlsx` — so parents can grab the whole year at once. Since it points at the live workbook, it's always in sync automatically; no separate export file to maintain.

### Updating the Community Pride Wall (homepage, no coding required)

The homepage shows a "Community Pride Wall" section (right before About Us) — two rows of community-made logo art flowing across the screen in opposite directions. It's fully automatic: the site reads whatever image files are sitting in `assets/images/branding/logo-variants/`, no spreadsheet or code edit involved.

**To add a new one:** drop the image file into `assets/images/branding/logo-variants/` (any filename works — no renaming needed) and push, or upload it via GitHub's website (**Add file → Upload files** on that folder, **Commit changes**). A GitHub Action (`.github/workflows/update-manifests.yml`) automatically regenerates `data/logo-variants.json` within about a minute of the push, and the homepage picks it up on the next page load — nothing else to touch.

**To remove one:** delete the image file from that folder the same way; the manifest updates automatically.

Naming convention for the current set: `variant-01.png` through `variant-19.png`, standardized from the original `Sankalp_Logo` folder. New uploads don't need to follow this pattern — any image filename is picked up automatically — but keeping the `variant-NN` style is a nice-to-have for tidiness if you're adding several at once.

### Updating the Announcement Marquee (homepage, no coding required)

The scrolling banner just under the homepage navigation (e.g. "Marathi Shala Admissions open now") is meant for occasional announcements, not a permanent fixture — it's driven entirely by `data/marquee.json` and defaults to **hidden** when that file is missing or turned off, not blank/empty.

**Easiest way:** open the admin panel, go to the **Home Page** tab, and use the "Announcement Marquee" card at the top — a checkbox to show/hide it, a text field for the announcement, and an optional link (e.g. `shala.html`) for what page it opens when clicked. Click Save; it's live in a minute or two.

**Manually:** `data/marquee.json` looks like:
```json
{
  "active": true,
  "text": "Marathi Shala Admissions open now",
  "link": "shala.html"
}
```
Set `active` to `false` to hide the banner entirely without losing the saved text (handy for turning it back on later with one click). `link` can be left as an empty string if the banner shouldn't be clickable.

### Setting up form delivery (Join Us / Become a Sponsor)

No Google Forms — `join.html` and `sponsor.html` are real HTML forms that submit via [Web3Forms](https://web3forms.com), a free email-relay service with no account or password required:

1. Go to https://web3forms.com and enter the email that should receive submissions (`sandeepj0208@gmail.com` or whichever inbox is best) — a free access key is emailed instantly. No sign-up, no password.
2. Open `assets/js/main.js` and paste the key into `CONFIG.WEB3FORMS_ACCESS_KEY` near the top of the file, replacing `'REPLACE_WITH_YOUR_WEB3FORMS_ACCESS_KEY'`.
3. Commit and push. Both forms start delivering to your inbox immediately — no other code changes needed.

Until a real key is set, submitting either form shows a friendly "not set up yet" message instead of failing silently.

### Updating Forms & Sign-ups (no coding required)

`forms.html` is a self-service page — anyone on the team can publish a new sign-up or survey without touching code, using `data/forms.xlsx`. Columns: `Title, Description, Link, Active, Order`.

**To add a new form:**

1. Create the form itself somewhere no-code — [Google Forms](https://forms.google.com) or [Tally](https://tally.so) both work well and are free. Build it there like any Google Form, then copy its shareable link.
2. Open `data/forms.xlsx` in Excel (or Google Sheets), add a new row: a short `Title`, a one-line `Description`, paste the `Link`, set `Active` to `Yes`, and give it an `Order` number (lower numbers show first).
3. Save the file, keeping the same filename (`forms.xlsx`).
4. Get it back into the repo (pick whichever fits the team):
   - **Easiest / recommended:** send the updated file (or just the new row's details) to whoever manages the GitHub repo — they drag the replacement file in and it's live in a minute.
   - **Fully self-service:** if a team member is added as a GitHub collaborator, they can go to `data/forms.xlsx` on github.com, click **Add file → Upload files**, drag in the updated workbook (same filename), and click **Commit changes** — all in the browser, no terminal, no git commands.

To take a form down, set its `Active` column to `No` instead of deleting the row (keeps the link on file for next time). If there are no active rows, the page just shows "No forms are open right now."

### Updating the Showcase (videos, photos & documents, no coding required)

`showcase.html` displays videos, photos, and documents — performances, student achievements, a community member's painting, an event's Aarti sheet, anything — driven entirely by `data/showcase.xlsx`. Columns: `Event, Title, Description, Category, YouTubeURL, ImageURL, DocumentURL, Active, Order`.

Each row is **exactly one** of a video, a photo, or a document — fill in only one of `YouTubeURL` / `ImageURL` / `DocumentURL` and leave the other two blank.

**To add a video:** paste any YouTube link into `YouTubeURL` (a normal `youtube.com/watch?v=...` link, a `youtu.be/...` short link, or a Shorts link all work). The site pulls the thumbnail from YouTube automatically — nothing to upload. Clicking it plays the video in a pop-up on the page.

**To add a photo** (a painting, a student's achievement, an event photo, etc.): save it into `assets/images/showcase/` (≤1200px, JPG preferred), then put the path in `ImageURL`, e.g. `assets/images/showcase/priya-painting.jpg`. Clicking it opens the photo full-size in a pop-up.

**To add a document** (an Aarti sheet, competition rules, an event program, etc.): save the PDF into `docs/showcase/`, then put the path in `DocumentURL`, e.g. `docs/showcase/ganpati-aarti-2026.pdf`. The card shows a document icon; clicking it opens the PDF in a new tab.

For any of the three, get the file into the repo the same way as `forms.xlsx` — hand it to the repo admin, or upload directly via GitHub's website if you're a collaborator (**Add file → Upload files** on the relevant folder, same filename, **Commit changes**).

**Grouping related items under one heading (e.g. a festival with multiple sessions):**

Use the `Event` column to group rows together — every row with the same `Event` text is shown under one shared heading, in its own mini-grid. For example, during Ganpati you might have a video and a document (the Aarti sheet) that belong together, and a separate competition with its own photos — give the first group's rows `Event = Ganpati 2026 — Aarti & Documents` and the second group's rows `Event = Ganpati 2026 — Competition`; they'll render as two distinct sections. Leave `Event` blank for anything that should just show in the default section with no heading (this is how the existing videos/photos work today).

**Columns shared by all rows, all groups:**

- `Title` (required), and optionally `Description` and `Category` (e.g. "Dance", "Art", "Documents").
- `Active` — set to `Yes` to publish, `No` to unpublish without losing the row.
- `Order` — **lower numbers show first**, controlling both the order of items within a group and which group appears first (whichever group's lowest `Order` item is smallest shows first). Use gaps (10, 20, 30…) to make it easy to slot new items in between later.

If there are no active rows, the page shows "Nothing to show yet."

### Updating Book a Performance (no coding required)

`book-a-performance.html` lists the living-room performance programs (Geet Ramayan, Abhangvani, and any future ones) as a menu of cards, with a request form below where someone picks a program, a date, and adds special instructions — submitted the same way as Join Us / Become a Sponsor (via Web3Forms, straight to your inbox).

**Two workbooks drive this page:**

- `data/programs.xlsx` — one row per program. Columns: `Title, Description, ImageURL, InstagramURL, PriceType, PriceDetails, Active, Order`.
  - `ImageURL` is the program's logo/photo, saved into `assets/images/programs/` — leave blank to show a generic music icon instead.
  - `InstagramURL` is optional — if filled in, the card shows a "View on Instagram" link.
  - `PriceType` — `Free` or `Paid`, shown as a badge next to the title (blank/anything else defaults to `Free`).
  - `PriceDetails` — optional, only shown when `PriceType` is `Paid`, e.g. `$100` or `Suggested donation`.
  - `Active` — set to `Yes` to publish, `No` to hide it (and its dropdown option) without deleting the row.
  - `Order` — lower numbers show first.
- `data/program-participants.xlsx` — one row per person. Columns: `Program, Name, Role, Order`.
  - `Program` must match a `Title` in `data/programs.xlsx` **exactly** — that's how a participant lands on the right card.
  - `Role` is free text (`Singer`, `Director`, `Harmonium`, etc.).
  - `Order` controls the order participants are listed within their program.

The site ships with two inactive example rows ("Geet Ramayan (example)" and "Abhangvani (example)") demonstrating the format — replace their placeholder text, add a logo, and set `Active` to `Yes` to publish, or delete them once real programs are added.

**Adding a new program later** — just add a new row to `programs.xlsx` (and any participant rows to `program-participants.xlsx`); no code or page changes needed, it appears in the menu and the request form's dropdown automatically.

**Request form fields:** Program (dropdown, auto-filled from `programs.xlsx`), Full Name, Email, Phone, Preferred Date, and Special Instructions. Clicking **"Request This Program"** on any card scrolls down and pre-selects that program in the dropdown. Submissions email the same way Join Us / Become a Sponsor do — see "Setting up form delivery" above; no separate setup needed since it reuses the same `CONFIG.WEB3FORMS_ACCESS_KEY`.

## Previewing changes

Double-click **Start Local Preview.command** — it opens the site in your browser with data loading correctly. (Opening an .html file directly won't load data; browsers block that.)

## Custom domain (www.sankalpmarathi.org)

1. Add a file named `CNAME` (no extension) at the repo root containing exactly: `www.sankalpmarathi.org`
2. Repo **Settings → Pages → Custom domain** → enter `www.sankalpmarathi.org` → Save, enable "Enforce HTTPS".
3. At your domain registrar, add a DNS CNAME record: `www` → `sankalpmarathimandal-debug.github.io`

## Admin panel — Flyer Builder

The admin panel (`admin-worker/`, deployed as the `sankalp-admin` Cloudflare
Worker) now includes a **🎨 Flyer Builder** tab, served at `/flyer` behind
the same login as the rest of the dashboard. It's a point-and-click tool for
making an event flyer — click to edit text, swap the background color or
image, add a hero photo, up to 5 sponsor logos, and an optional payment/RSVP
QR code — then "Download PDF" for a ready-to-share flyer, or "Save Editable
Copy" to keep an HTML version to reopen later. It runs entirely in the
browser (no GitHub commits, doesn't touch the live site). See
`admin-worker/SETUP.md` for details.

## Still to do

See `ROADMAP.md` for the running list of planned enhancements (currently:
cutting the custom domain over from Google Sites, plus a handful of smaller
cleanup items).

- **Cut the custom domain over from the old Google Sites site** — see the "Handoff guide" section near the top; this is the biggest outstanding item.
- ~~Add the missing GitHub Action for highlights/logo-variants automation~~ — done, see `.github/workflows/update-manifests.yml`. Keep an eye on the repo's **Actions** tab the first few times a photo is uploaded to either folder, just to confirm it's running cleanly.
- Delete any old/unused GitHub fine-grained tokens (github.com/settings/tokens?type=beta) once you've confirmed the current `GITHUB_TOKEN` secret is the only one in use.
- Shala admission form link (currently a mailto) in `shala.html`
- Confirm the Diwali date discrepancy noted above (Nov 8 vs Nov 15, 2026) and correct `data/shala-calendar.xlsx` if needed
- `data/forms.xlsx` ships with one inactive example row — replace or delete once real forms are added
- `data/showcase.xlsx` ships with two inactive example rows demonstrating Event grouping and a document card (`DocumentURL`/`ImageURL` point at placeholder files that don't exist yet) — replace or delete once real content is added
- `data/programs.xlsx` / `data/program-participants.xlsx` (Book a Performance page) ship with two inactive example rows ("Geet Ramayan", "Abhangvani") — replace the placeholder description/participants, add a logo photo, and set `Active` to `Yes` to publish each one
