# संकल्प - आपलं मराठी मंडळ (Sankalp Marathi Mandal)

Official website for Sankalp Marathi Mandal, Massachusetts.
Live at: https://sankalpmarathimandal-debug.github.io/Sankalp_SJ/ (custom domain: [sankalpmarathi.org](https://www.sankalpmarathi.org))

**How it works:** every page reads its content from an Excel workbook in `data/`. Edit the workbook → commit → the site updates. No Google Sheets dependency.

## Folder structure

```
Sankalp_SJ/
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
├── constitution.html   Constitution (embedded PDF)
│
├── data/               ← ALL SITE CONTENT (edit these to update the site)
│   ├── home-events.xlsx    Homepage event cards
│   ├── timeline.xlsx       Events page timeline
│   ├── testimonials.xlsx   Homepage "Community Voices"
│   ├── highlights.json     Homepage slider (auto-generated — don't edit)
│   ├── partners.xlsx       Homepage partner logos
│   ├── team.xlsx           Team page members
│   ├── shala-team.xlsx     Shala page team
│   ├── faq.xlsx            FAQ page questions
│   ├── shala-faq.xlsx      Shala page FAQs
│   ├── shala-calendar.xlsx Shala Calendar page (Year, Month, Day, Title, Type, Time, Notes)
│   ├── forms.xlsx          Forms & Sign-ups page (Title, Description, Link, Active, Order)
│   └── showcase.xlsx       Showcase page (Event, Title, Description, Category, YouTubeURL, ImageURL, DocumentURL, Active, Order)
│
├── assets/
│   ├── css/style.css       All styling
│   ├── js/main.js          All logic (reads the workbooks)
│   └── images/             Site images (branding, events, team, highlights, showcase…)
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
| Homepage event cards | `data/home-events.xlsx` |
| Event timeline | `data/timeline.xlsx` |
| Testimonials | `data/testimonials.xlsx` |
| Highlight photos | just add/remove photos in `assets/images/highlights/` and push — updates automatically |
| Community Pride Wall (homepage) | just add/remove image files in `assets/images/branding/logo-variants/` and push — updates automatically, see "Updating the Community Pride Wall" below |
| Partner logos | `data/partners.xlsx` + logo in `assets/images/partners/` |
| Team members | `data/team.xlsx` + photo in `assets/images/team/` |
| Shala team / FAQs | `data/shala-team.xlsx` / `data/shala-faq.xlsx` |
| Shala calendar | `data/shala-calendar.xlsx` — one row per date (see below) |
| General FAQs | `data/faq.xlsx` |
| Constitution | replace `docs/constitution.pdf` |
| Join Us / Become a Sponsor forms | `join.html` / `sponsor.html` — see "Setting up form delivery" below |
| Forms & Sign-ups (event RSVPs, surveys, etc.) | `data/forms.xlsx` — no coding, see "Updating Forms & Sign-ups" below |
| Showcase videos, photos & documents | `data/showcase.xlsx` — no coding, see "Updating the Showcase" below |
| Announcement banner | marquee text in `index.html` |

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

**To add a new one:** drop the image file into `assets/images/branding/logo-variants/` (any filename works — no renaming needed) and push, or upload it via GitHub's website (**Add file → Upload files** on that folder, **Commit changes**). A GitHub Action (`.github/workflows/update-logo-variants.yml`) automatically regenerates `data/logo-variants.json` within about a minute of the push, and the homepage picks it up on the next page load — nothing else to touch.

**To remove one:** delete the image file from that folder the same way; the manifest updates automatically.

Naming convention for the current set: `variant-01.png` through `variant-19.png`, standardized from the original `Sankalp_Logo` folder. New uploads don't need to follow this pattern — any image filename is picked up automatically — but keeping the `variant-NN` style is a nice-to-have for tidiness if you're adding several at once.

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

## Previewing changes

Double-click **Start Local Preview.command** — it opens the site in your browser with data loading correctly. (Opening an .html file directly won't load data; browsers block that.)

## Custom domain (www.sankalpmarathi.org)

1. Add a file named `CNAME` (no extension) at the repo root containing exactly: `www.sankalpmarathi.org`
2. Repo **Settings → Pages → Custom domain** → enter `www.sankalpmarathi.org` → Save, enable "Enforce HTTPS".
3. At your domain registrar, add a DNS CNAME record: `www` → `sankalpmarathimandal-debug.github.io`

## Still to do

- Shala admission form link (currently a mailto) in `shala.html`
- Confirm the Diwali date discrepancy noted above (Nov 8 vs Nov 15, 2026) and correct `data/shala-calendar.xlsx` if needed
- `data/forms.xlsx` ships with one inactive example row — replace or delete once real forms are added
- `data/showcase.xlsx` ships with two inactive example rows demonstrating Event grouping and a document card (`DocumentURL`/`ImageURL` point at placeholder files that don't exist yet) — replace or delete once real content is added
