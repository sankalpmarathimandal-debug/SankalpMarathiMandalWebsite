Subject: New Sankalp website — coming soon to sankalpmarathi.org

Namaskar all,

Over the past few weeks the Sankalp Web Team has rebuilt our website from the ground up, moving it off Google Sites onto a new setup that's faster, fully in our control, and — most importantly — one that any of us can update without needing to be "technical." We're working toward production go-live, so wanted to walk the core team through what's changed, why it's better for us long-term, and what's coming next.

## Where things stand right now

Our domain (sankalpmarathi.org) still points at the **old Google Sites version**. The new site is fully built, tested, and already being actively used by the team (uploading photos, editing event details, etc.) at a temporary address:
https://sankalpmarathimandal-debug.github.io/SankalpMarathiMandalWebsite/

Once we repoint sankalpmarathi.org, it'll serve the new site directly — nothing changes for visitors except everything gets better. Take a look at the temporary link above whenever you'd like a preview of what's coming.

## Old site vs. new site

| | Old (Google Sites) | New |
|---|---|---|
| Pages | Home, Event Timeline, Our Team, Marathi Shala, FAQs, Showcase, Constitution | Same 7, **plus** Shala Calendar, Forms & Sign-ups, Join Us, Become a Sponsor, and Book a Performance |
| Who can update content | Whoever has login access to the Google account running the site | Any core member — Excel-style editing through a private, password-protected Admin Panel, no Google or GitHub login needed |
| Event flyers | Made externally (Canva/Photoshop/etc.) by whoever had the skills, then uploaded manually | Built right inside the Admin Panel — pick colors/photo, add sponsors and a QR code, click "Download PDF" |
| Sign-up forms | Scattered Google Forms links, no central place to manage them | Self-service "Forms & Sign-ups" page — add a row to a spreadsheet and it's live |
| Join Us / Become a Sponsor | Google Forms | Real forms on our own site, delivered straight to our inbox |
| Performance bookings (Geet Ramayan, Abhangvani, etc.) | No dedicated way to request these | New "Book a Performance" page — browse programs, see who's involved, request a date, all in one place |
| Homepage highlights / Community Pride Wall | Manually maintained | Fully automatic — drop a photo in a folder and it appears, no editing required |
| Hosting cost | Free (Google) | Free (GitHub Pages + Cloudflare, both on free tiers) |
| Who "owns" it | Tied to whoever's personal Google account runs Sites | Lives in our own GitHub repository — a permanent, shared, versioned record the Mandal controls |

## What's genuinely new

- **Admin Panel** — a private dashboard where any of us can upload a replacement Excel file, add/remove photos, or edit table data directly in the browser, and it goes live in a minute or two. No coding, no GitHub account required.
- **Flyer Builder** — inside the Admin Panel. Pick a background, add the event's hero photo, up to five sponsor logos, and an optional payment/RSVP QR code, then download a print-ready PDF. This replaces needing someone with design software every time we need a flyer.
- **Book a Performance** — a public page listing our living-room performance teams (Geet Ramayan, Abhangvani, more to come), each with a photo, description, Instagram link, the participants and their role (singer, director, etc.), and a Free/Paid tag. Anyone can request a program for a specific date with special instructions, and it emails us directly. New programs can be added later with zero coding — just a new row in a spreadsheet.
- **Forms & Sign-ups page** — instead of hunting for scattered Google Forms links, every open sign-up (potlucks, RSVPs, surveys) lives in one place, self-managed.

## Community Pride Wall — our "Wall of Fame"

The homepage now has a **Community Pride Wall** — a continuously scrolling wall of creative logo art made by our own community members (currently **18 pieces** and growing). It sits right under a "Our Impact Stories" slider that already has **15 event/highlight photos** rotating on the homepage.

Both are fully automatic: to add a new piece of community art or a new highlight photo, someone just drops the image file into the right folder (or uploads it through the Admin Panel) — no editing, no code, no spreadsheet. It appears on the homepage within about a minute. This is a great way to spotlight community members' creativity and recent events without anyone needing to touch the website itself.

## Photos & media, generally

We've overhauled how photos and media work across the whole site:

- **Homepage highlights slider** and **Community Pride Wall** — auto-updating, described above.
- **Showcase page** — a dedicated gallery for performances, art, and achievements. Supports videos (just paste a YouTube link — thumbnail pulls automatically), photos, and documents (PDFs like Aarti sheets), and can group related items under one heading (e.g., a festival with multiple sessions).
- **Event Timeline & homepage event cards** — each event now carries its own photo, organized by year.
- **Book a Performance program cards** — logo/photo per program, sized so nothing gets awkwardly cropped.
- **Team & Shala Team pages** — individual photos per member, with a clean fallback (initials) for anyone without one yet.

All of these are managed the same simple way — the Admin Panel or a folder upload — so photo updates no longer depend on one person knowing how to edit the old Google Site.

## Why this benefits all of us

- **No single point of failure.** The old site depended on whoever had the Google Sites login knowing how to use that specific editor. Now any core member can update content through a simple dashboard — no bottleneck on one person.
- **Everything is recoverable.** Every change is tracked with full history. If something's edited by mistake, we can always see what changed and undo it — that was never possible with Google Sites.
- **New capabilities we didn't have before** — flyers, performance booking, self-service forms — without hiring anyone or buying software.
- **Zero ongoing cost**, same as before.

## About "GitHub" (in plain terms)

GitHub is just the storage system that hosts our website files for free and keeps a complete history of every change ever made — think of it as a shared, permanent filing cabinet for the site, rather than something tied to one person's personal account. Nobody needs a GitHub account to update everyday content — that's exactly what the Admin Panel is for. GitHub access only matters for two things:

1. Whoever wants to make deeper structural changes to the site itself (design, new page types, etc.).
2. Account continuity — see the punch list below, this is one of the items we want cleared before/around go-live.

## What's coming next

**The big one — going live on sankalpmarathi.org.** This is the last real step: repointing our domain from the old Google Sites to the new site. Needs whoever manages our domain registrar account to add one DNS record; everything on our side is already built and waiting.

Alongside that, a short punch list we're clearing before/around go-live:

1. **Finish real content for Book a Performance** — Geet Ramayan and Abhangvani are live with real descriptions and photos now; keep adding participants/roles as programs are confirmed, and this is the place to add any future performance teams too.
2. **Replace the remaining placeholder rows** — Showcase has two example entries (an Aarti sheet and a competition-winner demo) and Forms & Sign-ups has one example sign-up, all currently hidden — swap in real ones or delete when ready.
3. **Real Shala admission link** — currently a plain email link on the Marathi Shala page; replacing with a proper form.
4. **Document who holds the keys** — GitHub account, Cloudflare (Admin Panel) account, and domain registrar are each tied to one person's login right now. We want this written down so access never depends on a single individual.
5. **Keep growing the Community Pride Wall and highlight photos** — the more community art and event photos we add, the better the homepage looks. Anyone can contribute via the Admin Panel.

Longer term, once we're live, we're open to ideas from the team — more performance programs, more Showcase content, whatever helps the site reflect the community better.

Happy to walk anyone through the Admin Panel or Flyer Builder directly — it takes about five minutes to get comfortable with. Let me know if you have questions or want a quick demo before we flip the switch on the domain.

Dhanyawad,
Sankalp Web Team
