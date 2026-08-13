# Sankalp Website Admin — Setup Guide

This gives you a private webpage (e.g. `admin.sankalpmarathi.org`) where you
log in with a password and upload replacement Excel files, images, or PDFs.
It commits them straight to the `SankalpMarathiMandalWebsite` GitHub repo, and GitHub Pages
redeploys the live site automatically — no GitHub account needed for whoever
is uploading.

It runs on **Cloudflare Workers** (free tier is plenty for this).

Total setup time: ~15 minutes, one time only.

## 1. Create a GitHub token

This lets the admin panel commit files on your behalf.

1. Go to https://github.com/settings/tokens?type=beta (GitHub → Settings →
   Developer settings → Personal access tokens → Fine-grained tokens).
2. Click **Generate new token**.
3. Name it `sankalp-admin`, set an expiration (e.g. 1 year — you'll need to
   regenerate it when it expires).
4. Under **Repository access**, choose **Only select repositories** →
   `SankalpMarathiMandalWebsite`.
5. Under **Permissions → Repository permissions**, set **Contents** to
   **Read and write**.
6. Click **Generate token** and copy it immediately (you won't see it again).

## 2. Create a Cloudflare account (if you don't have one)

Go to https://dash.cloudflare.com/sign-up — free plan is fine.

## 3. Install Wrangler and log in

On your computer, in this `admin-worker` folder:

```bash
npm install
npx wrangler login
```

This opens a browser to authorize Wrangler against your Cloudflare account.

## 4. Set the secrets

Three secrets are needed (never stored in the code):

```bash
npx wrangler secret put GITHUB_TOKEN
# paste the token from Step 1

npx wrangler secret put ADMIN_PASSWORD
# pick a password you'll share only with trusted committee members

npx wrangler secret put SESSION_SECRET
# any long random string, e.g. output of: openssl rand -hex 32
```

## 5. Deploy

```bash
npx wrangler deploy
```

Wrangler prints a URL like `https://sankalp-admin.<your-subdomain>.workers.dev`.
Open it, log in with your `ADMIN_PASSWORD`, and you should see the dashboard.

## 6. (Optional) Put it on your own domain

In the Cloudflare dashboard: **Workers & Pages → sankalp-admin → Settings →
Domains & Routes → Add → Custom Domain**, e.g. `admin.sankalpmarathi.org`.
Cloudflare handles the DNS/SSL for you if your domain's nameservers are on
Cloudflare; otherwise add the CNAME it gives you at your registrar.

Until you do this, the `workers.dev` URL from Step 5 works fine.

**Note:** The site footer now links to `https://admin.sankalpmarathi.org`, so
this custom domain step is no longer optional — complete it, or update the
footer link on all pages to your actual `workers.dev` URL instead. Since the
admin link is now public, make sure `ADMIN_PASSWORD` is strong.

## What the dashboard can do

- **Excel data files** — replace any of the 11 workbooks (`home-events.xlsx`,
  `timeline.xlsx`, `testimonials.xlsx`, `partners.xlsx`, `team.xlsx`,
  `shala-team.xlsx`, `faq.xlsx`, `shala-faq.xlsx`, `shala-calendar.xlsx`,
  `forms.xlsx`, `showcase.xlsx`). Upload a file with the exact same name to
  replace it — same as the current "drag a new file into GitHub" workflow,
  just without needing GitHub access.
- **Constitution PDF** — replace `docs/constitution.pdf`.
- **Image/PDF folders** — upload new files or delete existing ones in:
  event images, team photos, shala images, showcase photos, partner logos,
  Community Pride Wall logos, homepage highlight photos, and showcase PDFs.
  Event images are organized by year in the repo (`assets/images/events/2026/…`),
  so that card has a Year field — fill it in before uploading.

Every change is a real Git commit (authored as you, via the token), so
nothing is lost — you can always see history on GitHub or revert a commit if
needed.

## Things to know

- **File size:** GitHub's Contents API (which this uses) comfortably handles
  images and PDFs up to a few MB, but has a hard ceiling around 25 MB per
  file. The site's own guidance is to keep images ≤1200px/JPG, so this
  shouldn't come up — if it ever does, let me know and I can extend the
  Worker to use GitHub's Git Data API instead, which removes that limit.
- **Excel uploads are whole-file replacements**, not row-by-row editing —
  you still edit the workbook in Excel/Google Sheets, then upload the
  finished file. Building an in-browser spreadsheet-row editor is possible
  as a v2 if that'd be more convenient.
- **Image filenames still need to be typed into the relevant Excel column**
  (e.g. `Photo`, `ImageURL`) for most folders — uploading a photo doesn't
  automatically wire it into a page. The two exceptions are the **Community
  Pride Wall** and **homepage highlights** folders, which are fully
  automatic (a GitHub Action regenerates the manifest whenever a file is
  added/removed there).
- **Rotating the password:** run `npx wrangler secret put ADMIN_PASSWORD`
  again with a new value, then `npx wrangler deploy`.
- **Revoking access:** delete the GitHub token (Step 1) any time to
  immediately cut off the admin panel's write access.
