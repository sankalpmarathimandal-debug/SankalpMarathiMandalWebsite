# Roadmap / Planned Enhancements

Running list of planned work for this site. Add new items as they come up.

## 1. Cut the custom domain over from Google Sites to GitHub Pages

The site is fully built and live on GitHub Pages at
https://sankalpmarathimandal-debug.github.io/SankalpMarathiMandalWebsite/,
but `www.sankalpmarathi.org` still points at the old Google Sites version.
Steps to finish the cutover:

1. Add a file named `CNAME` (no extension) at the repo root containing exactly: `www.sankalpmarathi.org`
2. Repo **Settings → Pages → Custom domain** → enter `www.sankalpmarathi.org` → Save, enable "Enforce HTTPS".
3. At the domain registrar, add/update a DNS CNAME record: `www` → `sankalpmarathimandal-debug.github.io` (remove the old record pointing at Google Sites).
4. Remove the custom domain mapping on the old Google Sites site so it doesn't conflict.

Owner: whoever has access to the domain registrar (not yet identified — see README).

## 2. Other outstanding items (from README "Still to do")

- Delete any old/unused GitHub fine-grained tokens once the current `GITHUB_TOKEN` secret is confirmed as the only one in use.
- Shala admission form link (currently a `mailto:`) in `shala.html` — replace with a real form.
- Confirm the Diwali date discrepancy (Nov 8 vs Nov 15, 2026) and correct `data/shala-calendar.xlsx` if needed.
- `data/forms.xlsx` ships with one inactive example row — replace or delete once real forms are added.
- `data/showcase.xlsx` ships with two inactive example rows (Event grouping + document card demo) — replace or delete once real content is added.
- `data/programs.xlsx` / `data/program-participants.xlsx` (Book a Performance page) ship with two inactive example rows ("Geet Ramayan", "Abhangvani") — replace the placeholder description/participants, add a logo photo to `assets/images/programs/`, and set `Active` to `Yes` to publish each one.
