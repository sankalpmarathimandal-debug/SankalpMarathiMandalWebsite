/* =====================================================
   SANKALP MARATHI MANDAL — SHARED JAVASCRIPT
   =====================================================
   ALL SITE DATA IS SERVED FROM EXCEL WORKBOOKS IN THIS
   REPO (converted from the original Google Sheets).
   To update content, edit the .xlsx files in /data —
   the site reads them directly in the browser (SheetJS):

     data/home-events.xlsx            → homepage event cards
     data/timeline.xlsx               → Event Timeline page
     data/testimonials.xlsx  → homepage Community Voices
     data/highlights.json    → homepage slider (AUTO — just
                               add photos to assets/images/highlights/)
     data/logo-variants.json → homepage Community Pride Wall (AUTO — just
                               add image files to
                               assets/images/branding/logo-variants/, any
                               filename works, no coding needed)
     data/partners.xlsx        → homepage partner logos
     data/team.xlsx             → Our Team page
     data/shala-team.xlsx        → Marathi Shala team section
     data/faq.xlsx            → FAQs page
     data/shala-faq.xlsx               → Marathi Shala FAQs
     data/shala-calendar.xlsx → Shala Calendar page (Year, Month, Day, Title, Type, Time, Notes)
     data/forms.xlsx          → Forms & Sign-ups page (Title, Description, Link, Active, Order)
                                 — add a row to publish any Google Forms / Tally
                                 link on the site with zero coding. See README.
     data/showcase.xlsx       → Showcase page (Event, Title, Description, Category,
                                 YouTubeURL, ImageURL, DocumentURL, Active, Order) —
                                 fill in exactly one of YouTubeURL / ImageURL /
                                 DocumentURL per row (photos in assets/images/showcase/,
                                 PDFs in docs/showcase/). Rows sharing the same Event
                                 name are grouped under one heading (e.g. "Ganpati 2026
                                 — Aarti & Documents" vs "Ganpati 2026 — Competition") —
                                 leave Event blank for the default ungrouped section.
                                 Ordered by priority within and across groups.

   Images live in assets/images/. Add new images there
   and reference them by relative path in the workbooks.

   join.html / sponsor.html use native <form> submissions sent
   to Web3Forms (no Google Forms, no backend). Set
   CONFIG.WEB3FORMS_ACCESS_KEY below with a free key from
   https://web3forms.com for submissions to be delivered.
   ===================================================== */

const CONFIG = {
  EVENTS_CSV: 'data/home-events.xlsx',
  TIMELINE_CSV: 'data/timeline.xlsx',
  TESTIMONIALS_CSV: 'data/testimonials.xlsx',
  HIGHLIGHTS_JSON: 'data/highlights.json',
  LOGO_VARIANTS_JSON: 'data/logo-variants.json',
  PARTNERS_CSV: 'data/partners.xlsx',
  TEAM_CSV: 'data/team.xlsx',
  SHALA_TEAM_CSV: 'data/shala-team.xlsx',
  FAQS_CSV: 'data/faq.xlsx',
  SHALA_FAQS_CSV: 'data/shala-faq.xlsx',
  SHALA_CALENDAR_CSV: 'data/shala-calendar.xlsx',
  FORMS_CSV: 'data/forms.xlsx',
  SHOWCASE_CSV: 'data/showcase.xlsx',

  SLIDER_INTERVAL: 4000,

  // FORM DELIVERY — no Google Forms. Join Us (join.html) and Become a
  // Sponsor (sponsor.html) submit via Web3Forms (a free, no-login email
  // relay for static sites). Get your key at https://web3forms.com —
  // enter your email, no account/password needed, a key arrives by email
  // instantly. Paste it below and both forms start working.
  WEB3FORMS_ACCESS_KEY: 'b5ba71b4-5c39-405f-9c67-1383a073f01f',
};

/* =====================================================
   NAV — mobile toggle + active link
   ===================================================== */
function initNav() {
  const toggle = document.querySelector('.nav-toggle');
  const links = document.querySelector('.nav-links');
  if (toggle && links) {
    toggle.addEventListener('click', () => {
      const open = links.classList.toggle('open');
      toggle.setAttribute('aria-expanded', open);
    });
    links.querySelectorAll('a').forEach(a => a.addEventListener('click', () => links.classList.remove('open')));
  }
  const here = location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.nav-links a').forEach(a => {
    if (a.getAttribute('href') === here) a.classList.add('active');
  });
}

/* =====================================================
   SCROLL REVEAL — gentle fade-up for sections
   ===================================================== */
function initReveal() {
  if (!('IntersectionObserver' in window)) return;
  const targets = document.querySelectorAll('.section-divider, .about-card, .info-card, .shala-highlight, .impact-slider-container, .partners-section, .doc-article');
  const io = new IntersectionObserver(entries => {
    entries.forEach(en => {
      if (en.isIntersecting) { en.target.classList.add('revealed'); io.unobserve(en.target); }
    });
  }, { threshold: 0.15 });
  targets.forEach(t => { t.classList.add('reveal'); io.observe(t); });
}

/* =====================================================
   BACK TO TOP
   ===================================================== */
function initBackToTop() {
  const btn = document.querySelector('.back-to-top');
  if (!btn) return;
  window.addEventListener('scroll', () => {
    btn.classList.toggle('visible', window.scrollY > 500);
  }, { passive: true });
  btn.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
}

/* =====================================================
   HELPERS
   ===================================================== */
function viewFullSize(imageUrl) { window.open(imageUrl, '_blank'); }

function badgeText(t) {
  return t === 'previous' ? 'Previous' : t === 'current' ? 'Happening Now' : 'Upcoming';
}

function escapeHtml(s) {
  return String(s || '').replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
}

/* Warn when the site is opened as a file:// page — data cannot load without a server */
function fileProtocolNotice() {
  if (location.protocol !== 'file:') return false;
  const msg = `<div style="text-align:center;padding:32px;color:#92400e;background:#fffbeb;border:1px dashed #f59e0b;border-radius:12px;max-width:640px;margin:0 auto;">
    <p style="font-weight:700;margin-bottom:8px;">Data can't load when the page is opened directly as a file.</p>
    <p style="font-size:14px;line-height:1.6;">Double-click <strong>Start Local Preview.command</strong> in the site folder to view the site with data,<br>or visit the live site once it's on GitHub Pages.</p>
  </div>`;
  ['timeline-loading','events-grid','testimonials-container','team-container','shala-team-container','faq-container','shala-faq-container']
    .forEach(id => { const el = document.getElementById(id); if (el) el.innerHTML = msg; });
  return true;
}

/* Load an Excel workbook and return rows of the first sheet as objects */
function loadSheet(path, onDone, onError) {
  if (location.protocol === 'file:') return;
  fetch(path)
    .then(r => { if (!r.ok) throw new Error(r.status); return r.arrayBuffer(); })
    .then(buf => {
      const wb = XLSX.read(buf, { type: 'array' });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
      onDone(rows.map(r => {
        const out = {};
        for (const k in r) out[String(k).trim()] = String(r[k]).trim();
        return out;
      }));
    })
    .catch(err => { console.error('Failed to load', path, err); if (onError) onError(err); });
}

const MONTH_NUM = { jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12 };
function monthNum(m) { return MONTH_NUM[(m || '').slice(0,3).toLowerCase()] || 0; }

/* Derive previous/current/future from Year+Month vs today */
function deriveType(e) {
  const now = new Date();
  const y = parseInt(e.Year) || now.getFullYear();
  const m = monthNum(e.Month);
  const nowKey = now.getFullYear() * 100 + (now.getMonth() + 1);
  const evKey = y * 100 + m;
  if (evKey < nowKey) return 'previous';
  if (evKey === nowKey) return 'current';
  return 'future';
}

/* =====================================================
   EVENT MODAL
   ===================================================== */
let modalTriggerElement = null;
const FOCUSABLE_ELEMENTS = 'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])';

function animateModalOpen(el) {
  const r = el.getBoundingClientRect();
  const ox = ((r.left + r.width / 2) / window.innerWidth) * 100;
  const oy = ((r.top + r.height / 2) / window.innerHeight) * 100;
  const mc = document.getElementById('modal-content');
  mc.style.transformOrigin = `${ox}% ${oy}%`;
  mc.classList.add('animating-in');
  setTimeout(() => mc.classList.remove('animating-in'), 400);
}

function openModal(e, clickedElement) {
  const modal = document.getElementById('modal');
  if (!modal) return;
  const flyer = (e.Flyer || e.ImageURL || '').trim();
  const media = document.getElementById('modal-media');
  modalTriggerElement = clickedElement;

  if (flyer) {
    media.innerHTML = `
      <div class="modal-media-container">
        <img src="${flyer}" alt="Event flyer for ${escapeHtml(e.Name)}" loading="eager">
        <button class="view-fullsize-btn" onclick="viewFullSize('${flyer}')" title="View Full Size" aria-label="View full size image">
          <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-5.197-5.197m0 0A7.5 7.5 0 105.196 5.196a7.5 7.5 0 0010.607 10.607zM10.5 7.5v6m3-3h-6" />
          </svg>
        </button>
      </div>`;
    media.style.display = 'flex';
  } else {
    media.innerHTML = '';
    media.style.display = 'none';
  }

  document.getElementById('modal-title').innerText = e.Name || '';
  document.getElementById('modal-desc').innerText  = e.Description || e.Summary || '';
  document.getElementById('modal-meta').innerHTML  = `
    <div class="meta-item">
      <svg class="meta-icon" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" />
      </svg>
      <span><strong>Date:</strong> ${escapeHtml(e.Month)} ${escapeHtml(e.Date || '')}${e.Date ? ',' : ''} ${escapeHtml(e.Year)}</span>
    </div>
    ${e.Time   ? `<div class="meta-item"><svg class="meta-icon" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"/></svg><span><strong>Time:</strong> ${escapeHtml(e.Time)}</span></div>` : ''}
    ${e.Venue  ? `<div class="meta-item"><svg class="meta-icon" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"/></svg><span><strong>Venue:</strong> ${escapeHtml(e.Venue)}${e.Location ? ', ' + escapeHtml(e.Location) : ''}</span></div>` : ''}
    ${e.Contact? `<div class="meta-item"><svg class="meta-icon" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"/></svg><span><strong>Contact:</strong> ${escapeHtml(e.Contact)}</span></div>` : ''}
  `;

  modal.classList.add('active');
  document.body.classList.add('modal-open');
  animateModalOpen(clickedElement);
  setTimeout(() => document.getElementById('modal-close-btn').focus(), 100);
  initFocusTrap();
}

function closeModal() {
  const modal = document.getElementById('modal');
  modal.style.opacity = '0';
  setTimeout(() => {
    modal.classList.remove('active');
    modal.style.opacity = '';
    document.body.classList.remove('modal-open');
    if (modalTriggerElement) { modalTriggerElement.focus(); modalTriggerElement = null; }
  }, 200);
}

function initFocusTrap() {
  const modal = document.getElementById('modal');
  const els   = modal.querySelectorAll(FOCUSABLE_ELEMENTS);
  const first = els[0], last = els[els.length - 1];
  modal.addEventListener('keydown', function(e) {
    if (e.key !== 'Tab') return;
    if (e.shiftKey) { if (document.activeElement === first) { last.focus();  e.preventDefault(); } }
    else            { if (document.activeElement === last)  { first.focus(); e.preventDefault(); } }
  });
}

document.addEventListener('keydown', e => {
  const modal = document.getElementById('modal');
  if (e.key === 'Escape' && modal?.classList.contains('active')) closeModal();
});

/* =====================================================
   EVENTS GRID (homepage)
   ===================================================== */
function loadEvents() {
  const grid = document.getElementById('events-grid');
  if (!grid) return;
  grid.innerHTML = '<div class="skeleton-card"></div><div class="skeleton-card"></div><div class="skeleton-card"></div>';

  loadSheet(CONFIG.EVENTS_CSV, rows => {
      grid.innerHTML = '';
      rows
        .filter(e => e.Name && e.Name.trim())
        .sort((a, b) => ({previous:0,current:1,future:2}[a.Type] - {previous:0,current:1,future:2}[b.Type]))
        .forEach(e => {
          const t = e.Type || deriveType(e);
          const card = document.createElement('article');
          card.className = `event-card ${t}`;
          card.setAttribute('tabindex', '0');
          card.setAttribute('role', 'button');
          card.setAttribute('aria-label', `View details for ${e.Name}`);
          card.innerHTML = `
            <div class="event-badge badge-${t}">${badgeText(t)}</div>
            <div class="event-image">
              ${e.ImageURL ? `<img src="${e.ImageURL}" alt="${escapeHtml(e.Name)} event image" loading="lazy">` : ''}
              <div class="date-overlay">
                <div class="date-month">${escapeHtml(e.Month)}</div>
                <div class="date-day">${escapeHtml(e.Date)}</div>
                <div class="date-year">${escapeHtml(e.Year)}</div>
              </div>
            </div>
            <div class="event-content">
              <h2 class="event-name">${escapeHtml(e.Name)}</h2>
              ${e.Venue    ? `<div class="event-venue-info"><svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" class="venue-icon location-building"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"/></svg><span>${escapeHtml(e.Venue)}</span></div>` : ''}
              ${e.Location ? `<div class="event-venue-info"><svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" class="venue-icon venue-pin"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 11a3 3 0 11-6 0 3 3 0 016 0z"/></svg><span>${escapeHtml(e.Location)}</span></div>` : ''}
              <p class="event-description">${escapeHtml(e.Summary || e.Description || '')}</p>
            </div>`;
          card.addEventListener('click', function() { openModal(e, this); });
          card.addEventListener('keydown', function(ev) { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); openModal(e, this); } });
          grid.appendChild(card);
        });
    },
    () => { grid.innerHTML = '<p style="text-align:center;color:var(--apple-gray);grid-column:1/-1;">Events could not be loaded right now.</p>'; }
  );
}

/* =====================================================
   EVENT TIMELINE (event-timeline.html)
   ===================================================== */
let timelineData = {};

function loadTimeline() {
  const timeline = document.getElementById('timeline');
  if (!timeline) return;
  const loading = document.getElementById('timeline-loading');

  loadSheet(CONFIG.TIMELINE_CSV, rows => {
      timelineData = {};
      rows.forEach(e => { if (!e.Name && e.Title) e.Name = e.Title; });
      rows.filter(e => e.Name && e.Name.trim()).forEach(e => {
        const y = e.Year || 'Other';
        (timelineData[y] = timelineData[y] || []).push(e);
      });
      Object.values(timelineData).forEach(list => list.sort((a, b) => {
        const d = monthNum(a.Month) - monthNum(b.Month);
        return d !== 0 ? d : (parseInt(a.Date) || 0) - (parseInt(b.Date) || 0);
      }));

      const select = document.getElementById('yearSelect');
      select.innerHTML = '';
      Object.keys(timelineData).sort((a, b) => b - a).forEach(y => {
        const o = document.createElement('option');
        o.value = y; o.textContent = y;
        select.appendChild(o);
      });
      select.onchange = renderTimelineYear;

      if (loading) loading.style.display = 'none';
      document.getElementById('yearSelector').style.display = 'block';
      renderTimelineYear();
    },
    () => { if (loading) loading.textContent = 'Error loading events. Please try again.'; }
  );
}

function renderTimelineYear() {
  const year = document.getElementById('yearSelect').value;
  const timeline = document.getElementById('timeline');
  timeline.innerHTML = '';
  if (!timelineData[year]) return;

  timelineData[year].forEach((e, i) => {
    const t = deriveType(e);
    const item = document.createElement('div');
    item.className = `timeline-item ${t}`;
    item.style.animationDelay = `${i * 0.1}s`;
    item.innerHTML = `
      <div class="timeline-marker"></div>
      <div class="timeline-content" tabindex="0" role="button" aria-label="View details for ${escapeHtml(e.Name)}">
        <div class="month-label">${escapeHtml(e.Month)} ${escapeHtml(e.Year)}</div>
        ${e.ImageURL ? `<img class="timeline-event-image" src="${e.ImageURL}" alt="${escapeHtml(e.Name)}" loading="lazy">` : ''}
        <div class="timeline-event-title">${escapeHtml(e.Name)}</div>
        <div class="timeline-event-summary">${escapeHtml(e.Summary || e.Description || '')}</div>
      </div>`;
    const card = item.querySelector('.timeline-content');
    card.addEventListener('click', function() { openModal(e, this); });
    card.addEventListener('keydown', function(ev) { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); openModal(e, this); } });
    timeline.appendChild(item);
  });
}

/* =====================================================
   TESTIMONIALS (homepage)
   ===================================================== */
function loadTestimonials() {
  const container = document.getElementById('testimonials-container');
  if (!container) return;
  loadSheet(CONFIG.TESTIMONIALS_CSV, rows => {
      container.innerHTML = '';
      const active = rows
        .filter(t => t.Active && t.Active.toUpperCase() === 'YES')
        .sort((a, b) => parseInt(a['Display Order'] || 0) - parseInt(b['Display Order'] || 0));
      active.forEach(t => {
        const card = document.createElement('div');
        card.className = 'testimonial-card';
        const showPic = t['Show Profile Picture']?.toUpperCase() === 'YES' && t['Profile Image URL'];
        card.innerHTML = `
          ${showPic ? `<div class="testimonial-profile"><img src="${t['Profile Image URL']}" alt="${escapeHtml(t.Name)}" loading="lazy" onerror="this.style.display='none'"><div class="testimonial-name">${escapeHtml(t.Name)}</div></div>` : `<div style="margin-bottom:16px;"><div class="testimonial-name">${escapeHtml(t.Name)}</div></div>`}
          <p class="testimonial-text">"${escapeHtml(t.Testimonial)}"</p>`;
        container.appendChild(card);
      });
      if (!active.length) container.innerHTML = '<p style="text-align:center;color:var(--apple-gray);">No testimonials available at this time.</p>';
  });
}

/* =====================================================
   IMPACT SLIDER (homepage)
   ===================================================== */
let slides = [], currentSlideIndex = 0, slideInterval;

function loadImpactSlider() {
  if (!document.getElementById('slider')) return;
  if (location.protocol === 'file:') return;
  fetch(CONFIG.HIGHLIGHTS_JSON)
    .then(r => { if (!r.ok) throw new Error(r.status); return r.json(); })
    .then(items => {
      slides = items.filter(s => s.photo);
      buildSlider();
      showSlide(0);
      startAutoSlide();
    })
    .catch(err => console.error('Failed to load highlights', err));
}

/* =====================================================
   COMMUNITY PRIDE WALL (homepage)
   Auto-loads from data/logo-variants.json, which GitHub
   Actions regenerates whenever a file is added to or removed
   from assets/images/branding/logo-variants/. No coding
   needed — just drop an image in that folder and push (or
   upload it via GitHub's web UI) and it flows in automatically.
   ===================================================== */
function renderLogoPride() {
  const row1 = document.getElementById('pride-row-1');
  const row2 = document.getElementById('pride-row-2');
  if (!row1 || !row2) return;
  if (location.protocol === 'file:') return;
  fetch(CONFIG.LOGO_VARIANTS_JSON)
    .then(r => { if (!r.ok) throw new Error(r.status); return r.json(); })
    .then(items => {
      const srcs = items.map(i => i.src).filter(Boolean);
      if (!srcs.length) return;
      const mid = Math.ceil(srcs.length / 2);
      fillPrideRow(row1, srcs.slice(0, mid));
      fillPrideRow(row2, srcs.length > mid ? srcs.slice(mid) : srcs.slice(0, mid));
    })
    .catch(err => console.error('Failed to load logo variants', err));
}

function fillPrideRow(el, srcs) {
  const doubled = [...srcs, ...srcs]; // duplicate so the marquee loops seamlessly
  el.innerHTML = doubled.map(src =>
    `<div class="pride-chip"><img src="${escapeHtml(src)}" alt="Community logo artwork" loading="lazy"></div>`
  ).join('');
}

function buildSlider() {
  const slider = document.getElementById('slider');
  const thumbs = document.getElementById('thumbs');
  slider.innerHTML = ''; thumbs.innerHTML = '';
  slides.forEach((s, i) => {
    const slide = document.createElement('div');
    slide.className = 'slide';
    slide.innerHTML = `<img class="bg-img" src="${s.photo}" alt="" loading="lazy"><img class="fg-img" src="${s.photo}" onclick="openPopup('${s.photo}')" alt="${escapeHtml(s.title || 'Highlight photo')}" loading="lazy">${s.title ? `<div class="slide-title">${escapeHtml(s.title)}</div>` : ''}`;
    slider.appendChild(slide);
    const thumb = document.createElement('div');
    thumb.className = 'thumbnail';
    thumb.innerHTML = `<img src="${s.photo}" alt="${escapeHtml(s.title || '')}" loading="lazy">`;
    thumb.onclick = () => showSlide(i);
    thumbs.appendChild(thumb);
  });
}

function showSlide(i) {
  document.querySelectorAll('.slide').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.thumbnail').forEach(t => t.classList.remove('active'));
  currentSlideIndex = i;
  document.querySelectorAll('.slide')[i]?.classList.add('active');
  document.querySelectorAll('.thumbnail')[i]?.classList.add('active');
}

function nextSlide() { showSlide((currentSlideIndex + 1) % slides.length); }
function prevSlide() { showSlide((currentSlideIndex - 1 + slides.length) % slides.length); }
function startAutoSlide() { if (slideInterval) clearInterval(slideInterval); slideInterval = setInterval(nextSlide, CONFIG.SLIDER_INTERVAL); }
function openPopup(src) { document.getElementById('popupImg').src = src; document.getElementById('popup').classList.add('active'); clearInterval(slideInterval); }
function closePopup() { document.getElementById('popup').classList.remove('active'); startAutoSlide(); }

/* =====================================================
   PARTNERS (homepage)
   ===================================================== */
function renderPartnerLogos() {
  const track = document.getElementById('logos-track');
  if (!track) return;
  loadSheet(CONFIG.PARTNERS_CSV, rows => {
      const active = rows.filter(p => p.Name && (p.Active || '').toLowerCase() === 'yes');
      track.innerHTML = active.map(logo => `
        <div class="logo-item">
          <div class="logo-item-image">
            <img src="${logo.ImageURL}" alt="${escapeHtml(logo.Name)}" loading="lazy" onerror="this.closest('.logo-item').style.display='none';">
          </div>
          <div class="logo-item-name">${escapeHtml(logo.Name)}</div>
        </div>`).join('');
  });
}

/* =====================================================
   TEAM (our-team.html + marathi-shala.html)
   ===================================================== */
function renderTeamGrid(csv, containerId) {
  const container = document.getElementById(containerId);
  if (!container) return;
  loadSheet(csv, rows => {
      const members = rows.filter(m => m.Name && m.Name.trim());
      const groups = [...new Set(members.map(m => m.Group))];
      container.innerHTML = groups.map(g => `
        <h2 class="timeline-year" style="text-align:center;margin-top:40px;">${escapeHtml(g)}</h2>
        <div class="team-grid" style="padding-top:20px;padding-bottom:20px;">
          ${members.filter(m => m.Group === g).map(m => {
            const initials = m.Name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
            const avatar = m.Photo
              ? `<img class="team-avatar" src="${m.Photo}" alt="${escapeHtml(m.Name)}" loading="lazy" onerror="this.outerHTML='<div class=\\'team-avatar\\'>${initials}</div>'">`
              : `<div class="team-avatar">${initials}</div>`;
            return `
              <div class="team-card">
                ${avatar}
                <div class="team-name">${escapeHtml(m.Name)}</div>
                <div class="team-role">${escapeHtml(m.Position)}</div>
              </div>`;
          }).join('')}
        </div>`).join('');
  });
}

/* =====================================================
   FORMS & SIGN-UPS (forms.html)
   Self-service list of external form links (Google Forms,
   Tally, etc.) — no coding required to add/remove one, just
   edit data/forms.xlsx. See README "Updating Forms & Sign-ups".
   ===================================================== */
function renderForms() {
  const container = document.getElementById('forms-container');
  if (!container) return;
  loadSheet(CONFIG.FORMS_CSV, rows => {
      const forms = rows
        .filter(f => f.Title && f.Title.trim() && f.Link && f.Link.trim())
        .filter(f => !f.Active || f.Active.toLowerCase() !== 'no')
        .sort((a, b) => parseInt(a.Order || 0) - parseInt(b.Order || 0));

      if (!forms.length) {
        container.innerHTML = `<p style="text-align:center;color:var(--apple-gray);max-width:500px;margin:0 auto;">No forms are open right now — check back soon!</p>`;
        return;
      }

      container.innerHTML = `<div class="info-cards">
        ${forms.map(f => `
          <div class="info-card">
            <div class="about-icon"><i class="fas fa-file-lines"></i></div>
            <h4>${escapeHtml(f.Title)}</h4>
            ${f.Description ? `<p>${escapeHtml(f.Description)}</p>` : ''}
            <p style="margin-top:14px;"><a href="${f.Link}" target="_blank" rel="noopener" style="color:var(--brand-gradient-start);font-weight:600;text-decoration:none;">Open Form <i class="fas fa-arrow-up-right-from-square" style="font-size:12px;"></i></a></p>
          </div>`).join('')}
      </div>`;
  });
}

/* =====================================================
   SHOWCASE (showcase.html)
   ===================================================== */
function extractYouTubeId(url) {
  if (!url) return null;
  const m = String(url).match(/(?:youtu\.be\/|youtube\.com\/(?:watch\?v=|embed\/|shorts\/|v\/))([A-Za-z0-9_-]{11})/);
  return m ? m[1] : null;
}

function showcaseMediaHtml(v) {
  const videoId = extractYouTubeId(v.YouTubeURL);
  if (videoId) {
    return `<div class="showcase-media showcase-clickable-media" tabindex="0" role="button" aria-label="Play video: ${escapeHtml(v.Title)}" onclick="openVideoModal('${videoId}')" onkeydown="if(event.key==='Enter')openVideoModal('${videoId}')">
               <img src="https://img.youtube.com/vi/${videoId}/hqdefault.jpg" alt="${escapeHtml(v.Title)}" loading="lazy">
               <div class="showcase-hover-icon"><i class="fas fa-play"></i></div>
             </div>`;
  }
  if (v.ImageURL && v.ImageURL.trim()) {
    const safeTitle = escapeHtml(v.Title).replace(/'/g, "\\'");
    return `<div class="showcase-media showcase-clickable-media" tabindex="0" role="button" aria-label="View image: ${escapeHtml(v.Title)}" onclick="openShowcaseImage('${v.ImageURL}', '${safeTitle}')" onkeydown="if(event.key==='Enter')openShowcaseImage('${v.ImageURL}', '${safeTitle}')">
               <img src="${v.ImageURL}" alt="${escapeHtml(v.Title)}" loading="lazy">
               <div class="showcase-hover-icon"><i class="fas fa-expand"></i></div>
             </div>`;
  }
  // Document (PDF or other file) — opens in a new tab, no modal needed.
  const isPdf = /\.pdf(\?|$)/i.test(v.DocumentURL || '');
  return `<a class="showcase-media showcase-clickable-media" href="${v.DocumentURL}" target="_blank" rel="noopener" aria-label="Open document: ${escapeHtml(v.Title)}">
             <i class="fas ${isPdf ? 'fa-file-pdf' : 'fa-file-lines'}" style="font-size:48px;color:var(--brand-gradient-start);"></i>
             <div class="showcase-hover-icon"><i class="fas fa-arrow-up-right-from-square"></i></div>
           </a>`;
}

function showcaseCardHtml(v) {
  return `
    <div class="showcase-card">
      ${showcaseMediaHtml(v)}
      <div class="showcase-body">
        ${v.Category ? `<div class="showcase-meta">${escapeHtml(v.Category)}</div>` : ''}
        <div class="showcase-title">${escapeHtml(v.Title)}</div>
        ${v.Description ? `<p class="showcase-desc">${escapeHtml(v.Description)}</p>` : ''}
      </div>
    </div>`;
}

function renderShowcase() {
  const container = document.getElementById('showcase-grid');
  if (!container) return;
  loadSheet(CONFIG.SHOWCASE_CSV, rows => {
      const items = rows
        .filter(v => v.Title && v.Title.trim() && (extractYouTubeId(v.YouTubeURL) || (v.ImageURL && v.ImageURL.trim()) || (v.DocumentURL && v.DocumentURL.trim())))
        .filter(v => !v.Active || v.Active.toLowerCase() !== 'no')
        .sort((a, b) => parseInt(a.Order || 0) - parseInt(b.Order || 0));

      if (!items.length) {
        container.innerHTML = `<p style="text-align:center;color:var(--apple-gray);">Nothing to show yet — check back soon!</p>`;
        return;
      }

      // Group by Event (blank Event = ungrouped, no heading). Groups appear in
      // the order their first (lowest-Order) item shows up, same pattern as
      // the FAQ page's category grouping.
      const groups = [...new Set(items.map(v => (v.Event || '').trim()))];
      container.innerHTML = groups.map(g => `
        ${g ? `<h2 class="timeline-year" style="text-align:center;margin:40px 0 20px;">${escapeHtml(g)}</h2>` : ''}
        <div class="showcase-grid">
          ${items.filter(v => (v.Event || '').trim() === g).map(showcaseCardHtml).join('')}
        </div>`).join('');
  });
}

function openVideoModal(videoId) {
  const modal = document.getElementById('video-modal');
  const frame = document.getElementById('video-modal-frame');
  if (!modal || !frame || !videoId) return;
  frame.src = `https://www.youtube.com/embed/${videoId}?autoplay=1&rel=0`;
  modal.classList.add('active');
  document.body.classList.add('modal-open');
}

function closeVideoModal() {
  const modal = document.getElementById('video-modal');
  const frame = document.getElementById('video-modal-frame');
  if (!modal || !frame) return;
  modal.classList.remove('active');
  frame.src = '';
  document.body.classList.remove('modal-open');
}

function openShowcaseImage(src, alt) {
  const popup = document.getElementById('showcase-popup');
  const img = document.getElementById('showcase-popup-img');
  if (!popup || !img || !src) return;
  img.src = src;
  img.alt = alt || 'Showcase image';
  popup.classList.add('active');
}

function closeShowcaseImage() {
  const popup = document.getElementById('showcase-popup');
  if (!popup) return;
  popup.classList.remove('active');
}

/* =====================================================
   SHALA CALENDAR (calendar.html)
   ===================================================== */
const CALENDAR_TYPE_COLORS = {
  class: '#764ba2',
  event: '#E36C18',
  holiday: '#B5171D',
  exam: '#10b981'
};
const CALENDAR_TYPE_LABELS = { class: 'Class', event: 'Event', holiday: 'Holiday', exam: 'Exam' };
const CALENDAR_DEFAULT_COLOR = '#6b7280';
const CALENDAR_MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];

let calendarEntries = [];
let calendarViewYear = null;
let calendarViewMonth = null; // 0-indexed
let calendarSelectedKey = null;

function calendarColorFor(type) {
  return CALENDAR_TYPE_COLORS[(type || '').toLowerCase()] || CALENDAR_DEFAULT_COLOR;
}

function loadShalaCalendar() {
  if (!document.getElementById('calGrid')) return;
  loadSheet(CONFIG.SHALA_CALENDAR_CSV, rows => {
    calendarEntries = rows
      .map(r => {
        const year = parseInt(r.Year);
        const month = monthNum(r.Month); // 1-12
        const day = parseInt(r.Day);
        return { year, month, day, title: r.Title, type: r.Type, time: r.Time, notes: r.Notes };
      })
      .filter(e => e.title && e.year && e.month && e.day);

    buildCalendarLegend();

    const now = new Date();
    let vy = now.getFullYear(), vm = now.getMonth();
    const hasEntries = (y, m) => calendarEntries.some(e => e.year === y && e.month === m + 1);
    if (!hasEntries(vy, vm)) {
      let found = false;
      for (let i = 1; i <= 24 && !found; i++) {
        let ty = vy, tm = vm + i;
        ty += Math.floor(tm / 12); tm = ((tm % 12) + 12) % 12;
        if (hasEntries(ty, tm)) { vy = ty; vm = tm; found = true; }
      }
      if (!found) {
        for (let i = 1; i <= 24 && !found; i++) {
          let ty = vy, tm = vm - i;
          ty += Math.floor(tm / 12); tm = ((tm % 12) + 12) % 12;
          if (hasEntries(ty, tm)) { vy = ty; vm = tm; found = true; }
        }
      }
    }
    calendarViewYear = vy; calendarViewMonth = vm;
    renderCalendarMonth();
  }, () => {
    const grid = document.getElementById('calGrid');
    if (grid) grid.innerHTML = '<p style="grid-column:1/-1;text-align:center;color:var(--apple-gray);">Calendar could not be loaded right now.</p>';
  });

  document.getElementById('calPrev')?.addEventListener('click', () => shiftCalendarMonth(-1));
  document.getElementById('calNext')?.addEventListener('click', () => shiftCalendarMonth(1));
}

function shiftCalendarMonth(delta) {
  calendarViewMonth += delta;
  calendarViewYear += Math.floor(calendarViewMonth / 12);
  calendarViewMonth = ((calendarViewMonth % 12) + 12) % 12;
  calendarSelectedKey = null;
  renderCalendarMonth();
}

function buildCalendarLegend() {
  const legend = document.getElementById('calLegend');
  if (!legend) return;
  legend.innerHTML = Object.keys(CALENDAR_TYPE_LABELS).map(t => `
    <div class="calendar-legend-item">
      <span class="calendar-legend-dot" style="background:${CALENDAR_TYPE_COLORS[t]}"></span>
      <span>${CALENDAR_TYPE_LABELS[t]}</span>
    </div>`).join('');
}

function renderCalendarMonth() {
  const label = document.getElementById('calMonthLabel');
  const grid = document.getElementById('calGrid');
  if (!label || !grid) return;
  label.textContent = `${CALENDAR_MONTH_NAMES[calendarViewMonth]} ${calendarViewYear}`;

  const monthEntries = calendarEntries.filter(e => e.year === calendarViewYear && e.month === calendarViewMonth + 1);
  const firstWeekday = new Date(calendarViewYear, calendarViewMonth, 1).getDay();
  const daysInMonth = new Date(calendarViewYear, calendarViewMonth + 1, 0).getDate();
  const totalCells = Math.ceil((firstWeekday + daysInMonth) / 7) * 7;

  const today = new Date();
  const todayKey = `${today.getFullYear()}-${today.getMonth() + 1}-${today.getDate()}`;

  let html = '';
  for (let i = 0; i < totalCells; i++) {
    const dayNum = i - firstWeekday + 1;
    let cellDate, inMonth;
    if (dayNum < 1) {
      cellDate = new Date(calendarViewYear, calendarViewMonth, dayNum);
      inMonth = false;
    } else if (dayNum > daysInMonth) {
      cellDate = new Date(calendarViewYear, calendarViewMonth, dayNum);
      inMonth = false;
    } else {
      cellDate = new Date(calendarViewYear, calendarViewMonth, dayNum);
      inMonth = true;
    }
    const y = cellDate.getFullYear(), m = cellDate.getMonth() + 1, d = cellDate.getDate();
    const key = `${y}-${m}-${d}`;
    const dayEntries = inMonth ? monthEntries.filter(e => e.day === d) : [];
    const classes = ['calendar-day'];
    if (!inMonth) classes.push('other-month');
    if (key === todayKey) classes.push('today');
    if (dayEntries.length) classes.push('has-events');
    if (key === calendarSelectedKey) classes.push('selected');

    const shown = dayEntries.slice(0, 2);
    const extra = dayEntries.length - shown.length;
    const chips = shown.map(e => `<div class="calendar-chip" style="background:${calendarColorFor(e.type)}">${escapeHtml(e.title)}</div>`).join('')
      + (extra > 0 ? `<div class="calendar-chip-more">+${extra} more</div>` : '');
    const dots = dayEntries.slice(0, 4).map(e => `<span class="calendar-day-dot" style="background:${calendarColorFor(e.type)}"></span>`).join('');

    html += `<div class="${classes.join(' ')}" data-key="${key}">
      <div class="calendar-day-num">${d}</div>
      <div class="calendar-day-chips">${chips}</div>
      <div class="calendar-day-dots">${dots}</div>
    </div>`;
  }
  grid.innerHTML = html;

  grid.querySelectorAll('.calendar-day.has-events').forEach(cell => {
    cell.addEventListener('click', () => {
      calendarSelectedKey = calendarSelectedKey === cell.dataset.key ? null : cell.dataset.key;
      renderCalendarMonth();
    });
  });

  renderCalendarAgenda(monthEntries);
}

function renderCalendarAgenda(monthEntries) {
  const title = document.getElementById('calAgendaTitle');
  const list = document.getElementById('calAgendaList');
  if (!title || !list) return;

  if (calendarSelectedKey) {
    const [y, m, d] = calendarSelectedKey.split('-').map(Number);
    const entries = monthEntries.filter(e => e.day === d);
    const dateLabel = new Date(y, m - 1, d).toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric' });
    title.textContent = dateLabel;
    list.innerHTML = entries.length ? entries.map(calendarAgendaItem).join('') : '<p class="calendar-agenda-empty">Nothing scheduled.</p>';
    return;
  }

  title.textContent = `${CALENDAR_MONTH_NAMES[calendarViewMonth]} at a glance`;
  const sorted = [...monthEntries].sort((a, b) => a.day - b.day);
  list.innerHTML = sorted.length
    ? sorted.map(e => calendarAgendaItem(e, true)).join('')
    : '<p class="calendar-agenda-empty">No classes or events scheduled this month yet.</p>';
}

function calendarAgendaItem(e, showDate) {
  const dateStr = showDate ? `${CALENDAR_MONTH_NAMES[e.month - 1].slice(0, 3)} ${e.day}` : '';
  const meta = [e.time, e.notes].filter(Boolean).join(' · ');
  return `<div class="calendar-agenda-item">
    <span class="calendar-agenda-dot" style="background:${calendarColorFor(e.type)}"></span>
    <div>
      ${dateStr ? `<div class="calendar-agenda-date">${escapeHtml(dateStr)}</div>` : ''}
      <div class="calendar-agenda-name">${escapeHtml(e.title)}</div>
      ${meta ? `<div class="calendar-agenda-meta">${escapeHtml(meta)}</div>` : ''}
    </div>
  </div>`;
}

/* =====================================================
   WEB3FORMS SUBMISSION (join.html + sponsor.html)
   No Google Forms — forms POST via fetch() to Web3Forms,
   a free no-login email relay. Set CONFIG.WEB3FORMS_ACCESS_KEY
   (get one free at https://web3forms.com) for delivery to work.
   ===================================================== */
function initWeb3Form(formId) {
  const form = document.getElementById(formId);
  if (!form) return;

  const keyInput = form.querySelector('input[name="access_key"]');
  const statusEl = form.querySelector('.form-status');
  const submitBtn = form.querySelector('.form-submit-btn');
  const card = form.closest('.form-card');
  const successEl = card ? card.querySelector('.form-success') : null;

  const key = CONFIG.WEB3FORMS_ACCESS_KEY;
  if (keyInput) keyInput.value = key || '';

  form.addEventListener('submit', function(e) {
    e.preventDefault();

    if (!key || key === 'REPLACE_WITH_YOUR_WEB3FORMS_ACCESS_KEY') {
      if (statusEl) {
        statusEl.textContent = 'Form isn\'t fully set up yet — a Web3Forms access key is needed. Please contact us directly in the meantime.';
        statusEl.className = 'form-status form-status-error';
      }
      return;
    }

    const originalBtnText = submitBtn ? submitBtn.textContent : '';
    if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = 'Sending...'; }
    if (statusEl) { statusEl.textContent = ''; statusEl.className = 'form-status'; }

    fetch('https://api.web3forms.com/submit', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
      body: JSON.stringify(Object.fromEntries(new FormData(form)))
    })
      .then(res => res.json())
      .then(data => {
        if (data.success) {
          if (card && successEl) {
            form.style.display = 'none';
            successEl.style.display = 'block';
          } else if (statusEl) {
            statusEl.textContent = 'Thank you — your submission was received.';
            statusEl.className = 'form-status form-status-success';
            form.reset();
          }
        } else {
          throw new Error(data.message || 'Submission failed');
        }
      })
      .catch(() => {
        if (statusEl) {
          statusEl.textContent = 'Something went wrong sending your message. Please try again in a moment.';
          statusEl.className = 'form-status form-status-error';
        }
      })
      .finally(() => {
        if (submitBtn) { submitBtn.disabled = false; submitBtn.textContent = originalBtnText; }
      });
  });
}

/* =====================================================
   FAQ ACCORDION (faqs.html + marathi-shala.html)
   ===================================================== */
function renderFaqs(csv, containerId, groupByCategory) {
  const container = document.getElementById(containerId);
  if (!container) return;
  loadSheet(csv, rows => {
      let faqs = rows
        .filter(f => f.Question && f.Question.trim())
        .filter(f => !f.Active || f.Active.toLowerCase() !== 'no' || csv === CONFIG.SHALA_FAQS_CSV)
        .sort((a, b) => parseInt(a.Order || 0) - parseInt(b.Order || 0));

      const item = f => `
        <div class="faq-item">
          <button class="faq-question" aria-expanded="false">${escapeHtml(f.Question)}</button>
          <div class="faq-answer"><div class="faq-answer-inner">${escapeHtml(f.Answer)}</div></div>
        </div>`;

      if (groupByCategory) {
        const cats = [...new Set(faqs.map(f => f.Category || 'General'))];
        container.innerHTML = cats.map(c => `
          <h2 class="timeline-year" style="font-size:22px;">${escapeHtml(c)}</h2>
          ${faqs.filter(f => (f.Category || 'General') === c).map(item).join('')}`).join('');
      } else {
        container.innerHTML = faqs.map(item).join('');
      }
      initFaqAccordion(container);
  });
}

function initFaqAccordion(scope) {
  (scope || document).querySelectorAll('.faq-item').forEach(item => {
    const btn = item.querySelector('.faq-question');
    const answer = item.querySelector('.faq-answer');
    btn.addEventListener('click', () => {
      const isOpen = item.classList.contains('open');
      document.querySelectorAll('.faq-item.open').forEach(o => {
        o.classList.remove('open');
        o.querySelector('.faq-answer').style.maxHeight = null;
        o.querySelector('.faq-question').setAttribute('aria-expanded', 'false');
      });
      if (!isOpen) {
        item.classList.add('open');
        answer.style.maxHeight = answer.scrollHeight + 'px';
        btn.setAttribute('aria-expanded', 'true');
      }
    });
  });
}

/* =====================================================
   INIT
   ===================================================== */
document.addEventListener('DOMContentLoaded', function() {
  initNav();
  initBackToTop();
  initReveal();
  fileProtocolNotice();
  loadEvents();
  loadTestimonials();
  loadImpactSlider();
  renderPartnerLogos();
  renderLogoPride();
  loadTimeline();
  renderTeamGrid(CONFIG.TEAM_CSV, 'team-container');
  renderTeamGrid(CONFIG.SHALA_TEAM_CSV, 'shala-team-container');
  renderFaqs(CONFIG.FAQS_CSV, 'faq-container', false);
  renderFaqs(CONFIG.SHALA_FAQS_CSV, 'shala-faq-container', true);
  renderShowcase();
  loadShalaCalendar();
  renderForms();
  initWeb3Form('joinForm');
  initWeb3Form('sponsorForm');

  const modal = document.getElementById('modal');
  modal?.addEventListener('click', function(e) { if (e.target === this) closeModal(); });

  const videoModal = document.getElementById('video-modal');
  videoModal?.addEventListener('click', function(e) { if (e.target === this) closeVideoModal(); });

  const showcasePopup = document.getElementById('showcase-popup');
  showcasePopup?.addEventListener('click', function() { closeShowcaseImage(); });

  document.addEventListener('keydown', function(e) {
    if (e.key !== 'Escape') return;
    if (videoModal?.classList.contains('active')) closeVideoModal();
    if (showcasePopup?.classList.contains('active')) closeShowcaseImage();
  });

  const yearEl = document.getElementById('currentYear');
  if (yearEl) yearEl.textContent = new Date().getFullYear();
});
