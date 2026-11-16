/* site.js — behaviour shared by every page (loaded with `defer`).
   • translations   (lang/<code>.json, applied to [data-i18n] and [data-i18n-attr])
   • theme toggle   (day / night)
   • drop-down menus, mobile drawer, active page highlight
   • reveal-on-scroll animation
   Page-specific scripts (home.js, about.js, projects.js, contact.js) build on top of `window.site`. */
(function () {
  'use strict';

  var root = document.documentElement;
  var page = document.body.getAttribute('data-page') || 'home';
  var dicts = {};            // cache: lang -> dictionary
  var current = window.__lang || 'en';
  var currentDict = {};
  var ready = false;         // true once a dictionary has been applied

  /* ---------- helpers -------------------------------------------------- */
  function get(obj, path) {
    return path.split('.').reduce(function (o, k) { return o != null && o[k] != null ? o[k] : undefined; }, obj);
  }
  /** Translate a key; `fallback` is used until a dictionary is loaded. */
  function t(key, fallback) {
    var v = get(currentDict, key);
    return v != null ? v : fallback;
  }
  function loadDict(lang) {
    if (dicts[lang]) return Promise.resolve(dicts[lang]);
    var p = (lang === window.__lang && window.__dictPromise)
      ? window.__dictPromise
      : fetch(window.__root + 'lang/' + lang + '.json').then(function (r) {
          if (!r.ok) throw new Error('lang file missing: ' + lang);
          return r.json();
        });
    return p.then(function (d) { dicts[lang] = d; return d; });
  }
  function store(key, value) { try { localStorage.setItem(key, value); } catch (e) {} }

  /* ---------- translations --------------------------------------------- */
  function applyDict(dict, lang) {
    currentDict = dict; current = lang; ready = true;

    document.querySelectorAll('[data-i18n]').forEach(function (el) {
      var val = get(dict, el.getAttribute('data-i18n'));
      if (val == null) return;
      if (val.indexOf('<') !== -1) el.innerHTML = val; else el.textContent = val;
    });
    // data-i18n-attr="placeholder:contact.form.email_ph;aria-label:ui.language"
    document.querySelectorAll('[data-i18n-attr]').forEach(function (el) {
      el.getAttribute('data-i18n-attr').split(';').forEach(function (pair) {
        var i = pair.indexOf(':'); if (i < 0) return;
        var val = get(dict, pair.slice(i + 1).trim());
        if (val != null) el.setAttribute(pair.slice(0, i).trim(), val);
      });
    });

    var title = get(dict, 'meta.' + page);
    if (title) document.title = title;

    root.setAttribute('lang', lang);
    root.setAttribute('dir', lang === 'ar' ? 'rtl' : 'ltr');
    root.classList.remove('i18n-pending');

    // language controls
    var label = document.getElementById('langLabel');
    if (label) label.textContent = lang.toUpperCase();
    document.querySelectorAll('[data-lang]').forEach(function (el) {
      var on = el.getAttribute('data-lang') === lang;
      if (el.classList.contains('lang-btn')) el.setAttribute('aria-pressed', on ? 'true' : 'false');
      else if (on) el.setAttribute('aria-current', 'true'); else el.removeAttribute('aria-current');
    });
    updateThemeLabel();
    document.dispatchEvent(new CustomEvent('langchange', { detail: { lang: lang, dict: dict } }));
  }

  function setLang(lang, opts) {
    if (['en', 'fr', 'de', 'ar'].indexOf(lang) < 0) lang = 'en';
    store('lang', lang);
    if (opts && opts.user) try { // only when the visitor picks a language: …?lang=fr stays shareable, plain URLs stay clean
      var u = new URL(location.href);
      if (lang === 'en') u.searchParams.delete('lang'); else u.searchParams.set('lang', lang);
      history.replaceState({}, '', u.pathname + u.search + location.hash);
    } catch (e) {}
    return loadDict(lang).then(function (d) { applyDict(d, lang); })
      .catch(function () { root.classList.remove('i18n-pending'); });
  }

  /* ---------- theme ---------------------------------------------------- */
  function updateThemeLabel() {
    var btn = document.getElementById('themeToggle');
    if (!btn) return;
    var dark = root.getAttribute('data-theme') === 'dark';
    btn.textContent = dark ? t('ui.theme_day', '☀︎ Day') : t('ui.theme_night', '☾ Night');
  }
  function toggleTheme() {
    var next = root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
    root.setAttribute('data-theme', next);
    root.style.colorScheme = next;
    store('theme', next);
    updateThemeLabel();
  }

  /* ---------- drop-down menus ------------------------------------------ */
  function closeMenus(except) {
    document.querySelectorAll('.submenu.show').forEach(function (m) {
      if (m === except) return;
      m.classList.remove('show');
      var owner = m.parentNode.querySelector('[aria-expanded]');
      if (owner) owner.setAttribute('aria-expanded', 'false');
    });
  }
  function toggleMenu(btn) {
    var menu = document.getElementById(btn.getAttribute('aria-controls'));
    if (!menu) return;
    var open = !menu.classList.contains('show');
    closeMenus(open ? menu : null);
    menu.classList.toggle('show', open);
    btn.setAttribute('aria-expanded', open ? 'true' : 'false');
  }

  /* ---------- mobile drawer -------------------------------------------- */
  function setDrawer(open) {
    root.classList.toggle('mobile-open', open);
    var drawer = document.getElementById('mobile-drawer');
    var btn = document.getElementById('hamburgerBtn');
    if (drawer) drawer.setAttribute('aria-hidden', open ? 'false' : 'true');
    if (btn) btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    if (open) { var c = document.getElementById('mobileCloseBtn'); if (c) setTimeout(function () { c.focus({ preventScroll: true }); }, 50); }
  }

  /* ---------- wiring --------------------------------------------------- */
  document.addEventListener('click', function (e) {
    var target = e.target;

    var langEl = target.closest('[data-lang]');
    if (langEl) {
      e.preventDefault();
      setLang(langEl.getAttribute('data-lang'), { user: true });
      closeMenus(); setDrawer(false);
      return;
    }
    if (target.closest('#themeToggle')) { toggleTheme(); return; }

    var chevBtn = target.closest('.chev, #langToggle');
    if (chevBtn) { toggleMenu(chevBtn); return; }

    if (target.closest('#hamburgerBtn')) { setDrawer(!root.classList.contains('mobile-open')); return; }
    if (target.closest('#mobileCloseBtn') || target.closest('#mobileBackdrop')) { setDrawer(false); return; }
    if (target.closest('.mobile-drawer a')) setDrawer(false);

    // click on the logo of the home page = smooth scroll to the top
    var brand = target.closest('a.brand');
    if (brand && page === 'home') { e.preventDefault(); window.scrollTo({ top: 0, behavior: 'smooth' }); }

    if (!target.closest('.menu')) closeMenus();
    else if (target.closest('.submenu a')) closeMenus();
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') { closeMenus(); setDrawer(false); }
  });
  window.addEventListener('scroll', function () { closeMenus(); }, { passive: true });
  window.addEventListener('resize', function () { if (window.innerWidth > 640) setDrawer(false); }, { passive: true });

  function init() {
    // active page in both navigations
    document.querySelectorAll('[data-nav="' + page + '"]').forEach(function (a) {
      if (a.tagName === 'A' && !a.classList.contains('brand')) a.setAttribute('aria-current', 'page');
    });

    // reveal on scroll
    var els = document.querySelectorAll('.reveal');
    var reduced = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (reduced || !('IntersectionObserver' in window)) {
      els.forEach(function (el) { el.classList.add('in'); });
    } else {
      var io = new IntersectionObserver(function (entries) {
        entries.forEach(function (en) { if (en.isIntersecting) { en.target.classList.add('in'); io.unobserve(en.target); } });
      }, { threshold: 0.15 });
      els.forEach(function (el) { io.observe(el); });
    }

    updateThemeLabel();
    setLang(current);
  }

  window.site = { t: t, setLang: setLang, get lang() { return current; }, get dict() { return currentDict; }, get ready() { return ready; } };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
