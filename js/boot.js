/* boot.js — runs first, blocking, in <head>.
   Decides theme + language BEFORE the page is painted so nothing flashes
   (no white flash in dark mode, no English flash in French/Arabic, no LTR→RTL jump),
   and starts downloading the translation file right away. */
(function () {
  var root = document.documentElement;
  // Site root URL (ends with '/'), derived from this script's own address so pages can live in sub-folders.
  window.__root = new URL('..', document.currentScript.src).href;
  var LANGS = ['en', 'fr', 'de', 'ar'];

  // Theme: saved choice → system preference → light
  var theme = 'light';
  try {
    var savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'dark' || savedTheme === 'light') theme = savedTheme;
    else if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) theme = 'dark';
  } catch (e) {}
  root.setAttribute('data-theme', theme);
  root.style.colorScheme = theme;

  // Language: ?lang=xx → saved choice → English
  var lang = 'en';
  try {
    var q = new URLSearchParams(location.search).get('lang');
    var saved = localStorage.getItem('lang');
    lang = (q || saved || 'en').toLowerCase();
    if (LANGS.indexOf(lang) < 0) lang = 'en';
    if (q) localStorage.setItem('lang', lang);
  } catch (e) {}
  root.setAttribute('lang', lang);
  root.setAttribute('dir', lang === 'ar' ? 'rtl' : 'ltr');

  // Start fetching the dictionary immediately; site.js picks the promise up.
  window.__lang = lang;
  window.__dictPromise = fetch(window.__root + 'lang/' + lang + '.json', { cache: 'no-cache' }).then(function (r) {
    if (!r.ok) throw new Error('lang file missing: ' + lang);
    return r.json();
  });
  window.__dictPromise.catch(function () {});

  // Hide translatable text until the dictionary is applied (English is already in the HTML).
  if (lang !== 'en') {
    root.classList.add('i18n-pending');
    setTimeout(function () { root.classList.remove('i18n-pending'); }, 2500); // fail-safe
  }
})();
