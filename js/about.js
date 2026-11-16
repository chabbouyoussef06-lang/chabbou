/* about.js — keeps the CV / certificate buttons pointing at the right PDF for the chosen
   language, and replaces a button by "Coming soon" when the PDF has not been uploaded yet. */
(function () {
  'use strict';
  var cfg = window.SITE_CONFIG || {};
  var exists = {}; // url -> Promise<boolean>

  function check(url) {
    if (!exists[url]) {
      exists[url] = fetch(url, { method: 'HEAD' }).then(function (r) { return r.ok; }).catch(function () { return false; });
    }
    return exists[url];
  }

  function refresh(lang) {
    document.querySelectorAll('[data-doc]').forEach(function (a) {
      var entry = (cfg.docs || {})[a.getAttribute('data-doc')];
      if (!entry) return;
      var url = window.__root + (cfg.docsFolder || '') + encodeURI(entry[lang] || entry.en);
      var label = a.querySelector('[data-label]');
      check(url).then(function (ok) {
        a.classList.toggle('is-missing', !ok);
        if (ok) { a.href = url; a.removeAttribute('aria-disabled'); a.removeAttribute('tabindex'); }
        else { a.removeAttribute('href'); a.setAttribute('aria-disabled', 'true'); a.setAttribute('tabindex', '-1'); }
        if (label) {
          label.setAttribute('data-i18n', ok ? 'ui.download' : 'ui.unavailable');
          label.textContent = window.site.t(ok ? 'ui.download' : 'ui.unavailable', ok ? 'Download' : 'Coming soon');
        }
      });
    });
  }

  document.addEventListener('langchange', function (e) { refresh(e.detail.lang); });
  if (window.site && window.site.ready) refresh(window.site.lang); // dictionary may already be applied
})();
