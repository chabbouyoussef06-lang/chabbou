/* contact.js - contact form.
   • With a Formspree address in js/config.js: sends the message (attachment included) from the page.
   • Without it: opens the visitor's e-mail app with the message pre-filled (mailto):
     the form still works, only the attachment field is hidden because mailto cannot carry files. */
(function () {
  'use strict';
  var cfg = window.SITE_CONFIG || {};
  var live = !!cfg.formEndpoint && !/X{4,}/.test(cfg.formEndpoint);
  // Clicking the e-mail address copies it instead of opening a mail app
  var mailLink = document.getElementById('copyMail');
  if (mailLink) {
    var original = mailLink.textContent.trim();
    mailLink.addEventListener('click', function (e) {
      e.preventDefault();
      var done = function () {
        mailLink.classList.add('is-copied');
        mailLink.textContent = '\u2713 ' + window.site.t('contact.info.copied', 'Copied to clipboard');
        clearTimeout(done.timer);
        done.timer = setTimeout(function () { mailLink.textContent = original; mailLink.classList.remove('is-copied'); }, 2200);
      };
      var fallback = function () {
        var ta = document.createElement('textarea');
        ta.value = original; ta.setAttribute('readonly', ''); ta.style.position = 'fixed'; ta.style.opacity = '0';
        document.body.appendChild(ta); ta.select();
        try { document.execCommand('copy'); done(); } catch (err) {}
        document.body.removeChild(ta);
      };
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(original).then(done, fallback);
      else fallback();
    });
  }

  var form = document.getElementById('contactForm');
  var status = document.getElementById('formStatus');
  if (!form) return;

  var t = function (k, fb) { return window.site.t(k, fb); };
  function say(kind, key, fallback) {
    status.className = 'status' + (kind ? ' ' + kind : '');
    status.setAttribute('data-i18n', key);
    status.textContent = t(key, fallback);
  }

  // Browser validation bubbles follow the OS language, not the page language: use the site's own wording.
  function problem(el) {
    if (el.validity.valueMissing) return t('contact.form.required', 'Please fill in this field.');
    if (el.validity.typeMismatch) return t('contact.form.email_invalid', 'Please enter a valid e-mail address.');
    return '';
  }
  Array.prototype.forEach.call(form.querySelectorAll('input, textarea'), function (el) {
    el.addEventListener('invalid', function () { el.setCustomValidity(''); el.setCustomValidity(problem(el)); });
    el.addEventListener('input', function () { el.setCustomValidity(''); });
  });

  // Not live yet → explain what will happen and hide the attachment field
  if (!live) {
    var att = document.getElementById('attachmentField');
    if (att) att.hidden = true;
    var note = document.getElementById('mailtoNote');
    if (note) note.hidden = false;
  }

  // Remember the visitor's address for next time (stored in their own browser only)
  var emailInput = document.getElementById('fromEmail');
  try { var saved = localStorage.getItem('visitor_email'); if (saved && emailInput) emailInput.value = saved; } catch (e) {}

  form.addEventListener('submit', function (e) {
    e.preventDefault();
    if (!form.checkValidity()) { form.reportValidity(); return; }
    var data = new FormData(form);
    try { localStorage.setItem('visitor_email', data.get('email') || ''); } catch (err) {}
    if (data.get('_gotcha')) return; // honeypot: bots fill the hidden field

    if (!live) {
      var body = data.get('message') + '\n\n- ' + data.get('email');
      location.href = 'mailto:' + cfg.contactEmail + '?subject=' + encodeURIComponent(data.get('subject')) + '&body=' + encodeURIComponent(body);
      say('ok', 'contact.form.mailto_opened', 'Your e-mail app should now be open.');
      return;
    }

    var button = form.querySelector('button.submit');
    button.disabled = true;
    say('', 'contact.form.sending', 'Sending…');
    fetch(cfg.formEndpoint, { method: 'POST', body: data, headers: { Accept: 'application/json' } })
      .then(function (r) {
        if (r.ok) { form.reset(); say('ok', 'contact.form.success', 'Message sent successfully. Thank you!'); }
        else say('err', 'contact.form.error', 'Sending failed. Please try again or email me directly.');
      })
      .catch(function () { say('err', 'contact.form.error_network', 'Network error. Please try again.'); })
      .then(function () { button.disabled = false; });
  });
})();
