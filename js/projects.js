/* projects.js — photo / video carousel of the realized project.
   Photos auto-advance; the video slide (3 films side by side) plays in a loop until the visitor moves on.
   Arrows, swipe, dots; pauses when the tab is hidden. */
(function () {
  'use strict';
  var box = document.getElementById('realizedMedia');
  if (!box) return;

  var track = box.querySelector('.media-track');
  var slides = Array.prototype.slice.call(track.children);
  var prev = box.querySelector('.media-nav.prev');
  var next = box.querySelector('.media-nav.next');
  var videos = Array.prototype.slice.call(box.querySelectorAll('video'));
  var master = videos[0];
  var dotsBox = box.querySelector('.media-dots');
  var AUTO_MS = 4500;
  var index = 0, auto = true, timer = null;

  // dots
  var dots = slides.map(function (_, i) {
    var b = document.createElement('button');
    b.type = 'button'; b.className = 'media-dot'; b.setAttribute('aria-label', (i + 1) + ' / ' + slides.length);
    b.addEventListener('click', function () { stopAuto(); show(i); });
    dotsBox.appendChild(b); return b;
  });

  function isVideo(i) { return !!slides[i].querySelector('video'); }
  function clearTimer() { if (timer) { clearInterval(timer); timer = null; } }
  function playVideos() {
    videos.forEach(function (v) { v.currentTime = 0; v.muted = true; v.loop = true; var p = v.play(); if (p && p.catch) p.catch(function () {}); });
  }
  function pauseVideos() { videos.forEach(function (v) { v.pause(); }); }
  // keep the three films in step after each loop
  if (master) master.addEventListener('timeupdate', function () {
    videos.forEach(function (v) { if (v !== master && Math.abs(v.currentTime - master.currentTime) > 0.25) v.currentTime = master.currentTime; });
  });

  function show(i) {
    pauseVideos();
    index = (i + slides.length) % slides.length;
    track.style.transform = 'translateX(-' + index * 100 + '%)';
    dots.forEach(function (d, n) { d.classList.toggle('active', n === index); });
    clearTimer();

    if (isVideo(index)) playVideos();            // stays here, looping, until the visitor moves on
    else if (auto) timer = setInterval(function () { show(index + 1); }, AUTO_MS);
  }
  function stopAuto() { auto = false; clearTimer(); }

  next.addEventListener('click', function () { stopAuto(); show(index + 1); });
  prev.addEventListener('click', function () { stopAuto(); show(index - 1); });

  var startX = null;
  box.addEventListener('touchstart', function (e) { startX = e.touches[0].clientX; }, { passive: true });
  box.addEventListener('touchend', function (e) {
    if (startX == null) return;
    var dx = e.changedTouches[0].clientX - startX; startX = null;
    if (Math.abs(dx) > 40) { stopAuto(); show(index + (dx < 0 ? 1 : -1)); }
  }, { passive: true });

  document.addEventListener('visibilitychange', function () {
    if (document.hidden) { clearTimer(); pauseVideos(); } else show(index);
  });

  show(0);

  /* "Rotate your phone" hint: phones in portrait only, disappears on touch, rotation or after a few seconds */
  var hint = box.querySelector('.rotate-overlay');
  if (hint && window.matchMedia('(max-width:640px)').matches && window.innerHeight > window.innerWidth) {
    var hide = function () { hint.classList.add('hidden'); };
    box.addEventListener('touchstart', hide, { passive: true, once: true });
    box.addEventListener('click', hide, { once: true });
    window.addEventListener('resize', function () { if (window.innerWidth > window.innerHeight) hide(); }, { passive: true });
    setTimeout(hide, 6000);
  } else if (hint) { hint.classList.add('hidden'); }
})();
