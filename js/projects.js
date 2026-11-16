/* projects.js — photo / video carousel of the realized project.
   Auto-advances (waits for the video to end), arrows, swipe, dots; pauses when the tab is hidden. */
(function () {
  'use strict';
  var box = document.getElementById('realizedMedia');
  if (!box) return;

  var track = box.querySelector('.media-track');
  var slides = Array.prototype.slice.call(track.children);
  var prev = box.querySelector('.media-nav.prev');
  var next = box.querySelector('.media-nav.next');
  var video = box.querySelector('video');
  var dotsBox = box.querySelector('.media-dots');
  var AUTO_MS = 4500;
  var index = 0, auto = true, timer = null;

  // dots
  var dots = slides.map(function (_, i) {
    var b = document.createElement('button');
    b.type = 'button'; b.className = 'media-dot'; b.setAttribute('aria-label', (i + 1) + ' / ' + slides.length);
    b.addEventListener('click', function () { stopAuto(); show(i, true); });
    dotsBox.appendChild(b); return b;
  });

  function isVideo(i) { return !!slides[i].querySelector('video'); }
  function clearTimer() { if (timer) { clearInterval(timer); timer = null; } }
  function onVideoEnded() { show(index + 1, false); }

  function show(i, manual) {
    if (video) { video.pause(); video.removeEventListener('ended', onVideoEnded); }
    index = (i + slides.length) % slides.length;
    track.style.transform = 'translateX(-' + index * 100 + '%)';
    dots.forEach(function (d, n) { d.classList.toggle('active', n === index); });
    clearTimer();

    if (isVideo(index) && video) {
      video.currentTime = 0; video.muted = true;
      video.loop = manual || !auto;            // a manual stop on the video = loop it
      var p = video.play(); if (p && p.catch) p.catch(function () {});
      if (!video.loop) video.addEventListener('ended', onVideoEnded);
    } else if (auto) {
      timer = setInterval(function () { show(index + 1, false); }, AUTO_MS);
    }
  }
  function stopAuto() {
    auto = false; clearTimer();
    if (video) { video.removeEventListener('ended', onVideoEnded); if (isVideo(index)) video.loop = true; }
  }

  next.addEventListener('click', function () { stopAuto(); show(index + 1, true); });
  prev.addEventListener('click', function () { stopAuto(); show(index - 1, true); });

  var startX = null;
  box.addEventListener('touchstart', function (e) { startX = e.touches[0].clientX; }, { passive: true });
  box.addEventListener('touchend', function (e) {
    if (startX == null) return;
    var dx = e.changedTouches[0].clientX - startX; startX = null;
    if (Math.abs(dx) > 40) { stopAuto(); show(index + (dx < 0 ? 1 : -1), true); }
  }, { passive: true });

  document.addEventListener('visibilitychange', function () {
    if (document.hidden) { clearTimer(); if (video) video.pause(); } else if (auto) show(index, false);
  });

  /* The clip is three portrait films side by side: paint each third on its own rounded canvas tile,
     exactly like the photo slides (works in every browser, no CSS masking). */
  var canvases = Array.prototype.slice.call(box.querySelectorAll('.media-tiles canvas'));
  // The three films are not exactly thirds (seams at x≈440 and x≈836 of 1280): crop inside each panel, away from the seams.
  var PANELS = [[0, 436], [446, 831], [840, 1280]];
  function paint(src, sw, sh) {
    if (!sw || !sh) return;
    var k1 = sw / 1280;
    canvases.forEach(function (c, k) {
      var ctx = c.getContext('2d'), x0 = PANELS[k][0] * k1, pw = (PANELS[k][1] - PANELS[k][0]) * k1;
      var ta = c.width / c.height, sa = pw / sh;
      var cw = sa > ta ? sh * ta : pw, ch = sa > ta ? sh : pw / ta;
      ctx.drawImage(src, x0 + (pw - cw) / 2, (sh - ch) / 2, cw, ch, 0, 0, c.width, c.height);
    });
  }
  if (video && canvases.length) {
    var poster = new Image();
    poster.onload = function () { if (video.paused && !video.currentTime) paint(poster, poster.naturalWidth, poster.naturalHeight); };
    poster.src = video.getAttribute('poster');
    var drawing = false;
    var frame = function () {
      paint(video, video.videoWidth, video.videoHeight);
      if (!video.paused && !video.ended) requestAnimationFrame(frame); else drawing = false;
    };
    var kick = function () { if (!drawing) { drawing = true; requestAnimationFrame(frame); } };
    video.addEventListener('play', kick);
    video.addEventListener('playing', kick);
    video.addEventListener('seeked', function () { paint(video, video.videoWidth, video.videoHeight); });
    video.addEventListener('loadeddata', function () { paint(video, video.videoWidth, video.videoHeight); });
  }

  show(0, false);

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
