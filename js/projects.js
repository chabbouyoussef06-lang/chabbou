/* projects.js — photo / video carousel of the realized project.
   • Photo slides (3 photos each) advance by themselves; the video slide (3 films) plays once, then the carousel
     goes back to the start.
   • With a mouse, resting the cursor on the media freezes everything: photos stay, videos loop.
   • Arrows only appear while the cursor is over the media (css); swipe and dots work everywhere. */
(function () {
  'use strict';
  var box = document.getElementById('realizedMedia');
  if (!box) return;

  var stage = box.querySelector('.media-stage') || box;
  var track = box.querySelector('.media-track');
  var slides = Array.prototype.slice.call(track.children);
  var prev = box.querySelector('.media-nav.prev');
  var next = box.querySelector('.media-nav.next');
  var videos = Array.prototype.slice.call(box.querySelectorAll('video'));
  var master = videos[0];
  var dotsBox = box.querySelector('.media-dots');
  var PHOTO_MS = 3500;                       // time each photo slide stays on screen
  var index = 0, timer = null, hovering = false;

  var dots = slides.map(function (_, i) {
    var b = document.createElement('button');
    b.type = 'button'; b.className = 'media-dot'; b.setAttribute('aria-label', (i + 1) + ' / ' + slides.length);
    b.addEventListener('click', function () { show(i); });
    dotsBox.appendChild(b); return b;
  });

  function isVideo(i) { return !!slides[i].querySelector('video'); }
  function clearTimer() { if (timer) { clearTimeout(timer); timer = null; } }

  function playVideos() {
    videos.forEach(function (v) { v.currentTime = 0; v.muted = true; v.loop = false; var p = v.play(); if (p && p.catch) p.catch(function () {}); });
  }
  function pauseVideos() { videos.forEach(function (v) { v.pause(); }); }

  // the three films share one length: when they are all over, loop if the cursor rests on them, otherwise start over
  videos.forEach(function (v) {
    v.addEventListener('ended', function () {
      if (!videos.every(function (x) { return x.ended; })) return;
      if (hovering) playVideos(); else show(index + 1);
    });
  });
  // keep the films in step
  if (master) master.addEventListener('timeupdate', function () {
    videos.forEach(function (v) { if (v !== master && !v.ended && Math.abs(v.currentTime - master.currentTime) > 0.25) v.currentTime = master.currentTime; });
  });

  function schedule() {                      // photos only: next slide after PHOTO_MS, unless the cursor rests on them
    clearTimer();
    if (isVideo(index)) return;
    timer = setTimeout(function () { if (hovering) schedule(); else show(index + 1); }, PHOTO_MS);
  }

  function show(i) {
    pauseVideos();
    index = (i + slides.length) % slides.length;
    track.style.transform = 'translateX(-' + index * 100 + '%)';
    dots.forEach(function (d, n) { d.classList.toggle('active', n === index); });
    if (isVideo(index)) { clearTimer(); playVideos(); } else schedule();
  }

  next.addEventListener('click', function () { show(index + 1); });
  prev.addEventListener('click', function () { show(index - 1); });

  // mouse only: touch has no "resting cursor"
  stage.addEventListener('pointerenter', function (e) { if (e.pointerType === 'mouse') hovering = true; });
  stage.addEventListener('pointerleave', function (e) {
    if (e.pointerType !== 'mouse') return;
    hovering = false; schedule();
  });

  var startX = null;
  box.addEventListener('touchstart', function (e) { startX = e.touches[0].clientX; }, { passive: true });
  box.addEventListener('touchend', function (e) {
    if (startX == null) return;
    var dx = e.changedTouches[0].clientX - startX; startX = null;
    if (Math.abs(dx) > 40) show(index + (dx < 0 ? 1 : -1));
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
