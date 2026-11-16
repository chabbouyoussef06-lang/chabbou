/* home.js — hero video fade-in, paper plane that follows the scroll, "scroll" cue. */
(function () {
  'use strict';
  var stage = document.querySelector('.flight-stage');
  var plane = document.querySelector('.plane');
  var cue = document.querySelector('.hero-scroll-cue');

  function clamp(v, min, max) { return Math.max(min, Math.min(max, v)); }

  function animatePlane() {
    if (!stage || !plane) return;
    var h = window.innerHeight, p = clamp(window.scrollY / h, 0, 1);
    var x = p * Math.min(0.65 * window.innerWidth, 900);
    var climb = 0, startClimb = 0.35;
    if (p > startClimb) {
      var pp = (p - startClimb) / (1 - startClimb), k = 3.2;
      climb = -(Math.expm1(k * pp) / Math.expm1(k)) * (0.55 * window.innerHeight);
    }
    plane.style.transform = 'translate(' + x + 'px,' + (-12 + climb) + 'px)rotate(' + Math.max(-25, -5 - p * 35) + 'deg)';
    plane.style.opacity = p < 1 ? 0.95 : 0;
  }
  function updateCue() { if (cue) cue.classList.toggle('is-hidden', window.scrollY > 60); }

  window.addEventListener('scroll', function () { animatePlane(); updateCue(); }, { passive: true });
  window.addEventListener('resize', animatePlane);
  animatePlane();

  // Hero video: fade in once playable, quietly fall back to the poster image if it cannot load.
  var v = document.getElementById('heroVideo');
  if (v) {
    v.addEventListener('loadeddata', function () { v.classList.add('loaded'); });
    v.addEventListener('error', function () { v.style.display = 'none'; });
    setTimeout(function () { if (!v.classList.contains('loaded')) v.style.display = 'none'; }, 3000);
  }
})();
