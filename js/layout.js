/* layout.js - the header, the mobile drawer and the footer exist ONCE, here.
   Every page includes this file twice, synchronously:
     <script src="js/layout.js" data-slot="header"></script>   (top of <body>)
     <script src="js/layout.js" data-slot="footer"></script>   (end of <body>)
   English text is written inline so the page is complete even before the
   translation file arrives; site.js translates it afterwards.
   To add a page or change a menu entry, edit this file only. */
(function () {
  var script = document.currentScript;
  var slot = script && script.getAttribute('data-slot');
  var R = window.__root || '';   // site root, set by boot.js
  var PAGES = { home: R, about: R + 'about/', skills: R + 'skills/', projects: R + 'projects/', contact: R + 'contact/' };

  var ABOUT = [
    [PAGES.about + '#profile', 'nav.about_menu.who', 'Profile'],
    [PAGES.about + '#background', 'nav.about_menu.background', 'Background'],
    [PAGES.about + '#cv', 'nav.about_menu.cv', 'Curriculum Vitae'],
    [PAGES.about + '#certs', 'nav.about_menu.certs', 'Certifications']
  ];
  var PROJECTS = [
    [PAGES.projects + '#realized', 'nav.projects_menu.realized', 'Projects realized'],
    [PAGES.projects + '#current', 'nav.projects_menu.current', 'Current projects'],
    [PAGES.projects + '#future', 'nav.projects_menu.future', 'Future projects']
  ];
  var LANGS = [['en', 'EN', 'English'], ['fr', 'FR', 'Français'], ['de', 'DE', 'Deutsch'], ['ar', 'AR', 'العربية']];

  function links(list) {
    return list.map(function (l) {
      return '<a href="' + l[0] + '" data-i18n="' + l[1] + '">' + l[2] + '</a>';
    }).join('');
  }
  function chev(id, key, label) {
    return '<button class="chev" type="button" aria-haspopup="true" aria-expanded="false" aria-controls="' + id +
      '" aria-label="' + label + '" data-i18n-attr="aria-label:' + key + '"></button>';
  }

  var HEADER =
    '<a class="skip-link" href="#main" data-i18n="ui.skip">Skip to content</a>' +
    '<div class="top-band"><div class="container"><header>' +
      '<button class="hamburger-btn" id="hamburgerBtn" type="button" aria-label="Open menu" data-i18n-attr="aria-label:ui.open_menu" aria-controls="mobile-drawer" aria-expanded="false"><span aria-hidden="true"></span></button>' +
      '<a class="brand" href="' + PAGES.home + '" data-nav="home"><img src="' + R + 'assets/images/signature.png" alt="Youssef Chabbou" class="signature" width="812" height="307" fetchpriority="high"></a>' +
      '<nav aria-label="Main">' +
        '<div class="menu"><a href="' + PAGES.about + '" data-nav="about" data-i18n="nav.about">About</a>' + chev('about-menu', 'nav.about', 'About') +
          '<div id="about-menu" class="submenu">' + links(ABOUT) + '</div></div>' +
        '<a href="' + PAGES.skills + '" data-nav="skills" data-i18n="nav.skills">Skills</a>' +
        '<div class="menu"><a href="' + PAGES.projects + '" data-nav="projects" data-i18n="nav.projects">Projects</a>' + chev('projects-menu', 'nav.projects', 'Projects') +
          '<div id="projects-menu" class="submenu">' + links(PROJECTS) + '</div></div>' +
        '<a href="' + PAGES.contact + '" data-nav="contact" data-i18n="nav.contact">Contact</a>' +
      '</nav>' +
      '<div class="menu lang-menu">' +
        '<button id="langToggle" class="toggle" type="button" aria-haspopup="true" aria-expanded="false" aria-controls="lang-menu" aria-label="Language" data-i18n-attr="aria-label:ui.language"><span id="langLabel">EN</span><i class="caret" aria-hidden="true"></i></button>' +
        '<div id="lang-menu" class="submenu">' +
          LANGS.map(function (l) { return '<a href="#" data-lang="' + l[0] + '">' + l[2] + '</a>'; }).join('') +
        '</div>' +
      '</div>' +
      '<button id="themeToggle" class="toggle" type="button">☾ Night</button>' +
    '</header></div></div>' +

    '<div class="mobile-backdrop" id="mobileBackdrop"></div>' +
    '<aside class="mobile-drawer" id="mobile-drawer" aria-hidden="true" aria-label="Menu">' +
      '<div class="drawer-head"><strong class="mono" data-i18n="ui.menu">Menu</strong>' +
        '<button class="toggle drawer-close" type="button" id="mobileCloseBtn" aria-label="Close menu" data-i18n-attr="aria-label:ui.close_menu">✕</button></div>' +
      '<nav aria-label="Mobile navigation">' +
        '<a href="' + PAGES.about + '" data-nav="about" data-i18n="nav.about">About</a><div class="mobile-sub">' + links(ABOUT) + '</div>' +
        '<a href="' + PAGES.skills + '" data-nav="skills" data-i18n="nav.skills">Skills</a>' +
        '<a href="' + PAGES.projects + '" data-nav="projects" data-i18n="nav.projects">Projects</a><div class="mobile-sub">' + links(PROJECTS) + '</div>' +
        '<a href="' + PAGES.contact + '" data-nav="contact" data-i18n="nav.contact">Contact</a>' +
      '</nav>' +
      '<div class="lang-grid" role="group" aria-label="Language" data-i18n-attr="aria-label:ui.language">' +
        LANGS.map(function (l) { return '<button class="lang-btn" type="button" data-lang="' + l[0] + '" aria-pressed="false">' + l[1] + '</button>'; }).join('') +
      '</div>' +
    '</aside>';

  var FOOTER =
    '<footer class="footer-band"><div class="footer-grid">' +
      '<div class="footer-brand-block"><span class="footer-brand-name">Youssef Chabbou</span>' +
        '<span class="footer-brand-tag" data-i18n="footer.brand_tag">Aerospace · Engineering · Design</span></div>' +
      '<div class="footer-nav"><span class="footer-nav-title" data-i18n="footer.explore">Explore</span>' +
        '<a href="' + PAGES.about + '" data-i18n="nav.about">About</a>' +
        '<a href="' + PAGES.skills + '" data-i18n="nav.skills">Skills</a>' +
        '<a href="' + PAGES.projects + '" data-i18n="nav.projects">Projects</a>' +
        '<a href="' + PAGES.contact + '" data-i18n="nav.contact">Contact</a></div>' +
      '<div class="footer-meta"><span class="footer-meta-line" data-i18n="footer.created_by">This website was entirely created by its owner,<br>Youssef Chabbou, using AI-Augmented Development.</span>' +
        '<span class="footer-copy">© 2025 Youssef Chabbou</span></div>' +
    '</div></footer>';

  if (script) script.insertAdjacentHTML('beforebegin', slot === 'footer' ? FOOTER : HEADER);
})();
