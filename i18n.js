// --- CONFIG ----------------------------------------------------
const I18N_DEFAULT = "en";               // langue par défaut
const I18N_SUPPORTED = ["en","fr","de","ar"]; // langues disponibles
const I18N_PATH = "/lang";               // dossier des JSON de traduction
const STORAGE_KEY = "site.lang.session"; // clé sessionStorage

// --- UTILS -----------------------------------------------------
function isRTL(lang){ return ["ar","fa","he","ur"].includes(lang); }

function getInitialLang(){
  // 1) URL ?lang=xx a priorité (utile si vous voulez forcer depuis un lien)
  const url = new URL(window.location.href);
  const urlLang = url.searchParams.get("lang");
  if (urlLang && I18N_SUPPORTED.includes(urlLang)) return urlLang;

  // 2) Mémoire de session (persiste tant que l’onglet reste ouvert)
  const saved = sessionStorage.getItem(STORAGE_KEY);
  if (saved && I18N_SUPPORTED.includes(saved)) return saved;

  // 3) Fallback : par défaut
  return I18N_DEFAULT;
}

function setLang(lang, {remember=true, applyNow=true} = {}){
  if (!I18N_SUPPORTED.includes(lang)) lang = I18N_DEFAULT;
  if (remember) sessionStorage.setItem(STORAGE_KEY, lang);
  document.documentElement.setAttribute("lang", lang);
  document.documentElement.setAttribute("dir", isRTL(lang) ? "rtl" : "ltr");
  document.documentElement.dataset.lang = lang; // pratique pour le CSS
  if (applyNow) applyTranslations(lang);
}

// Charge un JSON de traduction (mise en cache par le navigateur)
async function loadDict(lang){
  const url = `${I18N_PATH}/${lang}.json`;
  const res = await fetch(url, {cache:"no-cache"});
  if (!res.ok) throw new Error(`i18n: impossible de charger ${url}`);
  return res.json();
}

// Applique les traductions aux éléments portant [data-i18n]
// - innerText : data-i18n="section.key"
// - attributs  : data-i18n-attr="placeholder|aria-label|title"
async function applyTranslations(lang){
  try{
    const dict = await loadDict(lang);

    // Texte
    document.querySelectorAll("[data-i18n]").forEach(el=>{
      const key = el.getAttribute("data-i18n");
      const val = key.split(".").reduce((o,k)=> (o && o[k]!=null) ? o[k] : null, dict);
      if (typeof val === "string") el.textContent = val;
    });

    // Attributs
    document.querySelectorAll("[data-i18n-attr]").forEach(el=>{
      const attrs = el.getAttribute("data-i18n-attr").split("|").map(s=>s.trim());
      attrs.forEach(attr=>{
        const key = el.getAttribute(`data-i18n-${attr}`);
        if (!key) return;
        const val = key.split(".").reduce((o,k)=> (o && o[k]!=null) ? o[k] : null, dict);
        if (typeof val === "string") el.setAttribute(attr, val);
      });
    });

    // Optionnel : MAJ du titre de page si défini dans le JSON
    const pageTitle = dict?.meta?.title;
    if (typeof pageTitle === "string") document.title = pageTitle;

    // Met à jour l’état visuel des boutons/lang-switchers
    document.querySelectorAll("[data-lang-switch]").forEach(btn=>{
      btn.toggleAttribute("data-active", btn.getAttribute("data-lang-switch") === lang);
    });
  }catch(err){
    console.error(err);
  }
}

// Ecouteurs pour tout bouton/lien ayant data-lang-switch="xx"
function wireSwitchers(){
  document.addEventListener("click", (e)=>{
    const t = e.target.closest("[data-lang-switch]");
    if (!t) return;
    e.preventDefault();
    const lang = t.getAttribute("data-lang-switch");
    setLang(lang);
  });
}

// --- INIT ------------------------------------------------------
(function init(){
  wireSwitchers();
  const lang = getInitialLang();
  setLang(lang, {remember:true, applyNow:true});
})();
