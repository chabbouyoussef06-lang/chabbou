/* ===========================
   i18n.js — site-wide i18n
   Compatible avec :
   - Bouton desktop #langToggle + #lang-menu a[data-lang]
   - Drawer mobile .lang-btn[data-lang]
   - data-i18n / data-i18n-attr
   - Thème : updateThemeLabel() + i18nGet()
   - JSON: /lang/{xx}.json + /lang/pages/{page}.{xx}.json
   Persistance: sessionStorage (par onglet)
   =========================== */

// ---- CONFIG ----
const I18N_SUPPORTED = ["en", "fr", "de", "ar"];
const I18N_DEFAULT   = "en";
const I18N_DIR       = "lang";                 // dossier des JSON
const STORAGE_KEY    = "site.lang.session";    // persistance session

// ---- UTILS ----
const isRTL = (l) => ["ar", "fa", "he", "ur"].includes(l);

function pageKey() {
  // Ex: /, /index.html -> "index"; /about.html -> "about"
  const file = (location.pathname.split("/").pop() || "index.html");
  return file.split(".")[0] || "index";
}

function deepMerge(target, src){
  if (!src || typeof src !== "object") return target;
  const out = Array.isArray(target) ? target.slice() : {...target};
  for (const [k,v] of Object.entries(src)) {
    if (v && typeof v === "object" && !Array.isArray(v)) {
      out[k] = deepMerge(out[k] || {}, v);
    } else {
      out[k] = v;
    }
  }
  return out;
}

function get(obj, path){
  return path.split(".").reduce((o, k) => (o && o[k] != null) ? o[k] : undefined, obj);
}

async function fetchJson(url){
  const res = await fetch(url, { cache: "no-cache" });
  if (!res.ok) throw new Error(`${url} -> ${res.status}`);
  return res.json();
}

// ---- DICTS ----
async function loadDict(lang){
  // Base site-wide
  const baseUrl = `/${I18N_DIR}/${lang}.json`;
  // Page-specific (ex: /lang/pages/index.fr.json)
  const pageUrl = `/${I18N_DIR}/pages/${pageKey()}.${lang}.json`;

  let dict = {};
  try { dict = await fetchJson(baseUrl); } catch(_){}
  try { dict = deepMerge(dict, await fetchJson(pageUrl)); } catch(_){}

  return dict;
}

// ---- APPLY ----
function applyDirLang(lang){
  document.documentElement.setAttribute("lang", lang);
  document.documentElement.setAttribute("dir", isRTL(lang) ? "rtl" : "ltr");
  document.documentElement.dataset.lang = lang;
}

function updateActiveLangUI(lang){
  // Desktop toggle label
  const btn = document.getElementById("langToggle");
  if (btn){
    const map = {en:"EN", fr:"FR", de:"DE", ar:"AR"};
    btn.textContent = (map[lang] || "EN") + " ▾";
  }
  // Actifs sur éléments cliquables
  document.querySelectorAll("[data-lang]").forEach(el=>{
    el.toggleAttribute("data-active", el.getAttribute("data-lang") === lang);
  });
  // <select data-lang-select> si jamais tu en ajoutes
  const select = document.querySelector("[data-lang-select]");
  if (select && select.value !== lang) select.value = lang;
}

function translateDOM(dict){
  // Textes
  document.querySelectorAll("[data-i18n]").forEach(el=>{
    const key = el.getAttribute("data-i18n");
    const val = get(dict, key);
    if (typeof val === "string") el.textContent = val;
  });

  // Attributs (placeholder|title|aria-label|…)
  document.querySelectorAll("[data-i18n-attr]").forEach(el=>{
    const attrs = el.getAttribute("data-i18n-attr").split("|").map(s=>s.trim());
    for (const attr of attrs){
      const key = el.getAttribute(`data-i18n-${attr}`);
      if (!key) continue;
      const val = get(dict, key);
      if (typeof val === "string") el.setAttribute(attr, val);
    }
  });

  // <title> si fourni
  const title = get(dict, "meta.title");
  if (typeof title === "string") document.title = title;

  // Expose le dict pour i18nGet()/updateThemeLabel()
  window.__i18nDict = dict || {};
  if (typeof window.updateThemeLabel === "function") window.updateThemeLabel();
}

// ---- STATE MGMT ----
function pickInitialLang(){
  // 1) ?lang=xx a la priorité (utile pour liens partageables)
  try{
    const url = new URL(location.href);
    const q = url.searchParams.get("lang");
    if (q && I18N_SUPPORTED.includes(q)) return q;
  }catch(_){}

  // 2) sessionStorage (persiste sur toutes les pages de l’onglet)
  try{
    const saved = sessionStorage.getItem(STORAGE_KEY);
    if (saved && I18N_SUPPORTED.includes(saved)) return saved;
  }catch(_){}

  // 3) <html lang=".."> si cohérent
  const htmlLang = document.documentElement.getAttribute("lang");
  if (htmlLang && I18N_SUPPORTED.includes(htmlLang)) return htmlLang;

  // 4) fallback
  return I18N_DEFAULT;
}

async function setLang(lang, {remember=true, updateUrl=true} = {}){
  if (!I18N_SUPPORTED.includes(lang)) lang = I18N_DEFAULT;

  // Persistance de session (par onglet)
  if (remember){
    try{ sessionStorage.setItem(STORAGE_KEY, lang); }catch(_){}
  }

  // Optionnel : refléter dans l’URL sans recharger (pratique si tu partages un lien)
  if (updateUrl){
    try{
      const u = new URL(location.href);
      u.searchParams.set("lang", lang);
      history.replaceState({}, "", u.pathname + u.search + location.hash);
    }catch(_){}
  }

  applyDirLang(lang);

  // Charger et appliquer
  try{
    const dict = (lang === "en")
      ? {} // l'anglais = contenu déjà en dur
      : await loadDict(lang);
    translateDOM(dict);
  }catch(err){
    console.error("[i18n] load/apply failed:", err);
    // En cas d’échec, on retombe sur l’anglais (texte en dur)
    translateDOM({});
    applyDirLang("en");
    lang = "en";
    try{
      const u = new URL(location.href);
      u.searchParams.set("lang", "en");
      history.replaceState({}, "", u.pathname + u.search + location.hash);
    }catch(_){}
  }

  updateActiveLangUI(lang);
}

// Expose pour tes handlers déjà présents dans le HTML
window.setLangFromMenu = (l) => setLang(l, {remember:true, updateUrl:true});

// ---- INIT ----
(async function initI18N(){
  // Texte anglais en dur = "original" ; si tu veux garder un snapshot :
  document.querySelectorAll("[data-i18n]").forEach(el=>{
    if (!el.dataset.i18nOriginal) el.dataset.i18nOriginal = el.textContent;
  });

  const lang = pickInitialLang();
  await setLang(lang, {remember:true, updateUrl:false});

  // Wiring générique au cas où (sécurité)
  document.addEventListener("click", (e)=>{
    const el = e.target.closest("[data-lang]");
    if (!el) return;
    e.preventDefault();
    setLang(el.getAttribute("data-lang"));
  });

  const select = document.querySelector("[data-lang-select]");
  if (select){
    select.addEventListener("change", ()=> setLang(select.value));
  }
})();
