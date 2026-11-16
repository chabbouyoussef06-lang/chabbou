# Youssef Chabbou — Portfolio

Site statique (HTML / CSS / JS vanilla) publié avec GitHub Pages : https://chabbouyoussef06-lang.github.io/chabbou/

## Structure

```
index.html  about.html  skills.html  projects.html  contact.html
css/   site.css (jetons, en-tête, menus, pied de page) · home.css · subpage.css
js/    boot.js (thème/langue avant affichage) · layout.js (en-tête, menu mobile, pied de page)
       site.js (i18n, thème, menus) · config.js (réglages) · home.js · about.js · projects.js · contact.js
lang/  en.json · fr.json · de.json · ar.json  (mêmes clés dans les 4 fichiers)
assets/images · assets/video · assets/docs (PDF à ajouter)
```

## Traductions
Le HTML contient le texte anglais par défaut ; `data-i18n="section.cle"` pour le texte,
`data-i18n-attr="attribut:section.cle"` pour les attributs. Toute nouvelle clé doit être ajoutée aux 4 fichiers `lang/*.json`.

## Réglages (`js/config.js`)
- `formEndpoint` : remplacer `XXXXXXXX` par l'identifiant Formspree. Sans identifiant valide, le formulaire ouvre le client mail (mailto).
- `docs` : noms des PDF (CV, attestations) à déposer dans `assets/docs/`. Un fichier absent affiche « Coming soon ».

## Ajouter une page
Copier une page existante, définir `body[data-page]`, ajouter `meta.<page>` dans les JSON et le lien dans `js/layout.js`.
