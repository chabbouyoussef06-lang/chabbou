# Youssef Chabbou — Portfolio

Site statique (HTML / CSS / JS vanilla) publié avec GitHub Pages : https://chabbouyoussef06-lang.github.io/chabbou/

## Structure

```
index.html          → /
about/ skills/ projects/ contact/   (un dossier par page, chacun avec index.html → URLs courtes)
about.html, skills.html, … (redirections des anciennes adresses)
css/   site.css (jetons, en-tête, menus, pied de page) · home.css · subpage.css
js/    boot.js (thème/langue avant affichage) · layout.js (en-tête, menu mobile, pied de page)
       site.js (i18n, thème, menus) · config.js (réglages) · home.js · about.js · projects.js · contact.js
lang/  en.json · fr.json · de.json · ar.json  (mêmes clés dans les 4 fichiers)
assets/images (favicon*, signature.png = logo) · assets/video · assets/docs (PDF à ajouter)
favicon.ico (racine)
```

## Traductions
Le HTML contient le texte anglais par défaut ; `data-i18n="section.cle"` pour le texte,
`data-i18n-attr="attribut:section.cle"` pour les attributs. Toute nouvelle clé doit être ajoutée aux 4 fichiers `lang/*.json`.

## Réglages (`js/config.js`)
- `formEndpoint` : remplacer `XXXXXXXX` par l'identifiant Formspree. Sans identifiant valide, le formulaire ouvre le client mail (mailto).
- `docs` : noms des PDF (CV, attestations) à déposer dans `assets/docs/`. Un fichier absent affiche « Coming soon ».

## Ajouter une page
Copier une page existante, définir `body[data-page]`, ajouter `meta.<page>` dans les JSON et le lien dans `js/layout.js`.

## Adresses
- Accueil `/` · `/about/` · `/skills/` · `/projects/` · `/contact/` ; ancres : `/about/#cv`, `/projects/#future`…
- `?lang=fr|de|ar` force la langue (ajouté automatiquement quand on change de langue).
- Les chemins des scripts sont calculés depuis `js/boot.js` (`window.__root`) : une page peut vivre dans n'importe quel dossier.

## Cache
Après toute modification de `css/` ou `js/`, lancer `python3 tools/bust.py` : il ajoute `?v=<empreinte>` aux liens des pages pour que les navigateurs rechargent les nouveaux fichiers.
