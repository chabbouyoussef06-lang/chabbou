/* config.js — the only file you need to edit for things that live outside the code.
   ------------------------------------------------------------------------- */
window.SITE_CONFIG = {
  /* Contact form
     1. Create a free form on https://formspree.io  → you get an address like https://formspree.io/f/abcdwxyz
     2. Paste it below in place of the placeholder.
     Until then the form falls back to opening the visitor's e-mail app (mailto). */
  formEndpoint: 'https://formspree.io/f/XXXXXXXX',
  contactEmail: 'Youssefchabbou@outlook.com',

  /* Downloadable documents (About page). Put the PDF files in assets/docs/ using these names.
     A button whose file is missing shows "Coming soon" instead of a broken link. */
  docsFolder: 'assets/docs/',
  docs: {
    cvBasic:    { en: 'curriculum_vitae.pdf',          fr: 'cv.pdf',               de: 'lebenslauf.pdf',              ar: 'السيرة_الذاتية.pdf' },
    cvExtended: { en: 'extended_curriculum_vitae.pdf', fr: 'cv_detaille.pdf',      de: 'erweiterter_lebenslauf.pdf',  ar: 'السيرة_الذاتية_الموسعة.pdf' },
    certAdmin:  { en: 'certification_0722.pdf',        fr: 'certification_0722.pdf', de: 'certification_0722.pdf',    ar: 'certification_0722.pdf' },
    certAmr:    { en: 'certificate_amr.pdf',           fr: 'attestation_amr.pdf',  de: 'zertifikat_amr.pdf',          ar: 'شهادة_amr.pdf' }
  }
};
