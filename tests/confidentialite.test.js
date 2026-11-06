'use strict';

// Garde-fous de confidentialité : le code du site ne doit contenir aucun
// moyen d'envoyer des données sur le réseau, ni de les enregistrer.
//
// Ce sont des « détecteurs » : ils repèrent les façons habituelles d'envoyer
// ou d'enregistrer des données. La politique de sécurité (CSP) de la page
// bloque les requêtes (fetch, XHR, sendBeacon…) et tout chargement depuis un
// autre site. Mais elle ne bloque PAS les chargements depuis le site lui-même
// (image, script, feuille de style : 'self'), ni la navigation (liens,
// fenêtres, redirections), ni le changement de l'adresse de la page
// (history) : pour ces canaux, seuls ces détecteurs et la relecture du code
// protègent. Aucun détecteur ne peut déjouer un code volontairement camouflé
// (window['fe' + 'tch']) : le code est court et public, et doit être relu.
//
// Une fausse alerte sur un commentaire se corrige en reformulant le
// commentaire, pas en affaiblissant la règle.
const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const racine = path.join(__dirname, '..');
const lire = (nom) => fs.readFileSync(path.join(racine, nom), 'utf8');

// Tous les fichiers servis par le site (pas les tests ni les fichiers d'exemple).
function fichiersDuSite(dossier = '') {
  const liste = [];
  for (const entree of fs.readdirSync(path.join(racine, dossier), { withFileTypes: true })) {
    const relatif = dossier ? dossier + '/' + entree.name : entree.name;
    if (entree.isDirectory()) {
      const exclu = ['.git', 'node_modules'].includes(entree.name) || (dossier === '' && ['tests', 'exemples'].includes(entree.name));
      if (!exclu) liste.push(...fichiersDuSite(relatif));
    } else if (/\.(html?|css|[mc]?js|svg|json|webmanifest|xml)$/i.test(entree.name)) {
      liste.push(relatif);
    }
  }
  return liste;
}
const FICHIERS = fichiersDuSite();
const parType = (motif) => FICHIERS.filter((f) => motif.test(f));
const JS = /\.[mc]?js$/i;
const HTML = /\.html?$/i;

const CSP = "default-src 'none'; script-src 'self'; style-src 'self'; img-src 'self' data:; " +
            "connect-src 'none'; form-action 'none'; base-uri 'none'";

const INTERDITS = [
  ['fetch', /\bfetch\b/],
  ['XMLHttpRequest', /XMLHttpRequest/],
  ['WebSocket', /WebSocket/],
  ['sendBeacon', /sendBeacon/],
  ['EventSource', /EventSource/],
  ['WebRTC', /RTCPeerConnection|RTCDataChannel/],
  ['localStorage', /localStorage/],
  ['sessionStorage', /sessionStorage/],
  ['IndexedDB', /indexedDB/i],
  ['Cache API', /\bcaches\b/],
  ['service worker', /serviceWorker/],
  ['cookies', /document\s*\.\s*cookie|cookieStore/],
  ['WebSQL', /openDatabase/],
  ['import dynamique', /\bimport\s*\(/],
  ['Worker', /\b(Shared)?Worker\b/],
  ['envoi de formulaire', /\.submit\s*\(|requestSubmit/],
  ['eval', /\beval\s*\(/],
  ['new Function', /new\s+Function\b/],
  ['innerHTML / outerHTML', /(inner|outer)HTML/],
  ['insertAdjacentHTML', /insertAdjacentHTML/],
  ['document.write', /document\.write/],
  ['HTML construit à partir de texte', /DOMParser|createContextualFragment|setHTML|parseHTML|srcdoc/],
  ['Worklet (charge un script)', /Worklet|addModule/],
  ['fichiers du navigateur (OPFS, File System Access)', /getDirectory|createWritable|show(Save|Open|Directory)FilePicker/]
];

// Canaux que la CSP ne contrôle pas : ils pourraient emporter des données
// dans une adresse, vers le site lui-même ou ailleurs.
const INTERDITS_JS = [
  ['navigation (location)', /\blocation\b/],
  ['ouverture de fenêtre (open)', /\bopen\s*\(/],
  ['navigation (API Navigation)', /\bnavigation\s*\.\s*(navigate|reload|traverseTo|back|forward)\b/],
  ['adresse de la page modifiée (history)', /\bhistory\s*\.\s*(pushState|replaceState)\b/],
  ['postMessage', /postMessage/],
  ['chargement depuis le site lui-même (src, srcset, poster)', /\.(src|srcset|poster)\s*\+?=[^=]|\[\s*["'`](src|srcset|poster|href|srcdoc)["'`]\s*\]/],
  ['attribut posé par setAttribute (seuls aria-*, data-*, role, tabindex, scope, lang, translate, hidden)',
    /setAttribute(NS)?\s*\((?!\s*["'`](aria-|data-|(role|tabindex|scope|lang|translate|hidden)["'`]))/],
  ['attribut posé autrement (setAttributeNode, Object.assign, Reflect)', /setAttributeNode|createAttribute|\bObject\s*\.\s*assign\b|\bReflect\s*\./],
  ['adresse dans un objet (src:, href:…)', /\b(src|srcset|href|poster)\s*:/],
  ['titre de la page (gardé dans l’historique du navigateur)', /document\s*\.\s*title/],
  ['nom de la fenêtre (lisible par le site ouvert ensuite)', /\b(window|self|top|parent|globalThis)\s*\.\s*name\b/],
  ['adresse du fichier généré autre que blob:', /etat\.url\s*=(?!=)(?!\s*(URL\.createObjectURL\(\w+\)|null);)/],
  ['redirection (meta refresh)', /http-?equiv/i],
  ['CSS construit en JavaScript (url, image-set, @import)', /(?<!new\s+)\burl\s*\(|image-set|@import/i],
  ['console (garderait les données en mémoire)', /\bconsole\s*\./],
  ['adresse web', /https?:\/\//i],
  ['adresse sans protocole (//site)', /["'`]\/\/[a-z0-9]/i]
];

describe('Confidentialité : aucun envoi, aucun enregistrement', () => {
  it('les fichiers du site sont bien tous examinés', () => {
    for (const attendu of ['index.html', 'style.css', 'js/app.js', 'js/conversion.js', 'favicon.svg']) {
      assert.ok(FICHIERS.includes(attendu), attendu + ' non examiné');
    }
  });

  for (const fichier of FICHIERS) {
    it(fichier + ' ne contient aucune fonction réseau ou de stockage', () => {
      const contenu = lire(fichier);
      for (const [nom, motif] of INTERDITS) {
        assert.ok(!motif.test(contenu), nom + ' trouvé dans ' + fichier);
      }
    });
  }

  for (const fichier of parType(JS)) {
    it(fichier + ' n’ouvre aucune adresse (navigation, fenêtre, chargement, message)', () => {
      const contenu = lire(fichier);
      for (const [nom, motif] of INTERDITS_JS) {
        assert.ok(!motif.test(contenu), nom + ' trouvé dans ' + fichier);
      }
      // Seul lien créé par le code : celui du fichier généré (adresse blob: locale).
      const liens = contenu.match(/\.href\s*\+?=[^=]/g) || [];
      assert.ok(liens.length <= 1 && /lien\.href = etat\.url;/.test(contenu) === (liens.length === 1),
        'lien inattendu dans ' + fichier);
    });
  }

  it('la page déclare exactement la politique de sécurité (CSP) demandée', () => {
    const html = lire('index.html');
    const balise = html.match(/<meta\s+http-equiv="Content-Security-Policy"\s+content="([^"]*)"\s*>/i);
    assert.ok(balise, 'balise meta CSP absente');
    assert.equal(balise[1], CSP);
    // La CSP doit précéder tout chargement de ressource.
    assert.ok(html.indexOf('Content-Security-Policy') < html.search(/<(script|link)\b/i));
  });

  it('l’hébergeur (vercel.json) envoie la même politique de sécurité, sans l’assouplir', () => {
    const config = JSON.parse(lire('vercel.json'));
    const regle = (config.headers || []).find((r) => r.source === '/(.*)');
    assert.ok(regle, 'aucune règle d’en-têtes pour toutes les adresses dans vercel.json');
    const csp = regle.headers.find((h) => h.key.toLowerCase() === 'content-security-policy');
    assert.ok(csp, 'en-tête Content-Security-Policy absent de vercel.json');
    assert.equal(csp.value, CSP + "; frame-ancestors 'none'");
  });

  it('aucun script ni style en ligne dans la page', () => {
    for (const fichier of parType(HTML)) {
      const html = lire(fichier);
      assert.ok(!/<script(?![^>]*\ssrc\s*=)[^>]*>/i.test(html), 'script en ligne dans ' + fichier);
      assert.ok(!/<style\b/i.test(html), 'balise <style> dans ' + fichier);
      assert.ok(!/\sstyle\s*=/i.test(html), 'attribut style= dans ' + fichier);
      assert.ok(!/\son[a-z]+\s*=/i.test(html), 'attribut on…= (onclick, onload…) dans ' + fichier);
      assert.ok(!/javascript:/i.test(html), 'lien javascript: dans ' + fichier);
    }
  });

  it('toutes les ressources viennent du site lui-même (pas de CDN, de polices ou de statistiques externes)', () => {
    for (const fichier of parType(HTML)) {
      const html = lire(fichier);
      // Adresses externes, avec ou sans guillemets autour de l'attribut.
      assert.ok(!/\s(src|href)\s*=\s*["']?\s*(https?:)?\/\//i.test(html), 'ressource externe dans ' + fichier);
      assert.ok(!/<(iframe|object|embed|form|a)\b/i.test(html), 'iframe, object, embed, form ou lien dans ' + fichier);
      assert.ok(!/\srel\s*=\s*["']?[^"'>]*\b(dns-prefetch|preconnect|prefetch|prerender|modulepreload)\b/i.test(html),
        'préchargement dans ' + fichier);
      assert.ok(!/http-equiv\s*=\s*["']?\s*refresh/i.test(html), 'redirection automatique dans ' + fichier);
      assert.ok(!/\s(ping|srcset|poster|background)\s*=/i.test(html), 'attribut qui charge une ressource dans ' + fichier);
      assert.ok(!/speculationrules/i.test(html), 'règles de préchargement dans ' + fichier);
    }
    assert.ok(lire('index.html').match(/<(script|link)\b/gi).length >= 3);
    // Tout script chargé par une page doit faire partie des fichiers examinés
    // (pas un script rangé dans exemples/ ou tests/).
    for (const fichier of parType(HTML)) {
      for (const m of lire(fichier).matchAll(/<script\b[^>]*\ssrc\s*=\s*["']?([^"'\s>]+)/gi)) {
        const cible = path.posix.normalize(path.posix.join(path.posix.dirname(fichier), m[1].replace(/[?#].*$/, '')));
        assert.ok(FICHIERS.includes(cible), 'script non examiné par ces tests : ' + m[1]);
      }
    }
    for (const fichier of parType(/\.css$/i)) {
      // Les images « data: », écrites dans la feuille de style, et les renvois
      // vers la page elle-même (url(#…)) ne chargent rien.
      assert.ok(!/@import|image-set\s*\(|url\s*\((?!\s*["']?\s*(data:|#))/i.test(lire(fichier)), 'ressource chargée depuis ' + fichier);
    }
    for (const fichier of parType(/\.svg$/i)) {
      const svg = lire(fichier);
      assert.ok(!/<script|<foreignObject|\shref\s*=|url\s*\(|\son[a-z]+\s*=/i.test(svg), 'contenu actif dans ' + fichier);
    }
  });

  it('les données du fichier ne sont pas proposées à la traduction automatique', () => {
    const html = lire('index.html');
    for (const zone of [/<p id="resume-fichier"[^>]*translate="no"/, /<tbody translate="no">/,
      /<ul id="avertissements"[^>]*translate="no"/, /id="nom-telecharge" translate="no"/, /id="nom-numerote" translate="no"/,
      /id="annonce"[^>]*translate="no"/]) {
      assert.ok(zone.test(html), zone.source);
    }
  });

  it('affiche le message de confidentialité', () => {
    assert.ok(lire('index.html').includes("Votre fichier reste sur votre ordinateur. Rien n'est envoyé ni conservé."));
  });

  it('libère l’adresse du fichier généré et vide la mémoire après le téléchargement', () => {
    const app = lire('js/app.js');
    assert.ok(/URL\.revokeObjectURL\(/.test(app));
    // Le moteur JavaScript garde la dernière chaîne examinée par une expression régulière.
    assert.ok(/function oublierDerniereRecherche\(\) \{\s*\/\^\/\.test\(''\);/.test(app));
    const effacer = app.slice(app.indexOf('function effacerDonnees'), app.indexOf('function reinitialiser'));
    assert.ok(effacer.includes('oublierDerniereRecherche();'), 'effacerDonnees doit vider la dernière recherche');
  });
});
