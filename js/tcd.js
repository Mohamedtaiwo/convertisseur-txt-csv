/*
 * tcd.js — Tableau croisé dynamique
 *
 * À partir du tableau lu dans le fichier texte : préparation des données
 * (titres, nombres, dates), proposition automatique d'un tableau croisé,
 * calcul de son aperçu, et création d'un fichier Excel (.xlsx) contenant les
 * données et un vrai tableau croisé dynamique, qu'Excel calcule à l'ouverture.
 *
 * Fonctions pures, comme conversion.js : aucun accès au réseau ni à la page.
 * Le fichier .xlsx (un dossier compressé de fichiers XML) est fabriqué ici
 * même, sans aucune bibliothèque. Les adresses « schemas.openxmlformats.org »
 * qu'il contient ne sont que des noms de format exigés par Excel : rien n'est
 * jamais chargé depuis ces adresses.
 *
 * Utilisable dans le navigateur (objet global « TableauCroise ») comme dans
 * Node.js (require).
 */
(function (racine, fabrique) {
  'use strict';
  var api = fabrique();
  if (typeof module === 'object' && module !== null && module.exports) {
    module.exports = api;
  } else {
    racine.TableauCroise = api;
  }
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  /* ------------------------------------------------------------------ */
  /* Réglages                                                           */
  /* ------------------------------------------------------------------ */

  var MAX_ENREGISTREMENTS = 1048575; // lignes d'une feuille Excel, moins celle des titres
  var MAX_LONGUEUR_CELLULE = 32767;
  var MAX_LONGUEUR_ELEMENT = 255;   // longueur maximale d'un élément de tableau croisé
  var MAX_LIGNES_CHAMPS = 2;
  var MAX_VALEURS = 3;

  var CALCULS = ['somme', 'nombre', 'moyenne', 'min', 'max'];
  var NOMS_CALCULS = { somme: 'Somme', nombre: 'Nombre', moyenne: 'Moyenne', min: 'Min', max: 'Max' };
  var CALCULS_EXCEL = { somme: 'sum', nombre: 'count', moyenne: 'average', min: 'min', max: 'max' };
  var AFFICHAGES = { normal: '', total: 'percentOfTotal', ligne: 'percentOfRow', colonne: 'percentOfCol' };
  var TRIS = { libelle: true, decroissant: true, croissant: true };

  var MOIS = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août',
    'septembre', 'octobre', 'novembre', 'décembre'];

  // Une ligne qui commence par « Total », « Sous-total »… n'est pas une donnée :
  // la compter fausserait les sommes.
  var LIGNE_TOTAL = /^(?:sous[- ]?)?(?:total|totaux)(?![a-zà-ÿ])/i;

  // Noms de colonnes qui désignent des montants (à additionner) ou des codes
  // (à ne pas additionner), pour la proposition automatique.
  var MOTS_MONTANTS = /montant|total|prix|somme|quantit|qt[eé]|d[eé]bit|cr[eé]dit|chiffre|valeur|co[uû]t|solde|vente|recette|d[eé]pense|achat|paiement|factur|\bca\b|\bht\b|ttc|poids|heure|dur[eé]e|stock|effectif|budget|salaire|frais|remise/i;
  var MOTS_MOYENNES = /unitaire|moyen|taux|pourcent|\bprix\b|note\b|%/i;
  var MOTS_CODES = /(?:^|[^a-zà-ÿ])(?:id|n°|no|num(?:[eé]ro)?|code|cp|postal|t[eé]l(?:[eé]phone)?|portable|siret|siren|iban|bic|r[eé]f(?:[eé]rence)?|ann[eé]e|mois|jour|matricule|identifiant)(?![a-zà-ÿ])/i;
  var MOTS_CATEGORIES = /cat[eé]gorie|rayon|type|famille|groupe|r[eé]gion|ville|pays|service|d[eé]partement|client|fournisseur|produit|article|poste|nature|compte|vendeur|agence|magasin|secteur|classe|marque/i;
  var MOTS_ATTRIBUTS = /mode|statut|[eé]tat|genre|sexe|canal|origine|moyen/i;
  var MOTS_TEXTES_LIBRES = /libell|description|commentaire|remarque|d[eé]tail|adresse|e-?mail|courriel|observation|note/i;

  /* ------------------------------------------------------------------ */
  /* Petits utilitaires                                                 */
  /* ------------------------------------------------------------------ */

  // 0 → A, 25 → Z, 26 → AA… comme les colonnes d'Excel.
  function lettre(index) {
    var lettres = '';
    var n = index + 1;
    while (n > 0) {
      var reste = (n - 1) % 26;
      lettres = String.fromCharCode(65 + reste) + lettres;
      n = Math.floor((n - 1) / 26);
    }
    return lettres;
  }

  function deuxChiffres(n) {
    return (n < 10 ? '0' : '') + n;
  }

  function cellule(ligne, j) {
    var v = j < ligne.length ? ligne[j] : '';
    return v === null || v === undefined ? '' : String(v).trim();
  }

  function nbRemplies(ligne) {
    var n = 0;
    for (var j = 0; j < ligne.length; j++) {
      if (cellule(ligne, j) !== '') n++;
    }
    return n;
  }

  function premiereValeur(ligne) {
    for (var j = 0; j < ligne.length; j++) {
      var v = cellule(ligne, j);
      if (v !== '') return v;
    }
    return null;
  }

  /* ------------------------------------------------------------------ */
  /* Nombres et dates                                                   */
  /* ------------------------------------------------------------------ */

  var DEVISE_DEBUT = /^(?:€|\$|£|EUR|CHF)\s*/i;
  var DEVISE_FIN = /\s*(?:€|\$|£|EUR|euros?|CHF)$/i;
  var GROUPES_ESPACES = /^\d{1,3}(?:[   ']\d{3})+(?:[.,]\d+)?$/;

  // Lit un nombre écrit à la française ou à l'anglaise : -1 250,50 € /
  // 1.250,50 / 12.5 / (12,50) / 12,50- / 5 %. Renvoie null si ce n'en est pas un.
  function lireNombre(texte) {
    var t = String(texte).trim();
    if (t === '' || t.length > 40) return null;
    var signe = 1;
    var parentheses = /^\((.+)\)$/.exec(t);
    if (parentheses) {
      signe = -1;
      t = parentheses[1].trim();
    }
    var pourcent = false;
    if (t.charAt(t.length - 1) === '%') {
      pourcent = true;
      t = t.slice(0, -1).trim();
    }
    t = t.replace(DEVISE_FIN, '');
    var premier = t.charAt(0);
    if (premier === '-' || premier === '+' || premier === '−') {
      if (premier !== '+') signe = -signe;
      t = t.slice(1).trim();
    } else if (t.length > 1 && t.charAt(t.length - 1) === '-') {
      signe = -signe;
      t = t.slice(0, -1).trim();
    }
    t = t.replace(DEVISE_DEBUT, '');
    if (t === '') return null;
    if (GROUPES_ESPACES.test(t)) t = t.replace(/[   ']/g, '');
    var normal;
    if (/^\d+$/.test(t)) normal = t;
    else if (/^\d*[.,]\d+$/.test(t)) normal = t.replace(',', '.');
    else if (/^\d{1,3}(?:\.\d{3})+,\d+$/.test(t)) normal = t.replace(/\./g, '').replace(',', '.');
    else if (/^\d{1,3}(?:,\d{3})+\.\d+$/.test(t)) normal = t.replace(/,/g, '');
    else if (/^\d{1,3}(?:\.\d{3}){2,}$/.test(t)) normal = t.replace(/\./g, '');
    else if (/^\d{1,3}(?:,\d{3}){2,}$/.test(t)) normal = t.replace(/,/g, '');
    else return null;
    var valeur = signe * parseFloat(normal.charAt(0) === '.' ? '0' + normal : normal);
    if (!isFinite(valeur)) return null;
    if (valeur === 0) valeur = 0; // pas de « -0 »
    var entier = normal.indexOf('.') === -1;
    return {
      valeur: pourcent ? valeur / 100 : valeur,
      pourcent: pourcent,
      entier: entier && !pourcent,
      zeroInitial: entier && /^0\d/.test(normal) // code postal, numéro… : ce n'est pas un nombre
    };
  }

  var ORIGINE_EXCEL = Date.UTC(1899, 11, 30);
  var JOUR = 86400000;

  // Numéro de série d'Excel (jours depuis le 30/12/1899), avec l'heure en fraction.
  function serieExcel(a, m, j, h, mi, s) {
    h = h ? Number(h) : 0;
    mi = mi ? Number(mi) : 0;
    s = s ? Number(s) : 0;
    if (m < 1 || m > 12 || j < 1 || j > 31 || a < 1900 || a > 9999 || h > 23 || mi > 59 || s > 59) return null;
    var temps = Date.UTC(a, m - 1, j);
    var date = new Date(temps);
    if (date.getUTCMonth() !== m - 1 || date.getUTCDate() !== j) return null; // 31/02…
    var jours = (temps - ORIGINE_EXCEL) / JOUR;
    if (jours < 61) return null; // avant mars 1900, le calendrier d'Excel est décalé
    return jours + (h * 3600 + mi * 60 + s) / 86400;
  }

  var DATE_JMA = /^(\d{1,2})([\/.-])(\d{1,2})\2(\d{4}|\d{2})(?:[ T]+(\d{1,2})[:h](\d{2})(?::(\d{2}))?)?$/;
  var DATE_AMJ = /^(\d{4})([\/.-])(\d{1,2})\2(\d{1,2})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;

  // Lit une date : 05/01/2026, 5-1-26, 05.01.2026, 2026-01-05 (avec ou sans
  // heure). Renvoie son numéro de série Excel, ou null.
  function lireDate(texte) {
    var t = String(texte).trim();
    var m = DATE_JMA.exec(t);
    if (m) {
      var a = Number(m[4]);
      if (m[4].length === 2) a += a < 70 ? 2000 : 1900;
      return serieExcel(a, Number(m[3]), Number(m[1]), m[5], m[6], m[7]);
    }
    m = DATE_AMJ.exec(t);
    if (m) return serieExcel(Number(m[1]), Number(m[3]), Number(m[4]), m[5], m[6], m[7]);
    return null;
  }

  function dateDepuisSerie(serie) {
    var d = new Date(ORIGINE_EXCEL + Math.floor(serie + 1e-9) * JOUR);
    return { a: d.getUTCFullYear(), m: d.getUTCMonth() + 1, j: d.getUTCDate() };
  }

  function texteDate(serie, avecHeure) {
    var d = dateDepuisSerie(serie);
    var texte = deuxChiffres(d.j) + '/' + deuxChiffres(d.m) + '/' + d.a;
    if (avecHeure) {
      var minutes = Math.round((serie - Math.floor(serie + 1e-9)) * 1440);
      texte += ' ' + deuxChiffres(Math.floor(minutes / 60)) + ':' + deuxChiffres(minutes % 60);
    }
    return texte;
  }

  function isoDepuisSerie(serie) {
    return new Date(ORIGINE_EXCEL + Math.round(serie * JOUR / 1000) * 1000).toISOString().slice(0, 19);
  }

  /* ------------------------------------------------------------------ */
  /* Préparation des données                                            */
  /* ------------------------------------------------------------------ */

  // Lignes de titre au-dessus du tableau (« Relevé de compte — janvier 2026 ») :
  // elles ont nettement moins de cases remplies que le tableau lui-même.
  function debutDuTableau(lignes) {
    var limite = Math.min(lignes.length, 50);
    var remplies = [];
    var max = 0;
    for (var i = 0; i < limite; i++) {
      remplies.push(nbRemplies(lignes[i]));
      if (remplies[i] > max) max = remplies[i];
    }
    if (max < 2) return 0;
    var seuil = Math.max(2, Math.ceil(max / 2));
    for (i = 0; i < Math.min(limite, 10); i++) {
      if (remplies[i] >= seuil) return i;
    }
    return 0;
  }

  // La première ligne du tableau donne-t-elle les noms des colonnes ? Oui si
  // elle ne contient ni nombre ni date, et qu'elle est assez remplie.
  function aUneLigneDeTitres(lignes, debut, nbColonnes) {
    if (lignes.length - debut < 2) return false;
    var ligne = lignes[debut];
    var remplies = 0;
    for (var j = 0; j < ligne.length; j++) {
      var v = cellule(ligne, j);
      if (v === '') continue;
      remplies++;
      if (lireNombre(v) || lireDate(v) !== null) return false;
    }
    return remplies >= Math.max(1, Math.ceil(nbColonnes / 2));
  }

  function nettoyerNom(texte) {
    return String(texte).replace(/[\u0000-\u001F\u007F]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 200).trim();
  }

  function nomUnique(nom, utilises) {
    var candidat = nom;
    var k = 2;
    while (utilises[candidat.toLowerCase()]) {
      candidat = nom + ' (' + k + ')';
      k++;
    }
    utilises[candidat.toLowerCase()] = true;
    return candidat;
  }

  // Type d'une colonne (« nombre », « date » ou « texte ») et valeurs converties.
  function typer(lignes, j) {
    var nonVides = 0;
    var nombres = 0;
    var dates = 0;
    var zeros = 0;
    var pourcents = 0;
    var decimaux = 0;
    var heures = 0;
    var i;
    var v;
    for (i = 0; i < lignes.length; i++) {
      v = cellule(lignes[i], j);
      if (v === '') continue;
      nonVides++;
      var serie = lireDate(v);
      if (serie !== null) {
        dates++;
        if (serie !== Math.floor(serie)) heures++;
        continue;
      }
      var nombre = lireNombre(v);
      if (nombre) {
        nombres++;
        if (nombre.zeroInitial) zeros++;
        if (nombre.pourcent) pourcents++;
        if (!nombre.entier) decimaux++;
      }
    }
    var type = 'texte';
    var format = '';
    if (nonVides > 0 && dates >= 0.9 * nonVides) {
      type = 'date';
      format = heures ? 'dateheure' : 'date';
    } else if (nonVides > 0 && nombres >= 0.9 * nonVides && zeros === 0) {
      type = 'nombre';
      format = pourcents * 2 > nombres ? 'pourcent' : (decimaux ? 'decimal' : 'entier');
    }
    var valeurs = new Array(lignes.length);
    for (i = 0; i < lignes.length; i++) {
      v = cellule(lignes[i], j);
      if (v === '') {
        valeurs[i] = null;
      } else if (type === 'date') {
        var s = lireDate(v);
        valeurs[i] = s === null ? v : s;
      } else if (type === 'nombre') {
        var n = lireNombre(v);
        valeurs[i] = n ? n.valeur : v;
      } else {
        valeurs[i] = v;
      }
    }
    return { type: type, format: format, valeurs: valeurs };
  }

  // Colonnes « année », « trimestre » et « mois » calculées à partir d'une date,
  // pour regrouper les dates dans le tableau croisé.
  function partiesDeDates(series) {
    var annee = new Array(series.length);
    var trimestre = new Array(series.length);
    var mois = new Array(series.length);
    for (var i = 0; i < series.length; i++) {
      var v = series[i];
      if (typeof v !== 'number') {
        annee[i] = trimestre[i] = mois[i] = null;
        continue;
      }
      var d = dateDepuisSerie(v);
      annee[i] = String(d.a);
      trimestre[i] = d.a + '-T' + (Math.floor((d.m - 1) / 3) + 1);
      mois[i] = d.a + '-' + deuxChiffres(d.m) + ' (' + MOIS[d.m - 1] + ')';
    }
    return { annee: annee, trimestre: trimestre, mois: mois };
  }

  // Prépare le tableau lu dans le fichier (lignes de cellules texte).
  // options.entete : true / false pour imposer la présence d'une ligne de titres.
  function preparer(lignesBrutes, options) {
    options = options || {};
    var nbColonnes = 0;
    var i;
    var j;
    for (i = 0; i < lignesBrutes.length; i++) {
      if (lignesBrutes[i].length > nbColonnes) nbColonnes = lignesBrutes[i].length;
    }
    var debut = debutDuTableau(lignesBrutes);
    var entete = typeof options.entete === 'boolean' ? options.entete
      : aUneLigneDeTitres(lignesBrutes, debut, nbColonnes);
    var premiere = debut + (entete ? 1 : 0);

    var retenues = [];
    var lignesTotal = 0;
    var tronque = false;
    for (i = premiere; i < lignesBrutes.length; i++) {
      var valeur = premiereValeur(lignesBrutes[i]);
      if (valeur === null) continue;
      if (LIGNE_TOTAL.test(valeur)) {
        lignesTotal++;
        continue;
      }
      if (retenues.length === MAX_ENREGISTREMENTS) {
        tronque = true;
        break;
      }
      retenues.push(lignesBrutes[i]);
    }

    var donnees = {
      noms: [],
      types: [],
      formats: [],
      colonnes: [],
      origines: [],      // colonne de date d'où vient une colonne calculée (année, mois…)
      nbLignes: retenues.length,
      entete: entete,
      lignesTitre: debut,
      lignesTotal: lignesTotal,
      tronque: tronque
    };
    var utilises = {};
    for (j = 0; j < nbColonnes; j++) {
      var titre = entete ? nettoyerNom(cellule(lignesBrutes[debut], j)) : '';
      donnees.noms.push(nomUnique(titre || 'Colonne ' + lettre(j), utilises));
      var typage = typer(retenues, j);
      donnees.types.push(typage.type);
      donnees.formats.push(typage.format);
      donnees.colonnes.push(typage.valeurs);
      donnees.origines.push(null);
    }
    for (j = 0; j < nbColonnes; j++) {
      if (donnees.types[j] !== 'date') continue;
      var parties = partiesDeDates(donnees.colonnes[j]);
      var noms = { annee: ' (année)', trimestre: ' (trimestre)', mois: ' (mois)' };
      for (var partie in noms) {
        donnees.noms.push(nomUnique(donnees.noms[j] + noms[partie], utilises));
        donnees.types.push('texte');
        donnees.formats.push('');
        donnees.colonnes.push(parties[partie]);
        donnees.origines.push({ colonne: j, partie: partie });
      }
    }
    return donnees;
  }

  /* ------------------------------------------------------------------ */
  /* Clés, libellés et tri des éléments                                 */
  /* ------------------------------------------------------------------ */

  // Clé d'un élément : comme Excel, « Paris » et « PARIS » sont le même élément.
  function cle(v) {
    if (v === null || v === undefined) return '\u0000';
    if (typeof v === 'number') return '#' + v;
    return '"' + v.toLowerCase();
  }

  function rang(v) {
    if (typeof v === 'number') return 0;
    if (v === null || v === undefined) return 2;
    return 1;
  }

  // Ordre d'Excel : nombres et dates, puis textes (sans tenir compte des
  // majuscules), puis les cases vides.
  function comparerValeurs(a, b) {
    var ra = rang(a);
    var rb = rang(b);
    if (ra !== rb) return ra - rb;
    if (ra === 0) return a - b;
    if (ra === 2) return 0;
    var c = a.localeCompare(b, 'fr', { sensitivity: 'base' });
    if (c) return c;
    return a < b ? -1 : (a > b ? 1 : 0);
  }

  var FORMAT_NOMBRE = { minimumFractionDigits: 0, maximumFractionDigits: 10 };

  function formaterNombre(n, options) {
    try {
      return n.toLocaleString('fr-FR', options || FORMAT_NOMBRE);
    } catch (e) {
      return String(n).replace('.', ',');
    }
  }

  // Libellé d'un élément, tel qu'Excel l'affiche.
  function libelle(donnees, j, v) {
    if (v === null || v === undefined) return '(vide)';
    if (typeof v === 'number') {
      if (donnees.types[j] === 'date') return texteDate(v, donnees.formats[j] === 'dateheure');
      return formaterNombre(v);
    }
    return v;
  }

  // Affichage d'une valeur calculée (aperçu dans la page).
  function formaterValeur(v, format) {
    if (v === null || v === undefined) return '';
    if (format === 'pourcent') {
      return formaterNombre(v * 100, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' %';
    }
    if (format === 'entier') return formaterNombre(v, { maximumFractionDigits: 0 });
    if (format === 'date') return texteDate(v, false);
    return formaterNombre(v, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  /* ------------------------------------------------------------------ */
  /* Réglages du tableau croisé                                         */
  /* ------------------------------------------------------------------ */

  function champValide(donnees, j) {
    return typeof j === 'number' && j >= 0 && j < donnees.noms.length && Math.floor(j) === j;
  }

  // Réglages complets et cohérents, à partir de réglages partiels.
  function normaliserReglages(donnees, reglages) {
    reglages = reglages || {};
    var r = { lignes: [], colonnes: [], valeurs: [], affichage: 'normal', tri: 'libelle', filtre: { champ: -1, exclus: {} } };
    var pris = {};
    var i;
    var liste = reglages.lignes || [];
    for (i = 0; i < liste.length && r.lignes.length < MAX_LIGNES_CHAMPS; i++) {
      if (champValide(donnees, liste[i]) && !pris[liste[i]]) {
        r.lignes.push(liste[i]);
        pris[liste[i]] = true;
      }
    }
    if (r.lignes.length === 0) {
      for (i = 0; i < donnees.noms.length && r.lignes.length === 0; i++) r.lignes.push(i);
      pris[0] = true;
    }
    liste = reglages.colonnes || [];
    for (i = 0; i < liste.length && r.colonnes.length === 0; i++) {
      if (champValide(donnees, liste[i]) && !pris[liste[i]]) {
        r.colonnes.push(liste[i]);
        pris[liste[i]] = true;
      }
    }
    liste = reglages.valeurs || [];
    for (i = 0; i < liste.length && r.valeurs.length < MAX_VALEURS; i++) {
      var v = liste[i];
      if (v && champValide(donnees, v.champ)) {
        r.valeurs.push({ champ: v.champ, calcul: CALCULS.indexOf(v.calcul) >= 0 ? v.calcul : 'somme' });
      }
    }
    if (r.valeurs.length === 0) r.valeurs.push({ champ: r.lignes[0], calcul: 'nombre' });
    if (AFFICHAGES.hasOwnProperty(reglages.affichage)) r.affichage = reglages.affichage;
    if (TRIS.hasOwnProperty(reglages.tri)) r.tri = reglages.tri;
    var filtre = reglages.filtre;
    if (filtre && champValide(donnees, filtre.champ) && !pris[filtre.champ]) {
      r.filtre.champ = filtre.champ;
      var exclus = filtre.exclus || {};
      for (var k in exclus) {
        if (Object.prototype.hasOwnProperty.call(exclus, k) && exclus[k]) r.filtre.exclus[k] = true;
      }
    }
    return r;
  }

  // Noms des valeurs (« Somme de Montant »), différents entre eux et des
  // noms de colonnes, comme l'exige Excel.
  function nomsDesValeurs(donnees, reglages) {
    var utilises = {};
    for (var j = 0; j < donnees.noms.length; j++) utilises[donnees.noms[j].toLowerCase()] = true;
    var noms = [];
    for (var x = 0; x < reglages.valeurs.length; x++) {
      var v = reglages.valeurs[x];
      noms.push(nomUnique(NOMS_CALCULS[v.calcul] + ' de ' + donnees.noms[v.champ], utilises).slice(0, 255));
    }
    return noms;
  }

  // Format d'affichage d'une valeur calculée.
  function formatDeValeur(donnees, reglages, v) {
    if (reglages.affichage !== 'normal') return 'pourcent';
    if (v.calcul === 'nombre') return 'entier';
    if (donnees.types[v.champ] === 'date' && (v.calcul === 'min' || v.calcul === 'max')) return 'date';
    if (donnees.formats[v.champ] === 'pourcent') return 'pourcent';
    return 'decimal';
  }

  /* ------------------------------------------------------------------ */
  /* Proposition automatique                                            */
  /* ------------------------------------------------------------------ */

  function compterDistincts(colonne, limite) {
    var vus = Object.create(null);
    var n = 0;
    for (var i = 0; i < colonne.length; i++) {
      var k = cle(colonne[i]);
      if (!vus[k]) {
        vus[k] = true;
        n++;
        if (n > limite) return n;
      }
    }
    return n;
  }

  function remplissage(colonne) {
    var n = 0;
    for (var i = 0; i < colonne.length; i++) {
      if (colonne[i] !== null) n++;
    }
    return n;
  }

  function colonneDIdentifiants(colonne) {
    var vus = Object.create(null);
    var n = 0;
    for (var i = 0; i < colonne.length; i++) {
      var v = colonne[i];
      if (v === null) continue;
      if (typeof v !== 'number' || v !== Math.floor(v) || vus[v]) return false;
      vus[v] = true;
      n++;
    }
    return n > 5;
  }

  // Propose un tableau croisé pertinent : par exemple la somme des montants
  // par catégorie et par mois.
  function proposer(donnees) {
    var n = donnees.nbLignes;
    var nbChamps = donnees.noms.length;
    var j;
    var distincts = [];
    for (j = 0; j < nbChamps; j++) distincts.push(compterDistincts(donnees.colonnes[j], 1000));

    // Valeurs : les colonnes de nombres qui ressemblent à des montants.
    var candidats = [];
    for (j = 0; j < nbChamps; j++) {
      if (donnees.types[j] !== 'nombre' || MOTS_CODES.test(donnees.noms[j])) continue;
      if (colonneDIdentifiants(donnees.colonnes[j])) continue;
      var score = remplissage(donnees.colonnes[j]) / Math.max(1, n);
      if (MOTS_MONTANTS.test(donnees.noms[j])) score += 10;
      if (donnees.formats[j] === 'decimal') score += 3;
      candidats.push({ champ: j, score: score, montant: MOTS_MONTANTS.test(donnees.noms[j]) });
    }
    candidats.sort(function (a, b) { return b.score - a.score || a.champ - b.champ; });
    var valeurs = [];
    for (var c = 0; c < candidats.length && valeurs.length < MAX_VALEURS; c++) {
      if (valeurs.length > 0 && !candidats[c].montant) break;
      var nom = donnees.noms[candidats[c].champ];
      var calcul = MOTS_MOYENNES.test(nom) && !/total|montant/i.test(nom) ? 'moyenne' : 'somme';
      valeurs.push({ champ: candidats[c].champ, calcul: calcul });
    }
    // Les sommes d'abord (le tri des lignes se fait sur la première valeur).
    valeurs.sort(function (a, b) {
      return (a.calcul === 'somme' ? 0 : 1) - (b.calcul === 'somme' ? 0 : 1) || a.champ - b.champ;
    });

    // Lignes : une colonne de texte qui regroupe plusieurs lignes du fichier.
    var meilleure = -1;
    var meilleurScore = -Infinity;
    var premierTexte = -1;
    for (j = 0; j < nbChamps; j++) {
      if (donnees.types[j] !== 'texte' || donnees.origines[j]) continue;
      if (premierTexte === -1 && remplissage(donnees.colonnes[j]) > 0) premierTexte = j;
      var d = distincts[j];
      if (d < 2 || d > 100 || (d === n && n > 3) || d > n / 2 + 1) continue;
      var note = Math.min(n / d, 20) - (d > 30 ? (d - 30) / 10 : 0);
      if (MOTS_CATEGORIES.test(donnees.noms[j])) note += 6;
      else if (MOTS_ATTRIBUTS.test(donnees.noms[j])) note += 2;
      if (MOTS_TEXTES_LIBRES.test(donnees.noms[j])) note -= 5;
      if (MOTS_CODES.test(donnees.noms[j])) note -= 3;
      if (note > meilleurScore) {
        meilleurScore = note;
        meilleure = j;
      }
    }

    // Dates : regroupées par mois, trimestre ou année.
    // Les colonnes calculées d'une date se suivent : année, trimestre, mois.
    var regroupement = -1;
    for (j = 0; j + 2 < nbChamps && regroupement === -1; j++) {
      if (!donnees.origines[j] || donnees.origines[j].partie !== 'annee') continue;
      var essais = [j + 2, j + 1, j]; // mois, trimestre, année
      for (var e = 0; e < essais.length; e++) {
        if (distincts[essais[e]] >= 2 && distincts[essais[e]] <= 12) {
          regroupement = essais[e];
          break;
        }
      }
    }

    var reglages = { lignes: [], colonnes: [], valeurs: valeurs, affichage: 'normal', tri: 'libelle', filtre: { champ: -1, exclus: {} } };
    if (meilleure >= 0) {
      reglages.lignes.push(meilleure);
      reglages.tri = valeurs.length ? 'decroissant' : 'libelle';
      if (regroupement >= 0 && distincts[regroupement] * Math.max(1, valeurs.length) <= 24) {
        reglages.colonnes.push(regroupement);
      } else if (regroupement === -1) {
        // Deuxième colonne de catégories, peu variée : en colonnes.
        for (j = 0; j < nbChamps; j++) {
          if (j === meilleure || donnees.types[j] !== 'texte' || donnees.origines[j]) continue;
          if (distincts[j] >= 2 && distincts[j] <= 8 && distincts[j] < n / 2 &&
              !MOTS_TEXTES_LIBRES.test(donnees.noms[j]) && Math.max(1, valeurs.length) * distincts[j] <= 24) {
            reglages.colonnes.push(j);
            break;
          }
        }
      }
    } else if (regroupement >= 0) {
      reglages.lignes.push(regroupement);
    } else if (premierTexte >= 0) {
      reglages.lignes.push(premierTexte);
      reglages.tri = valeurs.length ? 'decroissant' : 'libelle';
    } else {
      for (j = 0; j < nbChamps; j++) {
        if (!valeurs.length || valeurs[0].champ !== j) {
          reglages.lignes.push(j);
          break;
        }
      }
    }
    if (!valeurs.length) reglages.valeurs.push({ champ: reglages.lignes[0] === undefined ? 0 : reglages.lignes[0], calcul: 'nombre' });
    return normaliserReglages(donnees, reglages);
  }

  /* ------------------------------------------------------------------ */
  /* Calcul de l'aperçu                                                 */
  /* ------------------------------------------------------------------ */

  var GRAND = '\u0001';
  var TOTAL = '\u0002';
  var SEP = '\u0003';
  var SEP_NIVEAUX = '\u0004';

  function nouveauxCumuls(nb) {
    var cumuls = [];
    for (var x = 0; x < nb; x++) cumuls.push({ s: 0, n: 0, c: 0, min: Infinity, max: -Infinity });
    return cumuls;
  }

  function cumuler(cumuls, colonnes, i) {
    for (var x = 0; x < colonnes.length; x++) {
      var v = colonnes[x][i];
      if (v === null || v === undefined) continue;
      var a = cumuls[x];
      a.c++;
      if (typeof v === 'number') {
        a.s += v;
        a.n++;
        if (v < a.min) a.min = v;
        if (v > a.max) a.max = v;
      }
    }
  }

  function resultat(a, calcul) {
    if (!a || a.c === 0) return null;
    switch (calcul) {
      case 'nombre': return a.c;
      case 'moyenne': return a.n ? a.s / a.n : null;
      case 'min': return a.n ? a.min : null;
      case 'max': return a.n ? a.max : null;
      default: return a.s;
    }
  }

  function Elements() {
    this.cles = [];
    this.valeurs = Object.create(null);
  }

  Elements.prototype.noter = function (k, v) {
    if (!(k in this.valeurs)) {
      this.valeurs[k] = v;
      this.cles.push(k);
    }
  };

  // Calcule le tableau croisé, pour l'aperçu dans la page.
  function calculer(donnees, reglagesDemandes) {
    var r = normaliserReglages(donnees, reglagesDemandes);
    var col = donnees.colonnes;
    var L1 = col[r.lignes[0]];
    var L2 = r.lignes.length > 1 ? col[r.lignes[1]] : null;
    var C = r.colonnes.length ? col[r.colonnes[0]] : null;
    var F = r.filtre.champ >= 0 ? col[r.filtre.champ] : null;
    var exclus = r.filtre.exclus;
    var V = [];
    for (var x = 0; x < r.valeurs.length; x++) V.push(col[r.valeurs[x].champ]);
    var nv = V.length;

    var cumuls = Object.create(null);
    var niveau1 = new Elements();
    var niveau2 = Object.create(null);
    var colonnes = new Elements();
    var enregistrements = 0;

    function ajouter(cleLigne, cleColonne, i) {
      var k = cleLigne + SEP + cleColonne;
      var c = cumuls[k] || (cumuls[k] = nouveauxCumuls(nv));
      cumuler(c, V, i);
    }

    for (var i = 0; i < donnees.nbLignes; i++) {
      if (F && exclus[cle(F[i])]) continue;
      enregistrements++;
      var k1 = cle(L1[i]);
      niveau1.noter(k1, L1[i]);
      var lignes = [k1, GRAND];
      if (L2) {
        var k2 = cle(L2[i]);
        (niveau2[k1] || (niveau2[k1] = new Elements())).noter(k2, L2[i]);
        lignes.push(k1 + SEP_NIVEAUX + k2);
      }
      var kc = TOTAL;
      if (C) {
        kc = cle(C[i]);
        colonnes.noter(kc, C[i]);
      }
      for (var l = 0; l < lignes.length; l++) {
        ajouter(lignes[l], kc, i);
        if (C) ajouter(lignes[l], TOTAL, i);
      }
    }

    function brut(cleLigne, cleColonne, x) {
      var c = cumuls[cleLigne + SEP + cleColonne];
      return resultat(c && c[x], r.valeurs[x].calcul);
    }

    function affiche(cleLigne, cleColonne, x) {
      var v = brut(cleLigne, cleColonne, x);
      if (v === null || r.affichage === 'normal') return v;
      var base;
      if (r.affichage === 'total') base = brut(GRAND, TOTAL, x);
      else if (r.affichage === 'ligne') base = brut(cleLigne, TOTAL, x);
      else base = brut(GRAND, cleColonne, x);
      return base ? v / base : null;
    }

    function ordonner(elements, prefixe) {
      var cles = elements.cles.slice();
      if (r.tri === 'libelle' || prefixe === null) {
        cles.sort(function (a, b) { return comparerValeurs(elements.valeurs[a], elements.valeurs[b]); });
      } else {
        var sens = r.tri === 'decroissant' ? -1 : 1;
        cles.sort(function (a, b) {
          var va = brut(prefixe + a, TOTAL, 0);
          var vb = brut(prefixe + b, TOTAL, 0);
          if (va === null || vb === null) return (va === null) - (vb === null);
          // À valeur égale, Excel range les libellés dans le même sens que le tri.
          return sens * (va - vb) || sens * comparerValeurs(elements.valeurs[a], elements.valeurs[b]);
        });
      }
      return cles;
    }

    var clesColonnes = C ? ordonner(colonnes, null) : [];
    var sortie = {
      valeurs: [],
      colonnes: [],
      lignes: [],
      total: null,
      enregistrements: enregistrements,
      nomsLignes: [],
      nomColonnes: C ? donnees.noms[r.colonnes[0]] : ''
    };
    var noms = nomsDesValeurs(donnees, r);
    for (x = 0; x < nv; x++) {
      sortie.valeurs.push({ nom: noms[x], calcul: r.valeurs[x].calcul, champ: r.valeurs[x].champ, format: formatDeValeur(donnees, r, r.valeurs[x]) });
    }
    for (x = 0; x < r.lignes.length; x++) sortie.nomsLignes.push(donnees.noms[r.lignes[x]]);
    for (var c = 0; c < clesColonnes.length; c++) {
      sortie.colonnes.push({ cle: clesColonnes[c], libelle: libelle(donnees, r.colonnes[0], colonnes.valeurs[clesColonnes[c]]) });
    }

    function ligneDeSortie(cleLigne, niveau, texte, groupe) {
      var cellules = [];
      for (var c = 0; c < clesColonnes.length; c++) {
        var valeursColonne = [];
        for (var x = 0; x < nv; x++) valeursColonne.push(affiche(cleLigne, clesColonnes[c], x));
        cellules.push(valeursColonne);
      }
      var totaux = [];
      for (x = 0; x < nv; x++) totaux.push(affiche(cleLigne, TOTAL, x));
      return { niveau: niveau, libelle: texte, groupe: groupe, cellules: cellules, totaux: totaux };
    }

    var cles1 = ordonner(niveau1, '');
    for (var a = 0; a < cles1.length; a++) {
      var k = cles1[a];
      sortie.lignes.push(ligneDeSortie(k, 0, libelle(donnees, r.lignes[0], niveau1.valeurs[k]), Boolean(L2)));
      if (L2) {
        var sous = niveau2[k];
        var cles2 = ordonner(sous, k + SEP_NIVEAUX);
        for (var b = 0; b < cles2.length; b++) {
          sortie.lignes.push(ligneDeSortie(k + SEP_NIVEAUX + cles2[b], 1,
            libelle(donnees, r.lignes[1], sous.valeurs[cles2[b]]), false));
        }
      }
    }
    sortie.total = ligneDeSortie(GRAND, 0, 'Total général', false);
    return sortie;
  }

  // Éléments distincts d'une colonne (pour la liste du filtre), dans l'ordre d'Excel.
  function elementsDistincts(donnees, j, limite) {
    var elements = new Elements();
    var colonne = donnees.colonnes[j];
    for (var i = 0; i < colonne.length; i++) {
      elements.noter(cle(colonne[i]), colonne[i]);
      if (limite && elements.cles.length > limite) return null;
    }
    var cles = elements.cles.slice().sort(function (a, b) {
      return comparerValeurs(elements.valeurs[a], elements.valeurs[b]);
    });
    var liste = [];
    for (var k = 0; k < cles.length; k++) liste.push({ cle: cles[k], libelle: libelle(donnees, j, elements.valeurs[cles[k]]) });
    return liste;
  }

  /* ------------------------------------------------------------------ */
  /* Fichier Excel : XML                                                */
  /* ------------------------------------------------------------------ */

  var SCHEMAS = 'http://schemas.openxmlformats.org/';
  var NS_PRINCIPAL = SCHEMAS + 'spreadsheetml/2006/main';
  var NS_RELATIONS = SCHEMAS + 'officeDocument/2006/relationships';
  var NS_PAQUET_RELATIONS = SCHEMAS + 'package/2006/relationships';
  var NS_TYPES = SCHEMAS + 'package/2006/content-types';
  var TYPE_TABLEUR = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
  var DEBUT_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

  var STYLE_DATE = 1;
  var STYLE_TITRE = 2;
  var STYLE_DECIMAL = 3;
  var STYLE_POURCENT = 4;
  var STYLE_DATE_HEURE = 5;

  // Retire les caractères interdits en XML (caractères de contrôle, moitiés
  // de caractères isolées).
  function nettoyerXml(t) {
    t = t.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '');
    if (!/[\uD800-\uDFFF]/.test(t)) return t;
    var sortie = '';
    for (var i = 0; i < t.length; i++) {
      var c = t.charCodeAt(i);
      if (c >= 0xD800 && c <= 0xDBFF) {
        var suivant = t.charCodeAt(i + 1);
        if (suivant >= 0xDC00 && suivant <= 0xDFFF) {
          sortie += t.charAt(i) + t.charAt(i + 1);
          i++;
        }
      } else if (c < 0xDC00 || c > 0xDFFF) {
        sortie += t.charAt(i);
      }
    }
    return sortie;
  }

  function couper(t, max) {
    if (t.length <= max) return t;
    var fin = max;
    var c = t.charCodeAt(fin - 1);
    if (c >= 0xD800 && c <= 0xDBFF) fin--;
    return t.slice(0, fin);
  }

  function texteXml(t) {
    return nettoyerXml(String(t)).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function attributXml(t) {
    return texteXml(t).replace(/"/g, '&quot;').replace(/\t/g, '&#9;').replace(/\n/g, '&#10;').replace(/\r/g, '&#13;');
  }

  function nombreXml(n) {
    return String(n);
  }

  var encodeur = null;

  function encoder(texte) {
    if (!encodeur) encodeur = new TextEncoder();
    return encodeur.encode(texte);
  }

  // Écrit un long texte par morceaux d'octets, pour ne jamais fabriquer une
  // seule chaîne géante.
  function Ecrivain() {
    this.morceaux = [];
    this.tampon = '';
  }

  Ecrivain.prototype.ecrire = function (texte) {
    this.tampon += texte;
    if (this.tampon.length > 1048576) this.vider();
  };

  Ecrivain.prototype.vider = function () {
    if (this.tampon) {
      this.morceaux.push(encoder(this.tampon));
      this.tampon = '';
    }
  };

  Ecrivain.prototype.fin = function () {
    this.vider();
    return this.morceaux;
  };

  function styleDeColonne(donnees, j) {
    if (donnees.types[j] === 'date') return donnees.formats[j] === 'dateheure' ? STYLE_DATE_HEURE : STYLE_DATE;
    if (donnees.types[j] === 'nombre') {
      if (donnees.formats[j] === 'pourcent') return STYLE_POURCENT;
      if (donnees.formats[j] === 'decimal') return STYLE_DECIMAL;
    }
    return 0;
  }

  function celluleTexte(ref, texte, style) {
    return '<c r="' + ref + '" t="inlineStr"' + (style ? ' s="' + style + '"' : '') +
      '><is><t xml:space="preserve">' + texteXml(couper(texte, MAX_LONGUEUR_CELLULE)) + '</t></is></c>';
  }

  function largeurs(donnees) {
    var cols = '';
    var echantillon = Math.min(donnees.nbLignes, 200);
    for (var j = 0; j < donnees.noms.length; j++) {
      var largeur = donnees.noms[j].length + 3;
      for (var i = 0; i < echantillon; i++) {
        var v = donnees.colonnes[j][i];
        var t = v === null ? '' : (typeof v === 'number' ? (donnees.types[j] === 'date' ? '00/00/0000 00:00' : formaterNombre(v)) : v);
        if (t.length + 2 > largeur) largeur = t.length + 2;
      }
      largeur = Math.max(8, Math.min(60, largeur));
      cols += '<col min="' + (j + 1) + '" max="' + (j + 1) + '" width="' + largeur + '" customWidth="1"/>';
    }
    return '<cols>' + cols + '</cols>';
  }

  function feuilleDonnees(donnees) {
    var w = new Ecrivain();
    var nbChamps = donnees.noms.length;
    var lettres = [];
    var styles = [];
    for (var j = 0; j < nbChamps; j++) {
      lettres.push(lettre(j));
      styles.push(styleDeColonne(donnees, j));
    }
    w.ecrire(DEBUT_XML + '<worksheet xmlns="' + NS_PRINCIPAL + '" xmlns:r="' + NS_RELATIONS + '">');
    w.ecrire('<dimension ref="A1:' + lettres[nbChamps - 1] + (donnees.nbLignes + 1) + '"/>');
    w.ecrire('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
      '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>');
    w.ecrire('<sheetFormatPr defaultRowHeight="15"/>' + largeurs(donnees) + '<sheetData>');
    var ligne = '<row r="1">';
    for (j = 0; j < nbChamps; j++) ligne += celluleTexte(lettres[j] + '1', donnees.noms[j], STYLE_TITRE);
    w.ecrire(ligne + '</row>');
    for (var i = 0; i < donnees.nbLignes; i++) {
      var r = i + 2;
      ligne = '<row r="' + r + '">';
      for (j = 0; j < nbChamps; j++) {
        var v = donnees.colonnes[j][i];
        if (v === null || v === undefined) continue;
        if (typeof v === 'number') {
          ligne += '<c r="' + lettres[j] + r + '"' + (styles[j] ? ' s="' + styles[j] + '"' : '') + '><v>' + nombreXml(v) + '</v></c>';
        } else {
          ligne += celluleTexte(lettres[j] + r, v, 0);
        }
      }
      w.ecrire(ligne + '</row>');
    }
    w.ecrire('</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>');
    return w.fin();
  }

  function feuilleTableau() {
    return DEBUT_XML + '<worksheet xmlns="' + NS_PRINCIPAL + '" xmlns:r="' + NS_RELATIONS + '">' +
      '<dimension ref="A3"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/><sheetData/>' +
      '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
  }

  // Ce que contient une colonne, tel que le décrit le cache du tableau croisé.
  function contenu(donnees, j) {
    var colonne = donnees.colonnes[j];
    var estDate = donnees.types[j] === 'date';
    var f = { s: false, n: false, d: false, m: false, entiers: true, min: Infinity, max: -Infinity };
    for (var i = 0; i < colonne.length; i++) {
      var v = colonne[i];
      if (v === null || v === undefined) {
        f.m = true;
      } else if (typeof v === 'number') {
        if (estDate) f.d = true;
        else f.n = true;
        if (v !== Math.floor(v)) f.entiers = false;
        if (v < f.min) f.min = v;
        if (v > f.max) f.max = v;
      } else {
        f.s = true;
      }
    }
    return f;
  }

  function attributsElements(f) {
    var a = '';
    if (!f.s && !f.m) a += ' containsSemiMixedTypes="0"';
    if (!f.s && !f.n && f.d) a += ' containsNonDate="0"';
    if (f.d) a += ' containsDate="1"';
    if (!f.s) a += ' containsString="0"';
    if (f.m) a += ' containsBlank="1"';
    if ((f.s ? 1 : 0) + (f.n ? 1 : 0) + (f.d ? 1 : 0) > 1) a += ' containsMixedTypes="1"';
    if (f.n) {
      a += ' containsNumber="1"';
      if (f.entiers) a += ' containsInteger="1"';
      if (!f.d) a += ' minValue="' + nombreXml(f.min) + '" maxValue="' + nombreXml(f.max) + '"';
    }
    if (f.d && !f.n) a += ' minDate="' + isoDepuisSerie(f.min) + '" maxDate="' + isoDepuisSerie(f.max) + '"';
    return a;
  }

  // Éléments partagés d'une colonne (liste des valeurs distinctes, dans
  // l'ordre où elles apparaissent).
  function elementsPartages(donnees, j) {
    var index = Object.create(null);
    var liste = [];
    var colonne = donnees.colonnes[j];
    for (var i = 0; i < colonne.length; i++) {
      var k = cle(colonne[i]);
      if (!(k in index)) {
        index[k] = liste.length;
        liste.push(colonne[i]);
      }
    }
    return { index: index, liste: liste };
  }

  function xmlElement(donnees, j, v) {
    if (v === null || v === undefined) return '<m/>';
    if (typeof v === 'number') {
      return donnees.types[j] === 'date' ? '<d v="' + isoDepuisSerie(v) + '"/>' : '<n v="' + nombreXml(v) + '"/>';
    }
    return '<s v="' + attributXml(couper(v, MAX_LONGUEUR_ELEMENT)) + '"/>';
  }

  function definitionCache(donnees, partages) {
    var nbChamps = donnees.noms.length;
    var champs = '';
    for (var j = 0; j < nbChamps; j++) {
      var attributs = attributsElements(contenu(donnees, j));
      var elements = '';
      if (partages[j]) {
        var liste = partages[j].liste;
        for (var k = 0; k < liste.length; k++) elements += xmlElement(donnees, j, liste[k]);
        elements = '<sharedItems' + attributs + ' count="' + liste.length + '">' + elements + '</sharedItems>';
      } else {
        elements = '<sharedItems' + attributs + '/>';
      }
      champs += '<cacheField name="' + attributXml(donnees.noms[j]) + '" numFmtId="' +
        (donnees.types[j] === 'date' ? 14 : 0) + '">' + elements + '</cacheField>';
    }
    return DEBUT_XML + '<pivotCacheDefinition xmlns="' + NS_PRINCIPAL + '" xmlns:r="' + NS_RELATIONS + '" r:id="rId1"' +
      ' refreshOnLoad="1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="0">' +
      '<cacheSource type="worksheet"><worksheetSource ref="A1:' + lettre(nbChamps - 1) + (donnees.nbLignes + 1) +
      '" sheet="Données"/></cacheSource><cacheFields count="' + nbChamps + '">' + champs + '</cacheFields></pivotCacheDefinition>';
  }

  function enregistrementsCache() {
    return DEBUT_XML + '<pivotCacheRecords xmlns="' + NS_PRINCIPAL + '" xmlns:r="' + NS_RELATIONS + '" count="0"/>';
  }

  var TRI_PAR_VALEUR = '<autoSortScope><pivotArea dataOnly="0" outline="0" fieldPosition="0"><references count="1">' +
    '<reference field="4294967294" count="1" selected="0"><x v="0"/></reference></references></pivotArea></autoSortScope>';

  function numFmtDeValeur(format) {
    if (format === 'pourcent') return 10;
    if (format === 'entier') return 3;
    if (format === 'date') return 14;
    return 4;
  }

  function definitionTableau(donnees, r, partages) {
    var nbChamps = donnees.noms.length;
    var axes = {};
    var x;
    for (x = 0; x < r.lignes.length; x++) axes[r.lignes[x]] = 'axisRow';
    for (x = 0; x < r.colonnes.length; x++) axes[r.colonnes[x]] = 'axisCol';
    if (r.filtre.champ >= 0) axes[r.filtre.champ] = 'axisPage';
    var donneesChamps = {};
    for (x = 0; x < r.valeurs.length; x++) donneesChamps[r.valeurs[x].champ] = true;

    var champs = '';
    for (var j = 0; j < nbChamps; j++) {
      var axe = axes[j];
      var a = '';
      if (axe) a += ' axis="' + axe + '"';
      if (donneesChamps[j]) a += ' dataField="1"';
      var masques = 0;
      if (axe === 'axisPage') {
        for (var k in r.filtre.exclus) {
          if (k in partages[j].index) masques++;
        }
        if (masques) a += ' multipleItemSelectionAllowed="1"';
      }
      a += ' showAll="0"';
      var triValeur = axe === 'axisRow' && r.tri !== 'libelle';
      if (axe === 'axisRow') a += r.tri === 'decroissant' ? ' sortType="descending"' : ' sortType="ascending"';
      if (axe === 'axisCol') a += ' sortType="ascending"';
      if (!axe) {
        champs += '<pivotField' + a + '/>';
        continue;
      }
      var liste = partages[j].liste;
      var ordre = [];
      for (var i = 0; i < liste.length; i++) ordre.push(i);
      ordre.sort(function (p, q) { return comparerValeurs(liste[p], liste[q]) || p - q; });
      var elements = '';
      for (i = 0; i < ordre.length; i++) {
        var masque = axe === 'axisPage' && r.filtre.exclus[cle(liste[ordre[i]])];
        elements += '<item x="' + ordre[i] + '"' + (masque ? ' h="1"' : '') + '/>';
      }
      elements += '<item t="default"/>';
      champs += '<pivotField' + a + '><items count="' + (ordre.length + 1) + '">' + elements + '</items>' +
        (triValeur ? TRI_PAR_VALEUR : '') + '</pivotField>';
    }

    var noms = nomsDesValeurs(donnees, r);
    var xml = DEBUT_XML + '<pivotTableDefinition xmlns="' + NS_PRINCIPAL + '" name="TableauCroise1" cacheId="1"' +
      ' applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0"' +
      ' applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Valeurs" updatedVersion="6"' +
      ' minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0"' +
      ' outline="1" outlineData="1" multipleFieldFilters="0"';
    var nomsLignes = [];
    for (x = 0; x < r.lignes.length; x++) nomsLignes.push(donnees.noms[r.lignes[x]]);
    xml += ' rowHeaderCaption="' + attributXml(couper(nomsLignes.join(' / '), 255)) + '"';
    if (r.colonnes.length) xml += ' colHeaderCaption="' + attributXml(couper(donnees.noms[r.colonnes[0]], 255)) + '"';
    xml += '><location ref="A3:B4" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"' +
      (r.filtre.champ >= 0 ? ' rowPageCount="1" colPageCount="1"' : '') + '/>';
    xml += '<pivotFields count="' + nbChamps + '">' + champs + '</pivotFields>';
    xml += '<rowFields count="' + r.lignes.length + '">';
    for (x = 0; x < r.lignes.length; x++) xml += '<field x="' + r.lignes[x] + '"/>';
    xml += '</rowFields>';
    var colonnesChamps = r.colonnes.slice();
    if (r.valeurs.length > 1) colonnesChamps.push(-2);
    if (colonnesChamps.length) {
      xml += '<colFields count="' + colonnesChamps.length + '">';
      for (x = 0; x < colonnesChamps.length; x++) xml += '<field x="' + colonnesChamps[x] + '"/>';
      xml += '</colFields>';
    }
    if (r.filtre.champ >= 0) xml += '<pageFields count="1"><pageField fld="' + r.filtre.champ + '" hier="-1"/></pageFields>';
    xml += '<dataFields count="' + r.valeurs.length + '">';
    for (x = 0; x < r.valeurs.length; x++) {
      var v = r.valeurs[x];
      xml += '<dataField name="' + attributXml(noms[x]) + '" fld="' + v.champ + '"' +
        (v.calcul !== 'somme' ? ' subtotal="' + CALCULS_EXCEL[v.calcul] + '"' : '') +
        (r.affichage !== 'normal' ? ' showDataAs="' + AFFICHAGES[r.affichage] + '"' : '') +
        ' baseField="0" baseItem="0" numFmtId="' + numFmtDeValeur(formatDeValeur(donnees, r, v)) + '"/>';
    }
    xml += '</dataFields><pivotTableStyleInfo name="PivotStyleMedium2" showRowHeaders="1" showColHeaders="1"' +
      ' showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>';
    return xml;
  }

  var STYLES = DEBUT_XML + '<styleSheet xmlns="' + NS_PRINCIPAL + '">' +
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font>' +
    '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="6">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>' +
    '</styleSheet>';

  function relations(liste) {
    var xml = DEBUT_XML + '<Relationships xmlns="' + NS_PAQUET_RELATIONS + '">';
    for (var i = 0; i < liste.length; i++) {
      xml += '<Relationship Id="rId' + (i + 1) + '" Type="' + NS_RELATIONS + '/' + liste[i][0] + '" Target="' + liste[i][1] + '"/>';
    }
    return xml + '</Relationships>';
  }

  var TYPES_CONTENU = DEBUT_XML + '<Types xmlns="' + NS_TYPES + '">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="' + TYPE_TABLEUR + 'sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="' + TYPE_TABLEUR + 'worksheet+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="' + TYPE_TABLEUR + 'worksheet+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="' + TYPE_TABLEUR + 'styles+xml"/>' +
    '<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="' + TYPE_TABLEUR + 'pivotTable+xml"/>' +
    '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="' + TYPE_TABLEUR + 'pivotCacheDefinition+xml"/>' +
    '<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="' + TYPE_TABLEUR + 'pivotCacheRecords+xml"/>' +
    '</Types>';

  var CLASSEUR = DEBUT_XML + '<workbook xmlns="' + NS_PRINCIPAL + '" xmlns:r="' + NS_RELATIONS + '">' +
    '<bookViews><workbookView activeTab="0"/></bookViews><sheets>' +
    '<sheet name="Tableau croisé" sheetId="1" r:id="rId1"/><sheet name="Données" sheetId="2" r:id="rId2"/>' +
    '</sheets><calcPr calcId="191029"/><pivotCaches><pivotCache cacheId="1" r:id="rId4"/></pivotCaches></workbook>';

  /* ------------------------------------------------------------------ */
  /* Fichier Excel : dossier compressé (ZIP, sans compression)          */
  /* ------------------------------------------------------------------ */

  var TABLE_CRC = (function () {
    var table = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
    return table;
  }());

  function crc32(crc, octets) {
    for (var i = 0; i < octets.length; i++) crc = TABLE_CRC[(crc ^ octets[i]) & 0xFF] ^ (crc >>> 8);
    return crc;
  }

  // Assemble les fichiers ({ nom, morceaux: [Uint8Array] }) en un ZIP.
  // Date fixe (1er janvier 1980) : le fichier ne dit pas quand il a été créé.
  function zipper(fichiers) {
    var sortie = [];
    var central = [];
    var position = 0;
    for (var f = 0; f < fichiers.length; f++) {
      var nom = encoder(fichiers[f].nom);
      var morceaux = fichiers[f].morceaux;
      var taille = 0;
      var crc = 0xFFFFFFFF;
      for (var m = 0; m < morceaux.length; m++) {
        taille += morceaux[m].length;
        crc = crc32(crc, morceaux[m]);
      }
      crc = (crc ^ 0xFFFFFFFF) >>> 0;

      var local = new Uint8Array(30 + nom.length);
      var vl = new DataView(local.buffer);
      vl.setUint32(0, 0x04034B50, true);
      vl.setUint16(4, 20, true);
      vl.setUint16(12, 0x21, true);
      vl.setUint32(14, crc, true);
      vl.setUint32(18, taille, true);
      vl.setUint32(22, taille, true);
      vl.setUint16(26, nom.length, true);
      local.set(nom, 30);
      sortie.push(local);
      for (m = 0; m < morceaux.length; m++) sortie.push(morceaux[m]);

      var entree = new Uint8Array(46 + nom.length);
      var vc = new DataView(entree.buffer);
      vc.setUint32(0, 0x02014B50, true);
      vc.setUint16(4, 20, true);
      vc.setUint16(6, 20, true);
      vc.setUint16(14, 0x21, true);
      vc.setUint32(16, crc, true);
      vc.setUint32(20, taille, true);
      vc.setUint32(24, taille, true);
      vc.setUint16(28, nom.length, true);
      vc.setUint32(42, position, true);
      entree.set(nom, 46);
      central.push(entree);
      position += local.length + taille;
    }
    var tailleCentral = 0;
    for (var c = 0; c < central.length; c++) {
      sortie.push(central[c]);
      tailleCentral += central[c].length;
    }
    var fin = new Uint8Array(22);
    var vf = new DataView(fin.buffer);
    vf.setUint32(0, 0x06054B50, true);
    vf.setUint16(8, fichiers.length, true);
    vf.setUint16(10, fichiers.length, true);
    vf.setUint32(12, tailleCentral, true);
    vf.setUint32(16, position, true);
    sortie.push(fin);
    return sortie;
  }

  // Fabrique le fichier .xlsx : une feuille « Tableau croisé » (le tableau
  // croisé dynamique, calculé par Excel à l'ouverture) et une feuille
  // « Données ». Renvoie une liste de morceaux d'octets (pour un Blob).
  function genererXlsx(donnees, reglagesDemandes) {
    if (!donnees.nbLignes) throw new Error('aucune donnée');
    var r = normaliserReglages(donnees, reglagesDemandes);
    var partages = {};
    var axes = r.lignes.concat(r.colonnes);
    if (r.filtre.champ >= 0) axes.push(r.filtre.champ);
    for (var x = 0; x < axes.length; x++) partages[axes[x]] = elementsPartages(donnees, axes[x]);

    var texte = function (nom, contenuXml) { return { nom: nom, morceaux: [encoder(contenuXml)] }; };
    return zipper([
      texte('[Content_Types].xml', TYPES_CONTENU),
      texte('_rels/.rels', relations([['officeDocument', 'xl/workbook.xml']])),
      texte('xl/workbook.xml', CLASSEUR),
      texte('xl/_rels/workbook.xml.rels', relations([
        ['worksheet', 'worksheets/sheet1.xml'],
        ['worksheet', 'worksheets/sheet2.xml'],
        ['styles', 'styles.xml'],
        ['pivotCacheDefinition', 'pivotCache/pivotCacheDefinition1.xml']
      ])),
      texte('xl/styles.xml', STYLES),
      texte('xl/worksheets/sheet1.xml', feuilleTableau()),
      texte('xl/worksheets/_rels/sheet1.xml.rels', relations([['pivotTable', '../pivotTables/pivotTable1.xml']])),
      { nom: 'xl/worksheets/sheet2.xml', morceaux: feuilleDonnees(donnees) },
      texte('xl/pivotTables/pivotTable1.xml', definitionTableau(donnees, r, partages)),
      texte('xl/pivotTables/_rels/pivotTable1.xml.rels', relations([['pivotCacheDefinition', '../pivotCache/pivotCacheDefinition1.xml']])),
      texte('xl/pivotCache/pivotCacheDefinition1.xml', definitionCache(donnees, partages)),
      texte('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels', relations([['pivotCacheRecords', 'pivotCacheRecords1.xml']])),
      texte('xl/pivotCache/pivotCacheRecords1.xml', enregistrementsCache())
    ]);
  }

  return {
    CALCULS: CALCULS,
    NOMS_CALCULS: NOMS_CALCULS,
    MAX_VALEURS: MAX_VALEURS,
    lireNombre: lireNombre,
    lireDate: lireDate,
    preparer: preparer,
    proposer: proposer,
    normaliserReglages: normaliserReglages,
    calculer: calculer,
    elementsDistincts: elementsDistincts,
    formaterValeur: formaterValeur,
    cle: cle,
    genererXlsx: genererXlsx
  };
}));
