/*
 * app.js — Convertisseur TXT vers CSV : fonctionnement de la page.
 *
 * Le fichier choisi est lu dans la mémoire du navigateur, converti, puis
 * proposé au téléchargement. Rien n'est envoyé sur le réseau et rien n'est
 * enregistré : après le téléchargement, les données sont effacées.
 */
(function () {
  'use strict';

  var C = window.ConversionCSV;
  var T = window.TableauCroise; // tableau croisé dynamique (peut manquer : le reste fonctionne)

  var TAILLE_MAX = 50 * 1024 * 1024; // 50 Mo
  var LIGNES_APERCU = 10;
  var COLONNES_APERCU = 50;
  var LONGUEUR_CELLULE_APERCU = 200;
  var LONGUEUR_VALEUR_COURTE = 25; // au-delà, la valeur peut passer à la ligne dans l'aperçu
  // Délai avant de libérer l'adresse du fichier généré : certains navigateurs
  // (Safari, notamment) en ont besoin un moment pour démarrer le
  // téléchargement. Les données de la page, elles, sont effacées tout de suite.
  var DELAI_LIBERATION = 40000;
  // Juste après le téléchargement, un clic de souris vient presque toujours
  // d'un double-clic sur « Télécharger », ou d'un second clic impatient : il
  // ne doit pas actionner le bouton « Recommencer » apparu entre-temps sous la
  // souris (ce qui effacerait l'écran qui explique où trouver le fichier).
  var DELAI_DOUBLE_CLIC = 2500;
  // Au-delà de cette longueur de texte, recalculer le tableau prend un moment.
  var TEXTE_LONG = 2000000;

  // Tableau croisé dynamique.
  var LIGNES_APERCU_TCD = 300;
  var MAX_ELEMENTS_FILTRE = 200;
  var SANS_CHAMP = '-1';
  var TYPE_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var TEXTES_CALCULS = { somme: 'Somme', nombre: 'Nombre', moyenne: 'Moyenne', min: 'Minimum', max: 'Maximum' };

  // Les deux outils proposés sur la page d'accueil.
  var NOMS_OUTILS = { csv: 'TXT en CSV', tcd: 'TXT en tableau croisé dynamique' };

  var INSECABLE = '\u00A0'; // espace insécable (avant « : », « ? »… et dans « »)

  var MESSAGES = {
    'vide': 'Ce fichier est vide. Choisissez un autre fichier.',
    'lignes-vides': 'Ce fichier ne contient que des lignes vides. Choisissez un autre fichier.',
    'binaire': "Ce fichier n'est pas un fichier texte. Choisissez un fichier dont le nom se termine par .txt.",
    'rtf': "Ce fichier est un document de traitement de texte (Word ou WordPad), pas un simple fichier texte. " +
      "Enregistrez-le d'abord au format " + guillemets('Texte (.txt)') + ', puis réessayez.',
    'outil': "Choisissez d'abord ce que vous voulez faire" + INSECABLE + ': un fichier CSV ou un tableau croisé dynamique.',
    'dossier': 'Vous avez déposé un dossier. Ouvrez-le, puis faites glisser le fichier .txt lui-même.',
    'trop-gros': 'Ce fichier est trop gros (plus de 50' + INSECABLE + 'Mo). Cet outil est prévu pour des fichiers plus petits.',
    'lecture': "Ce fichier n'a pas pu être lu. Réessayez, ou choisissez un autre fichier.",
    'demarrage': "L'outil n'a pas pu démarrer. Rechargez la page.",
    'navigateur': 'Votre navigateur est trop ancien pour cet outil. Essayez avec un navigateur à jour (Edge, Chrome, Firefox ou Safari).',
    'enregistrement': "Sur cet appareil, le navigateur ne permet pas d'enregistrer le fichier créé. Essayez sur un ordinateur."
  };

  var NOMS_ENCODAGES = {
    'utf-8': 'UTF-8',
    'utf-16le': 'UTF-16 (Unicode)',
    'utf-16be': 'UTF-16 (Unicode)',
    'windows-1252': 'Windows-1252 (ancien format de Windows)'
  };

  function $(id) {
    return document.getElementById(id);
  }

  var el = {
    zone: $('zone-depot'),
    boutonChoisir: $('bouton-choisir'),
    chargement: $('chargement'),
    champFichier: $('champ-fichier'),
    etapeOutil: $('etape-outil'),
    titreOutil: $('titre-outil'),
    cartesOutils: document.querySelectorAll('[data-outil]'),
    nomOutil: $('nom-outil'),
    sectionCsv: $('section-csv'),
    etapeChoix: $('etape-choix'),
    etapeResultat: $('etape-resultat'),
    etapeTermine: $('etape-termine'),
    erreur: $('message-erreur'),
    texteErreur: $('texte-erreur'),
    titreApercu: $('titre-apercu'),
    resume: $('resume-fichier'),
    separateur: $('choix-separateur'),
    aideDefilement: $('aide-defilement'),
    cadreTableau: $('cadre-tableau'),
    tableau: $('tableau-apercu'),
    legende: $('legende-apercu'),
    suite: $('suite-apercu'),
    avertissements: $('avertissements'),
    details: document.querySelector('.details'),
    encodage: $('detail-encodage'),
    lignesVides: $('detail-lignes-vides'),
    boutonTelecharger: $('bouton-telecharger'),
    titreTermine: $('titre-termine'),
    nomTelecharge: $('nom-telecharge'),
    nomNumerote: $('nom-numerote'),
    ouvrirExcel: $('ouvrir-excel'),
    ouvrirStandard: $('ouvrir-standard'),
    ouvrirTcd: $('ouvrir-tcd'),
    annonce: $('annonce'),
    // Tableau croisé dynamique
    sectionTcd: $('section-tcd'),
    panneauTcd: $('panneau-tcd'),
    resumeTcd: $('resume-tcd'),
    erreurTcd: $('erreur-tcd'),
    tcdEntete: $('tcd-entete'),
    tcdLignes1: $('tcd-lignes-1'),
    tcdLignes2: $('tcd-lignes-2'),
    tcdColonnes: $('tcd-colonnes'),
    tcdCalculs: [$('tcd-calcul-1'), $('tcd-calcul-2'), $('tcd-calcul-3')],
    tcdValeurs: [$('tcd-valeur-1'), $('tcd-valeur-2'), $('tcd-valeur-3')],
    tcdAffichage: $('tcd-affichage'),
    tcdTri: $('tcd-tri'),
    tcdFiltre: $('tcd-filtre'),
    tcdFiltreValeurs: $('tcd-filtre-valeurs'),
    tcdFiltreListe: $('tcd-filtre-liste'),
    tcdFiltreTrop: $('tcd-filtre-trop'),
    tcdToutCocher: $('tcd-tout-cocher'),
    tcdToutDecocher: $('tcd-tout-decocher'),
    tableauTcd: $('tableau-tcd'),
    suiteTcd: $('suite-tcd'),
    boutonTelechargerTcd: $('bouton-telecharger-tcd')
  };

  // Tout ce que la page sait du fichier en cours. Vidé après le téléchargement.
  var etat = {
    lecture: 0,          // numéro de la lecture en cours (une lecture annulée est ignorée)
    horsService: false,  // navigateur incompatible
    outil: null,         // outil choisi sur la page d'accueil : 'csv' ou 'tcd'
    nomFichier: '',
    nomCsv: '',
    texte: null,
    lignes: null,
    colonnes: 0,
    lignesVides: 0,
    longueurMax: 0,
    separateur: null,
    encodage: '',
    bom: false,
    illisibles: false,
    encodagesMelanges: false,
    encodageDouteux: false,
    finAbimee: false,
    plusieursFichiers: false,
    recalculEnAttente: null, // nouveau découpage du tableau pas encore fait
    tcd: null,           // tableau croisé : données préparées, réglages, filtre
    url: null,           // adresse temporaire du fichier généré (CSV ou Excel)
    minuterie: null,
    termineA: 0          // moment où l'écran « C'est fait » est apparu
  };

  /* ---- Petits utilitaires ---------------------------------------------- */

  function nombre(n) {
    return n.toLocaleString('fr-FR');
  }

  function pluriel(n, singulier, formePlurielle) {
    return nombre(n) + ' ' + (n > 1 ? (formePlurielle || singulier + 's') : singulier);
  }

  function guillemets(texte) {
    return '«' + INSECABLE + texte + INSECABLE + '»';
  }

  // Message lu par les lecteurs d'écran.
  var minuterieAnnonce = null;

  function annoncer(message) {
    clearTimeout(minuterieAnnonce);
    el.annonce.textContent = '';
    minuterieAnnonce = setTimeout(function () {
      minuterieAnnonce = null;
      el.annonce.textContent = message;
    }, 50);
  }

  function viderAnnonce() {
    clearTimeout(minuterieAnnonce);
    minuterieAnnonce = null;
    el.annonce.textContent = '';
  }

  // 0 → A, 25 → Z, 26 → AA… comme les colonnes d'Excel.
  function lettreColonne(index) {
    var lettres = '';
    var n = index + 1;
    while (n > 0) {
      var reste = (n - 1) % 26;
      lettres = String.fromCharCode(65 + reste) + lettres;
      n = Math.floor((n - 1) / 26);
    }
    return lettres;
  }

  // Raccourcit une valeur très longue pour l'aperçu (le CSV, lui, garde tout).
  function raccourcir(valeur) {
    if (valeur.length <= LONGUEUR_CELLULE_APERCU) return valeur;
    var fin = LONGUEUR_CELLULE_APERCU;
    var code = valeur.charCodeAt(fin - 1);
    if (code >= 0xD800 && code <= 0xDBFF) fin--; // ne pas couper un caractère en deux
    return valeur.slice(0, fin) + '…';
  }

  /* ---- Étapes et messages ---------------------------------------------- */

  function afficherEtape(nom) {
    if (nom === 'choix' && !etat.outil) nom = 'outil';
    el.etapeOutil.hidden = nom !== 'outil';
    el.etapeChoix.hidden = nom !== 'choix';
    el.etapeResultat.hidden = nom !== 'resultat';
    el.etapeTermine.hidden = nom !== 'termine';
  }

  function occupe(actif) {
    el.boutonChoisir.disabled = actif || etat.horsService;
    for (var i = 0; i < el.cartesOutils.length; i++) {
      var carte = el.cartesOutils[i];
      carte.disabled = etat.horsService || (carte.getAttribute('data-outil') === 'tcd' && !T);
    }
    el.boutonChoisir.textContent = actif ? 'Lecture du fichier…' : 'Choisir mon fichier .txt';
    el.zone.setAttribute('aria-busy', actif ? 'true' : 'false');
  }

  function afficherErreur(code, deplacerFocus) {
    // Le message d'erreur est lu par les lecteurs d'écran (role="alert") :
    // « Lecture du fichier… », encore en attente, ne doit pas le suivre.
    viderAnnonce();
    el.texteErreur.textContent = MESSAGES[code] || MESSAGES.lecture;
    el.erreur.hidden = false;
    afficherEtape('choix');
    if (deplacerFocus) el.erreur.focus();
  }

  function masquerErreur() {
    el.erreur.hidden = true;
    el.texteErreur.textContent = '';
  }

  /* ---- Effacement ------------------------------------------------------ */

  function libererFichierGenere() {
    if (etat.minuterie) {
      clearTimeout(etat.minuterie);
      etat.minuterie = null;
    }
    if (etat.url) {
      URL.revokeObjectURL(etat.url);
      etat.url = null;
    }
  }

  // Le moteur JavaScript garde en mémoire la dernière chaîne examinée par une
  // expression régulière (RegExp.input) : ce pourrait être le texte du
  // fichier. Une recherche réussie sur une chaîne vide la remplace.
  function oublierDerniereRecherche() {
    /^/.test('');
  }

  // Oublie le contenu du fichier : variables et tableau d'aperçu.
  function effacerDonnees() {
    etat.nomFichier = '';
    etat.nomCsv = '';
    etat.texte = null;
    etat.lignes = null;
    etat.colonnes = 0;
    etat.lignesVides = 0;
    etat.longueurMax = 0;
    etat.separateur = null;
    etat.encodage = '';
    etat.bom = false;
    etat.illisibles = false;
    etat.encodagesMelanges = false;
    etat.encodageDouteux = false;
    etat.finAbimee = false;
    etat.plusieursFichiers = false;
    etat.recalculEnAttente = null;
    el.champFichier.value = '';
    el.resume.textContent = '';
    el.tableau.tHead.textContent = '';
    el.tableau.tBodies[0].textContent = '';
    el.legende.textContent = '';
    el.suite.textContent = '';
    el.suite.hidden = true;
    el.avertissements.textContent = '';
    el.avertissements.hidden = true;
    el.encodage.textContent = '';
    el.lignesVides.textContent = '';
    el.aideDefilement.hidden = true;
    el.cadreTableau.removeAttribute('aria-describedby');
    el.cadreTableau.removeAttribute('aria-busy');
    viderTcd();
    oublierDerniereRecherche();
  }

  // Remet la page dans son état de départ.
  function reinitialiser(deplacerFocus) {
    etat.lecture++;
    libererFichierGenere();
    effacerDonnees();
    masquerErreur();
    occupe(false);
    viderAnnonce();
    el.nomTelecharge.textContent = '';
    el.nomNumerote.textContent = '';
    el.details.open = false;
    var excel = document.querySelector('input[name="format"][value="excel"]');
    if (excel) excel.checked = true;
    afficherEtape('choix');
    if (deplacerFocus) (etat.outil ? el.boutonChoisir : el.titreOutil).focus();
  }

  /* ---- Choix de l'outil ------------------------------------------------ */

  function choisirOutil(outil) {
    if (outil === 'tcd' && !T) return;
    reinitialiser(false);
    etat.outil = outil;
    el.nomOutil.textContent = NOMS_OUTILS[outil];
    afficherEtape('choix');
    el.boutonChoisir.focus();
    annoncer('Outil choisi' + INSECABLE + ': ' + NOMS_OUTILS[outil] + '.');
  }

  // Retour à la page d'accueil, sans rien garder du fichier.
  function changerOutil() {
    reinitialiser(false);
    etat.outil = null;
    el.nomOutil.textContent = '';
    afficherEtape('outil');
    el.titreOutil.focus();
  }

  /* ---- Lecture du fichier ---------------------------------------------- */

  function chargerFichier(fichier, plusieurs) {
    reinitialiser(false);
    if (!fichier) return;
    if (!etat.outil) {
      afficherErreur('outil', true);
      return;
    }
    if (fichier.size > TAILLE_MAX) {
      afficherErreur('trop-gros', true);
      return;
    }
    if (fichier.size === 0) {
      afficherErreur('vide', true);
      return;
    }
    var lecture = etat.lecture;
    occupe(true);
    annoncer('Lecture du fichier…');
    var lecteur = new FileReader();
    lecteur.onload = function () {
      if (lecture !== etat.lecture) return;
      var octets = new Uint8Array(lecteur.result);
      // Laisse le navigateur afficher « Lecture du fichier… » avant le calcul.
      setTimeout(function () {
        if (lecture !== etat.lecture) return;
        var resultat;
        try {
          resultat = C.analyser(octets);
        } catch (e) {
          resultat = { ok: false, erreur: 'lecture' };
        }
        octets = null;
        occupe(false);
        if (!resultat.ok) {
          oublierDerniereRecherche();
          afficherErreur(resultat.erreur, true);
          return;
        }
        etat.nomFichier = fichier.name;
        etat.nomCsv = C.nomFichierCsv(fichier.name);
        etat.texte = resultat.texte;
        etat.lignes = resultat.lignes;
        etat.colonnes = resultat.colonnes;
        etat.lignesVides = resultat.lignesVides;
        etat.longueurMax = resultat.longueurMax;
        etat.separateur = resultat.separateur;
        etat.encodage = resultat.encodage;
        etat.bom = resultat.bom;
        etat.illisibles = resultat.caracteresIllisibles;
        etat.encodagesMelanges = resultat.encodagesMelanges;
        etat.encodageDouteux = resultat.encodageDouteux;
        etat.finAbimee = resultat.finAbimee;
        etat.plusieursFichiers = Boolean(plusieurs);
        afficherResultat(true);
        annoncer('Fichier lu' + INSECABLE + ': ' + pluriel(etat.lignes.length, 'ligne') + ', ' +
          pluriel(etat.colonnes, 'colonne') + '.');
      }, 20);
    };
    lecteur.onerror = function () {
      if (lecture !== etat.lecture) return;
      occupe(false);
      afficherErreur('lecture', true);
    };
    lecteur.readAsArrayBuffer(fichier);
  }

  /* ---- Aperçu ---------------------------------------------------------- */

  function remplirTableau(lignes, colonnes) {
    var tete = el.tableau.tHead;
    var corps = el.tableau.tBodies[0];
    tete.textContent = '';
    corps.textContent = '';
    var nbColonnes = Math.min(colonnes, COLONNES_APERCU);

    var ligneTete = document.createElement('tr');
    var coin = document.createElement('td');
    coin.className = 'coin';
    ligneTete.appendChild(coin);
    for (var c = 0; c < nbColonnes; c++) {
      var th = document.createElement('th');
      th.scope = 'col';
      th.textContent = lettreColonne(c);
      ligneTete.appendChild(th);
    }
    tete.appendChild(ligneTete);

    for (var l = 0; l < lignes.length; l++) {
      var tr = document.createElement('tr');
      var numero = document.createElement('th');
      numero.scope = 'row';
      numero.textContent = String(l + 1);
      tr.appendChild(numero);
      for (var k = 0; k < nbColonnes; k++) {
        var td = document.createElement('td');
        var valeur = k < lignes[l].length ? raccourcir(lignes[l][k]) : '';
        td.textContent = valeur;
        if (valeur.length > LONGUEUR_VALEUR_COURTE) td.className = 'long';
        tr.appendChild(td);
      }
      corps.appendChild(tr);
    }
  }

  function afficherAvertissements() {
    var textes = [];
    var deuxPoints = INSECABLE + ': ';
    if (etat.plusieursFichiers) {
      textes.push('Vous avez déposé plusieurs fichiers' + deuxPoints + 'seul ' + guillemets(etat.nomFichier) + ' est utilisé.');
    }
    if (etat.encodageDouteux) {
      textes.push('Les accents de ce fichier ne sont peut-être pas bien reconnus (fichier créé par un ancien logiciel)' +
        deuxPoints + 'vérifiez-les dans le tableau.');
    }
    if (etat.encodagesMelanges) {
      textes.push("Une partie de ce fichier n'était pas enregistrée comme le reste" + deuxPoints +
        "l'outil a fait au mieux, mais vérifiez les accents une fois le fichier ouvert dans Excel.");
    }
    if (etat.illisibles) {
      textes.push('Quelques caractères de ce fichier sont abîmés' + deuxPoints + 'ils apparaissent sous la forme ' +
        guillemets('\uFFFD') + '.');
    }
    if (etat.finAbimee) {
      textes.push('La fin de ce fichier semble abîmée' + deuxPoints +
        'vérifiez sa dernière ligne une fois le fichier ouvert dans Excel.');
    }
    if (etat.colonnes > COLONNES_APERCU) {
      textes.push('Seules les ' + COLONNES_APERCU + ' premières colonnes sont montrées ici. Le fichier CSV les contiendra toutes.');
    }
    if (etat.lignes.length > C.LIMITES_EXCEL.lignes) {
      textes.push('Ce fichier a plus de ' + nombre(C.LIMITES_EXCEL.lignes) + ' lignes' + deuxPoints +
        "Excel n'affichera pas les dernières.");
    }
    if (etat.colonnes > C.LIMITES_EXCEL.colonnes) {
      textes.push('Ce fichier a plus de ' + nombre(C.LIMITES_EXCEL.colonnes) + ' colonnes' + deuxPoints +
        "Excel n'affichera pas les dernières.");
    }
    if (etat.longueurMax > C.LIMITES_EXCEL.caracteres) {
      textes.push('Une valeur de ce fichier dépasse ' + nombre(C.LIMITES_EXCEL.caracteres) + ' caractères' + deuxPoints +
        'Excel la coupera sur plusieurs lignes.');
    }
    el.avertissements.textContent = '';
    for (var i = 0; i < textes.length; i++) {
      var li = document.createElement('li');
      li.textContent = textes[i];
      el.avertissements.appendChild(li);
    }
    el.avertissements.hidden = textes.length === 0;
  }

  function afficherResultat(premierAffichage) {
    var lignes = etat.lignes;
    el.resume.textContent = 'Fichier ' + guillemets(etat.nomFichier) + INSECABLE + ': ' +
      pluriel(lignes.length, 'ligne') + ', ' + pluriel(etat.colonnes, 'colonne') + '.';
    el.separateur.value = etat.separateur;
    remplirTableau(lignes.slice(0, LIGNES_APERCU), etat.colonnes);

    if (lignes.length > LIGNES_APERCU) {
      el.legende.textContent = 'Les ' + LIGNES_APERCU + ' premières lignes de votre fichier';
    } else if (lignes.length === 1) {
      el.legende.textContent = 'La ligne de votre fichier';
    } else {
      el.legende.textContent = 'Les ' + lignes.length + ' lignes de votre fichier';
    }
    var reste = lignes.length - LIGNES_APERCU;
    el.suite.textContent = reste > 0
      ? '… et ' + pluriel(reste, 'autre ligne', 'autres lignes') + (reste > 1 ? ', qui seront' : ', qui sera') +
        (etat.outil === 'tcd' ? ' aussi prises en compte dans le tableau croisé.' : ' aussi dans le fichier CSV.')
      : '';
    el.suite.hidden = reste <= 0;

    afficherAvertissements();
    el.encodage.textContent = (NOMS_ENCODAGES[etat.encodage] || etat.encodage) + (etat.bom ? ', avec BOM' : '');
    el.lignesVides.textContent = nombre(etat.lignesVides);

    el.sectionCsv.hidden = etat.outil !== 'csv';
    el.sectionTcd.hidden = etat.outil !== 'tcd';
    // Outil « tableau croisé » : la proposition automatique est faite tout de suite.
    if (premierAffichage && etat.outil === 'tcd') preparerTcd();
    afficherEtape('resultat');
    mettreAJourAideDefilement();
    if (premierAffichage) {
      // Le tableau doit se voir tout de suite, même sur un petit écran.
      el.titreApercu.focus({ preventScroll: true });
      if (el.titreApercu.scrollIntoView) el.titreApercu.scrollIntoView({ block: 'start' });
    }
  }

  // Signale un tableau plus large que l'écran (fréquent sur téléphone).
  function mettreAJourAideDefilement() {
    var deborde = Boolean(etat.lignes) && !el.etapeResultat.hidden &&
      el.cadreTableau.scrollWidth > el.cadreTableau.clientWidth + 1;
    el.aideDefilement.hidden = !deborde;
    if (deborde) el.cadreTableau.setAttribute('aria-describedby', 'aide-defilement');
    else el.cadreTableau.removeAttribute('aria-describedby');
  }

  function changerSeparateur() {
    if (!etat.texte) return;
    var separateur = el.separateur.value;
    var long = etat.texte.length > TEXTE_LONG;
    var appliquer = function () {
      if (etat.recalculEnAttente !== appliquer) return; // annulé, remplacé ou déjà fait
      etat.recalculEnAttente = null;
      el.cadreTableau.removeAttribute('aria-busy');
      var resultat = C.decouper(etat.texte, separateur);
      if (resultat.lignes.length > 0) {
        etat.separateur = separateur;
        etat.lignes = resultat.lignes;
        etat.colonnes = resultat.colonnes;
        etat.lignesVides = resultat.lignesVides;
        etat.longueurMax = resultat.longueurMax;
      }
      afficherResultat(false);
      // Les colonnes ont changé : le tableau croisé repart d'une nouvelle proposition.
      if (etat.tcd) preparerTcd();
      annoncer('Tableau mis à jour' + INSECABLE + ': ' + pluriel(etat.colonnes, 'colonne') + '.');
    };
    etat.recalculEnAttente = appliquer;
    if (long) {
      // Gros fichier : on montre que le calcul est en cours avant de le lancer.
      el.resume.textContent = 'Mise à jour du tableau…';
      el.cadreTableau.setAttribute('aria-busy', 'true');
      setTimeout(appliquer, 20);
    } else {
      appliquer();
    }
  }

  /* ---- Téléchargement -------------------------------------------------- */

  function formatChoisi() {
    var choisi = document.querySelector('input[name="format"]:checked');
    return (choisi && choisi.value === 'standard') ? 'standard' : 'excel';
  }

  function telecharger() {
    if (etat.recalculEnAttente) etat.recalculEnAttente(); // un changement de séparateur pas encore appliqué
    if (!etat.lignes) return;
    var format = formatChoisi();
    var csv = C.genererCsv(etat.lignes, format);
    var fichier = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    csv = null;
    enregistrer(fichier, etat.nomCsv, format);
  }

  // Fait télécharger le fichier créé (CSV ou Excel), puis efface les données
  // de la page et affiche l'écran « C'est fait ». sorte : excel, standard ou tcd.
  function enregistrer(fichier, nom, sorte) {
    libererFichierGenere();
    etat.url = URL.createObjectURL(fichier);
    fichier = null;
    var lien = document.createElement('a');
    lien.href = etat.url;
    lien.download = nom;
    lien.hidden = true;
    document.body.appendChild(lien);
    lien.click();
    document.body.removeChild(lien);
    etat.minuterie = setTimeout(libererFichierGenere, DELAI_LIBERATION);

    effacerDonnees();
    el.nomTelecharge.textContent = nom;
    // Le navigateur numérote le fichier s'il en existe déjà un du même nom.
    var point = nom.lastIndexOf('.');
    el.nomNumerote.textContent = nom.slice(0, point) + ' (1)' + nom.slice(point);
    el.ouvrirExcel.hidden = sorte !== 'excel';
    el.ouvrirStandard.hidden = sorte !== 'standard';
    el.ouvrirTcd.hidden = sorte !== 'tcd';
    afficherEtape('termine');
    etat.termineA = Date.now();
    el.titreTermine.focus();
    annoncer('Le fichier ' + nom + ' a été téléchargé.');
  }

  /* ---- Tableau croisé dynamique ---------------------------------------- */

  // Oublie le tableau croisé : données préparées, et listes de choix (elles
  // contiennent les noms des colonnes du fichier).
  function viderTcd() {
    etat.tcd = null;
    if (!el.panneauTcd) return;
    var listes = [el.tcdLignes1, el.tcdLignes2, el.tcdColonnes, el.tcdFiltre].concat(el.tcdCalculs, el.tcdValeurs);
    for (var i = 0; i < listes.length; i++) listes[i].textContent = '';
    el.tcdFiltreListe.textContent = '';
    el.tcdFiltreValeurs.hidden = true;
    el.tcdFiltreTrop.hidden = true;
    el.tableauTcd.tHead.textContent = '';
    el.tableauTcd.tBodies[0].textContent = '';
    el.resumeTcd.textContent = '';
    el.suiteTcd.textContent = '';
    el.suiteTcd.hidden = true;
    masquerErreurTcd();
  }

  function montrerErreurTcd(message) {
    el.erreurTcd.textContent = message;
    el.erreurTcd.hidden = false;
  }

  function masquerErreurTcd() {
    el.erreurTcd.hidden = true;
    el.erreurTcd.textContent = '';
  }

  function ajouterOption(liste, valeur, texte) {
    var option = document.createElement('option');
    option.value = valeur;
    option.textContent = texte;
    liste.appendChild(option);
  }

  function remplirListeChamps(liste, noms, texteAucun) {
    liste.textContent = '';
    if (texteAucun) ajouterOption(liste, SANS_CHAMP, texteAucun);
    for (var j = 0; j < noms.length; j++) ajouterOption(liste, String(j), noms[j]);
  }

  // Prépare les données et applique la proposition automatique.
  // entete : true / false si la personne a coché ou décoché la case des titres.
  function preparerTcd(entete) {
    var donnees = T.preparer(etat.lignes, typeof entete === 'boolean' ? { entete: entete } : undefined);
    etat.tcd = { donnees: donnees, reglages: T.proposer(donnees), propose: true, elementsFiltre: null, enregistrements: 0, precedents: {} };
    el.tcdEntete.checked = donnees.entete;
    remplirChoixTcd();
    calculerTcd();
  }

  function remplirChoixTcd() {
    var noms = etat.tcd.donnees.noms;
    var r = etat.tcd.reglages;
    remplirListeChamps(el.tcdLignes1, noms, '');
    remplirListeChamps(el.tcdLignes2, noms, '(rien de plus)');
    remplirListeChamps(el.tcdColonnes, noms, '(aucune)');
    remplirListeChamps(el.tcdFiltre, noms, '(aucun)');
    for (var k = 0; k < el.tcdValeurs.length; k++) {
      remplirListeChamps(el.tcdValeurs[k], noms, k > 0 ? '(aucune)' : '');
      el.tcdCalculs[k].textContent = '';
      for (var c = 0; c < T.CALCULS.length; c++) ajouterOption(el.tcdCalculs[k], T.CALCULS[c], TEXTES_CALCULS[T.CALCULS[c]]);
      var v = r.valeurs[k];
      el.tcdValeurs[k].value = v ? String(v.champ) : SANS_CHAMP;
      el.tcdCalculs[k].value = v ? v.calcul : 'somme';
      el.tcdCalculs[k].disabled = !v;
    }
    el.tcdLignes1.value = String(r.lignes[0]);
    el.tcdLignes2.value = r.lignes.length > 1 ? String(r.lignes[1]) : SANS_CHAMP;
    el.tcdColonnes.value = r.colonnes.length ? String(r.colonnes[0]) : SANS_CHAMP;
    el.tcdFiltre.value = String(r.filtre.champ);
    el.tcdAffichage.value = r.affichage;
    el.tcdTri.value = r.tri;
    remplirFiltre();
  }

  // Réglages choisis dans la page.
  function lireReglagesTcd() {
    var r = {
      lignes: [Number(el.tcdLignes1.value)],
      colonnes: [],
      valeurs: [],
      affichage: el.tcdAffichage.value,
      tri: el.tcdTri.value,
      filtre: { champ: Number(el.tcdFiltre.value), exclus: etat.tcd.reglages.filtre.exclus }
    };
    if (el.tcdLignes2.value !== SANS_CHAMP) r.lignes.push(Number(el.tcdLignes2.value));
    if (el.tcdColonnes.value !== SANS_CHAMP) r.colonnes.push(Number(el.tcdColonnes.value));
    for (var k = 0; k < el.tcdValeurs.length; k++) {
      if (el.tcdValeurs[k].value !== SANS_CHAMP) {
        r.valeurs.push({ champ: Number(el.tcdValeurs[k].value), calcul: el.tcdCalculs[k].value });
      }
    }
    return T.normaliserReglages(etat.tcd.donnees, r);
  }

  // Liste des valeurs du filtre, à cocher ou décocher.
  function remplirFiltre() {
    var champ = Number(el.tcdFiltre.value);
    etat.tcd.filtreAffiche = champ;
    el.tcdFiltreListe.textContent = '';
    etat.tcd.elementsFiltre = null;
    el.tcdFiltreValeurs.hidden = true;
    el.tcdFiltreTrop.hidden = true;
    if (champ < 0) return;
    var elements = T.elementsDistincts(etat.tcd.donnees, champ, MAX_ELEMENTS_FILTRE);
    if (!elements) {
      el.tcdFiltreTrop.hidden = false;
      return;
    }
    etat.tcd.elementsFiltre = elements;
    var exclus = etat.tcd.reglages.filtre.exclus;
    for (var i = 0; i < elements.length; i++) {
      var etiquette = document.createElement('label');
      etiquette.className = 'case';
      var caseACocher = document.createElement('input');
      caseACocher.type = 'checkbox';
      caseACocher.value = String(i);
      caseACocher.checked = !exclus[elements[i].cle];
      var texte = document.createElement('span');
      texte.textContent = elements[i].libelle;
      etiquette.appendChild(caseACocher);
      etiquette.appendChild(texte);
      el.tcdFiltreListe.appendChild(etiquette);
    }
    el.tcdFiltreValeurs.hidden = false;
  }

  function changerCaseFiltre(e) {
    var elements = etat.tcd && etat.tcd.elementsFiltre;
    if (!elements || !e.target || e.target.type !== 'checkbox') return;
    var element = elements[Number(e.target.value)];
    if (!element) return;
    if (e.target.checked) delete etat.tcd.reglages.filtre.exclus[element.cle];
    else etat.tcd.reglages.filtre.exclus[element.cle] = true;
    calculerTcd();
  }

  function cocherTout(coche) {
    if (!etat.tcd || !etat.tcd.elementsFiltre) return;
    var elements = etat.tcd.elementsFiltre;
    var exclus = {};
    if (!coche) {
      for (var i = 0; i < elements.length; i++) exclus[elements[i].cle] = true;
    }
    etat.tcd.reglages.filtre.exclus = exclus;
    var cases = el.tcdFiltreListe.querySelectorAll('input');
    for (var k = 0; k < cases.length; k++) cases[k].checked = coche;
    calculerTcd();
  }

  // Une colonne ne peut servir qu'à un endroit (lignes, colonnes ou filtre) :
  // la choisir ailleurs la retire de son ancienne place.
  function retenirChoix(e) {
    if (etat.tcd) etat.tcd.precedents[e.target.id] = e.target.value;
  }

  function changerPlace(e) {
    if (!etat.tcd) return;
    var liste = e.target;
    var choix = liste.value;
    var places = [el.tcdLignes1, el.tcdLignes2, el.tcdColonnes, el.tcdFiltre];
    if (choix !== SANS_CHAMP) {
      for (var i = 0; i < places.length; i++) {
        if (places[i] === liste || places[i].value !== choix) continue;
        if (places[i] === el.tcdLignes1) {
          // Les lignes ne restent jamais vides : elles reprennent l'ancien choix.
          var ancien = etat.tcd.precedents[liste.id];
          places[i].value = ancien !== undefined && ancien !== SANS_CHAMP && ancien !== choix ? ancien : premierChampLibre(places);
        } else {
          places[i].value = SANS_CHAMP;
        }
      }
    }
    // Nouvelle colonne de filtre : toutes ses valeurs sont gardées au départ.
    if (Number(el.tcdFiltre.value) !== etat.tcd.filtreAffiche) {
      etat.tcd.reglages.filtre.exclus = {};
      remplirFiltre();
    }
    etat.tcd.precedents[liste.id] = liste.value;
    calculerTcd();
  }

  function premierChampLibre(places) {
    var pris = {};
    for (var i = 0; i < places.length; i++) pris[places[i].value] = true;
    for (var j = 0; j < etat.tcd.donnees.noms.length; j++) {
      if (!pris[String(j)]) return String(j);
    }
    return '0';
  }

  // Une nouvelle colonne de valeurs : somme pour des nombres, sinon nombre de lignes.
  function changerValeur(k) {
    if (!etat.tcd) return;
    var choix = el.tcdValeurs[k].value;
    el.tcdCalculs[k].disabled = choix === SANS_CHAMP;
    if (choix !== SANS_CHAMP) {
      el.tcdCalculs[k].value = etat.tcd.donnees.types[Number(choix)] === 'nombre' ? 'somme' : 'nombre';
    }
    calculerTcd();
  }

  function calculerTcd() {
    if (!etat.tcd) return;
    masquerErreurTcd();
    var r = lireReglagesTcd();
    etat.tcd.reglages = r;
    var resultat = T.calculer(etat.tcd.donnees, r);
    etat.tcd.enregistrements = resultat.enregistrements;
    afficherTcd(resultat);
    etat.tcd.propose = false;
  }

  function texteListe(mots) {
    if (mots.length <= 1) return mots.join('');
    return mots.slice(0, -1).join(', ') + ' et ' + mots[mots.length - 1];
  }

  function guillemetsTous(noms) {
    var sortie = [];
    for (var i = 0; i < noms.length; i++) sortie.push(guillemets(noms[i]));
    return sortie;
  }

  function resumerTcd(resultat) {
    var d = etat.tcd.donnees;
    var noms = [];
    for (var x = 0; x < resultat.valeurs.length; x++) noms.push(resultat.valeurs[x].nom);
    var phrase = (etat.tcd.propose ? 'Proposition automatique' + INSECABLE + ': ' : 'Tableau' + INSECABLE + ': ') +
      texteListe(guillemetsTous(noms)) + ' par ' + guillemetsTous(resultat.nomsLignes).join(' puis par ') +
      (resultat.nomColonnes ? ', avec ' + guillemets(resultat.nomColonnes) + ' en colonnes' : '') + '. ' +
      pluriel(resultat.enregistrements, 'ligne') + ' du fichier ' + (resultat.enregistrements > 1 ? 'sont comptées' : 'est comptée') + '.';
    var notes = [];
    if (d.lignesTitre) notes.push(pluriel(d.lignesTitre, 'ligne') + ' de titre au-dessus du tableau');
    if (d.lignesTotal) notes.push(pluriel(d.lignesTotal, 'ligne') + ' de total (pour ne pas compter deux fois)');
    if (notes.length) phrase += ' Ignorées' + INSECABLE + ': ' + texteListe(notes) + '.';
    if (d.tronque) phrase += ' Le fichier dépasse la taille d’une feuille Excel' + INSECABLE + ': les dernières lignes ne sont pas reprises.';
    if (etat.tcd.propose) phrase += ' Modifiez les choix ci-dessous si besoin.';
    return phrase;
  }

  function celluleTcd(ligne, texte, classe) {
    var td = document.createElement('td');
    td.textContent = texte;
    if (classe) td.className = classe;
    ligne.appendChild(td);
  }

  function afficherTcd(resultat) {
    el.resumeTcd.textContent = resumerTcd(resultat);
    var tete = el.tableauTcd.tHead;
    var corps = el.tableauTcd.tBodies[0];
    tete.textContent = '';
    corps.textContent = '';
    var valeurs = resultat.valeurs;
    var plusieurs = valeurs.length > 1;
    var x;

    var ligneTete = document.createElement('tr');
    var coin = document.createElement('th');
    coin.scope = 'col';
    coin.textContent = resultat.nomsLignes.join(' / ');
    ligneTete.appendChild(coin);
    function titreColonne(texte) {
      var th = document.createElement('th');
      th.scope = 'col';
      th.textContent = texte;
      ligneTete.appendChild(th);
    }
    for (var c = 0; c < resultat.colonnes.length; c++) {
      for (x = 0; x < valeurs.length; x++) {
        titreColonne(resultat.colonnes[c].libelle + (plusieurs ? ' – ' + valeurs[x].nom : ''));
      }
    }
    for (x = 0; x < valeurs.length; x++) {
      titreColonne(resultat.colonnes.length ? 'Total' + (plusieurs ? ' – ' + valeurs[x].nom : '') : valeurs[x].nom);
    }
    tete.appendChild(ligneTete);

    function ligneCorps(ligne, classe) {
      var tr = document.createElement('tr');
      if (classe) tr.className = classe;
      var th = document.createElement('th');
      th.scope = 'row';
      th.textContent = ligne.libelle;
      tr.appendChild(th);
      for (var c = 0; c < ligne.cellules.length; c++) {
        for (var x = 0; x < valeurs.length; x++) celluleTcd(tr, T.formaterValeur(ligne.cellules[c][x], valeurs[x].format), 'nombre');
      }
      for (x = 0; x < valeurs.length; x++) celluleTcd(tr, T.formaterValeur(ligne.totaux[x], valeurs[x].format), 'nombre total-ligne');
      corps.appendChild(tr);
    }

    if (resultat.enregistrements === 0) {
      var vide = document.createElement('tr');
      celluleTcd(vide, 'Aucune ligne ne correspond au filtre' + INSECABLE + ': cochez au moins une valeur.', 'vide');
      vide.firstChild.colSpan = ligneTete.childNodes.length;
      corps.appendChild(vide);
    } else {
      var lignes = resultat.lignes;
      var nb = Math.min(lignes.length, LIGNES_APERCU_TCD);
      for (var i = 0; i < nb; i++) {
        ligneCorps(lignes[i], lignes[i].niveau ? 'niveau-2' : (lignes[i].groupe ? 'groupe' : ''));
      }
      ligneCorps(resultat.total, 'total');
    }
    var reste = resultat.lignes.length - LIGNES_APERCU_TCD;
    el.suiteTcd.textContent = reste > 0
      ? '… et ' + pluriel(reste, 'autre ligne', 'autres lignes') + ' dans le fichier Excel.'
      : '';
    el.suiteTcd.hidden = reste <= 0;
  }

  function nomFichierXlsx() {
    return etat.nomCsv.replace(/\.csv$/i, '') + ' - tableau croisé.xlsx';
  }

  function telechargerTcd() {
    if (etat.recalculEnAttente) etat.recalculEnAttente(); // un changement de séparateur pas encore appliqué
    if (!etat.tcd) return;
    calculerTcd();
    if (!etat.tcd.enregistrements) {
      montrerErreurTcd('Aucune ligne ne correspond au filtre' + INSECABLE + ': cochez au moins une valeur, puis réessayez.');
      el.erreurTcd.scrollIntoView({ block: 'center' });
      return;
    }
    var morceaux;
    try {
      morceaux = T.genererXlsx(etat.tcd.donnees, etat.tcd.reglages);
    } catch (e) {
      montrerErreurTcd("Le fichier Excel n'a pas pu être créé (le fichier est peut-être trop gros pour ce navigateur). " +
        'Essayez plutôt le fichier CSV.');
      return;
    }
    var fichier = new Blob(morceaux, { type: TYPE_XLSX });
    morceaux = null;
    enregistrer(fichier, nomFichierXlsx(), 'tcd');
  }

  function activerTcd() {
    if (!T || !el.panneauTcd) {
      if (el.sectionTcd) el.sectionTcd.hidden = true;
      return;
    }
    el.boutonTelechargerTcd.addEventListener('click', clicSimple(telechargerTcd));
    el.tcdEntete.addEventListener('change', function () {
      if (etat.lignes) preparerTcd(el.tcdEntete.checked);
    });
    var places = [el.tcdLignes1, el.tcdLignes2, el.tcdColonnes, el.tcdFiltre];
    for (var i = 0; i < places.length; i++) {
      places[i].addEventListener('focus', retenirChoix);
      places[i].addEventListener('change', changerPlace);
    }
    el.tcdValeurs.forEach(function (liste, k) {
      liste.addEventListener('change', function () { changerValeur(k); });
    });
    var autres = el.tcdCalculs.concat([el.tcdAffichage, el.tcdTri]);
    for (i = 0; i < autres.length; i++) autres[i].addEventListener('change', calculerTcd);
    el.tcdFiltreListe.addEventListener('change', changerCaseFiltre);
    el.tcdToutCocher.addEventListener('click', function () { cocherTout(true); });
    el.tcdToutDecocher.addEventListener('click', function () { cocherTout(false); });
  }

  /* ---- Glisser-déposer ------------------------------------------------- */

  var survols = 0;

  function contientFichiers(evenement) {
    var types = evenement.dataTransfer && evenement.dataTransfer.types;
    if (!types) return false;
    for (var i = 0; i < types.length; i++) {
      if (types[i] === 'Files') return true;
    }
    return false;
  }

  // Un dossier lâché sur la page arrive comme un fichier vide : on le
  // reconnaît pour afficher un message clair.
  function estUnDossier(transfert) {
    var objets = transfert.items;
    if (!objets) return false;
    for (var i = 0; i < objets.length; i++) {
      if (objets[i].kind !== 'file') continue;
      var entree = objets[i].webkitGetAsEntry ? objets[i].webkitGetAsEntry() : null;
      return Boolean(entree && entree.isDirectory);
    }
    return false;
  }

  function finSurvol() {
    survols = 0;
    document.body.classList.remove('glisser');
  }

  function activerGlisserDeposer() {
    // Toute la page accepte le fichier : s'il est lâché à côté du cadre, le
    // navigateur ne doit pas l'ouvrir à la place de l'outil.
    window.addEventListener('dragenter', function (e) {
      if (!contientFichiers(e)) return;
      e.preventDefault();
      survols++;
      document.body.classList.add('glisser');
    });
    window.addEventListener('dragover', function (e) {
      if (!contientFichiers(e)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
    });
    window.addEventListener('dragleave', function (e) {
      if (!contientFichiers(e)) return;
      survols--;
      if (survols <= 0) finSurvol();
    });
    window.addEventListener('drop', function (e) {
      if (!contientFichiers(e)) return;
      e.preventDefault();
      finSurvol();
      if (etat.horsService) return;
      if (estUnDossier(e.dataTransfer)) {
        reinitialiser(false);
        afficherErreur('dossier', true);
        return;
      }
      var fichiers = e.dataTransfer.files;
      if (fichiers && fichiers.length > 0) chargerFichier(fichiers[0], fichiers.length > 1);
    });
  }

  /* ---- Démarrage ------------------------------------------------------- */

  function navigateurCompatible() {
    return Boolean(window.FileReader && window.Blob && window.URL && URL.createObjectURL &&
      URL.revokeObjectURL && window.TextDecoder && window.Uint8Array && window.Uint16Array);
  }

  // Certains appareils anciens (iPad restés sous iOS 12) ne savent pas
  // enregistrer un fichier créé par la page.
  function enregistrementPossible() {
    return 'download' in document.createElement('a');
  }

  // Un double-clic compte pour un seul clic : le second tomberait sur le
  // bouton apparu entre-temps sous la souris.
  function clicSimple(action) {
    return function (e) {
      if (e.detail > 1) return;
      action(e);
    };
  }

  function demarrer() {
    if (!C || !navigateurCompatible() || !enregistrementPossible()) {
      etat.horsService = true;
      el.chargement.hidden = true;
      occupe(false);
      afficherErreur(!C ? 'demarrage' : (navigateurCompatible() ? 'enregistrement' : 'navigateur'), false);
      return;
    }

    el.boutonChoisir.addEventListener('click', clicSimple(function () {
      el.champFichier.click();
    }));
    // Un clic n'importe où dans le cadre ouvre aussi le choix du fichier.
    el.zone.addEventListener('click', clicSimple(function (e) {
      if (e.target === el.boutonChoisir || e.target === el.champFichier || el.boutonChoisir.disabled) return;
      el.champFichier.click();
    }));
    el.champFichier.addEventListener('change', function () {
      var fichiers = el.champFichier.files;
      if (fichiers && fichiers.length > 0) chargerFichier(fichiers[0], fichiers.length > 1);
    });
    el.separateur.addEventListener('change', changerSeparateur);
    el.boutonTelecharger.addEventListener('click', clicSimple(telecharger));

    var boutonsRecommencer = document.querySelectorAll('[data-action="recommencer"]');
    for (var i = 0; i < boutonsRecommencer.length; i++) {
      boutonsRecommencer[i].addEventListener('click', clicSimple(function (e) {
        // Le clavier (detail = 0) n'est jamais concerné.
        if (e.detail === 1 && !el.etapeTermine.hidden && Date.now() - etat.termineA < DELAI_DOUBLE_CLIC) return;
        reinitialiser(true);
      }));
    }

    for (i = 0; i < el.cartesOutils.length; i++) {
      el.cartesOutils[i].addEventListener('click', clicSimple(function (e) {
        choisirOutil(e.currentTarget.getAttribute('data-outil'));
      }));
    }
    var boutonsChanger = document.querySelectorAll('[data-action="changer-outil"]');
    for (i = 0; i < boutonsChanger.length; i++) {
      boutonsChanger[i].addEventListener('click', clicSimple(function (e) {
        if (e.detail === 1 && !el.etapeTermine.hidden && Date.now() - etat.termineA < DELAI_DOUBLE_CLIC) return;
        changerOutil();
      }));
    }

    activerTcd();
    activerGlisserDeposer();
    window.addEventListener('resize', mettreAJourAideDefilement);

    // En quittant la page, on ne laisse rien derrière soi.
    window.addEventListener('pagehide', function () {
      reinitialiser(false);
    });

    // Tout est prêt : le bouton, désactivé dans la page tant que l'outil n'a
    // pas démarré, devient utilisable. Si une erreur survient avant ce point,
    // le message « Chargement… Si le bouton reste gris, rechargez la page »
    // reste affiché.
    occupe(false);
    el.chargement.hidden = true;
  }

  demarrer();
}());
