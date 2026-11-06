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

  var INSECABLE = '\u00A0'; // espace insécable (avant « : », « ? »… et dans « »)

  var MESSAGES = {
    'vide': 'Ce fichier est vide. Choisissez un autre fichier.',
    'lignes-vides': 'Ce fichier ne contient que des lignes vides. Choisissez un autre fichier.',
    'binaire': "Ce fichier n'est pas un fichier texte. Choisissez un fichier dont le nom se termine par .txt.",
    'rtf': "Ce fichier est un document de traitement de texte (Word ou WordPad), pas un simple fichier texte. " +
      "Enregistrez-le d'abord au format " + guillemets('Texte (.txt)') + ', puis réessayez.',
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
    annonce: $('annonce')
  };

  // Tout ce que la page sait du fichier en cours. Vidé après le téléchargement.
  var etat = {
    lecture: 0,          // numéro de la lecture en cours (une lecture annulée est ignorée)
    horsService: false,  // navigateur incompatible
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
    url: null,           // adresse temporaire du fichier CSV généré
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
    el.etapeChoix.hidden = nom !== 'choix';
    el.etapeResultat.hidden = nom !== 'resultat';
    el.etapeTermine.hidden = nom !== 'termine';
  }

  function occupe(actif) {
    el.boutonChoisir.disabled = actif || etat.horsService;
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
    if (deplacerFocus) el.boutonChoisir.focus();
  }

  /* ---- Lecture du fichier ---------------------------------------------- */

  function chargerFichier(fichier, plusieurs) {
    reinitialiser(false);
    if (!fichier) return;
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
        ' aussi dans le fichier CSV.'
      : '';
    el.suite.hidden = reste <= 0;

    afficherAvertissements();
    el.encodage.textContent = (NOMS_ENCODAGES[etat.encodage] || etat.encodage) + (etat.bom ? ', avec BOM' : '');
    el.lignesVides.textContent = nombre(etat.lignesVides);

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
    var nom = etat.nomCsv;
    var format = formatChoisi();
    var csv = C.genererCsv(etat.lignes, format);
    var fichier = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    csv = null;

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
    el.nomNumerote.textContent = nom.slice(0, -'.csv'.length) + ' (1).csv';
    el.ouvrirExcel.hidden = format !== 'excel';
    el.ouvrirStandard.hidden = format !== 'standard';
    afficherEtape('termine');
    etat.termineA = Date.now();
    el.titreTermine.focus();
    annoncer('Le fichier ' + nom + ' a été téléchargé.');
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
