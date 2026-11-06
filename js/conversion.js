/*
 * conversion.js — Convertisseur TXT vers CSV
 *
 * Toute la logique de conversion, sous forme de fonctions pures : elles
 * reçoivent des octets ou du texte et renvoient un résultat, sans rien lire
 * ni écrire ailleurs. Aucun accès au réseau, aucun enregistrement.
 *
 * Utilisable dans le navigateur (objet global « ConversionCSV ») comme dans
 * Node.js (require), ce qui permet de tester la logique avec node:test.
 */
(function (racine, fabrique) {
  'use strict';
  var api = fabrique();
  if (typeof module === 'object' && module !== null && module.exports) {
    module.exports = api;
  } else {
    racine.ConversionCSV = api;
  }
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  /* ------------------------------------------------------------------ */
  /* Réglages                                                           */
  /* ------------------------------------------------------------------ */

  // Séparateurs essayés par la détection automatique. En cas de quasi-égalité,
  // le premier de la liste l'emporte : le point-virgule passe avant la virgule,
  // car en français la virgule sert aussi pour les nombres décimaux.
  var SEPARATEURS_DETECTABLES = ['tabulation', 'point-virgule', 'barre', 'virgule', 'espaces'];

  // Tous les choix possibles (« aucun » : chaque ligne forme une seule colonne).
  var SEPARATEURS = SEPARATEURS_DETECTABLES.concat(['aucun']);

  var CARACTERES = { 'tabulation': '\t', 'point-virgule': ';', 'barre': '|', 'virgule': ',' };

  // Séparateurs fréquents dans un texte ordinaire : la détection est plus exigeante.
  var SEPARATEURS_FAIBLES = { 'virgule': true, 'espaces': true };

  // Formats de sortie.
  var FORMATS = {
    // Excel en français : point-virgule, UTF-8 avec BOM, et valeurs écrites de
    // façon qu'Excel les affiche telles quelles (voir aFigerPourExcel).
    excel: { separateur: ';', bom: true, figer: true },
    // CSV standard (RFC 4180) : virgule, UTF-8 sans BOM, valeurs inchangées.
    standard: { separateur: ',', bom: false, figer: false }
  };

  // Limites d'Excel. En ouvrant un CSV, Excel garde au plus 32 759 caractères
  // par case (mesuré dans Excel 16) et met la suite sur une nouvelle ligne.
  var LIMITES_EXCEL = { lignes: 1048576, colonnes: 16384, caracteres: 32759 };

  var LIGNES_ECHANTILLON = 200;  // lignes examinées pour deviner le séparateur
  var SEUIL_DETECTION = 0.5;     // score minimal pour retenir un séparateur
  var MARGE_EGALITE = 0.1;       // deux scores aussi proches sont considérés égaux
  var LARGEUR_TABULATION = 8;    // une tabulation avance jusqu'à la colonne multiple de 8 suivante

  /* ------------------------------------------------------------------ */
  /* Lecture des octets : encodage et nature du fichier                 */
  /* ------------------------------------------------------------------ */

  // Windows-1252 : caractères des octets 0x80 à 0x9F. Les autres octets
  // correspondent directement au même code Unicode (comme en Latin-1).
  // Windows-1252 n'utilise pas 0x81, 0x8D, 0x8F, 0x90 et 0x9D : les trouver
  // trahit un autre encodage (ancien DOS, ancien Mac…). Ils deviennent « � »
  // (0xFFFD) au lieu de caractères invisibles.
  var WINDOWS_1252 = [
    0x20AC, 0xFFFD, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021,
    0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0xFFFD, 0x017D, 0xFFFD,
    0xFFFD, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
    0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0xFFFD, 0x017E, 0x0178
  ];

  // Débuts de fichiers qui ne sont pas du texte.
  var SIGNATURES_BINAIRES = [
    [0x50, 0x4B, 0x03, 0x04], // ZIP : .docx, .xlsx, .odt, .zip…
    [0x50, 0x4B, 0x05, 0x06], // ZIP vide
    [0x25, 0x50, 0x44, 0x46], // PDF
    [0xD0, 0xCF, 0x11, 0xE0], // anciens .doc et .xls
    [0x89, 0x50, 0x4E, 0x47], // PNG
    [0xFF, 0xD8, 0xFF],       // JPEG
    [0x47, 0x49, 0x46, 0x38], // GIF
    [0x1F, 0x8B],             // GZIP
    [0x37, 0x7A, 0xBC, 0xAF, 0x27, 0x1C], // 7-Zip
    [0x52, 0x61, 0x72, 0x21, 0x1A, 0x07], // RAR
    [0x49, 0x49, 0x2A, 0x00], // TIFF (image numérisée)
    [0x4D, 0x4D, 0x00, 0x2A], // TIFF
    [0x00, 0x01, 0x00, 0x00, 0x00], // police de caractères TrueType
    [0x00, 0x00, 0x01, 0x00], // icône (.ico)
    [0x00, 0x00, 0x02, 0x00]  // curseur de souris (.cur)
  ];
  var SIGNATURE_RTF = [0x7B, 0x5C, 0x72, 0x74, 0x66]; // « {\rtf » : document WordPad ou Word

  function enOctets(donnees) {
    if (donnees instanceof Uint8Array) return donnees;
    if (donnees instanceof ArrayBuffer) return new Uint8Array(donnees);
    if (donnees && donnees.buffer instanceof ArrayBuffer) {
      return new Uint8Array(donnees.buffer, donnees.byteOffset, donnees.byteLength);
    }
    throw new TypeError('Octets attendus (Uint8Array ou ArrayBuffer).');
  }

  function commencePar(octets, debut, signature) {
    if (octets.length - debut < signature.length) return false;
    for (var i = 0; i < signature.length; i++) {
      if (octets[debut + i] !== signature[i]) return false;
    }
    return true;
  }

  function aSignatureBinaire(octets) {
    for (var s = 0; s < SIGNATURES_BINAIRES.length; s++) {
      if (commencePar(octets, 0, SIGNATURES_BINAIRES[s])) return true;
    }
    // Son WAV, vidéo AVI, image WEBP : « RIFF », la taille, puis le type.
    if (commencePar(octets, 0, [0x52, 0x49, 0x46, 0x46]) &&
        (commencePar(octets, 8, [0x57, 0x41, 0x56, 0x45]) ||
         commencePar(octets, 8, [0x41, 0x56, 0x49, 0x20]) ||
         commencePar(octets, 8, [0x57, 0x45, 0x42, 0x50]))) {
      return true;
    }
    // Vidéo MP4 ou MOV, son M4A, photo HEIC d'iPhone : la taille, puis « ftyp ».
    if (octets.length > 8 && octets[0] === 0x00 && commencePar(octets, 4, [0x66, 0x74, 0x79, 0x70])) return true;
    // Image BMP : « BM », la taille, puis quatre octets réservés, toujours nuls.
    if (commencePar(octets, 0, [0x42, 0x4D]) && commencePar(octets, 6, [0x00, 0x00, 0x00, 0x00])) return true;
    return false;
  }

  // Compte, à partir de « debut », les caractères UTF-8 valides de plusieurs
  // octets et les octets invalides (règles strictes de la norme Unicode).
  // « autres » compte les caractères valides qui ne sont pas des caractères
  // courants d'un texte français (lettres accentuées, guillemets, €…).
  function analyserUtf8(octets, debut) {
    var valides = 0;
    var invalides = 0;
    var autres = 0;
    var n = octets.length;
    var i = debut;
    while (i < n) {
      if (octets[i] < 0x80) { i++; continue; }
      var longueur = longueurUtf8(octets, i);
      if (longueur > 0) {
        valides++;
        if (!estCaractereFrancais(pointUtf8(octets, i, longueur))) autres++;
        i += longueur;
      } else {
        invalides++;
        i++;
      }
    }
    return { valides: valides, invalides: invalides, autres: autres };
  }

  // Code Unicode du caractère UTF-8 valide de « longueur » octets qui commence en i.
  function pointUtf8(octets, i, longueur) {
    var b = octets[i];
    var point = longueur === 2 ? (b & 0x1F) : longueur === 3 ? (b & 0x0F) : (b & 0x07);
    for (var j = 1; j < longueur; j++) point = (point << 6) | (octets[i + j] & 0x3F);
    return point;
  }

  // Lettres accentuées, œ, Ÿ, tirets, guillemets et apostrophes typographiques, points de suspension, €.
  function estCaractereFrancais(point) {
    return (point >= 0xA0 && point <= 0xFF) || point === 0x152 || point === 0x153 || point === 0x178 ||
      (point >= 0x2013 && point <= 0x2026) || point === 0x20AC;
  }

  // Longueur du caractère UTF-8 de plusieurs octets qui commence en i, ou 0
  // si les octets ne forment pas un caractère valide.
  function longueurUtf8(octets, i) {
    var b = octets[i];
    var longueur = 0;
    var min = 0x80;
    var max = 0xBF;
    if (b >= 0xC2 && b <= 0xDF) {
      longueur = 2;
    } else if (b >= 0xE0 && b <= 0xEF) {
      longueur = 3;
      if (b === 0xE0) min = 0xA0;
      if (b === 0xED) max = 0x9F;
    } else if (b >= 0xF0 && b <= 0xF4) {
      longueur = 4;
      if (b === 0xF0) min = 0x90;
      if (b === 0xF4) max = 0x8F;
    }
    if (longueur === 0 || i + longueur > octets.length) return 0;
    for (var k = 1; k < longueur; k++) {
      var suite = octets[i + k];
      var bas = k === 1 ? min : 0x80;
      var haut = k === 1 ? max : 0xBF;
      if (suite < bas || suite > haut) return 0;
    }
    return longueur;
  }

  // Reconnaît un texte UTF-16 sans BOM : un octet sur deux est nul.
  function devinerUtf16(octets) {
    var n = Math.min(octets.length, 4096);
    n -= n % 2;
    if (n < 4) return null;
    var nulsPairs = 0;
    var nulsImpairs = 0;
    for (var i = 0; i < n; i += 2) {
      if (octets[i] === 0) nulsPairs++;
      if (octets[i + 1] === 0) nulsImpairs++;
    }
    var paires = n / 2;
    if (nulsImpairs > paires * 0.4 && nulsPairs < paires * 0.05) return 'utf-16le';
    if (nulsPairs > paires * 0.4 && nulsImpairs < paires * 0.05) return 'utf-16be';
    return null;
  }

  // Détermine l'encodage : celui qu'annonce le BOM s'il y en a un ; sinon
  // UTF-8 s'il est valide, ou s'il contient aussi des octets d'un autre
  // encodage (deux fichiers mis bout à bout, par exemple) : quand les
  // caractères UTF-8 dominent, ou quand ce sont tous des caractères courants
  // en français ; sinon Windows-1252, l'ancien encodage de Windows.
  // Renvoie { encodage, debut (taille du BOM), bom, invalides }.
  function detecterEncodage(octets) {
    octets = enOctets(octets);
    if (commencePar(octets, 0, [0xEF, 0xBB, 0xBF])) {
      return { encodage: 'utf-8', debut: 3, bom: true, invalides: analyserUtf8(octets, 3).invalides };
    }
    if (commencePar(octets, 0, [0xFF, 0xFE])) return { encodage: 'utf-16le', debut: 2, bom: true, invalides: 0 };
    if (commencePar(octets, 0, [0xFE, 0xFF])) return { encodage: 'utf-16be', debut: 2, bom: true, invalides: 0 };
    var utf16 = devinerUtf16(octets);
    if (utf16) return { encodage: utf16, debut: 0, bom: false, invalides: 0 };
    var stats = analyserUtf8(octets, 0);
    if (stats.invalides === 0 || stats.valides > stats.invalides || (stats.valides > 0 && stats.autres === 0)) {
      return { encodage: 'utf-8', debut: 0, bom: false, invalides: stats.invalides };
    }
    return { encodage: 'windows-1252', debut: 0, bom: false, invalides: 0 };
  }

  function estUtf16(info) {
    return info.encodage === 'utf-16le' || info.encodage === 'utf-16be';
  }

  // Caractère de contrôle qu'un vrai texte ne contient presque jamais
  // (tabulation, retours à la ligne, saut de page et ESC sont permis).
  function estControle(code) {
    return code < 0x09 || (code > 0x0D && code < 0x20 && code !== 0x1B);
  }

  // Fin des données utiles : un fichier se termine parfois par une suite
  // d'octets nuls (fichier abîmé par une coupure de courant ou une clé USB
  // retirée trop tôt), de Ctrl+Z (0x1A, remplissage des anciens programmes)
  // ou de blancs. Renvoie le début de cette suite.
  function finSansRemplissage(octets, debut) {
    var fin = octets.length;
    while (fin > debut) {
      var b = octets[fin - 1];
      if (b === 0x00 || b === 0x1A || b === 0x20 || (b >= 0x09 && b <= 0x0D)) fin--;
      else break;
    }
    return fin;
  }

  // Le fichier semble abîmé à la fin : plusieurs octets nuls (en UTF-16,
  // plusieurs caractères nuls) dans la suite qui le termine.
  function finAbimee(octets, info) {
    var nuls = 0;
    var i;
    if (estUtf16(info)) {
      var fin = octets.length - ((octets.length - info.debut) % 2);
      while (fin - 2 >= info.debut && octets[fin - 1] === 0 && octets[fin - 2] === 0) {
        nuls++;
        fin -= 2;
      }
    } else {
      for (i = finSansRemplissage(octets, info.debut); i < octets.length; i++) {
        if (octets[i] === 0) nuls++;
      }
    }
    return nuls >= 2;
  }

  // « texte », « binaire » (image, son, PDF, document Word ou Excel…) ou « rtf ».
  function natureContenu(octets, info) {
    if (aSignatureBinaire(octets)) return 'binaire';
    if (commencePar(octets, info.debut, SIGNATURE_RTF)) return 'rtf';
    // Un vrai texte ne contient presque pas de caractères de contrôle.
    var examines = 0;
    var controles = 0;
    if (estUtf16(info)) {
      // En UTF-16, on examine les caractères (deux octets), pas les octets.
      var petitBoutiste = info.encodage === 'utf-16le';
      var finUtf16 = Math.min(octets.length - 1, info.debut + 8192);
      for (var p = info.debut; p < finUtf16; p += 2) {
        var u = petitBoutiste ? (octets[p] | (octets[p + 1] << 8)) : ((octets[p] << 8) | octets[p + 1]);
        examines++;
        if ((u !== 0 && estControle(u)) || u === 0xFFFE || u === 0xFFFF) controles++;
      }
    } else {
      var fin = Math.min(finSansRemplissage(octets, info.debut), info.debut + 8192);
      for (var i = info.debut; i < fin; i++) {
        examines++;
        if (estControle(octets[i])) controles++;
      }
    }
    return controles > Math.max(1, examines * 0.05) ? 'binaire' : 'texte';
  }

  // Construit une chaîne à partir de codes UTF-16, par paquets (la fonction
  // String.fromCharCode n'accepte pas un nombre illimité d'arguments).
  function chaineDepuisCodes(codes) {
    var morceaux = [];
    for (var i = 0; i < codes.length; i += 8192) {
      morceaux.push(String.fromCharCode.apply(null, codes.subarray(i, Math.min(i + 8192, codes.length))));
    }
    return morceaux.join('');
  }

  function codeWindows1252(b) {
    return (b >= 0x80 && b <= 0x9F) ? WINDOWS_1252[b - 0x80] : b;
  }

  // Dans un texte enregistré par un logiciel DOS (page de code 850) ou un
  // ancien Mac, « é », « è », « ç »… deviennent des octets que Windows-1252
  // affiche « ‚ », « Š », « ‡ », « Ž »… Un tel signe collé entre deux
  // lettres, dont une minuscule (« D‚bit »), trahit un autre encodage. (Des
  // noms en majuscules comme « DUŠAN » sont, eux, du vrai Windows-1252.)
  var SIGNES_DOS_MAC = { 0x82: true, 0x83: true, 0x87: true, 0x88: true, 0x8A: true, 0x8E: true };

  function estLettreAscii(b) {
    return (b >= 0x41 && b <= 0x5A) || (b >= 0x61 && b <= 0x7A);
  }

  function lettreDosOuMac(octets, debut) {
    for (var i = debut + 1; i < octets.length - 1; i++) {
      if (!SIGNES_DOS_MAC[octets[i]]) continue;
      var avant = octets[i - 1];
      var apres = octets[i + 1];
      if (estLettreAscii(avant) && estLettreAscii(apres) && (avant >= 0x61 || apres >= 0x61)) return true;
    }
    return false;
  }

  // Renvoie { texte, inconnus } : « inconnus » compte les octets que
  // Windows-1252 n'utilise pas (devenus « � »).
  function decoderWindows1252(octets, debut) {
    var codes = new Uint16Array(octets.length - debut);
    var inconnus = 0;
    for (var i = 0; i < codes.length; i++) {
      var code = codeWindows1252(octets[debut + i]);
      if (code === 0xFFFD) inconnus++;
      codes[i] = code;
    }
    return { texte: chaineDepuisCodes(codes), inconnus: inconnus };
  }

  // Fichier UTF-8 contenant aussi quelques octets d'un autre encodage (par
  // exemple deux fichiers mis bout à bout, l'un en UTF-8, l'autre enregistré
  // par un ancien logiciel) : les caractères UTF-8 valides sont lus en UTF-8,
  // les octets isolés en Windows-1252. Ainsi, aucun accent n'est perdu.
  // Renvoie { texte, inconnus }.
  function decoderUtf8Mixte(octets, debut) {
    var n = octets.length;
    var codes = new Uint16Array(n - debut);
    var k = 0;
    var inconnus = 0;
    var i = debut;
    while (i < n) {
      var b = octets[i];
      if (b < 0x80) {
        codes[k++] = b;
        i++;
        continue;
      }
      var longueur = longueurUtf8(octets, i);
      if (longueur === 0) {
        var code = codeWindows1252(b);
        if (code === 0xFFFD) inconnus++;
        codes[k++] = code;
        i++;
        continue;
      }
      var point = longueur === 2 ? (b & 0x1F) : longueur === 3 ? (b & 0x0F) : (b & 0x07);
      for (var j = 1; j < longueur; j++) point = (point << 6) | (octets[i + j] & 0x3F);
      if (point >= 0x10000) {
        point -= 0x10000;
        codes[k++] = 0xD800 + (point >> 10);
        codes[k++] = 0xDC00 + (point & 0x3FF);
      } else {
        codes[k++] = point;
      }
      i += longueur;
    }
    return { texte: chaineDepuisCodes(codes.subarray(0, k)), inconnus: inconnus };
  }

  function decoderUtf16(octets, debut, petitBoutiste) {
    var codes = new Uint16Array(Math.floor((octets.length - debut) / 2));
    for (var i = 0; i < codes.length; i++) {
      var a = octets[debut + 2 * i];
      var b = octets[debut + 2 * i + 1];
      codes[i] = petitBoutiste ? (a | (b << 8)) : ((a << 8) | b);
    }
    return chaineDepuisCodes(codes);
  }

  // Ctrl+Z (0x1A) marquait la fin des fichiers sous DOS, et la commande
  // « copy » de Windows en ajoute encore un. On l'enlève, avec ce qui le suit,
  // quand il ne reste ensuite que des blancs.
  function retirerMarqueFinDos(texte) {
    var fin = texte.length;
    while (fin > 0) {
      var c = texte.charCodeAt(fin - 1);
      if (c === 0x1A || c === 0x00 || c === 0x20 || (c >= 0x09 && c <= 0x0D)) fin--;
      else break;
    }
    var marque = texte.indexOf('\u001A', fin);
    return marque === -1 ? texte : texte.slice(0, marque);
  }

  // Fichier UTF-8 coupé au milieu de son dernier caractère (téléchargement
  // interrompu, logiciel qui tronque le texte) : renvoie la position où
  // commence ce caractère incomplet, ou -1. Seuls comptent les caractères
  // courants d'un texte français qui s'écrivent sur 3 ou 4 octets (’ … – — €,
  // qui commencent par E2, et les emoji, par F0 9F), et il faut au moins un
  // octet de suite : « é… » ou « é » seul en fin de fichier peuvent aussi être
  // du Windows-1252.
  function debutCaractereCoupe(octets, debut) {
    var n = octets.length;
    for (var p = n - 1; p >= debut && p >= n - 3; p--) {
      var b = octets[p];
      if (b >= 0x80 && b <= 0xBF) continue; // octet de suite
      if ((b !== 0xE2 && b !== 0xF0) || p === n - 1) return -1;
      if (n - p >= (b === 0xE2 ? 3 : 4)) return -1; // caractère complet
      if (b === 0xF0 && octets[p + 1] < 0x90) return -1;
      return p;
    }
    return -1;
  }

  // Transforme les octets en texte.
  // Renvoie { texte, encodage, bom, illisibles, melange, douteux, finCoupee } :
  // - illisibles : des caractères n'ont pas pu être lus (ils deviennent « � ») ;
  // - melange : fichier UTF-8 contenant aussi des octets d'un autre encodage ;
  // - douteux : fichier probablement enregistré par un ancien logiciel (DOS, Mac) ;
  // - finCoupee : fichier UTF-8 coupé au milieu de son dernier caractère.
  function decoder(octets, info) {
    octets = enOctets(octets);
    info = info || detecterEncodage(octets);
    var texte;
    var illisibles = false;
    var melange = false;
    var douteux = false;
    var finCoupee = false;
    var lu;
    if (info.encodage === 'utf-8') {
      if (info.invalides > 0) {
        // Le caractère coupé peut être suivi d'octets nuls (fichier abîmé).
        var finDonnees = octets.length;
        while (finDonnees > info.debut && octets[finDonnees - 1] === 0) finDonnees--;
        var coupe = debutCaractereCoupe(octets.subarray(0, finDonnees), info.debut);
        var fin = coupe === -1 ? octets.length : coupe;
        if (info.invalides > (coupe === -1 ? 0 : finDonnees - coupe)) {
          lu = decoderUtf8Mixte(octets.subarray(0, fin), info.debut);
          texte = lu.texte;
          melange = true;
          illisibles = lu.inconnus > 0;
        } else {
          texte = new TextDecoder('utf-8').decode(octets.subarray(info.debut, fin));
        }
        if (coupe !== -1) {
          texte += String.fromCharCode(0xFFFD);
          illisibles = true;
          finCoupee = true;
        }
      } else {
        texte = new TextDecoder('utf-8').decode(octets.subarray(info.debut));
      }
    } else if (info.encodage === 'windows-1252') {
      lu = decoderWindows1252(octets, info.debut);
      texte = lu.texte;
      douteux = lu.inconnus > 0 || lettreDosOuMac(octets, info.debut);
    } else {
      texte = decoderUtf16(octets, info.debut, info.encodage === 'utf-16le');
      if ((octets.length - info.debut) % 2 === 1) {
        texte += '\uFFFD'; // dernier octet orphelin
        illisibles = true;
      }
    }
    texte = retirerMarqueFinDos(texte);
    // Caractères invisibles parasites : octets nuls et BOM égarés.
    if (/[\u0000\uFEFF]/.test(texte)) texte = texte.replace(/[\u0000\uFEFF]/g, '');
    return {
      texte: texte,
      encodage: info.encodage,
      bom: info.bom,
      illisibles: illisibles,
      melange: melange,
      douteux: douteux,
      finCoupee: finCoupee
    };
  }

  /* ------------------------------------------------------------------ */
  /* Découpage du texte en lignes et en colonnes (règles RFC 4180)      */
  /* ------------------------------------------------------------------ */

  function estFinDeLigne(c) {
    return c === '\n' || c === '\r';
  }

  function estBlanc(c) {
    return c === ' ' || c === '\t';
  }

  // Position qui suit une suite d'espaces et de tabulations commençant en i.
  function finBlancs(texte, i) {
    while (i < texte.length && estBlanc(texte.charAt(i))) i++;
    return i;
  }

  // Longueur du séparateur qui commence à la position i (0 s'il n'y en a pas).
  function longueurSeparateur(texte, i, separateur) {
    var c = texte.charAt(i);
    if (separateur === 'espaces') {
      // Au moins deux espaces de suite, ou une tabulation.
      if (!estBlanc(c)) return 0;
      var j = finBlancs(texte, i);
      // Des espaces en fin de ligne ne séparent rien.
      if (j >= texte.length || estFinDeLigne(texte.charAt(j))) return 0;
      if (j - i >= 2 || c === '\t') return j - i;
      return 0;
    }
    var attendu = CARACTERES[separateur];
    return (attendu !== undefined && c === attendu) ? 1 : 0;
  }

  // Avance d'un caractère dans une valeur. En mode « espaces », une suite de
  // blancs qui ne sépare rien est franchie d'un coup : aucun de ses caractères
  // ne peut commencer un séparateur (sinon, de longues suites d'espaces en fin
  // de ligne rendraient le découpage très lent).
  function avancer(texte, i, separateur) {
    return (separateur === 'espaces' && estBlanc(texte.charAt(i))) ? finBlancs(texte, i) : i + 1;
  }

  // Lit une valeur entre guillemets, à partir du caractère qui suit le
  // guillemet ouvrant. "" représente un guillemet ; séparateurs et retours à
  // la ligne sont permis à l'intérieur. Renvoie { valeur, fin, fermeture }
  // (fermeture : position qui suit le guillemet fermant), ou null si les
  // guillemets sont mal formés (jamais fermés, ou du texte collé après le
  // guillemet fermant) : la valeur est alors lue telle quelle, sans rien perdre.
  function lireEntreGuillemets(texte, i, separateur, estMarge) {
    var n = texte.length;
    var morceaux = [];
    var debut = i;
    for (;;) {
      var q = texte.indexOf('"', i);
      if (q === -1) return null;
      if (texte.charAt(q + 1) === '"') {
        morceaux.push(texte.slice(debut, q + 1));
        i = q + 2;
        debut = i;
        continue;
      }
      morceaux.push(texte.slice(debut, q));
      var m = q + 1;
      while (m < n && estMarge(texte.charAt(m)) && longueurSeparateur(texte, m, separateur) === 0) {
        m = avancer(texte, m, separateur);
      }
      if (m >= n || estFinDeLigne(texte.charAt(m)) || longueurSeparateur(texte, m, separateur) > 0) {
        // Retours à la ligne uniformisés (\n), comme dans Excel.
        return { valeur: morceaux.join('').replace(/\r\n?/g, '\n'), fin: m, fermeture: q + 1 };
      }
      return null;
    }
  }

  function estLigneVide(champs) {
    for (var i = 0; i < champs.length; i++) {
      if (champs[i] !== '') return false;
    }
    return true;
  }

  // Colonnes à l'écran des positions données (rangées dans l'ordre, toutes
  // sur la ligne qui commence en debutLigne), comme dans le Bloc-notes :
  // une tabulation avance jusqu'à la colonne multiple de 8 suivante, un
  // caractère codé sur deux unités (emoji…) ne compte qu'une fois, et un
  // accent séparé de sa lettre (texte « décomposé », venu d'un Mac ou d'un
  // PDF) ne prend pas de place.
  function colonnesAffichees(texte, debutLigne, positions) {
    var colonnes = new Array(positions.length);
    var colonne = 0;
    var p = debutLigne;
    for (var k = 0; k < positions.length; k++) {
      while (p < positions[k]) {
        var code = texte.charCodeAt(p);
        if (code === 0x09) colonne += LARGEUR_TABULATION - (colonne % LARGEUR_TABULATION);
        else if ((code < 0xDC00 || code > 0xDFFF) && (code < 0x0300 || code > 0x036F)) colonne++;
        p++;
      }
      colonnes[k] = colonne;
    }
    return colonnes;
  }

  // Découpe le texte. Les lignes vides (ou ne contenant que des espaces ou
  // des séparateurs) sont ignorées et comptées. Les valeurs sans guillemets
  // sont débarrassées des espaces au début et à la fin.
  // avecPositions : note aussi, pour chaque ligne d'au moins deux valeurs
  // tenant sur une seule ligne du fichier, la colonne à l'écran où commence
  // et finit chaque valeur (sinon null).
  // Renvoie { lignes, colonnes, lignesVides, lignesGuillemetsIncoherents,
  // longueurMax, positions }.
  function decouperTexte(texte, separateur, maxLignes, avecPositions) {
    var guillemetsActifs = separateur !== 'aucun';
    var n = texte.length;
    var lignes = [];
    var positions = avecPositions ? [] : null;
    var colonnes = 0;
    var lignesVides = 0;
    var lignesGuillemetsIncoherents = 0;
    var longueurMax = 0;
    var i = 0;

    // Espaces (et tabulations, sauf si c'est le séparateur) autour d'une valeur.
    function estMarge(c) {
      return c === ' ' || (c === '\t' && separateur !== 'tabulation');
    }

    while (i < n && lignes.length < maxLignes) {
      var debutLigne = i;
      var champs = [];
      var bornes = avecPositions ? [] : null;
      var surPlusieursLignes = false;
      var incoherent = false;
      for (;;) {
        var debut = i;
        while (debut < n && estMarge(texte.charAt(debut))) debut++;
        var valeur = null;
        var fin = debut;
        var debutValeur = debut;
        var finValeur = debut;
        if (guillemetsActifs && texte.charAt(debut) === '"') {
          var lu = lireEntreGuillemets(texte, debut + 1, separateur, estMarge);
          if (lu) {
            valeur = lu.valeur;
            fin = lu.fin;
            finValeur = lu.fermeture;
            if (valeur.indexOf('\n') !== -1) surPlusieursLignes = true;
          } else {
            incoherent = true;
          }
        }
        if (valeur === null) {
          fin = debut;
          while (fin < n && !estFinDeLigne(texte.charAt(fin)) &&
                 longueurSeparateur(texte, fin, separateur) === 0) {
            fin = avancer(texte, fin, separateur);
          }
          var brut = texte.slice(debut, fin);
          valeur = brut.trim();
          if (bornes && valeur !== '') {
            debutValeur = debut + brut.indexOf(valeur);
            finValeur = debutValeur + valeur.length;
          }
        }
        champs.push(valeur);
        if (bornes) bornes.push(debutValeur, finValeur);
        if (valeur.length > longueurMax) longueurMax = valeur.length;
        i = fin;
        if (i >= n) break;
        var c = texte.charAt(i);
        if (estFinDeLigne(c)) {
          i += (c === '\r' && texte.charAt(i + 1) === '\n') ? 2 : 1;
          break;
        }
        i += Math.max(1, longueurSeparateur(texte, i, separateur));
      }
      if (incoherent) lignesGuillemetsIncoherents++;
      if (estLigneVide(champs)) {
        lignesVides++;
        continue;
      }
      lignes.push(champs);
      if (positions) {
        positions.push((champs.length >= 2 && !surPlusieursLignes) ? colonnesAffichees(texte, debutLigne, bornes) : null);
      }
      if (champs.length > colonnes) colonnes = champs.length;
    }
    return {
      lignes: lignes,
      colonnes: colonnes,
      lignesVides: lignesVides,
      lignesGuillemetsIncoherents: lignesGuillemetsIncoherents,
      longueurMax: longueurMax,
      positions: positions
    };
  }

  // Colonnes sous lesquelles se trouvent les « nombre » valeurs d'une ligne
  // (positions q), ou -1 si l'une d'elles est à cheval sur deux colonnes, ou
  // null si l'une d'elles n'est sous aucune colonne ou partage sa colonne
  // avec une autre.
  function placerValeurs(q, nombre, debuts, fins) {
    var complet = debuts.length;
    var cibles = [];
    var colonne = 0;
    for (var k = 0; k < nombre; k++) {
      var d = q[2 * k];
      var f = q[2 * k + 1];
      while (colonne < complet && fins[colonne] <= d) colonne++;
      if (colonne >= complet || debuts[colonne] >= f) return null;       // sous aucune colonne
      if (colonne + 1 < complet && debuts[colonne + 1] < f) return -1;   // à cheval sur deux colonnes
      if (k > 0 && colonne <= cibles[k - 1]) return null;                // deux valeurs sous la même colonne
      cibles.push(colonne);
    }
    return cibles;
  }

  // Colonnes à l'écran où commencent les valeurs d'une ligne de « nombre »
  // valeurs, sauf la dernière quand il y en a au moins trois (souvent un
  // montant aligné à droite, qui commence plus ou moins loin).
  function signature(q, nombre) {
    var s = '';
    for (var k = 0; k < (nombre >= 3 ? nombre - 1 : nombre); k++) s += q[2 * k] + ',';
    return s;
  }

  // Recale les lignes incomplètes sur les colonnes des lignes qui ont
  // « complet » valeurs. Renvoie false, sans rien changer, si ces lignes ne
  // donnent pas les colonnes du tableau : un titre, une ligne de total ou un
  // libellé contenant deux espaces de suite peuvent avoir plus de valeurs
  // que le tableau, mais leurs espaces ne tombent pas entre ses colonnes.
  function recalerSur(lignes, positions, complet) {
    var i;
    var j;
    var k;
    var cibles;
    var debuts = [];
    var fins = [];
    var premiere = -1; // première ligne complète : l'en-tête, en général
    // Une ligne complète située sous toutes les lignes incomplètes (ligne de
    // total, remarque) ne donne pas les colonnes du tableau.
    var derniere = lignes.length - 1;
    while (derniere >= 0 && !(positions[derniere] && lignes[derniere].length < complet)) derniere--;
    for (j = 0; j < complet; j++) {
      debuts.push(Infinity);
      fins.push(-Infinity);
    }
    for (i = 0; i < derniere; i++) {
      var p = positions[i];
      if (!p || lignes[i].length !== complet) continue;
      if (premiere === -1) premiere = i;
      for (j = 0; j < complet; j++) {
        if (p[2 * j] < debuts[j]) debuts[j] = p[2 * j];
        if (p[2 * j + 1] > fins[j]) fins[j] = p[2 * j + 1];
      }
    }
    if (premiere === -1) return false;
    for (j = 1; j < complet; j++) {
      if (fins[j - 1] > debuts[j]) return false; // colonnes qui se chevauchent
    }

    // Les lignes incomplètes situées sous la première ligne complète doivent
    // tomber sous ses colonnes. Une ligne à cheval sur deux colonnes (où
    // qu'elle soit) montre que ces colonnes coupent celles du tableau.
    var utilisees = [];
    var aCheval = 0;
    var placees = 0;
    var frequences = {};
    var forme = 0; // nombre de valeurs le plus courant parmi ces lignes
    for (j = 0; j < complet; j++) utilisees.push(false);
    for (i = 0; i < lignes.length; i++) {
      var n = lignes[i].length;
      if (!positions[i] || n >= complet) continue;
      cibles = placerValeurs(positions[i], n, debuts, fins);
      if (cibles === -1) aCheval++;
      if (i < premiere) continue;
      frequences[n] = (frequences[n] || 0) + 1;
      if (!forme || frequences[n] > frequences[forme]) forme = n;
      if (cibles && cibles !== -1) {
        placees++;
        for (k = 0; k < cibles.length; k++) utilisees[cibles[k]] = true;
      }
    }
    if (placees === 0) return false; // rien à recaler en dessous (ligne de total en bas…)
    for (j = 0; j < complet; j++) {
      if (!utilisees[j] && aCheval > 0) return false;
    }
    // Au-dessus de la première ligne complète, une ligne du tableau (mêmes
    // colonnes de départ que les lignes à recaler) : la ligne complète est au
    // milieu du tableau (libellé contenant deux espaces), ce n'est pas l'en-tête.
    // On compare les colonnes de départ et les colonnes d'arrivée, et seulement
    // avec les lignes que le recalage déplacerait vraiment : une ligne de
    // coordonnées imprimée sur la grille du tableau (« Titulaire   M. DUPONT
    // Jean   Compte n° … ») n'est pas une ligne du tableau.
    var signatures = {};
    for (i = premiere + 1; i < lignes.length; i++) {
      if (!positions[i] || lignes[i].length !== forme) continue;
      cibles = placerValeurs(positions[i], forme, debuts, fins);
      if (cibles && cibles !== -1 && cibles[forme - 1] !== forme - 1) {
        signatures[signature(positions[i], forme) + '|' + cibles.join(',')] = true;
      }
    }
    for (i = 0; i < premiere; i++) {
      if (!positions[i] || lignes[i].length !== forme) continue;
      cibles = placerValeurs(positions[i], forme, debuts, fins);
      if (cibles && cibles !== -1 && signatures[signature(positions[i], forme) + '|' + cibles.join(',')]) return false;
    }

    // Les lignes du dessus (titre, coordonnées…) restent telles quelles.
    for (i = premiere + 1; i < lignes.length; i++) {
      var ligne = lignes[i];
      if (!positions[i] || ligne.length >= complet) continue;
      cibles = placerValeurs(positions[i], ligne.length, debuts, fins);
      if (!cibles || cibles === -1) continue;
      var recalee = [];
      for (j = 0; j <= cibles[cibles.length - 1]; j++) recalee.push(''); // genererCsv complète la ligne
      for (k = 0; k < ligne.length; k++) recalee[cibles[k]] = ligne[k];
      lignes[i] = recalee;
    }
    return true;
  }

  // Colonnes alignées par des espaces : une case vide n'y est qu'un espace
  // plus large, que le découpage fait disparaître (les valeurs suivantes
  // glissent alors d'une colonne vers la gauche : un crédit passe sous
  // « Débit »). On replace donc les valeurs des lignes incomplètes sous les
  // colonnes des lignes complètes, d'après leur position à l'écran, mais
  // seulement quand c'est sans ambiguïté : chaque valeur doit se trouver sous
  // une seule colonne, et dans l'ordre. Sinon, la ligne reste telle quelle.
  // On essaie d'abord les lignes les plus larges ; si ce sont un titre ou un
  // total, les suivantes.
  // Limites connues (le résultat est alors celui d'un simple découpage, comme
  // si rien n'était recalé) : tableau sans en-tête dont les premières lignes
  // sont incomplètes ; colonne vide tout le mois (aucun crédit) avec une ligne
  // de totaux à cheval sur les colonnes. Cas rares où une colonne peut rester
  // partagée : deux tableaux imprimés sur la même grille, ou un titre de
  // plusieurs valeurs qui tombe exactement sur les colonnes du tableau.
  function recalerColonnes(lignes, positions) {
    var largeurs = [];
    var vues = {};
    for (var i = 0; i < lignes.length; i++) {
      var n = lignes[i].length;
      if (positions[i] && !vues[n]) {
        vues[n] = true;
        largeurs.push(n);
      }
    }
    largeurs.sort(function (a, b) { return b - a; });
    for (var w = 0; w < largeurs.length && w < 4 && largeurs[w] >= 3; w++) {
      if (recalerSur(lignes, positions, largeurs[w])) return;
    }
  }

  // Y a-t-il des lignes d'au moins deux valeurs, mais moins que les plus complètes ?
  function aDesLignesIncompletes(lignes) {
    var complet = 0;
    var i;
    for (i = 0; i < lignes.length; i++) {
      if (lignes[i].length > complet) complet = lignes[i].length;
    }
    for (i = 0; i < lignes.length; i++) {
      if (lignes[i].length >= 2 && lignes[i].length < complet) return true;
    }
    return false;
  }

  // Découpe le texte avec le séparateur choisi.
  // options.maxLignes : nombre maximal de lignes à lire.
  // Renvoie { lignes, colonnes, lignesVides, lignesGuillemetsIncoherents, longueurMax }.
  function decouper(texte, separateur, options) {
    if (SEPARATEURS.indexOf(separateur) === -1) {
      throw new Error('Séparateur inconnu : ' + separateur);
    }
    var maxLignes = (options && options.maxLignes) || Infinity;
    var resultat = decouperTexte(texte, separateur, maxLignes, false);
    // Colonnes alignées par des espaces : les positions à l'écran ne sont
    // calculées (second passage) que s'il y a des lignes à recaler.
    if (separateur === 'espaces' && aDesLignesIncompletes(resultat.lignes)) {
      resultat = null; // libère le premier découpage pendant le second
      resultat = decouperTexte(texte, separateur, maxLignes, true);
      recalerColonnes(resultat.lignes, resultat.positions);
    }
    return {
      lignes: resultat.lignes,
      colonnes: resultat.colonnes,
      lignesVides: resultat.lignesVides,
      lignesGuillemetsIncoherents: resultat.lignesGuillemetsIncoherents,
      longueurMax: resultat.longueurMax
    };
  }

  /* ------------------------------------------------------------------ */
  /* Détection automatique du séparateur                                */
  /* ------------------------------------------------------------------ */

  // Des virgules qui ne sont que des virgules décimales (« 12,40 »,
  // « 1 250,00 ») ne séparent pas de colonnes. Un texte de part et d'autre
  // d'une virgule (en-tête « id,age ») prouve en revanche que si.
  function virgulesDecimales(lignes) {
    var frontieres = 0;
    var decimales = 0;
    for (var l = 0; l < lignes.length; l++) {
      var ligne = lignes[l];
      for (var c = 1; c < ligne.length; c++) {
        var gauche = ligne[c - 1];
        var droite = ligne[c];
        frontieres++;
        if (/[^\d\s]$/.test(gauche) && /^[^\d\s]/.test(droite)) return false;
        // Un nombre seul entre deux virgules (« 12,23,34 ») : ce sont des séparateurs.
        if (c + 1 < ligne.length && /^\d+$/.test(droite)) return false;
        // Partie entière d'un nombre, et non fin d'une date ou d'une heure
        // (« 05/01/2026,12 », « 01/2026,45 », « 08:00,12 »). « Loyer:650,00 »
        // ou « 1kg/2,99 », eux, sont bien des nombres décimaux.
        if (/\d$/.test(gauche) && !/(?:^|\D)\d{1,2}(?:\/\d{1,2}(?:\/\d{2,4})?|\/\d{4}|:\d{2}(?::\d{2})?)$/.test(gauche) &&
            /^\d{2}(?!\d)/.test(droite)) decimales++;
      }
    }
    return frontieres > 0 && decimales / frontieres >= 0.8;
  }

  // Note de 0 à 1 : un bon séparateur coupe presque toutes les lignes, en un
  // nombre de colonnes régulier, avec des guillemets bien placés.
  function noter(resultat, separateur) {
    var lignes = resultat.lignes;
    var n = lignes.length;
    if (n === 0) return 0;
    var frequences = {};
    var coupees = 0;
    var meilleure = 0;
    for (var i = 0; i < n; i++) {
      var k = lignes[i].length;
      if (k < 2) continue;
      coupees++;
      frequences[k] = (frequences[k] || 0) + 1;
      if (frequences[k] > meilleure) meilleure = frequences[k];
    }
    if (coupees === 0) return 0;
    var regularite = meilleure / n;
    // Virgules et espaces abondent dans un texte ordinaire : on exige alors le
    // même nombre de colonnes sur au moins la moitié des lignes.
    if (SEPARATEURS_FAIBLES[separateur] && regularite < 0.5) return 0;
    if (separateur === 'virgule' && virgulesDecimales(lignes)) return 0;
    var score = (coupees / n + regularite) / 2;
    // Des guillemets qui « tombent mal » trahissent un mauvais séparateur.
    score *= 1 - 0.5 * (resultat.lignesGuillemetsIncoherents / n);
    // Des valeurs coupées au milieu d'espaces alignés aussi (texte en colonnes
    // alignées contenant des nombres à virgule, par exemple). Les espaces qui
    // complètent une valeur entre guillemets jusqu'à une largeur fixe
    // (« "DUPONT      " ») ne comptent pas.
    if (separateur !== 'tabulation' && separateur !== 'espaces') {
      var cellules = 0;
      var suspectes = 0;
      for (var l = 0; l < n; l++) {
        for (var c = 0; c < lignes[l].length; c++) {
          var v = lignes[l][c];
          if (v === '') continue;
          cellules++;
          if (/ {2}|\t/.test(v.trim())) suspectes++;
        }
      }
      if (cellules > 0) score *= 1 - suspectes / cellules;
    }
    return score;
  }

  // Colonnes alignées à la main avec la touche Tab (Bloc-notes) : selon la
  // longueur des valeurs, il faut une ou deux tabulations, mais chaque
  // colonne commence au même endroit à l'écran. Dans un vrai fichier à
  // tabulations (export d'Excel), les cases vides décalent au contraire les
  // valeurs à l'écran.
  function alignementParTabulations(resultat) {
    var departs = [];
    var comparees = 0;
    for (var i = 0; i < resultat.lignes.length; i++) {
      var p = resultat.positions[i];
      if (!p) continue;
      for (var k = 1; k < resultat.lignes[i].length; k++) {
        if (departs[k] === undefined) departs[k] = p[2 * k];
        else if (departs[k] !== p[2 * k]) return false;
      }
      comparees++;
    }
    return comparees >= 2;
  }

  // Renvoie le code du séparateur le plus probable (« aucun » si rien ne convient).
  function detecterSeparateur(texte) {
    var scores = {};
    var meilleur = 0;
    var s;
    for (s = 0; s < SEPARATEURS_DETECTABLES.length; s++) {
      var code = SEPARATEURS_DETECTABLES[s];
      var score = noter(decouperTexte(texte, code, LIGNES_ECHANTILLON, false), code);
      scores[code] = score;
      if (score > meilleur) meilleur = score;
    }
    if (meilleur < SEUIL_DETECTION) return 'aucun';
    // La tabulation, prioritaire en cas d'égalité, céderait à tort la place
    // aux colonnes alignées par plusieurs tabulations.
    if (scores.espaces === meilleur && scores.tabulation < meilleur && texte.indexOf('\t\t') !== -1 &&
        alignementParTabulations(decouperTexte(texte, 'espaces', LIGNES_ECHANTILLON, true))) {
      scores.tabulation = 0;
    }
    for (s = 0; s < SEPARATEURS_DETECTABLES.length; s++) {
      if (scores[SEPARATEURS_DETECTABLES[s]] >= meilleur - MARGE_EGALITE) return SEPARATEURS_DETECTABLES[s];
    }
    return 'aucun';
  }

  /* ------------------------------------------------------------------ */
  /* Écriture du CSV                                                    */
  /* ------------------------------------------------------------------ */

  // Nombres qu'Excel lit tels quels, avec ou sans espaces autour : -12,40 /
  // - 12,40 / 1 250,00 € / -5 % / +5 / -,5 / -1,5E3. (Des chiffres groupés par
  // une espace insécable restent du texte dans Excel, mais s'affichent tels
  // quels : inutile de les protéger.)
  var NOMBRE_SIMPLE = /^[+-]? *(?:\d{1,3}(?:[ \u00A0\u202F]\d{3})+|\d+|(?=,\d))(?:,\d+)?(?:[eE][+-]?\d{1,3})? ?(?:\u20AC|%)?$/;

  // Nombre entier, éventuellement entouré d'espaces.
  var NOMBRE_ENTIER = /^[ \t]*\d+[ \t]*$/;

  // Valeurs qu'Excel transformerait en ouvrant le CSV, et qui seraient donc
  // affichées autrement que dans le fichier d'origine :
  // - un texte qui commence par = + - @ est pris pour une formule
  //   (« - Pain » affiche #NOM?, « -1-2 » affiche -3) ;
  // - les zéros du début disparaissent (06000 → 6000, 0612345678 → 612345678) ;
  // - un long nombre est arrondi ou affiché en notation scientifique (SIRET,
  //   numéro de carte, +33612345678 → 3,36E+10).
  // Les nombres ordinaires et les dates ne sont pas concernés.
  function aFigerPourExcel(v) {
    if (v === '') return false;
    var c = v.charAt(0);
    // Pris pour une formule (une espace devant suffirait à l'éviter).
    if (c === '=' || c === '@') return true;
    // Excel ignore les espaces autour d'un nombre : on les retire pour l'examiner.
    var t = (estBlanc(c) || estBlanc(v.charAt(v.length - 1))) ? v.replace(/^[ \t]+|[ \t]+$/g, '') : v;
    if ((c === '+' || c === '-') && !NOMBRE_SIMPLE.test(t)) return true;
    var s = t.charAt(0);
    var signe = s === '+' || s === '-';
    if (!signe && (s < '0' || s > '9')) return false;
    var corps = signe ? t.replace(/^[+-] */, '') : t;
    if (!/^\d+$/.test(corps)) return false;
    if (corps.length >= 2 && corps.charAt(0) === '0') return true;   // 06000, 0612345678, 000123
    if (corps.length >= 11) return true;                            // SIRET, EAN, longues références
    if (s === '+' && corps.length >= 8) return true;                // +33612345678
    return false;
  }

  // Colonnes de codes : codes postaux, téléphones, références de même longueur
  // (06000, 75001…). Une colonne en est une si l'un de ses nombres entiers doit
  // garder ses zéros du début ou tous ses chiffres, si tous ses nombres entiers
  // ont le même nombre de chiffres, et si elle ne contient aucun autre nombre
  // (-600, 7,5…). Dans ces colonnes, tous les nombres entiers sont écrits comme
  // du texte : la colonne reste homogène (tri, alignement). Une valeur isolée
  // (numéro de compte dans l'en-tête, « Page 01 ») ne change pas une colonne
  // de quantités : elle seule est protégée, et les sommes restent justes.
  function colonnesDeCodes(lignes, colonnes) {
    var codes = [];
    var chiffres = [];
    var i;
    var j;
    for (j = 0; j < colonnes; j++) {
      codes.push(false);
      chiffres.push(0);
    }
    for (i = 0; i < lignes.length; i++) {
      var ligne = lignes[i];
      for (j = 0; j < ligne.length; j++) {
        var v = String(ligne[j]);
        if (codes[j] === null || v === '') continue;
        var t = (estBlanc(v.charAt(0)) || estBlanc(v.charAt(v.length - 1))) ? v.replace(/^[ \t]+|[ \t]+$/g, '') : v;
        if (/^\d+$/.test(t)) {
          if (chiffres[j] === 0) chiffres[j] = t.length;
          if (chiffres[j] !== t.length) codes[j] = null;
          else if (!codes[j] && aFigerPourExcel(v)) codes[j] = true;
        } else if (NOMBRE_SIMPLE.test(t) && !aFigerPourExcel(v)) {
          codes[j] = null; // un autre nombre : colonne de quantités ou de montants
        }
      }
    }
    for (j = 0; j < colonnes; j++) codes[j] = codes[j] === true;
    return codes;
  }

  // Valeur écrite pour Excel. La forme ="…" s'affiche exactement comme du
  // texte. Mais une formule est limitée à 8 192 caractères, et Excel n'affiche
  // pas sur plusieurs lignes le résultat d'une formule : dans ces deux cas,
  // une espace devant suffit à ce qu'un texte ne soit pas pris pour une formule.
  function pourExcel(v, colonneDeCodes) {
    if (!aFigerPourExcel(v) && !(colonneDeCodes && NOMBRE_ENTIER.test(v))) return v;
    if (v.indexOf('\n') === -1 && v.length < 8100) {
      var formule = formuleTexte(v);
      if (formule.length <= 8192) return formule;
    }
    var c = v.charAt(0);
    return (c === '=' || c === '+' || c === '-' || c === '@') ? ' ' + v : v;
  }

  // Écrit la valeur sous la forme ="…", qu'Excel affiche exactement comme du
  // texte. Un texte dans une formule est limité à 255 caractères : une valeur
  // plus longue est écrite en morceaux reliés par &.
  function formuleTexte(v) {
    var morceaux = [];
    for (var i = 0; i < v.length;) {
      var f = Math.min(i + 200, v.length);
      var code = v.charCodeAt(f - 1);
      if (f < v.length && code >= 0xD800 && code <= 0xDBFF) f--; // ne pas couper un caractère en deux
      morceaux.push('"' + v.slice(i, f).replace(/"/g, '""') + '"');
      i = f;
    }
    return '=' + morceaux.join('&');
  }

  // Génère le texte CSV : une ligne par enregistrement, séparées par CRLF ;
  // une valeur est mise entre guillemets si elle contient le séparateur, un
  // guillemet, un retour à la ligne, ou des espaces au début ou à la fin.
  // Toutes les lignes reçoivent le même nombre de colonnes.
  function genererCsv(lignes, format) {
    var reglage = FORMATS[format];
    if (!reglage) throw new Error('Format inconnu : ' + format);
    var separateur = reglage.separateur;
    var aProteger = separateur === ';' ? /[;"\r\n]|^\s|\s$/ : /[,"\r\n]|^\s|\s$/;
    var colonnes = 0;
    var i;
    var j;
    for (i = 0; i < lignes.length; i++) {
      if (lignes[i].length > colonnes) colonnes = lignes[i].length;
    }
    var codes = reglage.figer ? colonnesDeCodes(lignes, colonnes) : null;
    var sortie = new Array(lignes.length);
    for (i = 0; i < lignes.length; i++) {
      var ligne = lignes[i];
      var cellules = new Array(colonnes);
      for (j = 0; j < colonnes; j++) {
        var v = j < ligne.length ? String(ligne[j]) : '';
        if (codes) v = pourExcel(v, codes[j]);
        // Un fichier qui commence par « ID » est pris par Excel pour un
        // fichier SYLK : des guillemets évitent ce message d'erreur.
        var sylk = i === 0 && j === 0 && v.slice(0, 2) === 'ID';
        cellules[j] = (sylk || aProteger.test(v)) ? '"' + v.replace(/"/g, '""') + '"' : v;
      }
      sortie[i] = cellules.join(separateur);
    }
    return (reglage.bom ? '\uFEFF' : '') + sortie.join('\r\n') + (lignes.length > 0 ? '\r\n' : '');
  }

  // « relevé.txt » → « relevé.csv » ; un nom sans extension reçoit « .csv ».
  // Une extension en double (« notes.txt.txt », fréquent quand Windows cache
  // les extensions) est retirée entièrement.
  function nomFichierCsv(nomOriginal) {
    var nom = String(nomOriginal || '').replace(/^.*[\\\/]/, '');
    nom = nom.replace(/(\.(txt|text|csv|tsv|tab|dat|prn|asc))+$/i, '');
    nom = nom.replace(/[\s.]+$/, '');
    return (nom || 'fichier') + '.csv';
  }

  /* ------------------------------------------------------------------ */
  /* Fonctions principales                                              */
  /* ------------------------------------------------------------------ */

  function echec(code) {
    return { ok: false, erreur: code };
  }

  // Lit les octets d'un fichier et prépare le tableau (sans générer le CSV).
  // options.separateur : un code de SEPARATEURS, ou « auto » (par défaut).
  // En cas de problème : { ok: false, erreur: 'vide' | 'lignes-vides' | 'binaire' | 'rtf' }.
  function analyser(octets, options) {
    octets = enOctets(octets);
    if (octets.length === 0) return echec('vide');
    var info = detecterEncodage(octets);
    var nature = natureContenu(octets, info);
    if (nature !== 'texte') return echec(nature);
    var decodage = decoder(octets, info);
    var texte = decodage.texte;
    if (texte.length === 0) return echec('vide');
    // Rien que des espaces, des retours à la ligne, des séparateurs ou des guillemets.
    if (/^[\s;,|"]*$/.test(texte)) return echec('lignes-vides');
    var demande = options && options.separateur;
    var automatique = !demande || demande === 'auto';
    var separateur = automatique ? detecterSeparateur(texte) : demande;
    var resultat = decouper(texte, separateur);
    if (resultat.lignes.length === 0) return echec('lignes-vides');
    return {
      ok: true,
      texte: texte,
      encodage: decodage.encodage,
      bom: decodage.bom,
      caracteresIllisibles: decodage.illisibles,
      encodagesMelanges: decodage.melange,
      encodageDouteux: decodage.douteux,
      // Fichier terminé par plusieurs octets nuls, ou coupé au milieu d'un
      // caractère : sa fin est sans doute abîmée.
      finAbimee: finAbimee(octets, info) || decodage.finCoupee,
      separateur: separateur,
      separateurAutomatique: automatique,
      lignes: resultat.lignes,
      colonnes: resultat.colonnes,
      lignesVides: resultat.lignesVides,
      longueurMax: resultat.longueurMax
    };
  }

  // Conversion complète : octets du fichier .txt → texte du fichier CSV.
  // options : { separateur: 'auto' | code, format: 'excel' (par défaut) | 'standard' }.
  function convertir(octets, options) {
    var resultat = analyser(octets, options);
    if (!resultat.ok) return resultat;
    resultat.format = (options && options.format) || 'excel';
    resultat.csv = genererCsv(resultat.lignes, resultat.format);
    return resultat;
  }

  return {
    SEPARATEURS: SEPARATEURS.slice(),
    LIMITES_EXCEL: {
      lignes: LIMITES_EXCEL.lignes,
      colonnes: LIMITES_EXCEL.colonnes,
      caracteres: LIMITES_EXCEL.caracteres
    },
    analyser: analyser,
    convertir: convertir,
    detecterEncodage: detecterEncodage,
    decoder: decoder,
    decouper: decouper,
    detecterSeparateur: detecterSeparateur,
    genererCsv: genererCsv,
    nomFichierCsv: nomFichierCsv
  };
}));
