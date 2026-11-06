'use strict';

// Tests de la logique de conversion. Lancer avec : node --test
const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const C = require('../js/conversion.js');

const encodeur = new TextEncoder();
const utf8 = (texte) => encodeur.encode(texte);
const octets = (...liste) => Uint8Array.from(liste);
const concat = (...parties) => Uint8Array.from(parties.flatMap((p) => Array.from(p)));

// Encode un texte en Windows-1252 (caractères latins et quelques signes courants).
const SPECIAUX_1252 = { '€': 0x80, '’': 0x92, '“': 0x93, '”': 0x94, 'Œ': 0x8C, 'œ': 0x9C, '…': 0x85, '–': 0x96, 'Š': 0x8A, 'Ž': 0x8E };
function windows1252(texte) {
  return Uint8Array.from(Array.from(texte, (ch) => {
    if (SPECIAUX_1252[ch] !== undefined) return SPECIAUX_1252[ch];
    const code = ch.codePointAt(0);
    if (code > 0xFF || (code >= 0x80 && code < 0xA0)) throw new Error('Caractère non géré : ' + ch);
    return code;
  }));
}

function utf16le(texte, avecBom) {
  const corps = Buffer.from(texte, 'utf16le');
  return new Uint8Array(avecBom ? Buffer.concat([Buffer.from([0xFF, 0xFE]), corps]) : corps);
}

const lignesDe = (texte, separateur) => C.decouper(texte, separateur).lignes;

// Ce qu'Excel affiche pour une valeur du CSV : ="…" (morceaux reliés par &)
// est un texte écrit sous forme de formule.
function afficheParExcel(valeur) {
  if (!valeur.startsWith('="')) return valeur;
  const morceaux = valeur.slice(1).match(/"(?:[^"]|"")*"/g);
  assert.equal(morceaux.join('&'), valeur.slice(1), 'formule texte mal formée : ' + valeur);
  const textes = morceaux.map((m) => m.slice(1, -1).replace(/""/g, '"'));
  // Excel limite chaque texte d'une formule à 255 caractères.
  for (const t of textes) assert.ok(t.length <= 255, 'morceau de ' + t.length + ' caractères');
  return textes.join('');
}

// Cellules du CSV généré, telles qu'Excel les afficherait.
function cellulesExcel(lignes) {
  const csv = C.genererCsv(lignes, 'excel');
  return C.decouper(csv.slice(1), 'point-virgule').lignes.map((l) => l.map(afficheParExcel));
}

describe('Détection du séparateur', () => {
  it('reconnaît la tabulation', () => {
    assert.equal(C.detecterSeparateur('Nom\tVille\tÂge\nDupont\tParis\t71\nMartin\tLyon\t68\n'), 'tabulation');
  });

  it('reconnaît le point-virgule', () => {
    assert.equal(C.detecterSeparateur('Nom;Ville;Âge\nDupont;Paris;71\nMartin;Lyon;68\n'), 'point-virgule');
  });

  it('reconnaît la virgule', () => {
    assert.equal(C.detecterSeparateur('nom,ville,age\nDupont,Paris,71\nMartin,Lyon,68\n'), 'virgule');
  });

  it('reconnaît la barre verticale', () => {
    assert.equal(C.detecterSeparateur('Nom|Ville|Âge\nDupont|Paris|71\nMartin|Lyon|68\n'), 'barre');
  });

  it('reconnaît plusieurs espaces (colonnes alignées)', () => {
    const texte = 'Nom        Ville           Âge\n' +
                  'Dupont     Paris           71\n' +
                  'Le Gall    Saint-Malo      68\n';
    assert.equal(C.detecterSeparateur(texte), 'espaces');
    assert.deepEqual(lignesDe(texte, 'espaces'), [
      ['Nom', 'Ville', 'Âge'],
      ['Dupont', 'Paris', '71'],
      ['Le Gall', 'Saint-Malo', '68']
    ]);
  });

  it('laisse une seule colonne quand il n’y a pas de séparateur', () => {
    const texte = 'Bonjour, ceci est une note.\nUne autre ligne sans rien\nEncore une, avec, deux virgules.\n';
    assert.equal(C.detecterSeparateur(texte), 'aucun');
    assert.deepEqual(lignesDe(texte, 'aucun').map((l) => l.length), [1, 1, 1]);
  });

  it('préfère le point-virgule quand les nombres ont une virgule décimale', () => {
    assert.equal(C.detecterSeparateur('Article;Prix;Rayon\nPain;1,20;Boulangerie\nLait;0,95;Crèmerie\n'), 'point-virgule');
  });

  it('préfère les espaces pour des colonnes alignées contenant des nombres à virgule', () => {
    const texte = 'Article       Prix\nPain          1,20\nLait          0,95\nBeurre        2,35\n';
    assert.equal(C.detecterSeparateur(texte), 'espaces');
  });

  it('reconnaît la virgule quand des valeurs entre guillemets contiennent des points-virgules', () => {
    const texte = 'nom,adresse\n"Dupont","1 rue A; bât. B"\n"Martin","2 rue C; bât. D"\n';
    assert.equal(C.detecterSeparateur(texte), 'virgule');
  });

  it('reconnaît le point-virgule quand des valeurs entre guillemets contiennent des virgules', () => {
    const texte = 'Nom;Adresse\n"Dupont, Jean";"1 rue A, Paris"\n"Martin, Paul";"2 rue C, Lyon"\n';
    assert.equal(C.detecterSeparateur(texte), 'point-virgule');
  });

  it('reconnaît la tabulation quand les valeurs contiennent des virgules', () => {
    assert.equal(C.detecterSeparateur('Nom\tAdresse\nDupont, Jean\t1 rue A, Paris\nMartin, Paul\t2 rue C, Lyon\n'), 'tabulation');
  });

  it('permet de choisir le séparateur à la main', () => {
    const r = C.analyser(utf8('a;b,c\nd;e,f\n'), { separateur: 'virgule' });
    assert.equal(r.separateur, 'virgule');
    assert.equal(r.separateurAutomatique, false);
    assert.deepEqual(r.lignes, [['a;b', 'c'], ['d;e', 'f']]);
  });
});

describe('Détection : virgules décimales, tabulations du Bloc-notes, exports complétés par des espaces', () => {
  it('une liste de montants n’est pas coupée à la virgule décimale', () => {
    for (const texte of ['12,40\n1 250,00\n-23,15\n67,80\n', 'Pain 1,20\nLait 0,95\nBeurre 2,35\nCafé 3,99\n',
      'Courses 45,30 €\nEssence 60,00 €\n']) {
      assert.equal(C.detecterSeparateur(texte), 'aucun', JSON.stringify(texte));
    }
  });

  it('une liste tapée à la main avec des points-virgules irréguliers reste au point-virgule', () => {
    const texte = 'Pain;1,20\nLait;demi-écrémé;0,95\nBeurre;doux;2,35\nCafé;3,99\nŒufs;bio;x6;2,80\n';
    assert.equal(C.detecterSeparateur(texte), 'point-virgule');
    assert.deepEqual(lignesDe(texte, 'point-virgule')[1], ['Lait', 'demi-écrémé', '0,95']);
  });

  it('un vrai fichier à virgules reste à la virgule (en-tête, petits nombres, nombres à point)', () => {
    for (const texte of ['id,age\n1,20\n2,35\n3,41\n', 'x,y\n10,20\n30,40\n', '1,2,3\n4,5,6\n7,8,9\n',
      'bread,1.25\nmilk,0.95\n', 'nom,note\nZoé,12\nLéa,15\nTom,09\n']) {
      assert.equal(C.detecterSeparateur(texte), 'virgule', JSON.stringify(texte));
    }
  });

  it('colonnes alignées avec la touche Tab dans le Bloc-notes (une ou deux tabulations)', () => {
    const noms = ['Durand', 'Martin', 'Bernard', 'Petit', 'Robert', 'Richard', 'Moreau', 'Simon', 'Laurent', 'Michel'];
    let texte = 'Nom\t\tTéléphone\n';
    noms.forEach((nom, k) => { texte += nom + (nom.length < 8 ? '\t\t' : '\t') + '06 39 98 45 ' + (10 + k) + '\n'; });
    texte += 'Rousseau\t01 99 00 33 44\n';
    const r = C.analyser(utf8(texte));
    assert.equal(r.separateur, 'espaces');
    assert.equal(r.colonnes, 2);
    assert.deepEqual(r.lignes[1], ['Durand', '06 39 98 45 10']);
    assert.deepEqual(r.lignes[11], ['Rousseau', '01 99 00 33 44']);
  });

  it('des nombres seuls entre virgules, ou après une date, restent séparés par la virgule', () => {
    for (const texte of ['05/01/2026,12,23,34,41,45,7\n08/01/2026,03,15,22,38,49,2\n12/01/2026,07,11,29,33,47,9\n',
      '12,23,34,41,45\n03,15,22,38,49\n07,11,29,33,47\n', '05/01/2026,62\n06/01/2026,61\n07/01/2026,63\n']) {
      assert.equal(C.detecterSeparateur(texte), 'virgule', JSON.stringify(texte));
    }
    assert.equal(C.detecterSeparateur('12,40 15,30\n8,20 9,10\n1,50 2,75\n'), 'aucun');
    let idAge = 'id,age\n';
    for (let i = 1; i <= 10; i++) idAge += i + ',' + (20 + i) + '\n';
    assert.equal(C.detecterSeparateur(idAge), 'virgule');
  });

  it('un fichier à tabulations tapé à la main (espace avant la tabulation, lignes inégales) reste à la tabulation', () => {
    assert.equal(C.detecterSeparateur('Nom \tAdresse \tVille \tRemarque\nDupont \t12, rue des Lilas \tParis\n' +
      'Martin \t3, place du Marché \tLyon \tvoisin\nDurand \t5, avenue Foch\n' +
      'Petit \t7, rue Haute \tNice \tclés chez elle \tappeler le soir\nRoux \t9, bd Victor-Hugo \tTours\n'), 'tabulation');
    assert.equal(C.detecterSeparateur('Nom\tPrénom\tVille\nDupont\t\tParis\nMartin\tJean\nDurand\tLéa\nPetit\t\tNice\n'), 'tabulation');
  });

  it('montant collé à son libellé par « : » ou « / » : pas coupé à la virgule décimale', () => {
    assert.equal(C.detecterSeparateur('Loyer:650,00\nEDF:45,67\nEau:23,15\nTéléphone:19,99\n'), 'aucun');
    assert.equal(C.detecterSeparateur('Tomates 1kg/2,99\nPommes 1kg/3,49\nPoires 1kg/2,79\n'), 'aucun');
  });

  it('heures, mois ou dates suivis d’une valeur : séparés par la virgule', () => {
    for (const texte of ['08:00,12\n09:00,15\n10:00,18\n11:00,21\n', '01/2026,45\n02/2026,52\n03/2026,48\n',
      '05/01,12\n06/01,15\n07/01,18\n']) {
      assert.equal(C.detecterSeparateur(texte), 'virgule', JSON.stringify(texte));
    }
  });

  it('colonnes alignées dont les libellés contiennent un point-virgule : restent aux espaces', () => {
    assert.equal(C.detecterSeparateur('Article              Rayon; allée      Prix\nPain                 Boulangerie; 1    1,20\n' +
      'Lait                 Crèmerie; 2       0,95\nBeurre               Crèmerie; 2       2,35\n'), 'espaces');
  });

  it('un export à tabulations d’Excel, avec des cases vides, reste à la tabulation', () => {
    const texte = 'Nom\tPrénom\tTéléphone\nDupont\t\t0612\nMartin\tJean\t0613\nDurand\t\t0614\n';
    assert.equal(C.detecterSeparateur(texte), 'tabulation');
    assert.deepEqual(lignesDe(texte, 'tabulation')[1], ['Dupont', '', '0612']);
  });

  it('export aux valeurs entre guillemets complétées par des espaces', () => {
    const pointVirgule = '"CB BOULANGERIE      ";"-12,40    ";"05/01/2026"\n' +
                         '"PRLV ASSURANCE      ";"-45,67    ";"06/01/2026"\n' +
                         '"VIR PENSION         ";"1250,00   ";"07/01/2026"\n';
    assert.equal(C.detecterSeparateur(pointVirgule), 'point-virgule');
    const virgule = '"DUPONT              ","JEAN          ","PARIS         ",12.40\n' +
                    '"MARTIN              ","PAUL          ","LYON          ",3.50\n';
    assert.equal(C.detecterSeparateur(virgule), 'virgule');
    assert.equal(lignesDe(virgule, 'virgule')[0].length, 4);
  });
});

describe('Colonnes alignées par des espaces : cases vides', () => {
  it('relevé avec colonnes Débit et Crédit : chaque montant reste dans sa colonne', () => {
    const texte = 'Date        Libellé                     Débit       Crédit\n' +
                  '05/01/2026  Boulangerie du Marché       12,40\n' +
                  '07/01/2026  Pension de retraite                     1 250,00\n' +
                  '12/01/2026  Pharmacie Centrale          23,15\n';
    assert.equal(C.detecterSeparateur(texte), 'espaces');
    // (Une ligne recalée s'arrête à sa dernière valeur : le CSV complète les lignes.)
    assert.deepEqual(lignesDe(texte, 'espaces'), [
      ['Date', 'Libellé', 'Débit', 'Crédit'],
      ['05/01/2026', 'Boulangerie du Marché', '12,40'],
      ['07/01/2026', 'Pension de retraite', '', '1 250,00'],
      ['12/01/2026', 'Pharmacie Centrale', '23,15']
    ]);
  });

  it('montants alignés à droite', () => {
    const texte = 'Date        Libellé                    Débit      Crédit\n' +
                  '05/01/2026  Boulangerie                12,40\n' +
                  '07/01/2026  Pension de retraite                1 250,00\n';
    assert.deepEqual(lignesDe(texte, 'espaces').slice(1), [
      ['05/01/2026', 'Boulangerie', '12,40'],
      ['07/01/2026', 'Pension de retraite', '', '1 250,00']
    ]);
  });

  it('première case vide (ligne en retrait) et case vide au milieu', () => {
    assert.deepEqual(lignesDe('Date        Libellé                Montant\n' +
                              '05/01/2026  Boulangerie            -12,40\n' +
                              '            Report                 100,00\n', 'espaces')[2], ['', 'Report', '100,00']);
    assert.deepEqual(lignesDe('Article          Rayon           Quantité   Prix\n' +
                              'Pain             Boulangerie     2          1,20\n' +
                              'Piles                            4          5,90\n', 'espaces')[2], ['Piles', '', '4', '5,90']);
  });

  it('ne déplace rien quand l’alignement n’est pas sûr', () => {
    const textes = [
      'Article  Prix  Rayon\nPain  1,20  Boulangerie\nBeurre doux  2,35  Crèmerie\n',
      'Beurre doux  2,35  Crèmerie\nLait  0,95\n',
      'Date        Libellé                         Montant\n05/01/2026  Boulangerie                      -12,40\n' +
        '06/01/2026  Prélèvement assurance habitation -45,67\n',
      'Nom        Ville      Âge\nDupont     Paris      71\nDe La Fontaine-Martin  Lyon  68\n'
    ];
    const attendus = [
      [['Article', 'Prix', 'Rayon'], ['Pain', '1,20', 'Boulangerie'], ['Beurre doux', '2,35', 'Crèmerie']],
      [['Beurre doux', '2,35', 'Crèmerie'], ['Lait', '0,95']],
      [['Date', 'Libellé', 'Montant'], ['05/01/2026', 'Boulangerie', '-12,40'],
        ['06/01/2026', 'Prélèvement assurance habitation -45,67']],
      [['Nom', 'Ville', 'Âge'], ['Dupont', 'Paris', '71'], ['De La Fontaine-Martin', 'Lyon', '68']]
    ];
    textes.forEach((texte, k) => assert.deepEqual(lignesDe(texte, 'espaces'), attendus[k]));
  });

  it('une valeur entre deux colonnes, ou des colonnes qui se chevauchent : rien ne bouge', () => {
    assert.deepEqual(lignesDe('Produit      Prix     Stock\nPain         1,20     10\nLait         0,95     25\nŒufs              x6\n',
      'espaces')[3], ['Œufs', 'x6']);
    assert.deepEqual(lignesDe('Beurre doux     2,35   Crèmerie\nCafé moulu  1,20   Boulangerie\nBeurre doux      0,95  Épicerie\n' +
      'Beurre doux             2,35\n', 'espaces')[3], ['Beurre doux', '2,35']);
  });

  it('une ligne de total, un titre ou un libellé à double espace ne dérangent pas un tableau complet', () => {
    const tableaux = [
      // ligne de total plus large que le tableau, en bas
      'Date opération  Libellé                      Montant\n05/01/2026      Boulangerie                  -12,40\n' +
        '07/01/2026      Virement pension retraite    1 250,00\n12/01/2026      Pharmacie                    -23,15\n\n' +
        'Total des débits :  35,55    Total des crédits :  1 250,00\n',
      // titre plus large que le tableau, en haut
      'RELEVÉ  N° 0123456  du 01/01/2026  au 31/01/2026\nDate        Libellé                Montant\n' +
        '05/01/2026  CB BOULANGERIE         -12,40\n07/01/2026  VIR PENSION RETRAITE   1 250,00\n08/01/2026  EDF                    -45,67\n',
      // libellé contenant deux espaces au milieu du tableau
      'Date        Libellé                     Montant\n05/01/2026  Boulangerie du Marché       -12,40\n' +
        '06/01/2026  VIR SEPA  M. DUPONT JEAN    250,00\n07/01/2026  EDF                         -45,67\n08/01/2026  Pharmacie                   -23,15\n',
      // même chose avec un en-tête court et des libellés courts
      'Date        Lib.                    Montant\n05/01/2026  Café                    -814,12\n01/01/2026  VIR SEPA  DUPONT        -505,37\n' +
        '01/01/2026  Pain                    456,79\n07/01/2026  Bus                     529,76\n'
    ];
    const attendus = [
      [['Date opération', 'Libellé', 'Montant'], ['05/01/2026', 'Boulangerie', '-12,40'],
        ['07/01/2026', 'Virement pension retraite', '1 250,00'], ['12/01/2026', 'Pharmacie', '-23,15'],
        ['Total des débits :', '35,55', 'Total des crédits :', '1 250,00']],
      [['RELEVÉ', 'N° 0123456', 'du 01/01/2026', 'au 31/01/2026'], ['Date', 'Libellé', 'Montant'],
        ['05/01/2026', 'CB BOULANGERIE', '-12,40'], ['07/01/2026', 'VIR PENSION RETRAITE', '1 250,00'], ['08/01/2026', 'EDF', '-45,67']],
      [['Date', 'Libellé', 'Montant'], ['05/01/2026', 'Boulangerie du Marché', '-12,40'],
        ['06/01/2026', 'VIR SEPA', 'M. DUPONT JEAN', '250,00'], ['07/01/2026', 'EDF', '-45,67'], ['08/01/2026', 'Pharmacie', '-23,15']],
      [['Date', 'Lib.', 'Montant'], ['05/01/2026', 'Café', '-814,12'], ['01/01/2026', 'VIR SEPA', 'DUPONT', '-505,37'],
        ['01/01/2026', 'Pain', '456,79'], ['07/01/2026', 'Bus', '529,76']]
    ];
    // Les montants restent sous « Montant » : rien n'est déplacé (la ligne
    // au double espace, elle, est coupée comme avant).
    tableaux.forEach((texte, k) => assert.deepEqual(lignesDe(texte, 'espaces'), attendus[k]));
  });

  it('Débit / Crédit sous une ligne de coordonnées imprimée sur la grille du tableau', () => {
    const tableau = 'Date        Libellé                Débit       Crédit\n05/01/2026  CB BOULANGERIE         12,40\n' +
      '07/01/2026  VIR PENSION RETRAITE               1 250,00\n08/01/2026  PRLV EDF               45,67\n';
    for (const tete of ['Titulaire   M. DUPONT Jean         Compte n° 0123456789',
      'Titulaire   M. DUPONT Jean         Cpte 0123', 'Titulaire   M. DUPONT Jean         Agence Lyon']) {
      assert.deepEqual(lignesDe(tete + '\n\n' + tableau, 'espaces')[3], ['07/01/2026', 'VIR PENSION RETRAITE', '', '1 250,00'], tete);
    }
  });

  it('Débit / Crédit suivi d’une ligne de totaux aussi large, ou d’une phrase de conclusion', () => {
    const releve = 'Date        Libellé                Débit       Crédit\n05/01/2026  CB BOULANGERIE         12,40\n' +
      '07/01/2026  VIR PENSION RETRAITE               1 250,00\n08/01/2026  PRLV EDF               45,67\n\n';
    for (const fin of ['Total des débits :  58,07    Total des crédits :  1 250,00', 'Document à conserver.  Merci de votre confiance.']) {
      assert.deepEqual(lignesDe(releve + fin + '\n', 'espaces')[2], ['07/01/2026', 'VIR PENSION RETRAITE', '', '1 250,00'], fin);
    }
  });

  it('Débit / Crédit avec des lignes d’en-tête au-dessus, ou un pied de page très large', () => {
    assert.deepEqual(lignesDe('Titulaire :  M. DUPONT Jean\nCompte n° :  0123456789\n\nDate        Libellé                Débit       Crédit\n' +
      '05/01/2026  CB BOULANGERIE         12,40\n07/01/2026  VIR PENSION RETRAITE               1 250,00\n08/01/2026  PRLV EDF               45,67\n',
    'espaces'), [
      ['Titulaire :', 'M. DUPONT Jean'], ['Compte n° :', '0123456789'], ['Date', 'Libellé', 'Débit', 'Crédit'],
      ['05/01/2026', 'CB BOULANGERIE', '12,40'], ['07/01/2026', 'VIR PENSION RETRAITE', '', '1 250,00'], ['08/01/2026', 'PRLV EDF', '45,67']
    ]);
    const avecPied = lignesDe('Date        Libellé                Débit       Crédit\n05/01/2026  CB BOULANGERIE         12,40\n' +
      '07/01/2026  VIR PENSION RETRAITE               1 250,00\n08/01/2026  PRLV EDF               45,67\n\n' +
      'Ancien solde :  1 000,00    Nouveau solde :  2 191,93    Page  1  /  1\n', 'espaces');
    assert.deepEqual(avecPied[2], ['07/01/2026', 'VIR PENSION RETRAITE', '', '1 250,00']);
  });

  it('accents décomposés (texte venu d’un Mac) et emoji : la position à l’écran reste juste', () => {
    const nfd = ('Date        Libelle                              Debit    Credit\n' +
      '09/01/2026  Prélèvement électricité régularisé   67,80\n' +
      '10/01/2026  Pension                                       1 250,00\n').normalize('NFD');
    const lignes = lignesDe(nfd, 'espaces');
    assert.equal(lignes[1].length, 3, 'le débit reste sous « Debit »');
    assert.deepEqual(lignes[2], ['10/01/2026', 'Pension', '', '1 250,00']);
    assert.deepEqual(lignesDe('Nom          Ville        Âge\nZoé 😀       Paris        12\nLéa                       9\n', 'espaces')[2],
      ['Léa', '', '9']);
  });

  it('une ligne très large au-dessus de nombreuses lignes courtes reste rapide et sans cases inutiles', () => {
    const texte = 'a  '.repeat(1000) + '\n' + 'a  b\n'.repeat(50000);
    const debut = process.hrtime.bigint();
    const lignes = lignesDe(texte, 'espaces');
    assert.ok(Number(process.hrtime.bigint() - debut) / 1e6 < 2000);
    assert.ok(lignes.slice(1).every((l) => l.length === 2));
  });
});

describe('Encodage et accents', () => {
  const texte = 'Prénom;Ville\nHélène;Besançon\nFrançois;Évreux\nNoël;Châlons\n';

  it('lit un fichier UTF-8', () => {
    const r = C.analyser(utf8(texte));
    assert.equal(r.encodage, 'utf-8');
    assert.equal(r.caracteresIllisibles, false);
    assert.equal(r.encodagesMelanges, false);
    assert.deepEqual(r.lignes[1], ['Hélène', 'Besançon']);
  });

  it('lit un fichier Windows-1252 (é, è, à, ç, œ, €, apostrophe et guillemets courbes)', () => {
    const r = C.analyser(windows1252('Article;Prix\nBœuf à la crème;12,50 €\nGâteau “maison”;3 €\nL’été…;1 €\n'));
    assert.equal(r.encodage, 'windows-1252');
    assert.equal(r.encodageDouteux, false);
    assert.deepEqual(r.lignes, [['Article', 'Prix'], ['Bœuf à la crème', '12,50 €'], ['Gâteau “maison”', '3 €'], ['L’été…', '1 €']]);
  });

  it('décode les octets 0x80 à 0x9F selon Windows-1252, pas Latin-1', () => {
    const d = C.decoder(octets(0x80, 0x8C, 0x9C, 0x92, 0x85, 0x96, 0x20, 0xE9));
    assert.equal(d.encodage, 'windows-1252');
    assert.equal(d.texte, '€Œœ’…– é');
  });

  it('signale un fichier d’un ancien logiciel (DOS, Mac) au lieu de faire disparaître des lettres', () => {
    // « Crème;Été » enregistré par un logiciel DOS (page de code 850) : è = 0x8A, É = 0x90, é = 0x82.
    const r = C.analyser(concat(utf8('Nom;Saison\nCr'), [0x8A], utf8('me;'), [0x90], utf8('t'), [0x82], utf8('\n')));
    assert.equal(r.encodage, 'windows-1252');
    assert.equal(r.encodageDouteux, true);
    assert.equal(r.lignes[1][1].charAt(0), '\uFFFD');
  });

  it('retire le BOM UTF-8', () => {
    const r = C.analyser(octets(0xEF, 0xBB, 0xBF, ...utf8('Nom;Âge\nZoé;7\n')));
    assert.equal(r.encodage, 'utf-8');
    assert.equal(r.bom, true);
    assert.equal(r.lignes[0][0], 'Nom');
  });

  it('lit un fichier UTF-16 avec BOM (export « Texte Unicode » d’Excel)', () => {
    const r = C.analyser(utf16le('Nom\tVille\nHélène\tBesançon\n', true));
    assert.equal(r.encodage, 'utf-16le');
    assert.equal(r.separateur, 'tabulation');
    assert.deepEqual(r.lignes[1], ['Hélène', 'Besançon']);
  });

  it('lit un fichier UTF-16 sans BOM', () => {
    const r = C.analyser(utf16le('Nom;Ville\nZoé;Nîmes\n', false));
    assert.equal(r.encodage, 'utf-16le');
    assert.deepEqual(r.lignes[1], ['Zoé', 'Nîmes']);
  });

  it('fichier UTF-8 contenant quelques octets Windows-1252 : aucun accent perdu', () => {
    const mixte = concat(utf8('Prénom;Ville\nHélène;Besançon\nFrançois;'), [0xC9], utf8('vreux\n'));
    const r = C.analyser(mixte);
    assert.equal(r.encodage, 'utf-8');
    assert.equal(r.encodagesMelanges, true);
    assert.equal(r.caracteresIllisibles, false);
    assert.deepEqual(r.lignes[1], ['Hélène', 'Besançon']);
    assert.deepEqual(r.lignes[2], ['François', 'Évreux']);
  });

  it('en-tête UTF-8 suivi de données Windows-1252, ou deux exports de même taille mis bout à bout', () => {
    const entete = C.analyser(concat(utf8('Prénom;Ville\n'), windows1252('Hélène;Besançon\nFrançois;Évreux\nZoé;Nîmes\n')));
    assert.equal(entete.encodagesMelanges, true);
    assert.deepEqual(entete.lignes[0], ['Prénom', 'Ville']);
    assert.deepEqual(entete.lignes[2], ['François', 'Évreux']);
    const releve = 'Date;Libellé;Débit;Crédit\n05/01;Pharmacie;12,40;\n07/01;Pension;;1 250,00\n';
    const boutABout = C.analyser(concat(windows1252(releve), utf8(releve)));
    assert.equal(boutABout.encodagesMelanges, true);
    assert.deepEqual(boutABout.lignes[0], ['Date', 'Libellé', 'Débit', 'Crédit']);
    assert.deepEqual(boutABout.lignes[3], ['Date', 'Libellé', 'Débit', 'Crédit']);
  });

  it('un texte Windows-1252 dont deux octets forment par hasard un caractère UTF-8 rare reste en Windows-1252', () => {
    // « É» » (C9 BB) et « É » + espace insécable (C9 A0) sont des caractères UTF-8 valides, mais pas français.
    const r = C.analyser(windows1252('Titre;Note\r\nRÉSUMÉ\u00A0: été;«CAFÉ»\r\n'));
    assert.equal(r.encodage, 'windows-1252');
    assert.equal(r.encodagesMelanges, false);
    assert.deepEqual(r.lignes[1], ['RÉSUMÉ\u00A0: été', '«CAFÉ»']);
  });

  it('une ligne UTF-8 dont le seul accent est « € » ou « ’ », suivie de lignes Windows-1252', () => {
    for (const s of ['Montant en €', 'L’été']) {
      const r = C.analyser(concat(utf8(s + '\r\n'), windows1252('Café;Thé\r\nCrème;Pâté\r\n')));
      assert.equal(r.encodagesMelanges, true, s);
      assert.deepEqual(r.lignes[0], [s]);
      assert.deepEqual(r.lignes[1], ['Café', 'Thé']);
    }
  });

  it('un fichier mélangé qui se termine par « é » ou « é… » en Windows-1252 : rien n’est perdu', () => {
    const r = C.analyser(concat(utf8('Prénom;Ville\n'), windows1252('Zoé;Café')));
    assert.deepEqual(r.lignes[1], ['Zoé', 'Café']);
    assert.equal(r.finAbimee, false);
    const r2 = C.analyser(concat(utf8('Prénom;Note\r\n'), windows1252('Zoé;café…')));
    assert.deepEqual(r2.lignes[1], ['Zoé', 'café…']);
    assert.equal(r2.finAbimee, false);
    for (const fin of [[0xE0, 0x80], [0xED, 0xA0]]) {
      assert.equal(C.analyser(concat(utf8('Prénom;Ville\nZoé;'), fin)).finAbimee, false, fin.join(' '));
    }
  });

  it('un caractère coupé suivi d’octets nuls (coupure de courant) : « \uFFFD » et fin abîmée', () => {
    const r = C.analyser(concat(utf8('Libellé;Montant\r\nPension;1 250,00 €\r\nLoyer;-650,00 '), [0xE2, 0x82], new Uint8Array(512)));
    assert.deepEqual(r.lignes[2], ['Loyer', '-650,00 \uFFFD']);
    assert.equal(r.caracteresIllisibles, true);
    assert.equal(r.finAbimee, true);
    assert.equal(r.encodagesMelanges, false);
  });

  it('un texte Windows-1252 ordinaire n’est ni « mélangé » ni « douteux »', () => {
    const r = C.analyser(windows1252('Libellé;Note\nL’été “chaud”…;12 €\nBœuf;œuf\u00A0frais\nVIR SEPA DUŠAN PETROVIC;3 €\n'));
    assert.equal(r.encodage, 'windows-1252');
    assert.equal(r.encodagesMelanges, false);
    assert.equal(r.encodageDouteux, false);
  });

  it('signale un texte DOS où « é » se trouve au milieu des mots (« D‚bit »)', () => {
    const r = C.analyser(concat(utf8('Libell'), [0x82], utf8(';D'), [0x82], utf8('bit;Cr'), [0x82], utf8('dit\nPain;1,20;\n')));
    assert.equal(r.encodageDouteux, true);
  });

  it('fichier UTF-8 coupé au milieu de son dernier caractère : « \uFFFD » et fin abîmée', () => {
    const r = C.analyser(concat(utf8('Libellé;Montant\nLoyer;-650,00 '), [0xE2, 0x82]));
    assert.deepEqual(r.lignes[1], ['Loyer', '-650,00 \uFFFD']);
    assert.equal(r.finAbimee, true);
    assert.equal(r.caracteresIllisibles, true);
    assert.equal(r.encodagesMelanges, false);
  });

  it('fichier UTF-16 terminé par des caractères nuls : fin abîmée', () => {
    const abime = C.analyser(concat(utf16le('Libellé;Montant\r\nPension;1 250,00\r\nLoyer;-65', true), new Uint8Array(512)));
    assert.equal(abime.finAbimee, true);
    assert.deepEqual(abime.lignes[2], ['Loyer', '-65']);
    assert.equal(C.analyser(utf16le('Libellé;Montant\r\nPension;1 250,00\r\n', true)).finAbimee, false);
    assert.equal(C.analyser(concat(utf16le('Libellé;Montant\r\nPension;1 250,00', true), [0, 0])).finAbimee, false);
    // Un caractère nul et un octet orphelin : pas une fin abîmée.
    assert.equal(C.analyser(concat(utf16le('Libellé;Montant\r\nPension;1 250,00', true), [0, 0, 0])).finAbimee, false);
  });

  it('BOM UTF-8 suivi d’un texte Windows-1252', () => {
    const r = C.analyser(concat([0xEF, 0xBB, 0xBF], windows1252('Hélène;Besançon\nZoé;Nîmes\n')));
    assert.deepEqual(r.lignes, [['Hélène', 'Besançon'], ['Zoé', 'Nîmes']]);
  });

  it('le CSV « Excel » commence par le BOM et garde les accents en UTF-8', () => {
    const r = C.convertir(windows1252('Prénom;Ville\nHélène;Besançon\n'));
    const sortie = utf8(r.csv);
    assert.deepEqual(Array.from(sortie.slice(0, 3)), [0xEF, 0xBB, 0xBF]);
    assert.equal(r.csv, '\uFEFFPrénom;Ville\r\nHélène;Besançon\r\n');
    assert.ok(Buffer.from(sortie).includes(Buffer.from([0xC3, 0xA9])), 'é encodé en UTF-8');
  });

  it('le CSV standard utilise la virgule, sans BOM', () => {
    const r = C.convertir(utf8('Prénom;Ville\nHélène;Besançon\n'), { format: 'standard' });
    assert.equal(r.csv, 'Prénom,Ville\r\nHélène,Besançon\r\n');
  });
});

describe('Guillemets, séparateurs et retours à la ligne dans les valeurs (RFC 4180)', () => {
  const source = 'Nom;Remarque;Adresse\n' +
                 '"Dupont; Jean";"Il a dit ""bonjour""";"12 rue des Lilas\n75001 Paris"\n' +
                 'Martin;"";"Rue de la Gare\r\nLyon"\n';

  it('lit les séparateurs, guillemets doublés et retours à la ligne entre guillemets', () => {
    assert.deepEqual(lignesDe(source, 'point-virgule'), [
      ['Nom', 'Remarque', 'Adresse'],
      ['Dupont; Jean', 'Il a dit "bonjour"', '12 rue des Lilas\n75001 Paris'],
      ['Martin', '', 'Rue de la Gare\nLyon']
    ]);
  });

  it('un retour chariot seul (ancien Mac) entre guillemets devient un retour à la ligne', () => {
    assert.deepEqual(lignesDe('a;"ligne 1\rligne 2";b\n', 'point-virgule'), [['a', 'ligne 1\nligne 2', 'b']]);
  });

  it('protège ces valeurs dans le CSV généré', () => {
    const lignes = lignesDe(source, 'point-virgule');
    assert.equal(C.genererCsv(lignes, 'excel'),
      '\uFEFFNom;Remarque;Adresse\r\n' +
      '"Dupont; Jean";"Il a dit ""bonjour""";"12 rue des Lilas\n75001 Paris"\r\n' +
      'Martin;;"Rue de la Gare\nLyon"\r\n');
    assert.equal(C.genererCsv(lignes, 'standard'),
      'Nom,Remarque,Adresse\r\n' +
      'Dupont; Jean,"Il a dit ""bonjour""","12 rue des Lilas\n75001 Paris"\r\n' +
      'Martin,,"Rue de la Gare\nLyon"\r\n');
  });

  it('met entre guillemets une valeur contenant une virgule en CSV standard', () => {
    assert.equal(C.genererCsv([['Dupont, Jean', '12,5']], 'standard'), '"Dupont, Jean","12,5"\r\n');
  });

  it('accepte des espaces autour d’une valeur entre guillemets', () => {
    assert.deepEqual(lignesDe('a ;  "b;c"  ; d\n', 'point-virgule'), [['a', 'b;c', 'd']]);
  });

  it('laisse tel quel un guillemet au milieu d’une valeur', () => {
    assert.deepEqual(lignesDe('Écran 27";ok\nIl a dit "oui";non\n', 'point-virgule'),
      [['Écran 27"', 'ok'], ['Il a dit "oui"', 'non']]);
  });

  it('garde tels quels des guillemets mal formés, sans rien perdre', () => {
    assert.deepEqual(lignesDe('"Les Misérables" de Victor Hugo;1862\n', 'point-virgule'),
      [['"Les Misérables" de Victor Hugo', '1862']]);
  });

  it('un guillemet jamais fermé ne « mange » pas la suite du fichier', () => {
    assert.deepEqual(lignesDe('"Le Petit Prince;Saint-Exupéry\nDune;Herbert\n', 'point-virgule'),
      [['"Le Petit Prince', 'Saint-Exupéry'], ['Dune', 'Herbert']]);
  });

  it('en mode « rien », les guillemets sont gardés tels quels', () => {
    assert.deepEqual(lignesDe('"a;b"\n"c\n', 'aucun'), [['"a;b"'], ['"c']]);
  });

  it('aller-retour : relire le CSV généré redonne exactement les mêmes valeurs', () => {
    const morceaux = ['a', 'é', ' ', ';', ',', '"', '\n', '\r\n', '\t', '|', 'Zoé', '12,5', '€', 'ID',
      '-', '+', '=', '@', '0', '06000', '- Pain', '😀'];
    let graine = 42;
    const hasard = () => { graine = (graine * 1103515245 + 12345) % 2147483648; return graine / 2147483648; };
    const lignes = [];
    for (let l = 0; l < 200; l++) {
      const ligne = [];
      const largeur = 1 + Math.floor(hasard() * 5);
      for (let c = 0; c < largeur; c++) {
        let valeur = '';
        const longueur = Math.floor(hasard() * 6);
        for (let k = 0; k < longueur; k++) valeur += morceaux[Math.floor(hasard() * morceaux.length)];
        ligne.push(valeur);
      }
      if (ligne[0].trim() === '') ligne[0] = 'L' + l + ligne[0]; // jamais de ligne entièrement vide
      lignes.push(ligne);
    }
    // Lecture de référence : les retours à la ligne entre guillemets deviennent \n.
    const attendu = lignes.map((l) => l.map((v) => v.replace(/\r\n?/g, '\n')));
    const largeurMax = Math.max(...lignes.map((l) => l.length));
    const complete = (l) => l.concat(Array(largeurMax - l.length).fill(''));
    const standard = C.analyser(utf8(C.genererCsv(lignes, 'standard')), { separateur: 'virgule' });
    assert.deepEqual(standard.lignes.map(complete), attendu.map(complete), 'format standard');
    // Excel : un texte sur plusieurs lignes qui commence par = + - @ reçoit une
    // espace devant (le résultat d'une formule ne s'affiche pas sur plusieurs lignes).
    const vuDansExcel = (v) => (v.includes('\n') && /^[=+\-@]/.test(v) ? ' ' + v : v);
    assert.deepEqual(cellulesExcel(lignes).map(complete), attendu.map((l) => complete(l.map(vuDansExcel))), 'format Excel');
  });
});

describe('Excel : les valeurs s’affichent telles qu’elles sont dans le fichier', () => {
  it('une liste à tirets (« - Pain ») n’est pas prise pour une formule', () => {
    const lignes = [['- Pain'], ['-Lait'], ['+ penser au cadeau'], ['= total à voir'], ['@SOMME(1;1)'], ['-1-2']];
    const csv = C.genererCsv(lignes, 'excel');
    assert.ok(csv.includes('"=""- Pain"""'), csv);
    assert.deepEqual(cellulesExcel(lignes), lignes);
  });

  it('les zéros du début, les longs numéros et les numéros internationaux sont gardés', () => {
    const valeurs = ['06000', '0612345678', '000123', '+33612345678', '73282932000074', '4970123456789012'];
    const csv = C.genererCsv([valeurs], 'excel');
    for (const v of valeurs) assert.ok(csv.includes('"=""' + v + '"""'), v);
    assert.deepEqual(cellulesExcel([valeurs]), [valeurs]);
  });

  it('les nombres ordinaires et les dates restent des nombres et des dates', () => {
    const valeurs = ['-12,40', '- 12,40', '1 250,00', '-1 250,00 €', '-5 %', '+5', '0', '0,5', '1234567890', '05/01/2026'];
    assert.equal(C.genererCsv([valeurs], 'excel'), '\uFEFF' + valeurs.join(';') + '\r\n');
  });

  it('une longue valeur est découpée en morceaux de moins de 255 caractères', () => {
    // L'emoji tombe à la limite d'un morceau : il ne doit pas être coupé en deux.
    const valeur = '- ' + 'é'.repeat(197) + '😀' + 'x'.repeat(400) + '"fin"';
    const formule = C.decouper(C.genererCsv([[valeur]], 'excel').slice(1), 'point-virgule').lignes[0][0];
    assert.ok(formule.split('&').length >= 3);
    assert.equal(afficheParExcel(formule), valeur);
  });

  it('les espaces autour d’une valeur ne changent rien : -12,40 reste un nombre, 06000 garde son zéro', () => {
    assert.equal(C.genererCsv([['-12,40    ', '06000 ', ' 0612345678']], 'excel'),
      '\uFEFF"-12,40    ";"=""06000 """;"="" 0612345678"""\r\n');
  });

  it('dans une colonne de codes (codes postaux…), tous les nombres entiers sont écrits comme du texte', () => {
    assert.equal(C.genererCsv([['CP', 'Âge'], ['75001', '71'], ['06000', '68'], ['2A004', '9']], 'excel'),
      '\uFEFFCP;Âge\r\n"=""75001""";71\r\n"=""06000""";68\r\n2A004;9\r\n');
    assert.equal(C.genererCsv([['Qté'], ['002'], ['010'], ['100']], 'excel'),
      '\uFEFFQté\r\n"=""002"""\r\n"=""010"""\r\n"=""100"""\r\n');
    // Codes entourés d'espaces : même longueur une fois les espaces retirés.
    assert.equal(C.genererCsv([[' 06000'], ['75001']], 'excel'), '\uFEFF"="" 06000"""\r\n"=""75001"""\r\n');
  });

  it('un code isolé (numéro de compte en en-tête) ne change pas une colonne de montants ou de quantités', () => {
    const releve = C.convertir(utf8('Titulaire :  Mme MARTIN\r\nCompte n° :  00012345678\r\n\r\nLibellé        Montant\r\n' +
      'Retraite       1450\r\nLoyer          -620\r\nRemboursement  40\r\n'));
    assert.equal(releve.csv, '\uFEFFTitulaire :;Mme MARTIN\r\nCompte n° :;"=""00012345678"""\r\nLibellé;Montant\r\n' +
      'Retraite;1450\r\nLoyer;-620\r\nRemboursement;40\r\n');
    assert.equal(C.genererCsv([['Magasin :', '0042'], ['Article', 'Quantité'], ['Vis', '100'], ['Clous', '250'], ['Écrous', '50']], 'excel'),
      '\uFEFFMagasin :;"=""0042"""\r\nArticle;Quantité\r\nVis;100\r\nClous;250\r\nÉcrous;50\r\n');
  });

  it('un numéro international précédé d’une espace garde son « + »', () => {
    assert.equal(C.genererCsv([[' +3361234567']], 'excel'), '\uFEFF"="" +3361234567"""\r\n');
  });

  it('accepte aussi des nombres ou des cases vides (null) dans les lignes', () => {
    assert.equal(C.genererCsv([[1, 2]], 'excel'), '\uFEFF1;2\r\n');
  });

  it('une formule de 8 192 caractères exactement reste une formule ; au-delà, une espace devant', () => {
    const premiereCellule = (v) => C.decouper(C.genererCsv([[v]], 'excel').slice(1), 'point-virgule').lignes[0][0];
    // Les guillemets doublés allongent la formule : on cherche une valeur qui donne exactement 8 192 caractères.
    let valeur = null;
    for (let q = 0; q < 400 && !valeur; q++) {
      for (let n = 7000; n < 8100; n++) {
        const essai = '- ' + '"'.repeat(q) + 'a'.repeat(n - 2 - q);
        const f = premiereCellule(essai);
        if (f.startsWith('=') && f.length === 8192) { valeur = essai; break; }
        if (f.length > 8192) break;
      }
    }
    assert.ok(valeur, 'aucune valeur de 8 192 caractères trouvée');
    assert.ok(premiereCellule(valeur).startsWith('="'));
    assert.equal(premiereCellule(valeur + '"'), ' ' + valeur + '"');
  });

  it('valeur sur plusieurs lignes ou trop longue pour une formule : une espace devant suffit', () => {
    assert.equal(C.genererCsv([['- Pain\n- Lait']], 'excel'), '\uFEFF" - Pain\n- Lait"\r\n');
    const longue = '- ' + 'mot '.repeat(2250);
    assert.equal(C.decouper(C.genererCsv([[longue]], 'excel').slice(1), 'point-virgule').lignes[0][0], ' ' + longue);
    // Avec beaucoup de guillemets, la formule dépasserait 8 192 caractères.
    const guillemets = '- ' + 'ab"'.repeat(2332);
    assert.equal(C.decouper(C.genererCsv([[guillemets]], 'excel').slice(1), 'point-virgule').lignes[0][0], ' ' + guillemets);
  });

  it('nombres avec exposant ou sans partie entière : restent des nombres', () => {
    const valeurs = ['-1,5E3', '+1,5E3', '-,5', '-  5', '1,5E3'];
    assert.equal(C.genererCsv([valeurs], 'excel'), '\uFEFF' + valeurs.join(';') + '\r\n');
  });

  it('le format CSV standard ne change aucune valeur', () => {
    assert.equal(C.genererCsv([['- Pain', '06000', '=1+1']], 'standard'), '- Pain,06000,=1+1\r\n');
  });
});

describe('Lignes et fins de ligne', () => {
  it('accepte les fins de ligne Windows (CRLF), Unix (LF) et anciens Mac (CR)', () => {
    assert.deepEqual(lignesDe('a;b\r\nc;d\ne;f\rg;h', 'point-virgule'),
      [['a', 'b'], ['c', 'd'], ['e', 'f'], ['g', 'h']]);
  });

  it('ignore les lignes vides (ou faites d’espaces) et les compte', () => {
    const r = C.decouper('\n\nNom;Âge\n   \n\nZoé;7\n;;\n\n', 'point-virgule');
    assert.deepEqual(r.lignes, [['Nom', 'Âge'], ['Zoé', '7']]);
    assert.equal(r.lignesVides, 6);
  });

  it('enlève les espaces autour des valeurs', () => {
    assert.deepEqual(lignesDe('  Dupont  ;  Jean \n', 'point-virgule'), [['Dupont', 'Jean']]);
  });

  it('garde les valeurs vides au milieu et en fin de ligne', () => {
    assert.deepEqual(lignesDe('a;;c;\n', 'point-virgule'), [['a', '', 'c', '']]);
  });

  it('complète les lignes courtes pour que toutes aient le même nombre de colonnes', () => {
    assert.equal(C.genererCsv([['a', 'b', 'c'], ['d'], ['e', 'f']], 'excel'), '\uFEFFa;b;c\r\nd;;\r\ne;f;\r\n');
  });

  it('évite le piège SYLK d’Excel (fichier qui commence par « ID »)', () => {
    assert.equal(C.genererCsv([['ID', 'Nom'], ['1', 'Zoé']], 'standard'), '"ID",Nom\r\n1,Zoé\r\n');
  });

  it('indique la longueur de la plus longue valeur (Excel coupe une case au-delà de 32 759 caractères)', () => {
    assert.equal(C.LIMITES_EXCEL.caracteres, 32759);
    assert.equal(C.decouper('a;' + 'b'.repeat(40000) + '\nc;d\n', 'point-virgule').longueurMax, 40000);
  });

  it('retire la marque de fin de fichier des anciens programmes DOS (Ctrl+Z)', () => {
    const sansRetour = C.analyser(concat(utf8('Libellé;Montant\r\nPension;1250,00'), [0x1A]));
    assert.deepEqual(sansRetour.lignes[1], ['Pension', '1250,00']);
    const apresRetour = C.analyser(concat(utf8('a;b\r\nc;d\r\n'), [0x1A]));
    assert.equal(apresRetour.lignes.length, 2);
    assert.deepEqual(C.analyser(octets(0x1A)), { ok: false, erreur: 'vide' });
  });
});

describe('Rapidité : pas de blocage sur les longues suites d’espaces', () => {
  const chrono = (fonction) => { const debut = process.hrtime.bigint(); fonction(); return Number(process.hrtime.bigint() - debut) / 1e6; };

  it('200 000 espaces en fin de ligne (mode « plusieurs espaces »)', () => {
    assert.ok(chrono(() => C.decouper('x' + ' '.repeat(200000) + '\n', 'espaces')) < 1000);
    assert.ok(chrono(() => C.decouper('"x"' + ' '.repeat(200000) + '\n', 'espaces')) < 1000);
  });

  it('détection sur un export à largeur fixe (lignes complétées par des espaces)', () => {
    const texte = ('a;b;i' + ' '.repeat(5000) + '\n').repeat(300);
    assert.ok(chrono(() => C.detecterSeparateur(texte)) < 1000);
  });
});

describe('Fichiers vides ou qui ne sont pas du texte', () => {
  it('fichier de 0 octet', () => {
    assert.deepEqual(C.analyser(new Uint8Array(0)), { ok: false, erreur: 'vide' });
  });

  it('fichier qui ne contient qu’un BOM', () => {
    assert.deepEqual(C.analyser(octets(0xEF, 0xBB, 0xBF)), { ok: false, erreur: 'vide' });
  });

  it('fichier fait uniquement de lignes vides ou d’espaces', () => {
    assert.deepEqual(C.analyser(utf8('\r\n\r\n   \n\t\n')), { ok: false, erreur: 'lignes-vides' });
  });

  it('fichier fait uniquement de séparateurs', () => {
    assert.deepEqual(C.analyser(utf8(';;;\n;;\n')), { ok: false, erreur: 'lignes-vides' });
  });

  it('refuse un document Word, Excel, PDF, une image, un son, une vidéo ou une archive', () => {
    const plein = (taille, valeur) => new Uint8Array(taille).fill(valeur);
    const avecDebut = (debut, taille, valeur) => { const b = plein(taille, valeur); b.set(debut); return b; };
    const exemples = {
      'docx/xlsx': octets(0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00),
      'pdf': utf8('%PDF-1.7\n%âãÏÓ\n'),
      'png': octets(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A),
      'contrôles': octets(0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x10, 0x11, 0x12),
      'bmp blanc': avecDebut([0x42, 0x4D, 0x36, 0x53, 0x07, 0x00, 0, 0, 0, 0, 0x36, 0, 0, 0], 50000, 0xFF),
      'tiff blanc': avecDebut([0x49, 0x49, 0x2A, 0x00, 0x08, 0, 0, 0], 50000, 0xFF),
      'wav': avecDebut([...utf8('RIFF'), 0x24, 0x4E, 0, 0, ...utf8('WAVEfmt ')], 20000, 0x80),
      'mp4/heic': avecDebut([0x00, 0x00, 0x00, 0x18, ...utf8('ftypheic')], 20000, 0x41),
      '7z': avecDebut([0x37, 0x7A, 0xBC, 0xAF, 0x27, 0x1C], 1000, 0x41),
      'icône': avecDebut([0, 0, 1, 0, 1, 0, 32, 32], 5000, 0xFF),
      'curseur': avecDebut([0, 0, 2, 0, 1, 0, 32, 32], 5000, 0xFF),
      'police': avecDebut([0x00, 0x01, 0x00, 0x00, 0x00, 0x0F], 1000, 0x41)
    };
    for (const [nom, contenu] of Object.entries(exemples)) {
      assert.deepEqual(C.analyser(contenu), { ok: false, erreur: 'binaire' }, nom);
    }
  });

  it('refuse un binaire qui ressemble à de l’UTF-16 (un octet sur deux nul)', () => {
    const b = new Uint8Array(8000);
    for (let i = 0; i < b.length; i += 2) b[i] = 1 + (i % 7);
    assert.deepEqual(C.analyser(b), { ok: false, erreur: 'binaire' });
  });

  it('accepte un texte qui commence comme une signature de fichier', () => {
    for (const texte of ['Format: WAVE;Durée\r\nchanson.wav;3:12\r\n', 'RIFF;WAVE\nx;y\n', 'BMW;Série 3\r\nAudi;A4\r\n',
      'Notaftyp;b\r\nc;d\r\n']) {
      assert.equal(C.analyser(utf8(texte)).ok, true, texte);
    }
  });

  it('accepte un texte dont la fin est remplie d’octets nuls (fichier abîmé) et le signale', () => {
    const r = C.analyser(concat(utf8('Nom;Ville\r\nZoé;Nîmes\r\nLéa;Albi'), [0, 0, 0]));
    assert.equal(r.ok, true);
    assert.equal(r.finAbimee, true);
    assert.deepEqual(r.lignes[2], ['Léa', 'Albi']);
    const long = new Uint8Array(4096);
    long.set(utf8('Nom;Ville\r\n'.repeat(150)));
    assert.equal(C.analyser(long).finAbimee, true);
    assert.equal(C.analyser(concat(utf8('a;b\n'), [0])).finAbimee, false);
    assert.deepEqual(C.analyser(new Uint8Array(100)), { ok: false, erreur: 'vide' });
  });

  it('accepte un petit fichier terminé par plusieurs Ctrl+Z ou complété par des Ctrl+Z (anciens programmes)', () => {
    assert.equal(C.analyser(concat(utf8('Nom;Ville\r\nZoé;Nîmes\r\n'), [0x1A, 0x1A])).lignes.length, 2);
    const remplissage = new Uint8Array(128).fill(0x1A);
    remplissage.set(utf8('a;b\r\nc;d\r\ne;f\r\n'));
    const r = C.analyser(remplissage);
    assert.equal(r.lignes.length, 3);
    assert.equal(r.finAbimee, false);
    // Des octets nuls suivis d'un retour à la ligne : toujours signalés.
    assert.equal(C.analyser(concat(utf8('Libellé;Montant\r\nLoyer;-65'), new Uint8Array(512), utf8('\r\n'))).finAbimee, true);
  });

  it('refuse un document mis en forme (RTF)', () => {
    assert.deepEqual(C.analyser(utf8('{\\rtf1\\ansi Bonjour}')), { ok: false, erreur: 'rtf' });
  });
});

describe('Nom du fichier téléchargé', () => {
  const cas = {
    'relevé.txt': 'relevé.csv',
    'Notes.TXT': 'Notes.csv',
    'liste.clients.2026.txt': 'liste.clients.2026.csv',
    'export.tsv': 'export.csv',
    'notes.txt.txt': 'notes.csv',
    'export.csv.txt': 'export.csv',
    'sans-extension': 'sans-extension.csv',
    'archive.2026': 'archive.2026.csv',
    '.txt': 'fichier.csv',
    '': 'fichier.csv',
    'C:\\Documents\\courses.txt': 'courses.csv'
  };
  for (const [original, attendu] of Object.entries(cas)) {
    it(JSON.stringify(original) + ' → ' + attendu, () => {
      assert.equal(C.nomFichierCsv(original), attendu);
    });
  }
});

describe('Fichiers d’exemple fournis', () => {
  const lire = (nom) => new Uint8Array(fs.readFileSync(path.join(__dirname, '..', 'exemples', nom)));

  it('contacts-utf8.txt : UTF-8, points-virgules, guillemets et ligne vide', () => {
    const r = C.convertir(lire('contacts-utf8.txt'));
    assert.equal(r.encodage, 'utf-8');
    assert.equal(r.separateur, 'point-virgule');
    assert.equal(r.lignes.length, 7);
    assert.equal(r.colonnes, 5);
    assert.equal(r.lignesVides, 1);
    assert.equal(r.lignes[1][4], 'Préfère être appelée le matin; jamais le dimanche');
    assert.equal(r.lignes[3][4], 'A écrit : "Merci pour tout !"');
    assert.equal(r.lignes[4][4], 'Nouvelle adresse :\n12 rue des Lilas');
    assert.equal(r.lignes[5][4], '');
  });

  it('inventaire-windows1252.txt : Windows-1252, tabulations, fins de ligne Windows', () => {
    const r = C.convertir(lire('inventaire-windows1252.txt'));
    assert.equal(r.encodage, 'windows-1252');
    assert.equal(r.encodageDouteux, false);
    assert.equal(r.separateur, 'tabulation');
    assert.equal(r.lignes.length, 7);
    assert.equal(r.lignesVides, 1);
    assert.deepEqual(r.lignes[3], ['Bœuf haché', '1', '6,90 €', 'Boucherie']);
    assert.deepEqual(r.lignes[6], ['Café moulu « pur arabica »', '2', '3,99 €', 'Épicerie']);
    assert.ok(r.csv.startsWith('\uFEFFArticle;Quantité;Prix unitaire;Rayon\r\n'));
  });

  it('budget-espaces-alignes.txt : UTF-8 avec BOM, colonnes alignées par des espaces', () => {
    const r = C.convertir(lire('budget-espaces-alignes.txt'));
    assert.equal(r.encodage, 'utf-8');
    assert.equal(r.bom, true);
    assert.equal(r.separateur, 'espaces');
    assert.deepEqual(r.lignes[0], ['Relevé de compte — janvier 2026']);
    assert.deepEqual(r.lignes[1], ['Date', 'Libellé', 'Montant']);
    assert.deepEqual(r.lignes[3], ['07/01/2026', 'Pension de retraite', '1 250,00']);
    assert.deepEqual(r.lignes[4], ['12/01/2026', 'Pharmacie Centrale', '-23,15']);
    // Les montants négatifs restent des nombres dans Excel.
    assert.ok(r.csv.includes('\r\n12/01/2026;Pharmacie Centrale;-23,15\r\n'));
  });
});
