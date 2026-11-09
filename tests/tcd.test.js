'use strict';

// Tests du tableau croisé dynamique (js/tcd.js). Le fichier .xlsx produit a
// aussi été vérifié dans Excel lui-même : ouverture sans réparation, et
// résultats identiques à ceux calculés ici.
const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const zlib = require('node:zlib');
const T = require('../js/tcd.js');

const releve = [
  ['Relevé de compte — premier trimestre 2026'],
  ['Date', 'Catégorie', 'Libellé', 'Montant', 'Mode'],
  ['05/01/2026', 'Alimentation', 'Boulangerie', '-12,40', 'Carte'],
  ['07/01/2026', 'Revenus', 'Pension', '1 250,00', 'Virement'],
  ['12/01/2026', 'Santé', 'Pharmacie', '-23,15', 'Carte'],
  ['03/02/2026', 'Alimentation', 'Marché', '-45,60', 'Carte'],
  ['07/02/2026', 'Revenus', 'Pension', '1 250,00', 'Virement'],
  ['15/02/2026', 'alimentation', 'Épicerie', '-8,90', 'Carte'],
  ['02/03/2026', 'Santé', 'Médecin', '-30,00', 'Virement'],
  ['07/03/2026', 'Revenus', 'Pension', '1 250,00', 'Virement'],
  ['20/03/2026', 'Alimentation', 'Supermarché', '', 'Carte'],
  ['Total', '', '', '3 629,95', '']
];

// Lecture d'un fichier ZIP (sans compression) : { nom: Buffer }.
function lireZip(morceaux) {
  const zip = Buffer.concat(morceaux.map((m) => Buffer.from(m)));
  const fin = zip.lastIndexOf(Buffer.from([0x50, 0x4B, 0x05, 0x06]));
  assert.ok(fin >= 0, 'fin de ZIP introuvable');
  const nombre = zip.readUInt16LE(fin + 10);
  let p = zip.readUInt32LE(fin + 16);
  const fichiers = {};
  for (let i = 0; i < nombre; i++) {
    assert.equal(zip.readUInt32LE(p), 0x02014B50);
    const crc = zip.readUInt32LE(p + 16);
    const taille = zip.readUInt32LE(p + 24);
    const longueurNom = zip.readUInt16LE(p + 28);
    const position = zip.readUInt32LE(p + 42);
    const nom = zip.toString('utf8', p + 46, p + 46 + longueurNom);
    assert.equal(zip.readUInt32LE(position), 0x04034B50, 'en-tête local de ' + nom);
    const debut = position + 30 + zip.readUInt16LE(position + 26) + zip.readUInt16LE(position + 28);
    const contenu = zip.subarray(debut, debut + taille);
    assert.equal(zlib.crc32(contenu), crc, 'somme de contrôle de ' + nom);
    fichiers[nom] = contenu;
    p += 46 + longueurNom + zip.readUInt16LE(p + 30) + zip.readUInt16LE(p + 32);
  }
  return fichiers;
}

const texteDe = (fichiers, nom) => fichiers[nom].toString('utf8');

describe('Lecture des nombres et des dates', () => {
  it('nombres écrits à la française ou à l’anglaise', () => {
    const cas = {
      '-12,40': -12.4, '1 250,00': 1250, '1 250,50 €': 1250.5, '1.250,50': 1250.5, '1,250.50': 1250.5,
      '12.5': 12.5, '(12,50)': -12.5, '12,50-': -12.5, '+5': 5, '5 %': 0.05, '€ 3,99': 3.99, ',5': 0.5, '1 000 000': 1000000
    };
    for (const [texte, attendu] of Object.entries(cas)) {
      const n = T.lireNombre(texte);
      assert.ok(n, texte);
      assert.ok(Math.abs(n.valeur - attendu) < 1e-9, texte + ' → ' + n.valeur);
    }
    for (const texte of ['', 'abc', '01 99 00 12 34', '12 3', '1,2,3', '05/01/2026', '--5']) {
      assert.equal(T.lireNombre(texte), null, texte);
    }
    assert.equal(T.lireNombre('06000').zeroInitial, true);
    assert.equal(T.lireNombre('0,45').zeroInitial, false);
  });

  it('dates, avec numéro de série d’Excel', () => {
    assert.equal(T.lireDate('05/01/2026'), 46027);
    assert.equal(T.lireDate('5-1-26'), 46027);
    assert.equal(T.lireDate('2026-01-05'), 46027);
    assert.equal(T.lireDate('05.01.2026 12:00'), 46027.5);
    for (const texte of ['31/02/2026', '13/13/2026', '1.250', '05/01', 'hier']) assert.equal(T.lireDate(texte), null, texte);
  });
});

describe('Préparation des données', () => {
  const d = T.preparer(releve);

  it('saute le titre au-dessus du tableau, repère les noms de colonnes et ignore la ligne de total', () => {
    assert.equal(d.lignesTitre, 1);
    assert.equal(d.entete, true);
    assert.equal(d.lignesTotal, 1);
    assert.equal(d.nbLignes, 9);
    assert.deepEqual(d.noms.slice(0, 5), ['Date', 'Catégorie', 'Libellé', 'Montant', 'Mode']);
  });

  it('reconnaît les types et convertit les valeurs', () => {
    assert.deepEqual(d.types.slice(0, 5), ['date', 'texte', 'texte', 'nombre', 'texte']);
    assert.equal(d.colonnes[3][1], 1250);
    assert.equal(d.colonnes[3][8], null);
    assert.equal(d.colonnes[0][0], 46027);
  });

  it('ajoute année, trimestre et mois pour regrouper les dates', () => {
    assert.deepEqual(d.noms.slice(5), ['Date (année)', 'Date (trimestre)', 'Date (mois)']);
    assert.equal(d.colonnes[5][0], '2026');
    assert.equal(d.colonnes[6][4], '2026-T1');
    assert.equal(d.colonnes[7][3], '2026-02 (février)');
  });

  it('codes à zéro initial (codes postaux) : texte, pas nombre', () => {
    const p = T.preparer([['Ville', 'CP'], ['Nice', '06000'], ['Paris', '75001']]);
    assert.equal(p.types[1], 'texte');
    assert.equal(p.colonnes[1][0], '06000');
  });

  it('sans ligne de titres : noms « Colonne A »… ; la case peut l’imposer', () => {
    const lignes = [['Nice', '12'], ['Paris', '15'], ['Nice', '3']];
    const p = T.preparer(lignes);
    assert.equal(p.entete, false);
    assert.deepEqual(p.noms, ['Colonne A', 'Colonne B']);
    assert.equal(p.nbLignes, 3);
    const q = T.preparer([['Ville', 'Ville'], ['Nice', 'Nice']], { entete: true });
    assert.deepEqual(q.noms, ['Ville', 'Ville (2)']);
  });
});

describe('Proposition automatique', () => {
  it('relevé : somme des montants par catégorie, mois en colonnes, du plus grand au plus petit', () => {
    const d = T.preparer(releve);
    const r = T.proposer(d);
    assert.deepEqual(r.lignes.map((j) => d.noms[j]), ['Catégorie']);
    assert.deepEqual(r.colonnes.map((j) => d.noms[j]), ['Date (mois)']);
    assert.deepEqual(r.valeurs, [{ champ: 3, calcul: 'somme' }]);
    assert.equal(r.tri, 'decroissant');
  });

  it('inventaire : somme des quantités et moyenne des prix unitaires par rayon', () => {
    const d = T.preparer([['Article', 'Quantité', 'Prix unitaire', 'Rayon'],
      ['Crème', '2', '1,85 €', 'Crèmerie'], ['Pâtes', '3', '1,20 €', 'Épicerie'], ['Riz', '1', '2,10 €', 'Épicerie'],
      ['Lait', '6', '0,95 €', 'Crèmerie'], ['Bœuf', '1', '6,90 €', 'Boucherie'], ['Café', '2', '3,99 €', 'Épicerie']]);
    const r = T.proposer(d);
    assert.deepEqual(r.lignes.map((j) => d.noms[j]), ['Rayon']);
    assert.deepEqual(r.valeurs.map((v) => v.calcul + ' ' + d.noms[v.champ]), ['somme Quantité', 'moyenne Prix unitaire']);
  });

  it('sans colonne de nombres : nombre de lignes', () => {
    const d = T.preparer([['Nom', 'Ville'], ['A', 'Nice'], ['B', 'Paris'], ['C', 'Nice'], ['D', 'Nice'], ['E', 'Paris']]);
    const r = T.proposer(d);
    assert.equal(d.noms[r.lignes[0]], 'Ville');
    assert.equal(r.valeurs[0].calcul, 'nombre');
  });
});

describe('Calcul du tableau croisé', () => {
  const d = T.preparer(releve);
  const champ = (nom) => d.noms.indexOf(nom);

  it('sommes par élément et total général ; « Alimentation » et « alimentation » sont le même élément', () => {
    const res = T.calculer(d, { lignes: [champ('Catégorie')], valeurs: [{ champ: champ('Montant'), calcul: 'somme' }], tri: 'libelle' });
    assert.deepEqual(res.lignes.map((l) => l.libelle), ['Alimentation', 'Revenus', 'Santé']);
    assert.ok(Math.abs(res.lignes[0].totaux[0] - -66.9) < 1e-9);
    assert.ok(Math.abs(res.total.totaux[0] - 3629.95) < 1e-9);
    assert.equal(res.enregistrements, 9);
  });

  it('nombre (cases remplies), moyenne, minimum, maximum', () => {
    const res = T.calculer(d, { lignes: [champ('Catégorie')], valeurs: [
      { champ: champ('Montant'), calcul: 'nombre' }, { champ: champ('Montant'), calcul: 'moyenne' }, { champ: champ('Montant'), calcul: 'max' }], tri: 'libelle' });
    const alim = res.lignes[0].totaux;
    assert.equal(alim[0], 3); // la ligne au montant vide ne compte pas
    assert.ok(Math.abs(alim[1] - -22.3) < 1e-9);
    assert.ok(Math.abs(alim[2] - -8.9) < 1e-9);
    assert.deepEqual(res.valeurs.map((v) => v.nom), ['Nombre de Montant', 'Moyenne de Montant', 'Max de Montant']);
  });

  it('colonnes (mois regroupés) et pourcentages du total, de la ligne, de la colonne', () => {
    const base = { lignes: [champ('Mode')], colonnes: [champ('Date (mois)')], valeurs: [{ champ: champ('Montant'), calcul: 'somme' }], tri: 'libelle' };
    const res = T.calculer(d, base);
    assert.deepEqual(res.colonnes.map((c) => c.libelle), ['2026-01 (janvier)', '2026-02 (février)', '2026-03 (mars)']);
    const virement = res.lignes[1];
    assert.equal(virement.libelle, 'Virement');
    assert.equal(virement.cellules[2][0], 1220);
    const total = T.calculer(d, Object.assign({}, base, { affichage: 'total' }));
    assert.equal(total.total.totaux[0], 1);
    const ligne = T.calculer(d, Object.assign({}, base, { affichage: 'ligne' }));
    assert.equal(ligne.lignes[0].totaux[0], 1);
    const colonne = T.calculer(d, Object.assign({}, base, { affichage: 'colonne' }));
    assert.equal(colonne.total.cellules[0][0], 1);
  });

  it('filtre : les éléments exclus ne comptent pas', () => {
    const res = T.calculer(d, { lignes: [champ('Catégorie')], valeurs: [{ champ: champ('Montant'), calcul: 'somme' }],
      filtre: { champ: champ('Mode'), exclus: { [T.cle('Carte')]: true } }, tri: 'libelle' });
    assert.equal(res.enregistrements, 4);
    assert.deepEqual(res.lignes.map((l) => l.libelle), ['Revenus', 'Santé']);
  });

  it('tri par valeur ; à valeur égale, ordre des libellés dans le sens du tri (comme Excel)', () => {
    const p = T.preparer([['Rayon', 'Prix'], ['B', '2'], ['A', '2'], ['C', '5'], ['D', '1']]);
    const dec = T.calculer(p, { lignes: [0], valeurs: [{ champ: 1, calcul: 'somme' }], tri: 'decroissant' });
    assert.deepEqual(dec.lignes.map((l) => l.libelle), ['C', 'B', 'A', 'D']);
    const cro = T.calculer(p, { lignes: [0], valeurs: [{ champ: 1, calcul: 'somme' }], tri: 'croissant' });
    assert.deepEqual(cro.lignes.map((l) => l.libelle), ['D', 'A', 'B', 'C']);
  });

  it('deux niveaux de lignes, avec sous-totaux', () => {
    const res = T.calculer(d, { lignes: [champ('Mode'), champ('Catégorie')], valeurs: [{ champ: champ('Montant'), calcul: 'somme' }], tri: 'libelle' });
    assert.deepEqual(res.lignes.map((l) => [l.niveau, l.libelle]),
      [[0, 'Carte'], [1, 'Alimentation'], [1, 'Santé'], [0, 'Virement'], [1, 'Revenus'], [1, 'Santé']]);
    assert.ok(Math.abs(res.lignes[0].totaux[0] - -90.05) < 1e-9);
  });

  it('les réglages incohérents sont corrigés (même colonne à deux places, valeurs manquantes)', () => {
    const r = T.normaliserReglages(d, { lignes: [1, 1, 99], colonnes: [1], valeurs: [], filtre: { champ: 1 } });
    assert.deepEqual(r.lignes, [1]);
    assert.deepEqual(r.colonnes, []);
    assert.equal(r.filtre.champ, -1);
    assert.deepEqual(r.valeurs, [{ champ: 1, calcul: 'nombre' }]);
  });
});

describe('Fichier Excel (.xlsx)', () => {
  const d = T.preparer(releve);
  const champ = (nom) => d.noms.indexOf(nom);
  const reglages = {
    lignes: [champ('Catégorie')], colonnes: [champ('Date (mois)')],
    valeurs: [{ champ: champ('Montant'), calcul: 'somme' }, { champ: champ('Montant'), calcul: 'nombre' }],
    affichage: 'total', tri: 'decroissant', filtre: { champ: champ('Mode'), exclus: { [T.cle('Carte')]: true } }
  };
  const fichiers = lireZip(T.genererXlsx(d, reglages));

  it('ZIP valide, avec toutes les parties d’un classeur à tableau croisé', () => {
    for (const nom of ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml',
      'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/_rels/sheet1.xml.rels',
      'xl/pivotTables/pivotTable1.xml', 'xl/pivotTables/_rels/pivotTable1.xml.rels',
      'xl/pivotCache/pivotCacheDefinition1.xml', 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
      'xl/pivotCache/pivotCacheRecords1.xml']) {
      assert.ok(fichiers[nom], nom + ' manquant');
      assert.ok(texteDe(fichiers, nom).startsWith('<?xml'), nom);
    }
    assert.match(texteDe(fichiers, 'xl/workbook.xml'), /<sheet name="Tableau croisé" sheetId="1" r:id="rId1"\/><sheet name="Données"/);
  });

  it('feuille « Données » : titres, nombres, dates et textes', () => {
    const feuille = texteDe(fichiers, 'xl/worksheets/sheet2.xml');
    assert.match(feuille, /<dimension ref="A1:H10"\/>/);
    assert.match(feuille, /<c r="A2" s="1"><v>46027<\/v><\/c>/);
    assert.match(feuille, /<c r="D3" s="3"><v>1250<\/v><\/c>/);
    assert.match(feuille, /<c r="B2" t="inlineStr"><is><t xml:space="preserve">Alimentation<\/t><\/is><\/c>/);
    assert.match(feuille, /<c r="H4" t="inlineStr"><is><t xml:space="preserve">2026-01 \(janvier\)<\/t>/);
    assert.ok(!/<c r="D10"/.test(feuille), 'case vide non écrite');
  });

  it('cache : source, champs, éléments (sans distinction de majuscules), actualisation à l’ouverture', () => {
    const cache = texteDe(fichiers, 'xl/pivotCache/pivotCacheDefinition1.xml');
    assert.match(cache, /refreshOnLoad="1"/);
    assert.match(cache, /<worksheetSource ref="A1:H10" sheet="Données"\/>/);
    assert.match(cache, /<cacheField name="Catégorie" numFmtId="0"><sharedItems count="3"><s v="Alimentation"\/><s v="Revenus"\/><s v="Santé"\/><\/sharedItems>/);
    assert.match(cache, /<cacheField name="Date" numFmtId="14"><sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0" minDate="2026-01-05T00:00:00" maxDate="2026-03-20T00:00:00"\/>/);
    assert.match(cache, /<cacheField name="Montant" numFmtId="0"><sharedItems containsString="0" containsBlank="1" containsNumber="1" minValue="-45.6" maxValue="1250"\/>/);
  });

  it('tableau : lignes, colonnes, valeurs, filtre, pourcentage et tri par valeur', () => {
    const t = texteDe(fichiers, 'xl/pivotTables/pivotTable1.xml');
    assert.match(t, /<rowFields count="1"><field x="1"\/><\/rowFields>/);
    assert.match(t, /<colFields count="2"><field x="7"\/><field x="-2"\/><\/colFields>/);
    assert.match(t, /<pageFields count="1"><pageField fld="4" hier="-1"\/><\/pageFields>/);
    assert.match(t, /<dataField name="Somme de Montant" fld="3" showDataAs="percentOfTotal" baseField="0" baseItem="0" numFmtId="10"\/>/);
    assert.match(t, /<dataField name="Nombre de Montant" fld="3" subtotal="count" showDataAs="percentOfTotal"/);
    assert.match(t, /<pivotField axis="axisRow" showAll="0" sortType="descending"><items count="4">.*<\/items><autoSortScope>/);
    // Filtre : « Carte » masqué, les autres gardés.
    assert.match(t, /<pivotField axis="axisPage" multipleItemSelectionAllowed="1" showAll="0"><items count="3"><item x="0" h="1"\/><item x="1"\/><item t="default"\/><\/items><\/pivotField>/);
    assert.match(t, /rowPageCount="1" colPageCount="1"/);
  });

  it('caractères spéciaux : échappés, et caractères interdits en XML retirés', () => {
    const p = T.preparer([['Nom <&>', 'Quantité'], ['A "B" & <C>\u0001', '1'], ['D\uD800', '2']]);
    const f = lireZip(T.genererXlsx(p, T.proposer(p)));
    const feuille = texteDe(f, 'xl/worksheets/sheet2.xml');
    assert.ok(feuille.includes('Nom &lt;&amp;&gt;'));
    assert.ok(feuille.includes('A "B" &amp; &lt;C&gt;</t>'));
    assert.ok(!/[\u0001\uD800]/.test(feuille));
    assert.ok(texteDe(f, 'xl/pivotCache/pivotCacheDefinition1.xml').includes('name="Nom &lt;&amp;&gt;"'));
  });

  it('aucune date ni information personnelle dans le ZIP (date fixe : 1er janvier 1980)', () => {
    const zip = Buffer.concat(T.genererXlsx(d, reglages).map((m) => Buffer.from(m)));
    assert.equal(zip.readUInt16LE(10), 0);
    assert.equal(zip.readUInt16LE(12), 0x21);
  });

  it('refuse un tableau sans aucune ligne', () => {
    const vide = T.preparer([['Titre A', 'Titre B'], ['Total', '12']], { entete: true });
    assert.equal(vide.nbLignes, 0);
    assert.throws(() => T.genererXlsx(vide, {}));
  });
});
