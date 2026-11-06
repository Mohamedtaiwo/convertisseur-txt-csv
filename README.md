# Convertisseur TXT vers CSV

Un petit outil web, gratuit et simple, pour transformer un fichier texte (`.txt`) en fichier `.csv` qui s'ouvre directement dans Excel.

**Utiliser l'outil : [convertisseur-txt-csv.vercel.app](https://convertisseur-txt-csv.vercel.app)**

> **Votre fichier reste sur votre ordinateur. Rien n'est envoyé ni conservé.**

## À quoi sert cet outil ?

Beaucoup de logiciels et de sites (banque, mutuelle, carnet d'adresses, inventaire…) permettent d'exporter des données dans un fichier texte (`.txt`). Ouvert tel quel dans Excel, ce fichier s'affiche souvent mal : tout est entassé dans une seule colonne, ou les accents sont remplacés par des signes bizarres (`Ã©` au lieu de `é`).

Cet outil le transforme en fichier `.csv` qu'Excel ouvre proprement : une colonne par information, et des accents corrects.

## Comment l'utiliser

1. Ouvrez [la page de l'outil](https://convertisseur-txt-csv.vercel.app) dans votre navigateur (Edge, Chrome, Firefox ou Safari, sur ordinateur, tablette ou téléphone).
2. Cliquez sur **Choisir mon fichier .txt** et choisissez votre fichier. Vous pouvez aussi le faire glisser sur la page.
3. **Vérifiez le tableau** : les 10 premières lignes s'affichent, rangées en colonnes A, B, C… comme dans Excel.
4. Cliquez sur **Télécharger le fichier CSV**. Le fichier garde le nom de l'original : `releve.txt` devient `releve.csv`.
5. Pour l'ouvrir dans Excel, cliquez dessus dans la liste des téléchargements du navigateur, ou double-cliquez dessus dans le dossier *Téléchargements*. Sur téléphone ou tablette, il se trouve dans l'application *Fichiers*.

Si vous transformez plusieurs fois le même fichier, le navigateur numérote les nouveaux : `releve (1).csv`, `releve (2).csv`… Prenez le plus récent.

Pour convertir un autre fichier, cliquez sur **Recommencer**.

### Si le tableau est mal découpé

L'outil devine tout seul comment les colonnes sont séparées dans votre fichier. S'il se trompe (tout est dans une seule colonne, ou des informations sont coupées au mauvais endroit), choisissez le bon séparateur dans la liste sous le tableau : tabulations, points-virgules, virgules, barres verticales, plusieurs espaces, ou rien (une seule colonne). Le tableau se met à jour aussitôt.

### Quel format choisir ?

| Choix | Séparateur | Encodage | Pour qui ? |
|---|---|---|---|
| **Excel (recommandé)** | point-virgule `;` | UTF-8 avec BOM | Excel en français : un double-clic suffit ; les accents s'affichent bien et les codes (codes postaux, téléphones) gardent leurs zéros. |
| **CSV standard** | virgule `,` | UTF-8 sans BOM | Les autres logiciels (format CSV standard, RFC 4180). Ne convient pas à Excel en français. |

## Confidentialité

L'outil est conçu pour que vos données ne quittent pas votre appareil.

- **Tout se passe dans votre navigateur.** Aucun serveur ne reçoit ni ne traite vos fichiers : le site n'est fait que de fichiers fixes (HTML, CSS et JavaScript). Votre fichier est lu, converti et téléchargé sur place.
- **Rien n'est envoyé.** Le code de la page ne contient aucune fonction d'envoi, et des tests automatiques le vérifient. En plus, une règle de sécurité (*Content-Security-Policy*) interdit à la page de charger quoi que ce soit depuis un autre site ou de lui envoyer des données (requêtes, formulaires, images, polices, scripts).
- **Rien n'est conservé.** Ni cookies, ni stockage dans le navigateur (localStorage, sessionStorage, IndexedDB, cache). Après le téléchargement, les données sont effacées de la page, et l'adresse temporaire du fichier généré est libérée peu après (au plus 40 secondes, le temps que le navigateur l'enregistre ; aussitôt si vous cliquez sur **Recommencer** ou quittez la page). Le bouton **Recommencer** remet tout à zéro, et quitter la page efface tout. Seule trace, comme pour tout téléchargement : le navigateur garde le nom du fichier CSV dans sa liste des téléchargements.
- **Aucun mouchard.** Pas de statistiques de visite, pas de publicité, pas de polices ni de scripts venant d'autres sites.
- **Vous pouvez le vérifier.** Une fois la page ouverte, coupez Internet : l'outil fonctionne toujours. Le code est court et public.

À savoir :

- La page elle-même est hébergée par Vercel. Comme tout hébergeur, Vercel peut enregistrer la visite de la page (adresse IP, date et heure). Le contenu de vos fichiers, lui, ne lui est jamais transmis.
- Les extensions installées dans le navigateur et ses services en ligne (comme la traduction automatique des pages) échappent aux règles de la page. Le tableau d'aperçu est marqué pour ne pas être traduit.

## Ce que l'outil sait faire

- **Encodage** détecté automatiquement : UTF-8 (avec ou sans BOM), sinon Windows-1252 (l'ancien encodage de Windows), pour que les accents restent corrects. Les fichiers UTF-16 (« Texte Unicode » d'Excel) sont aussi reconnus, ainsi que la plupart des fichiers qui mélangent UTF-8 et Windows-1252 (deux fichiers mis bout à bout, par exemple), avec un avertissement. Un fichier enregistré par un ancien logiciel (DOS, ancien Mac), dont les accents ne peuvent pas être reconnus, est en général signalé.
- **Séparateurs** : tabulation, point-virgule, virgule, barre verticale `|`, plusieurs espaces (colonnes alignées, y compris avec des cases vides, ou alignées avec la touche Tab dans le Bloc-notes). Détection automatique, modifiable à la main. Une virgule décimale (`12,40`) n'est pas prise pour un séparateur.
- **Guillemets**, selon les règles CSV standard (RFC 4180) : une valeur entre guillemets peut contenir des séparateurs, des retours à la ligne et des guillemets doublés (`""`). Un guillemet mal placé est gardé tel quel, sans rien perdre.
- **Lignes vides** : ignorées ; leur nombre est indiqué dans « Détails techniques ».
- **Fichier vide, dossier, document Word ou PDF, image, son ou vidéo** : refusés, avec un message simple. Les marques de fin des anciens fichiers DOS (Ctrl+Z) sont retirées ; un fichier qui se termine par des octets nuls (fin abîmée) est accepté, avec un avertissement. Un fichier coupé au milieu de son dernier caractère est en général signalé, mais pas toujours (par exemple s'il ne contient aucun autre accent).
- **Taille** : jusqu'à 50 Mo (un gros fichier peut demander quelques secondes).

## Bon à savoir avec Excel

En ouvrant un fichier CSV, Excel interprète lui-même certaines valeurs. Il supprime les zéros du début (`06000` devient `6000`), arrondit les longs numéros, et prend pour une formule un texte qui commence par `=`, `+`, `-` ou `@` (une liste `- Pain`, `- Lait` afficherait `#NOM?`).

Pour l'éviter, le format **Excel** écrit ces valeurs-là sous la forme `="06000"`, qu'Excel affiche exactement comme dans le fichier d'origine :

- codes postaux et numéros qui commencent par 0 (`06000`, `0612345678`) ;
- numéros de 11 chiffres ou plus (SIRET, références) et numéros internationaux (`+33612345678`) ;
- textes qui commencent par `=`, `+`, `-` ou `@`.

Dans une colonne de codes de même longueur (codes postaux, téléphones…), tous les nombres entiers sont écrits de la même façon, pour que la colonne reste homogène (tri, alignement) : `75001` comme `06000`. Ces valeurs sont du texte : Excel ne les compte pas dans les sommes. Un code isolé, comme un numéro de compte en haut du relevé, ne change pas une colonne de montants ou de quantités : seul ce code est protégé. Un texte sur plusieurs lignes, ou très long, qui commence par `=`, `+`, `-` ou `@` est simplement précédé d'une espace.

Les nombres ordinaires (`-12,40`, `1 250,00 €`) et les dates restent des nombres et des dates, pour que les calculs fonctionnent ; Excel les affiche à sa façon (`12,40` devient `12,4`). Il transforme encore en date quelques valeurs courtes comme `3-4` ou `1/2`, et en nombre des références comme `2E45`.

La forme `="06000"` est prévue pour ouvrir le fichier dans Excel par un double-clic (ou par **Fichier › Ouvrir**). L'import d'Excel **Données › À partir d'un fichier texte/CSV**, LibreOffice ou certains aperçus l'affichent telle quelle. Pour ces usages, ou pour garder telles quelles des valeurs comme `3-4`, téléchargez le fichier au format **CSV standard**, qui écrit toutes les valeurs telles quelles, puis, dans Excel, utilisez **Données › À partir d'un fichier texte/CSV** et choisissez le type **Texte** pour la colonne concernée.

Si vous modifiez le tableau dans Excel, enregistrez-le avec **Fichier › Enregistrer sous**, au format **Classeur Excel (.xlsx)** : enregistré en CSV par Excel, il perdrait ses zéros du début et ses listes à tirets à la prochaine ouverture.

Avec le format **CSV standard** ouvert dans un tableur, une valeur qui commence par `=` ou `@`, ou par `+` ou `-` suivi de texte, peut être prise pour une formule (un calcul, un lien, ou l'erreur `#NOM?`). Si le fichier contient des textes écrits par d'autres personnes (libellés de virement, par exemple), ne cliquez pas sur un lien inattendu dans une cellule.

## Pour les curieux et les développeurs

| Fichier | Rôle |
|---|---|
| `index.html`, `style.css` | La page. |
| `js/conversion.js` | Toute la logique de conversion, en fonctions pures (aucun accès au réseau ni à la page). |
| `js/app.js` | L'interface : choix du fichier, aperçu, téléchargement, effacement des données. |
| `exemples/` | Trois fichiers `.txt` de démonstration, aux données fictives. |
| `tests/` | Les tests automatiques. |
| `vercel.json` | Les règles de sécurité envoyées par l'hébergeur (Vercel) : les mêmes que dans la page, plus l'interdiction d'afficher l'outil à l'intérieur d'un autre site. |

Fichiers d'exemple :

- `contacts-utf8.txt` : UTF-8, points-virgules, valeurs entre guillemets contenant un point-virgule, des guillemets et un retour à la ligne ;
- `inventaire-windows1252.txt` : Windows-1252 avec fins de ligne Windows, tabulations, une ligne vide ;
- `budget-espaces-alignes.txt` : UTF-8 avec BOM, colonnes alignées par des espaces.

Aucune dépendance, aucune étape de construction. Les tests utilisent le module `node:test` intégré à Node.js (version 18.6 ou plus récente) :

```sh
node --test
```

Pour utiliser l'outil sans Internet, téléchargez le dépôt (bouton **Code › Download ZIP**), décompressez-le et ouvrez `index.html` dans votre navigateur.

## Licence

[MIT](LICENSE)
