# Insertion Texte — Outil d'insertion de clauses dans les prospectus

Outil desktop destiné au département juridique d'un asset manager pour automatiser
l'insertion de clauses réglementaires dans des prospectus et documents réglementaires au format `.docx`.

---

## Contexte

Le département juridique gère des centaines de fonds. Lors de chaque campagne
réglementaire (ex. : clauses LMT sur la gestion de la liquidité, clauses ESG, etc.),
les juristes doivent insérer de nouvelles clauses dans les prospectus de chaque fonds.

Avant cet outil, le processus était entièrement manuel :
1. Ouvrir le `.docx` dans Word
2. Naviguer jusqu'à la bonne section
3. Coller la clause depuis le presse-papier
4. Sauvegarder

Cet outil automatise les étapes répétitives tout en laissant le juriste décider
précisément du point d'insertion, fichier par fichier.

---

## Fonctionnalités

### Onglet Insertion

- **Sélection de dossier** : charge en une fois tous les `.docx` d'un dossier réseau
- **Navigation fichier par fichier** : boutons ◀ Précédent / Suivant ▶ avec compteur `x / n`
- **Fichiers originaux intacts** : aucune modification sur les `.docx` source — deux fichiers de sortie sont créés par insertion dans des sous-dossiers dédiés
- **Sélection du code d'insertion** : combobox pré-rempli avec les clauses configurées — tous les champs se chargent automatiquement et restent éditables
- **Style texte `auto`** : le corps de clause hérite automatiquement du style du paragraphe d'ancrage (`Normal`, `APU_Default`, `Descriptif`…). Si l'ancre est un titre, l'outil remonte aux paragraphes voisins pour trouver le style de corps réel — évite d'insérer le texte en police titre. Un style explicite peut être forcé clause par clause.
- **Taille de police** : spinbox 0–72 pt pour le sous-titre et pour le texte (0 = hérite du style)
- **Recherche de section** : saisir un mot-clé pour localiser la section cible ; les paragraphes correspondants sont surlignés en jaune dans la listbox
- **Affichage tout** : affiche l'intégralité des paragraphes du document sans filtre
- **Point d'insertion** : cliquer sur un paragraphe dans la listbox = insérer la clause **après** ce paragraphe
- **Prévisualisation HTML** : rendu du document en temps réel à droite ; le paragraphe sélectionné est mis en évidence (fond jaune, barre orange) — la vue est centrée sur la cible, toujours visible sans scroll
- **Vue complète** : bouton dans le panneau preview pour afficher le document entier ; cliquer à nouveau sur un paragraphe de la listbox repasse en vue fenêtrée
- **Compatibilité documents structurés** : les paragraphes dans des tableaux (`w:td`) ou des contrôles de contenu Word (`w:sdt`) sont détectés et affichés
- **Mode révision Word (Track Changes)** : la version `_track_changes/` est balisée `<w:ins>` dans le XML OOXML — elle apparaît en mode révision dans Word, avec le nom de l'auteur et la date
- **Champ Auteur** : le nom saisi est affiché dans les bulles de révision Word (défaut : « Juriste »)
- **Mise à jour automatique des dates** : à chaque insertion, l'outil met également à jour la date du jour dans les deux fichiers de sortie — deux patterns détectés automatiquement :
  - Corps du document : paragraphe contenant `Date de publication : JJ/MM/AAAA`
  - Footer Word : paragraphe contenant `mise à jour le JJ/MM/AAAA`
  Un popup avertit si aucun pattern n'est trouvé dans le fichier courant.
- **Insérer** : produit les deux fichiers de sortie, met à jour les dates, journalise l'opération, passe au fichier suivant
- **Passer sans insérer** : avancer au fichier suivant sans modifier ni créer de fichier
- **Historique** : les 20 dernières insertions sont affichées en bas de l'écran

### Onglet Clauses

Éditeur complet des codes d'insertion stockés dans `clauses.json` :

| Action | Description |
|---|---|
| **Sélectionner** | Clic sur un code → formulaire pré-rempli |
| **Modifier** | Éditer code, sous-titre ou texte → Enregistrer |
| **Renommer** | Changer le nom dans le champ Code → Enregistrer (l'ordre est conservé) |
| **Nouveau** | Bouton `+ Nouveau` → formulaire vide |
| **Dupliquer** | Copie le code sélectionné avec le suffixe `_copie` |
| **Supprimer** | Confirmation avant suppression définitive |
| **Annuler** | Restaure les valeurs sauvegardées |

- **Aperçu en temps réel** : un rendu HTML en bas du formulaire se régénère à chaque frappe — sous-titre dans son format (gras / souligné / heading / puce), extrait du corps avec note de style (`[style hérité]` si `auto`, nom du style sinon)

Toute modification est immédiatement écrite dans `clauses.json` et synchronisée
avec le combobox de l'onglet Insertion.

---

## Structure du projet

```
Insertion_texte/
├── main.py                  # Point d'entrée — lance la fenêtre Tkinter
├── clauses.json             # Base de données des clauses (codes d'insertion)
├── requirements.txt         # Dépendances Python
├── .gitignore
│
├── src/
│   ├── __init__.py
│   ├── app.py               # Interface graphique (Tkinter) — logique UI complète
│   ├── docx_handler.py      # Manipulation .docx (lecture, insertion, prévisualisation HTML)
│   ├── logger.py            # Journalisation CSV des insertions
│   └── paths.py             # Chemins runtime (compatible PyInstaller)
│
└── logs/
    └── insertions.csv       # Créé automatiquement au premier lancement
```

### Dossiers de sortie (générés à l'exécution dans le dossier de travail)

```
[dossier_prospectus]/
├── fonds_A.docx             # Fichiers originaux — jamais modifiés
├── fonds_B.docx
│
├── _track_changes/          # Créé automatiquement lors de la première insertion
│   ├── fonds_A.docx         # Insertion en mode révision Word (<w:ins>)
│   └── fonds_B.docx
│
└── _texte_brut/             # Créé automatiquement lors de la première insertion
    ├── fonds_A.docx         # Insertion en texte direct, sans balisage de révision
    └── fonds_B.docx
```

---

## Format de clauses.json

Chaque entrée est un **code d'insertion** identifié par une clé unique.
La convention de nommage recommandée est `CAMPAGNE_TYPEFONDS`
(ex. : `LMT_OPCVM`, `ESG_FIA`).

```json
{
  "LMT_OPCVM": {
    "subtitle": "Outils de gestion de la liquidité (LMT)",
    "subtitle_type": "style",
    "subtitle_style": "Heading 3",
    "subtitle_bullet": "•",
    "subtitle_indent": 1,
    "subtitle_font_size": 0,
    "text": "La société de gestion peut, dans des circonstances...",
    "text_style": "Normal",
    "text_font_size": 11
  }
}
```

| Champ | Obligatoire | Description |
|---|---|---|
| `subtitle` | Non | Texte du sous-titre. Laisser `""` si pas de sous-titre. |
| `subtitle_type` | Non | Format du sous-titre : `"bold"` (gras), `"underline"` (souligné), `"style"` (style Word nommé) ou `"puce"` (puce avec indentation). Défaut : `"bold"`. |
| `subtitle_style` | Non | Nom du style Word appliqué quand `subtitle_type = "style"` (ex. `"Heading 3"`, `"Titre 2"`). |
| `subtitle_bullet` | Non | Caractère de puce quand `subtitle_type = "puce"` (ex. `"•"`, `"–"`, `"→"`). |
| `subtitle_indent` | Non | Niveau d'indentation de la puce (1, 2 ou 3). Correspond respectivement à 0.5, 1 et 1.5 pouce. |
| `subtitle_font_size` | Non | Taille de police du sous-titre en points (ex. `12`). `0` = hérite du style. |
| `text` | Oui | Corps de la clause. Texte brut, sans mise en forme. |
| `text_style` | Non | Style Word du corps de clause. `"auto"` (défaut) = hérite du style du paragraphe d'ancrage au moment de l'insertion. Valeur explicite possible : `"Normal"`, `"Corps de texte"`, `"APU_Default"`… |
| `text_font_size` | Non | Taille de police du corps de clause en points. `0` = hérite du style. |

---

## Format du journal (logs/insertions.csv)

Le fichier est encodé en UTF-8, séparateur point-virgule, exploitable directement dans Excel.

| Colonne | Exemple |
|---|---|
| `date` | `2026-03-08` |
| `heure` | `14:32:07` |
| `fichier` | `prospectus_fonds_A.docx` |
| `code_insertion` | `LMT_OPCVM` |
| `paragraphe_index` | `42` |
| `sous_titre` | `Outils de gestion de la liquidité (LMT)` |
| `extrait_clause` | `La société de gestion peut, dans des circonstances...` |

---

## Installation

### Prérequis système

- Python 3.11+
- Tkinter (non inclus dans certaines distributions Linux)
- Debian/Ubuntu :

```bash
sudo apt install python3-tk
```

### Installation des dépendances

```bash
# Créer et activer le venv
python3 -m venv .venv
source .venv/bin/activate

# Installer les dépendances
pip install -r requirements.txt
```

### Dépendances (`requirements.txt`)

| Package | Version testée | Rôle |
|---|---|---|
| `python-docx` | 1.2.0 | Lecture et écriture des fichiers `.docx` |
| `tkinterweb` | 4.24.1 | Rendu HTML dans Tkinter (prévisualisation) |
| `pyinstaller` | 6.19.0 | Packaging en exécutable standalone |

---

## Lancement

```bash
source .venv/bin/activate
python main.py
```

---

## Packaging en exécutable (optionnel)

Pour distribuer l'outil aux juristes sans installation Python :

```bash
source .venv/bin/activate
pyinstaller --onefile --windowed main.py
```

L'exécutable est généré dans `dist/main`. Copier également `clauses.json` à côté
de l'exécutable — `src/paths.py` détecte automatiquement si l'application tourne
en mode frozen (`sys.frozen`) et résout les chemins (`clauses.json`, `logs/`)
relativement au dossier de l'exécutable plutôt qu'à la racine du projet source.

---

## Notes techniques

### Deux fichiers de sortie par insertion

L'original n'est jamais modifié. Chaque insertion produit deux copies indépendantes
depuis le fichier source :

| Dossier | Contenu | Usage |
|---|---|---|
| `_track_changes/` | Insertions balisées `<w:ins>` | Relecture dans Word avec mode révision |
| `_texte_brut/` | Insertions en paragraphes directs | Version finale acceptée, sans traces de révision |

Les dossiers sont créés automatiquement dans le dossier de travail à la première insertion.

### Insertion en mode révision Word (Track Changes)

L'insertion utilise le mécanisme natif de révision d'Office Open XML. Chaque nouveau
paragraphe est balisé avec `<w:ins>` portant un identifiant de révision unique, le nom
de l'auteur et la date UTC :

```xml
<w:p>
  <w:pPr>
    [<w:pStyle w:val="Heading3"/>]    <!-- ID de style, PAS le nom affiché "Heading 3" -->
    [<w:ind w:left="720"/>]           <!-- si sous-titre de type "À puce" -->
    <w:rPr>
      <w:ins w:id="N" w:author="Juriste" w:date="2026-03-10T14:32:00Z"/>
    </w:rPr>
  </w:pPr>
  <w:ins w:id="N+1" w:author="Juriste" w:date="2026-03-10T14:32:00Z">
    <w:r>
      [<w:rPr>
        [<w:b/>]                       <!-- si sous-titre de type "Gras" -->
        [<w:u w:val="single"/>]        <!-- si sous-titre de type "Souligné" -->
        [<w:sz w:val="24"/>            <!-- ex. 12pt → 24 demi-points -->
         <w:szCs w:val="24"/>]
      </w:rPr>]
      <w:t>texte inséré</w:t>
    </w:r>
  </w:ins>
</w:p>
```

Quand le document est ouvert dans Word, la clause apparaît surlignée (couleur de révision),
avec une bulle latérale indiquant l'auteur. Le juriste peut accepter ou refuser la modification.

Les identifiants de révision `w:id` doivent être uniques dans le document. L'outil scanne
le XML complet pour trouver le maximum existant et incrémente à partir de là.

La taille de police est exprimée en OOXML en **demi-points** : `<w:sz w:val="24"/>` = 12 pt.
`w:sz` s'applique aux scripts latins, `w:szCs` aux scripts complexes (arabe, hébreu, etc.).

### Format du sous-titre

Quatre modes sont disponibles, configurables par clause et éditables à la volée :

| Mode | Rendu dans Word | Configuration |
|---|---|---|
| **Gras** | Texte en gras, style hérité du document | `subtitle_type: "bold"` |
| **Souligné** | Texte souligné simple, style hérité | `subtitle_type: "underline"` |
| **Style Word** | Style nommé Word (ex. Heading 3) | `subtitle_type: "style"` + `subtitle_style: "Heading 3"` |
| **À puce** | Caractère de puce + indentation gauche | `subtitle_type: "puce"` + `subtitle_bullet` + `subtitle_indent` |

L'indentation des puces est exprimée en twips (1/1440 de pouce) :
niveau 1 = 720 twips (0,5"), niveau 2 = 1440 twips (1"), niveau 3 = 2160 twips (1,5").

### Résolution des identifiants de style

Word distingue le **nom affiché** d'un style (ex. `"Heading 3"`, `"APU_Heading 3"`) de
son **identifiant OOXML** (ex. `"Heading3"`, `"APUHeading3"`). L'attribut `<w:pStyle w:val="..."/>`
attend l'identifiant, pas le nom affiché — sinon Word ignore silencieusement le style.

Le champ `subtitle_style` de `clauses.json` et le combobox de l'interface stockent
le **nom affiché** (c'est ce que Word expose dans ses menus). L'outil résout
automatiquement ce nom en identifiant OOXML au moment de l'insertion via
`_resolve_style_id(doc, style_name)`, qui consulte les métadonnées du document cible
(`doc.styles`). La correspondance est donc correcte pour tous les templates
(classique, APU, RSV…) sans configuration manuelle des IDs.

### Insertion dans le .docx

L'insertion de paragraphes à une position arbitraire n'est pas supportée nativement
par `python-docx`. On manipule directement le XML Office Open XML sous-jacent via `lxml` :
chaque nouveau paragraphe est construit comme un élément `<w:p>` et inséré via
`addnext()` sur l'élément ancre. Pour respecter l'ordre final (sous-titre puis corps),
les éléments sont insérés dans l'ordre inversé.

### Documents structurés (tableaux et content controls)

`python-docx` expose `doc.paragraphs` qui ne retourne que les paragraphes enfants
directs de `<w:body>`. Les paragraphes imbriqués dans des tableaux (`w:tbl > w:tr > w:tc`)
ou des contrôles de contenu Word (`w:sdt > w:sdtContent`) n'y apparaissent pas — la
prévisualisation serait vide pour ces documents.

L'outil résout cela avec `collect_paragraphs(doc)` :

```python
list(doc.element.body.iter(qn('w:p')))
```

`iter()` fait une traversée en profondeur du XML et retourne **tous** les `<w:p>`
dans l'ordre de lecture, quelle que soit leur profondeur d'imbrication.

La liste résultante (`flat_paras`) est la référence unique utilisée pour :
- l'affichage dans la listbox
- la recherche de mot-clé
- la prévisualisation HTML (avec highlight)
- l'ancre d'insertion (`addnext()`)

Les deux types de documents sont ainsi traités de façon identique.

### Prévisualisation HTML

`tkinterweb` utilise le moteur `tkhtml3` qui **ne supporte pas JavaScript**.
Le scroll programmatique (`yview_moveto`) est peu fiable car les hauteurs rendues
sont inconnues. La mise en évidence du paragraphe sélectionné est donc résolue
par une **vue fenêtrée** : on génère un fragment HTML contenant uniquement
4 paragraphes avant la cible et 25 après. La cible est ainsi toujours en haut
du contenu affiché — aucun scroll nécessaire.

Le bouton **Vue complète** régénère le document entier (`build_html`). Cliquer
à nouveau sur un paragraphe dans la listbox repasse en vue fenêtrée (`build_html_window`).

### Style `auto` — héritage du style d'ancrage

Quand `text_style == "auto"`, le style appliqué au corps de clause est déterminé
dynamiquement au moment de l'insertion, en lisant le style OOXML du paragraphe
d'ancrage (`w:pStyle w:val`).

**Cas particulier — ancre sur un titre :** si le paragraphe sélectionné est un
titre (`Heading`, `Titre`, `APU_Heading`, `RSV_Heading`…), appliquer son style au
corps de clause produirait du texte en police titre. L'outil détecte ce cas et
remonte aux paragraphes voisins pour trouver le premier style de corps de texte
(scan arrière puis avant). Si aucun style de corps n'est trouvé, Word applique
son style par défaut.

Les documents peuvent mélanger plusieurs familles de styles dans le même fichier
(`Normal`, `APU_Default`, `Default_RSV`, `Descriptif`…) : `auto` s'adapte à
chaque point d'insertion sans configuration supplémentaire.

### Mise à jour automatique des dates

À chaque insertion, l'outil cherche et remplace la date du jour dans deux
emplacements du document (détection automatique, les deux peuvent coexister) :

| Emplacement | Pattern recherché | Exemple |
|---|---|---|
| Corps du document | Paragraphe contenant `Date de publication` | `Date de publication : 01/01/2026` |
| Footer Word | Paragraphe contenant `mise à jour le` | `Dernière mise à jour le 28/11/2025` |

La date remplacée doit être au format `JJ/MM/AAAA` et être contenue dans un seul
run OOXML (cas standard pour les dates générées automatiquement).

**En mode Track Changes** (`_track_changes/`) : l'ancienne date est balisée
`<w:del>`, la nouvelle `<w:ins>`. Les propriétés de caractère du run original
(`w:rPr` : police, taille, gras…) sont conservées dans tous les fragments.

```xml
<!-- Exemple de remplacement Track Changes dans un run "28/11/2025" → "20/03/2026" -->
<w:del w:id="N" w:author="Juriste" w:date="2026-03-20T...">
  <w:r><w:rPr>...</w:rPr><w:delText>28/11/2025</w:delText></w:r>
</w:del>
<w:ins w:id="N+1" w:author="Juriste" w:date="2026-03-20T...">
  <w:r><w:rPr>...</w:rPr><w:t>20/03/2026</w:t></w:r>
</w:ins>
```

**En mode texte brut** (`_texte_brut/`) : remplacement direct dans `<w:t>`.

Si aucun des deux patterns n'est trouvé dans le fichier, un popup avertit
l'utilisateur (la clause est quand même insérée).

### Compatibilité PyInstaller

`src/paths.py` expose les chemins utilisés par l'application (`clauses.json`,
dossier `logs/`) de façon compatible avec un exécutable PyInstaller :

```python
# sys.frozen = True quand l'app tourne depuis un exe PyInstaller
if getattr(sys, "frozen", False):
    RUNTIME_DIR = Path(sys.executable).parent   # dossier du .exe
else:
    RUNTIME_DIR = Path(__file__).parent.parent  # racine du projet (dev)
```

Lors du packaging, `clauses.json` doit être copié manuellement à côté de l'exécutable.
