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
- **Sauvegarde automatique** : avant toute modification, une copie `_old.docx` est créée dans `_sauvegardes/`
- **Sélection du code d'insertion** : combobox pré-rempli avec les clauses configurées — tous les champs se chargent automatiquement et restent éditables
- **Recherche de section** : saisir un mot-clé pour localiser la section cible ; les paragraphes correspondants sont surlignés en jaune dans la listbox
- **Affichage tout** : affiche l'intégralité des paragraphes du document sans filtre
- **Point d'insertion** : cliquer sur un paragraphe dans la listbox = insérer la clause **après** ce paragraphe
- **Prévisualisation HTML** : rendu du document en temps réel à droite ; le paragraphe sélectionné est mis en évidence (fond jaune, barre orange) avec scroll automatique
- **Compatibilité documents structurés** : les paragraphes dans des tableaux (`w:td`) ou des contrôles de contenu Word (`w:sdt`) sont détectés et affichés — fonctionne même si la quasi-totalité du contenu est dans des structures imbriquées
- **Mode révision Word (Track Changes)** : toute insertion est balisée `<w:ins>` dans le XML OOXML — elle apparaît en mode révision dans Word, avec le nom de l'auteur et la date
- **Champ Auteur** : le nom saisi est affiché dans la bulle de révision Word (défaut : « Juriste »)
- **Insertion** : écrase le `.docx` original, journalise l'opération, passe au fichier suivant
- **Passer sans insérer** : avancer au fichier suivant sans modifier le document courant
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
│   ├── docx_handler.py      # Manipulation .docx (lecture, insertion Track Changes, HTML)
│   └── logger.py            # Journalisation CSV des insertions
│
└── logs/
    └── insertions.csv       # Créé automatiquement au premier lancement
```

### Dossier de travail (généré à l'exécution)

```
[dossier_prospectus]/
├── fonds_A.docx             # Fichiers originaux (modifiés en place)
├── fonds_B.docx
└── _sauvegardes/
    ├── fonds_A_old.docx     # Copies de sécurité créées avant chaque modification
    └── fonds_B_old.docx
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
    "text": "La société de gestion peut, dans des circonstances...",
    "text_style": "Normal"
  }
}
```

| Champ | Obligatoire | Description |
|---|---|---|
| `subtitle` | Non | Texte du sous-titre. Laisser `""` si pas de sous-titre. |
| `subtitle_type` | Non | Format du sous-titre : `"bold"` (gras), `"style"` (style Word nommé) ou `"puce"` (puce avec indentation). Défaut : `"bold"`. |
| `subtitle_style` | Non | Nom du style Word appliqué quand `subtitle_type = "style"` (ex. `"Heading 3"`, `"Titre 2"`). |
| `subtitle_bullet` | Non | Caractère de puce quand `subtitle_type = "puce"` (ex. `"•"`, `"–"`, `"→"`). |
| `subtitle_indent` | Non | Niveau d'indentation de la puce (1, 2 ou 3). Correspond respectivement à 0.5, 1 et 1.5 pouce. |
| `text` | Oui | Corps de la clause. Texte brut, sans mise en forme. |
| `text_style` | Non | Style Word du corps de clause (ex. `"Normal"`, `"Corps de texte"`). Défaut : `"Normal"`. |

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
| `mammoth` | 1.11.0 | Présent dans le venv, non utilisé en production |
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
de l'exécutable pour que les clauses soient éditables sans repackager.

---

## Notes techniques

### Insertion en mode révision Word (Track Changes)

L'insertion utilise le mécanisme natif de révision d'Office Open XML. Chaque nouveau
paragraphe est balisé avec `<w:ins>` portant un identifiant de révision unique, le nom
de l'auteur et la date UTC :

```xml
<w:p>
  <w:pPr>
    [<w:pStyle w:val="Heading 3"/>]   <!-- si sous-titre de type "Style Word" -->
    [<w:ind w:left="720"/>]           <!-- si sous-titre de type "À puce" -->
    <w:rPr>
      <w:ins w:id="N" w:author="Juriste" w:date="2026-03-10T14:32:00Z"/>
    </w:rPr>
  </w:pPr>
  <w:ins w:id="N+1" w:author="Juriste" w:date="2026-03-10T14:32:00Z">
    <w:r>
      [<w:rPr><w:b/></w:rPr>]         <!-- si sous-titre de type "Gras" -->
      <w:t>texte inséré</w:t>
    </w:r>
  </w:ins>
</w:p>
```

Quand le document est ouvert dans Word, la clause apparaît surlignée (couleur de révision),
avec une bulle latérale indiquant l'auteur. Le juriste peut accepter ou refuser la modification.

Les identifiants de révision `w:id` doivent être uniques dans le document. L'outil scanne
le XML complet pour trouver le maximum existant et incrémente à partir de là.

### Format du sous-titre

Trois modes sont disponibles, configurables par clause et éditables à la volée :

| Mode | Rendu dans Word | Configuration |
|---|---|---|
| **Gras** | Texte en gras, style hérité du document | `subtitle_type: "bold"` |
| **Style Word** | Style nommé Word (ex. Heading 3) | `subtitle_type: "style"` + `subtitle_style: "Heading 3"` |
| **À puce** | Caractère de puce + indentation gauche | `subtitle_type: "puce"` + `subtitle_bullet` + `subtitle_indent` |

L'indentation des puces est exprimée en twips (1/1440 de pouce) :
niveau 1 = 720 twips (0,5"), niveau 2 = 1440 twips (1"), niveau 3 = 2160 twips (1,5").

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
- le calcul de la fraction de scroll
- l'ancre d'insertion (`addnext()`)

Les deux types de documents sont ainsi traités de façon identique.

### Prévisualisation HTML

`tkinterweb` utilise le moteur `tkhtml3` qui **ne supporte pas JavaScript**.
La mise en évidence du paragraphe sélectionné est donc réalisée côté serveur :
on régénère le HTML complet avec la classe CSS `.highlight` sur le paragraphe cible,
puis on scrolle via `yview_moveto()`. La position de scroll est calculée par une
heuristique pondérée par la longueur de chaque paragraphe (plus précis qu'un simple
ratio d'index).