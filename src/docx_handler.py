# =============================================================================
# src/docx_handler.py — Manipulation des fichiers .docx
#
# Ce module gère toutes les opérations sur les documents Word :
#   - Sauvegarde de sécurité avant modification
#   - Lecture et recherche de paragraphes
#   - Génération du HTML de prévisualisation
#   - Insertion de nouvelles clauses dans le corps du document
# =============================================================================

import html as _html
import shutil
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement

# Correspondance entre les noms de styles Word (EN et FR) et les balises HTML.
# Utilisé pour reproduire la hiérarchie des titres dans la prévisualisation.
_HEADING_TAGS: dict[str, str] = {
    "heading 1": "h1", "titre 1": "h1",
    "heading 2": "h2", "titre 2": "h2",
    "heading 3": "h3", "titre 3": "h3",
    "heading 4": "h4", "titre 4": "h4",
}


# -----------------------------------------------------------------------------
# Prévisualisation
# -----------------------------------------------------------------------------

def build_html(doc: Document, highlight_idx: int | None = None) -> str:
    """
    Génère un fragment HTML à partir des paragraphes du document.

    Chaque paragraphe non vide reçoit un attribut id="para-{i}" où i est son
    index dans doc.paragraphs. Cela permet à l'interface de faire le lien entre
    la listbox (qui affiche les index) et la prévisualisation HTML.

    Le paragraphe highlight_idx reçoit class="highlight" pour être mis en
    évidence visuellement (fond jaune + barre latérale orange).

    Args:
        doc: Document python-docx ouvert.
        highlight_idx: Index du paragraphe à mettre en évidence, ou None.

    Returns:
        Chaîne HTML (fragment <body>, sans les balises <html>/<head>).
    """
    parts = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text
        if not text.strip():
            # On ignore les paragraphes vides (sauts de ligne, séparateurs)
            continue

        # Détermination de la balise HTML selon le style Word du paragraphe
        style = (para.style.name or "").lower() if para.style else ""
        tag = next((v for k, v in _HEADING_TAGS.items() if k in style), "p")

        cls = ' class="highlight"' if i == highlight_idx else ""
        escaped = _html.escape(text)  # Échappe les caractères spéciaux HTML (<, >, &, etc.)
        parts.append(f'<{tag} id="para-{i}"{cls}>{escaped}</{tag}>')

    return "\n".join(parts)


# -----------------------------------------------------------------------------
# Gestion des fichiers
# -----------------------------------------------------------------------------

def backup_and_open(filepath: str, backup_dir: Path) -> Document:
    """
    Crée une copie de sécurité du fichier avant toute modification, puis ouvre
    le document original.

    La sauvegarde est nommée <nom_du_fichier>_old.docx et placée dans backup_dir.
    Si backup_dir n'existe pas, il est créé automatiquement.

    Args:
        filepath: Chemin absolu du fichier .docx à ouvrir.
        backup_dir: Dossier de destination pour la copie de sécurité.

    Returns:
        Instance Document python-docx prête à être manipulée.
    """
    path = Path(filepath)
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup = backup_dir / (path.stem + "_old" + path.suffix)
    shutil.copy2(path, backup)
    return Document(filepath)


def save_document(doc: Document, filepath: str) -> None:
    """
    Sauvegarde le document à son emplacement d'origine (écrase le fichier).

    Args:
        doc: Document python-docx modifié.
        filepath: Chemin de destination (généralement le fichier d'origine).
    """
    doc.save(filepath)


# -----------------------------------------------------------------------------
# Recherche et navigation dans les paragraphes
# -----------------------------------------------------------------------------

def search_paragraphs(doc: Document, keyword: str) -> list[tuple[int, str]]:
    """
    Recherche un mot-clé (insensible à la casse) dans tous les paragraphes.

    Args:
        doc: Document python-docx ouvert.
        keyword: Terme à rechercher.

    Returns:
        Liste de tuples (index, texte) pour chaque paragraphe contenant le mot-clé.
    """
    keyword_lower = keyword.strip().lower()
    return [
        (i, para.text)
        for i, para in enumerate(doc.paragraphs)
        if keyword_lower and keyword_lower in para.text.lower()
    ]


def get_paragraphs_around(doc: Document, center_idx: int, context: int = 4) -> list[tuple[int, str]]:
    """
    Retourne les paragraphes situés autour d'un index central.

    Utilisé pour afficher le contexte autour d'un résultat de recherche :
    on montre quelques paragraphes avant et après la section trouvée.

    Args:
        doc: Document python-docx ouvert.
        center_idx: Index du paragraphe central (résultat de recherche).
        context: Nombre de paragraphes à inclure de chaque côté (défaut : 4).

    Returns:
        Liste de tuples (index, texte) dans l'ordre du document.
    """
    start = max(0, center_idx - context)
    end = min(len(doc.paragraphs), center_idx + context + 1)
    return [(i, doc.paragraphs[i].text) for i in range(start, end)]


def get_all_paragraphs(doc: Document) -> list[tuple[int, str]]:
    """
    Retourne tous les paragraphes non vides du document avec leur index.

    Args:
        doc: Document python-docx ouvert.

    Returns:
        Liste de tuples (index, texte) pour chaque paragraphe non vide.
    """
    return [(i, p.text) for i, p in enumerate(doc.paragraphs) if p.text.strip()]


# -----------------------------------------------------------------------------
# Insertion de clause
# -----------------------------------------------------------------------------

def _make_paragraph_element(text: str, bold: bool = False) -> OxmlElement:
    """
    Construit un élément XML Word <w:p> contenant un run de texte.

    python-docx ne propose pas de méthode native pour insérer un paragraphe
    à une position arbitraire dans le document. On manipule donc directement
    le XML sous-jacent (Office Open XML) via lxml.

    Structure produite :
        <w:p>
          <w:r>
            [<w:rPr><w:b/></w:rPr>]   ← uniquement si bold=True
            <w:t xml:space="preserve">texte</w:t>
          </w:r>
        </w:p>

    Args:
        text: Texte du paragraphe.
        bold: Si True, applique le style gras au run.

    Returns:
        Élément OxmlElement <w:p> prêt à être inséré dans le corps du document.
    """
    new_p = OxmlElement("w:p")
    new_r = OxmlElement("w:r")

    if bold:
        # Propriétés du run : mise en gras pour les sous-titres de clause
        rPr = OxmlElement("w:rPr")
        b = OxmlElement("w:b")
        rPr.append(b)
        new_r.append(rPr)

    new_t = OxmlElement("w:t")
    new_t.text = text
    # xml:space="preserve" est requis quand le texte commence ou finit par une espace,
    # sinon Word supprime les espaces en bordure à l'affichage.
    if text.startswith(" ") or text.endswith(" "):
        new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    new_r.append(new_t)
    new_p.append(new_r)
    return new_p


def insert_clause_after(doc: Document, para_idx: int, subtitle: str, text: str) -> None:
    """
    Insère une clause (sous-titre optionnel + corps de texte) dans le document,
    immédiatement après le paragraphe désigné par para_idx.

    Technique d'insertion :
        On utilise addnext() de lxml, qui insère un élément XML juste après un
        élément de référence dans l'arbre. Pour obtenir l'ordre final :
            [paragraphe_ancre] → [sous-titre] → [corps de clause]
        on insère les éléments dans l'ordre INVERSÉ, chacun après l'ancre :
            addnext(corps)    → ancre | corps
            addnext(sous-titre) → ancre | sous-titre | corps

    Args:
        doc: Document python-docx ouvert et modifiable.
        para_idx: Index du paragraphe après lequel insérer la clause.
        subtitle: Sous-titre de la clause (affiché en gras). Ignoré si vide.
        text: Corps de la clause à insérer.
    """
    anchor = doc.paragraphs[para_idx]._p  # Élément XML <w:p> de référence

    # Construction de la liste des éléments à insérer, dans l'ordre final souhaité
    items = []
    if subtitle and subtitle.strip():
        items.append({"text": subtitle.strip(), "bold": True})
    items.append({"text": text.strip(), "bold": False})

    # Insertion en ordre inversé pour que addnext produise le bon ordre final
    for item in reversed(items):
        new_p = _make_paragraph_element(item["text"], bold=item["bold"])
        anchor.addnext(new_p)
