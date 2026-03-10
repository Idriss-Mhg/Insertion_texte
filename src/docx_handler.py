# =============================================================================
# src/docx_handler.py — Manipulation des fichiers .docx
#
# Ce module gère toutes les opérations sur les documents Word :
#   - Sauvegarde de sécurité avant modification
#   - Lecture et recherche de paragraphes
#   - Génération du HTML de prévisualisation
#   - Insertion de nouvelles clauses en mode révision (Track Changes / <w:ins>)
# =============================================================================

import html as _html
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Correspondance entre les noms de styles Word (EN et FR) et les balises HTML.
# Utilisé pour reproduire la hiérarchie des titres dans la prévisualisation.
_HEADING_TAGS: dict[str, str] = {
    "heading 1": "h1", "titre 1": "h1",
    "heading 2": "h2", "titre 2": "h2",
    "heading 3": "h3", "titre 3": "h3",
    "heading 4": "h4", "titre 4": "h4",
}

# Niveaux d'indentation pour les puces, en twips (1 twip = 1/1440 pouce).
# Niveau 1 = 0.5 pouce, niveau 2 = 1 pouce, niveau 3 = 1.5 pouce.
INDENT_LEVELS: dict[int, int] = {1: 720, 2: 1440, 3: 2160}


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
            continue
        style = (para.style.name or "").lower() if para.style else ""
        tag = next((v for k, v in _HEADING_TAGS.items() if k in style), "p")
        cls = ' class="highlight"' if i == highlight_idx else ""
        escaped = _html.escape(text)
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
    """
    path = Path(filepath)
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup = backup_dir / (path.stem + "_old" + path.suffix)
    shutil.copy2(path, backup)
    return Document(filepath)


def save_document(doc: Document, filepath: str) -> None:
    """Sauvegarde le document à son emplacement d'origine (écrase le fichier)."""
    doc.save(filepath)


# -----------------------------------------------------------------------------
# Recherche et navigation dans les paragraphes
# -----------------------------------------------------------------------------

def search_paragraphs(doc: Document, keyword: str) -> list[tuple[int, str]]:
    """Retourne la liste des (index, texte) des paragraphes contenant le mot-clé."""
    keyword_lower = keyword.strip().lower()
    return [
        (i, para.text)
        for i, para in enumerate(doc.paragraphs)
        if keyword_lower and keyword_lower in para.text.lower()
    ]


def get_paragraphs_around(doc: Document, center_idx: int, context: int = 4) -> list[tuple[int, str]]:
    """Retourne les paragraphes situés autour d'un index central (pour le contexte de recherche)."""
    start = max(0, center_idx - context)
    end = min(len(doc.paragraphs), center_idx + context + 1)
    return [(i, doc.paragraphs[i].text) for i in range(start, end)]


def get_all_paragraphs(doc: Document) -> list[tuple[int, str]]:
    """Retourne tous les paragraphes non vides du document avec leur index."""
    return [(i, p.text) for i, p in enumerate(doc.paragraphs) if p.text.strip()]


# -----------------------------------------------------------------------------
# Insertion de clause en mode révision (Track Changes)
# -----------------------------------------------------------------------------

def _next_revision_id(doc: Document) -> int:
    """
    Calcule le prochain identifiant de révision disponible dans le document.

    Chaque élément <w:ins> / <w:del> doit porter un w:id unique. On scanne
    le XML brut pour trouver le maximum existant et on retourne max + 1.
    """
    existing_ids = [int(m) for m in re.findall(r'w:id="(\d+)"', doc.element.xml)]
    return (max(existing_ids) + 1) if existing_ids else 1


def _make_tracked_paragraph(
    text: str,
    author: str,
    date_str: str,
    rev_id: int,
    bold: bool = False,
    style_name: str | None = None,
    indent_twips: int = 0,
) -> tuple[OxmlElement, int]:
    """
    Construit un <w:p> dont le contenu est balisé comme insertion Track Changes.

    Structure XML produite :
        <w:p>
          <w:pPr>
            [<w:pStyle w:val="..."/>]       ← si style_name fourni
            [<w:ind w:left="..."/>]         ← si indent_twips > 0
            <w:rPr>
              <w:ins w:id="N" w:author="..." w:date="..."/>
            </w:rPr>
          </w:pPr>
          <w:ins w:id="N+1" w:author="..." w:date="...">
            <w:r>
              [<w:rPr><w:b/></w:rPr>]      ← si bold=True
              <w:t>texte</w:t>
            </w:r>
          </w:ins>
        </w:p>

    Args:
        text: Texte du paragraphe.
        author: Auteur affiché dans la bulle de révision Word.
        date_str: Date ISO 8601 UTC (ex. "2026-03-10T14:32:00Z").
        rev_id: Premier ID de révision disponible (deux IDs consommés).
        bold: Applique le gras au run (sous-titres en mode "Gras").
        style_name: Nom du style Word à appliquer (ex. "Heading 3", "Normal").
                    None = pas de style explicite (hérite du document).
        indent_twips: Indentation gauche en twips pour les puces (0 = aucune).

    Returns:
        Tuple (élément <w:p>, prochain rev_id disponible).
    """
    new_p = OxmlElement("w:p")

    # ── Propriétés de paragraphe ──────────────────────────────────────────────
    pPr = OxmlElement("w:pPr")

    if style_name:
        # Applique un style Word nommé (Heading 3, Normal, Titre 2, etc.)
        pStyle = OxmlElement("w:pStyle")
        pStyle.set(qn("w:val"), style_name)
        pPr.append(pStyle)

    if indent_twips > 0:
        # Indentation gauche pour les paragraphes à puce
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(indent_twips))
        pPr.append(ind)

    # Marque le signe de paragraphe (¶) lui-même comme inséré
    pRpr = OxmlElement("w:rPr")
    ins_pmark = OxmlElement("w:ins")
    ins_pmark.set(qn("w:id"),     str(rev_id))
    ins_pmark.set(qn("w:author"), author)
    ins_pmark.set(qn("w:date"),   date_str)
    pRpr.append(ins_pmark)
    pPr.append(pRpr)
    new_p.append(pPr)

    # ── Run de texte encapsulé dans <w:ins> ───────────────────────────────────
    ins_run = OxmlElement("w:ins")
    ins_run.set(qn("w:id"),     str(rev_id + 1))
    ins_run.set(qn("w:author"), author)
    ins_run.set(qn("w:date"),   date_str)

    new_r = OxmlElement("w:r")
    if bold:
        rPr = OxmlElement("w:rPr")
        rPr.append(OxmlElement("w:b"))
        new_r.append(rPr)

    new_t = OxmlElement("w:t")
    new_t.text = text
    if text.startswith(" ") or text.endswith(" "):
        new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    new_r.append(new_t)

    ins_run.append(new_r)
    new_p.append(ins_run)

    return new_p, rev_id + 2


def insert_clause_after(
    doc: Document,
    para_idx: int,
    subtitle: str,
    text: str,
    author: str,
    subtitle_config: dict | None = None,
    text_style: str | None = None,
) -> None:
    """
    Insère une clause en mode révision Word (Track Changes) immédiatement
    après le paragraphe désigné par para_idx.

    Format du sous-titre contrôlé par subtitle_config :
        {"type": "bold"}                    → texte gras (défaut)
        {"type": "style", "style": "Heading 3"}   → style Word nommé
        {"type": "puce",  "bullet": "•", "indent": 1} → puce avec indentation

    Args:
        doc: Document python-docx ouvert.
        para_idx: Index du paragraphe d'ancrage (insertion après).
        subtitle: Texte du sous-titre (ignoré si vide).
        text: Corps de la clause.
        author: Nom affiché dans la bulle de révision.
        subtitle_config: Dict décrivant le format du sous-titre (voir ci-dessus).
        text_style: Style Word du corps de clause (ex. "Normal"). None = défaut.
    """
    anchor = doc.paragraphs[para_idx]._p
    date_str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    rev_id = _next_revision_id(doc)

    cfg = subtitle_config or {}
    sub_type = cfg.get("type", "bold")

    items = []

    if subtitle and subtitle.strip():
        if sub_type == "style":
            items.append({
                "text": subtitle.strip(),
                "bold": False,
                "style_name": cfg.get("style", "Heading 3"),
                "indent_twips": 0,
            })
        elif sub_type == "puce":
            bullet = cfg.get("bullet", "•")
            level = int(cfg.get("indent", 1))
            items.append({
                "text": f"{bullet}\t{subtitle.strip()}",
                "bold": False,
                "style_name": None,
                "indent_twips": INDENT_LEVELS.get(level, 720),
            })
        else:  # "bold" (défaut)
            items.append({
                "text": subtitle.strip(),
                "bold": True,
                "style_name": None,
                "indent_twips": 0,
            })

    items.append({
        "text": text.strip(),
        "bold": False,
        "style_name": text_style or None,
        "indent_twips": 0,
    })

    # Insertion en ordre inversé pour que addnext produise l'ordre final correct
    for item in reversed(items):
        new_p, rev_id = _make_tracked_paragraph(
            item["text"], author, date_str, rev_id,
            bold=item["bold"],
            style_name=item["style_name"],
            indent_twips=item["indent_twips"],
        )
        anchor.addnext(new_p)