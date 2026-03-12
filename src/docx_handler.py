# =============================================================================
# src/docx_handler.py — Manipulation des fichiers .docx
#
# Ce module gère toutes les opérations sur les documents Word :
#   - Sauvegarde de sécurité avant modification
#   - Lecture et recherche de paragraphes (y compris dans tableaux et SDT)
#   - Génération du HTML de prévisualisation
#   - Insertion de nouvelles clauses en mode révision (Track Changes / <w:ins>)
#
# Deux types de documents sont supportés :
#   1. Documents "classiques" : paragraphes directement dans <w:body>
#   2. Documents "structurés" : contenu dans des tableaux (<w:td>) ou des
#      contrôles de contenu Word (<w:sdt>/<w:sdtContent>)
#
# python-docx expose uniquement les paragraphes de premier niveau via
# doc.paragraphs. Ce module utilise collect_paragraphs() pour traverser le
# XML complet et capturer tous les paragraphes quelle que soit leur profondeur.
# =============================================================================

import html as _html
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Niveaux d'indentation pour les puces, en twips (1 twip = 1/1440 pouce).
# Niveau 1 = 0.5", niveau 2 = 1", niveau 3 = 1.5"
INDENT_LEVELS: dict[int, int] = {1: 720, 2: 1440, 3: 2160}


# =============================================================================
# Helpers bas niveau sur les éléments <w:p> lxml
#
# Ces fonctions opèrent directement sur des éléments lxml (pas des objets
# python-docx), ce qui permet de traiter des paragraphes à n'importe quelle
# profondeur dans le document XML.
# =============================================================================

def _para_text(p_elem) -> str:
    """
    Extrait le texte brut d'un élément <w:p> lxml en concaténant tous ses <w:t>.

    Inclut le texte des runs normaux et des insertions Track Changes (<w:ins>),
    mais PAS le texte des suppressions (<w:delText>), ce qui reflète l'état
    "accepté" du document.
    """
    return ''.join(t.text or '' for t in p_elem.iter(qn('w:t')))


def _para_html_tag(p_elem) -> str:
    """
    Retourne la balise HTML (h1–h4 ou p) correspondant au style du paragraphe.

    Travaille sur l'ID de style Word (attribut w:val de w:pStyle), pas sur
    le nom complet. Les IDs typiques sont : "Heading1", "Titre2", "1", etc.

    Note : l'ID de style ≠ nom de style. "Heading1" (ID) → "Heading 1" (nom).
    Cette correspondance est suffisante pour la prévisualisation HTML.
    """
    pPr = p_elem.find(qn('w:pPr'))
    if pPr is not None:
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is not None:
            sid = pStyle.get(qn('w:val'), '').lower().replace(' ', '')
            for lvl in range(1, 5):
                if any(k in sid for k in (f'heading{lvl}', f'titre{lvl}', f'title{lvl}')):
                    return f'h{lvl}'
                # Certains templates utilisent "1", "2", etc. comme ID de style
                if sid == str(lvl):
                    return f'h{lvl}'
    return 'p'


# =============================================================================
# Collecte des paragraphes (traversée complète du XML)
# =============================================================================

def collect_paragraphs(doc: Document) -> list:
    """
    Retourne la liste plate de tous les éléments <w:p> du corps du document,
    en ordre de lecture, y compris ceux imbriqués dans :
      - des tableaux          : w:body > w:tbl > w:tr > w:tc > w:p
      - des content controls  : w:body > w:sdt > w:sdtContent > w:p
      - des structures mixtes : content controls contenant des tableaux, etc.

    À utiliser à la place de doc.paragraphs, qui ne retourne que les paragraphes
    enfants directs de w:body (paragraphes de premier niveau).

    Returns:
        Liste d'éléments lxml <w:p> dans l'ordre du document. Les indices de
        cette liste sont utilisés comme référence dans toute l'application
        (listbox, highlight, insertion).
    """
    return list(doc.element.body.iter(qn('w:p')))


# =============================================================================
# Prévisualisation HTML
# =============================================================================

def build_html(
    doc: Document,
    highlight_idx: int | None = None,
    flat_paras: list | None = None,
) -> str:
    """
    Génère un fragment HTML à partir des paragraphes du document.

    Chaque paragraphe non vide reçoit un attribut id="para-{i}" où i est son
    index dans flat_paras. Cela permet à l'interface de faire le lien entre
    la listbox (qui affiche les index) et la prévisualisation HTML.

    Le paragraphe highlight_idx reçoit class="highlight" pour être mis en
    évidence visuellement (fond jaune + barre latérale orange).

    Args:
        doc: Document python-docx (utilisé uniquement si flat_paras est None).
        highlight_idx: Index du paragraphe à mettre en évidence, ou None.
        flat_paras: Liste plate des éléments <w:p> produite par collect_paragraphs().
                    Si None, se replie sur [p._p for p in doc.paragraphs] (rétrocompat).

    Returns:
        Chaîne HTML (fragment <body>, sans les balises <html>/<head>).
    """
    paras = flat_paras if flat_paras is not None else [p._p for p in doc.paragraphs]
    parts = []
    for i, p_elem in enumerate(paras):
        text = _para_text(p_elem)
        if not text.strip():
            continue
        tag = _para_html_tag(p_elem)
        cls = ' class="highlight"' if i == highlight_idx else ""
        escaped = _html.escape(text)
        parts.append(f'<{tag} id="para-{i}"{cls}>{escaped}</{tag}>')
    return "\n".join(parts)


# =============================================================================
# Gestion des fichiers
# =============================================================================

def backup_and_open(filepath: str, backup_dir: Path) -> Document:
    """
    Crée une copie de sécurité du fichier avant toute modification, puis ouvre
    le document original.

    La sauvegarde est nommée <nom_du_fichier>_old.docx et placée dans backup_dir.
    Si backup_dir n'existe pas, il est créé automatiquement.

    La copie n'est créée qu'une seule fois (si elle n'existe pas encore), afin de
    toujours conserver l'original vierge — même si le fichier est rechargé plusieurs
    fois (navigation Précédent/Suivant ou plusieurs sessions sur le même dossier).
    """
    path = Path(filepath)
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup = backup_dir / (path.stem + "_old" + path.suffix)
    if not backup.exists():
        shutil.copy2(path, backup)
    return Document(filepath)


def save_document(doc: Document, filepath: str) -> None:
    """Sauvegarde le document à son emplacement d'origine (écrase le fichier)."""
    doc.save(filepath)


# =============================================================================
# Recherche et navigation dans les paragraphes
# =============================================================================

def search_paragraphs(
    doc: Document,
    keyword: str,
    flat_paras: list | None = None,
) -> list[tuple[int, str]]:
    """
    Retourne la liste des (index, texte) des paragraphes contenant le mot-clé.

    Args:
        doc: Document python-docx (utilisé uniquement si flat_paras est None).
        keyword: Terme à rechercher (insensible à la casse).
        flat_paras: Liste plate produite par collect_paragraphs().
    """
    paras = flat_paras if flat_paras is not None else [p._p for p in doc.paragraphs]
    keyword_lower = keyword.strip().lower()
    results = []
    for i, p_elem in enumerate(paras):
        text = _para_text(p_elem)
        if keyword_lower and keyword_lower in text.lower():
            results.append((i, text))
    return results


def get_paragraphs_around(
    doc: Document,
    center_idx: int,
    context: int = 4,
    flat_paras: list | None = None,
) -> list[tuple[int, str]]:
    """
    Retourne les paragraphes situés dans une fenêtre de ±context autour
    de center_idx (utilisé pour afficher le contexte d'une occurrence de recherche).

    Args:
        doc: Document python-docx (utilisé uniquement si flat_paras est None).
        center_idx: Index central (paragraphe trouvé par la recherche).
        context: Nombre de paragraphes à afficher de chaque côté.
        flat_paras: Liste plate produite par collect_paragraphs().
    """
    paras = flat_paras if flat_paras is not None else [p._p for p in doc.paragraphs]
    start = max(0, center_idx - context)
    end = min(len(paras), center_idx + context + 1)
    return [(i, _para_text(paras[i])) for i in range(start, end)]


def get_all_paragraphs(
    doc: Document,
    flat_paras: list | None = None,
) -> list[tuple[int, str]]:
    """
    Retourne tous les paragraphes non vides du document avec leur index.

    Args:
        doc: Document python-docx (utilisé uniquement si flat_paras est None).
        flat_paras: Liste plate produite par collect_paragraphs().
    """
    paras = flat_paras if flat_paras is not None else [p._p for p in doc.paragraphs]
    result = []
    for i, p_elem in enumerate(paras):
        text = _para_text(p_elem)
        if text.strip():
            result.append((i, text))
    return result


# =============================================================================
# Insertion de clause en mode révision (Track Changes)
# =============================================================================

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
    underline: bool = False,
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
              [<w:rPr>
                [<w:b/>]                   ← si bold=True
                [<w:u w:val="single"/>]    ← si underline=True
              </w:rPr>]
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
        underline: Applique le soulignement simple au run (sous-titres en mode "Souligné").
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
    if bold or underline:
        # On regroupe toutes les propriétés de caractère dans un seul <w:rPr>
        rPr = OxmlElement("w:rPr")
        if bold:
            rPr.append(OxmlElement("w:b"))
        if underline:
            # w:u w:val="single" = soulignement simple (le plus courant dans Word)
            u = OxmlElement("w:u")
            u.set(qn("w:val"), "single")
            rPr.append(u)
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
    flat_paras: list | None = None,
) -> None:
    """
    Insère une clause en mode révision Word (Track Changes) immédiatement
    après le paragraphe désigné par para_idx.

    Utilise flat_paras[para_idx] comme élément ancre si flat_paras est fourni,
    ce qui permet d'insérer après des paragraphes dans des tableaux ou des
    content controls. Sinon, se replie sur doc.paragraphs[para_idx]._p.

    Format du sous-titre contrôlé par subtitle_config :
        {"type": "bold"}                              → texte gras (défaut)
        {"type": "underline"}                         → texte souligné
        {"type": "style", "style": "Heading 3"}       → style Word nommé
        {"type": "puce",  "bullet": "•", "indent": 1} → puce avec indentation

    Args:
        doc: Document python-docx ouvert.
        para_idx: Index dans flat_paras (ou doc.paragraphs si flat_paras=None).
        subtitle: Texte du sous-titre (ignoré si vide).
        text: Corps de la clause.
        author: Nom affiché dans la bulle de révision.
        subtitle_config: Dict décrivant le format du sous-titre (voir ci-dessus).
        text_style: Style Word du corps de clause (ex. "Normal"). None = défaut.
        flat_paras: Liste plate produite par collect_paragraphs(). Recommandé.
    """
    # Élément XML ancre — l'insertion se fait via addnext() juste après lui
    if flat_paras is not None:
        anchor = flat_paras[para_idx]
    else:
        anchor = doc.paragraphs[para_idx]._p

    date_str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    rev_id = _next_revision_id(doc)

    cfg = subtitle_config or {}
    sub_type = cfg.get("type", "bold")

    # Construire la liste des paragraphes à insérer (sous-titre + corps)
    items = []

    if subtitle and subtitle.strip():
        if sub_type == "style":
            items.append({
                "text": subtitle.strip(),
                "bold": False, "underline": False,
                "style_name": cfg.get("style", "Heading 3"),
                "indent_twips": 0,
            })
        elif sub_type == "puce":
            bullet = cfg.get("bullet", "•")
            level = int(cfg.get("indent", 1))
            items.append({
                "text": f"{bullet}\t{subtitle.strip()}",
                "bold": False, "underline": False,
                "style_name": None,
                "indent_twips": INDENT_LEVELS.get(level, 720),
            })
        elif sub_type == "underline":
            items.append({
                "text": subtitle.strip(),
                "bold": False, "underline": True,
                "style_name": None,
                "indent_twips": 0,
            })
        else:  # "bold" (défaut)
            items.append({
                "text": subtitle.strip(),
                "bold": True, "underline": False,
                "style_name": None,
                "indent_twips": 0,
            })

    items.append({
        "text": text.strip(),
        "bold": False, "underline": False,
        "style_name": text_style or None,
        "indent_twips": 0,
    })

    # Insertion en ordre inversé pour que addnext() produise l'ordre final correct.
    # addnext(X) insère X immédiatement après l'ancre. En insérant dans l'ordre
    # [corps, sous-titre] (inversé), le résultat dans le document est [sous-titre, corps].
    for item in reversed(items):
        new_p, rev_id = _make_tracked_paragraph(
            item["text"], author, date_str, rev_id,
            bold=item["bold"],
            underline=item["underline"],
            style_name=item["style_name"],
            indent_twips=item["indent_twips"],
        )
        anchor.addnext(new_p)