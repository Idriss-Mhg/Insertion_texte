# =============================================================================
# src/app.py — Interface graphique principale (Tkinter)
#
# L'application est structurée en deux onglets :
#
#   Onglet "Insertion" :
#     Workflow principal. L'utilisateur sélectionne un dossier contenant des
#     fichiers .docx, navigue de fichier en fichier, recherche une section,
#     sélectionne le point d'insertion et insère la clause choisie.
#
#     À chaque clic "Insérer" :
#       1. Insertion de la clause (Track Changes puis texte brut)
#       2. Mise à jour automatique des dates de publication dans les deux
#          fichiers de sortie (corps "Date de publication" + footer "mise à
#          jour le"). Popup d'avertissement si aucune date trouvée.
#
#   Onglet "Clauses" :
#     Éditeur de clauses avec aperçu HTML en temps réel. Permet de créer,
#     modifier, renommer, dupliquer et supprimer les codes d'insertion
#     stockés dans clauses.json.
#
# Gestion des deux types de documents .docx rencontrés :
#   - Documents "classiques"  : paragraphes directement dans <w:body>
#   - Documents "structurés"  : contenu dans des tableaux (<w:td>) ou des
#     contrôles de contenu Word (<w:sdt>/<w:sdtContent>)
#
#   La variable self._flat_paras (produite par docx_handler.collect_paragraphs)
#   est la liste de référence des paragraphes. Elle est passée à toutes les
#   fonctions docx_handler et remplace doc.paragraphs (premier niveau seulement).
#
# Style "auto" pour le corps de clause :
#   Quand text_style == "auto", le style est hérité du paragraphe d'ancrage.
#   Exception : si l'ancre est un titre (Heading, APU_Heading…), on remonte
#   aux paragraphes voisins pour trouver le style de corps de texte réel
#   (évite d'insérer le corps de clause en police titre).
#
# Dépendances externes : tkinterweb (prévisualisation HTML), python-docx, lxml.
# =============================================================================

import html as _html
import json
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tkinterweb import HtmlFrame

from src import docx_handler, logger
from src.paths import CLAUSES_PATH

# -----------------------------------------------------------------------------
# Constantes
# -----------------------------------------------------------------------------

# Noms des sous-dossiers de sortie créés dans le dossier de travail
TC_FOLDER_NAME    = "_track_changes"   # insertions avec Track Changes (w:ins)
PLAIN_FOLDER_NAME = "_texte_brut"      # insertions en texte direct, sans révision

# CSS injecté dans la prévisualisation HTML du document.
# tkinterweb utilise un moteur HTML basique (tkhtml3) : on reste sur du CSS simple.
_PREVIEW_CSS = """
body  { font-family: Georgia, serif; font-size: 13px; margin: 16px;
        line-height: 1.6; color: #222; }
h1, h2, h3, h4 { font-family: Arial, sans-serif; color: #1a1a4e;
                  margin-top: 1.2em; }
p     { margin: 0.5em 0; text-align: justify; }
.highlight { background: #fff3cd; border-left: 3px solid #f0ad4e;
             padding-left: 6px; }
.ellipsis { color: #aaa; font-style: italic; text-align: center;
            border-top: 1px dashed #ccc; border-bottom: 1px dashed #ccc;
            padding: 3px 0; margin: 6px 0; }
"""


# -----------------------------------------------------------------------------
# Helpers I/O pour clauses.json
# -----------------------------------------------------------------------------

def _load_clauses() -> dict:
    """Lit et retourne le contenu de clauses.json sous forme de dict."""
    with open(CLAUSES_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _save_clauses(data: dict) -> None:
    """Écrase clauses.json avec le contenu du dict fourni (encodage UTF-8, indenté)."""
    with open(CLAUSES_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# -----------------------------------------------------------------------------
# Classe principale
# -----------------------------------------------------------------------------

class App(tk.Tk):
    """
    Fenêtre principale de l'application.

    Hérite de tk.Tk pour être elle-même la racine Tkinter.
    Toute la logique UI et métier est centralisée ici pour garder
    la surface publique minimale (un seul point d'entrée : main.py).
    """

    def __init__(self):
        super().__init__()
        self.title("Insertion texte — Prospectus")
        self.geometry("1100x760")
        self.resizable(True, True)

        # Clauses chargées depuis clauses.json au démarrage
        self._clauses: dict = _load_clauses()

        # Document .docx actuellement ouvert (instance python-docx Document)
        self._doc = None

        # Widget HtmlFrame de prévisualisation dans l'onglet Clauses.
        # Initialisé à None ici pour permettre les appels anticipés dans
        # _on_type_select avant que le widget soit réellement créé.
        self._clause_preview = None

        # Liste plate de tous les éléments <w:p> du document courant, en ordre
        # de lecture. Inclut les paragraphes dans les tableaux (w:td) et les
        # contrôles de contenu (w:sdt), contrairement à doc.paragraphs qui ne
        # retourne que les paragraphes de premier niveau.
        # Produite par docx_handler.collect_paragraphs() à chaque chargement.
        self._flat_paras: list = []

        # Dossier de travail sélectionné par l'utilisateur
        self._folder: Path | None = None

        # Dossiers de sortie créés dans _folder lors de la première insertion
        self._tc_dir: Path | None = None     # versions Track Changes
        self._plain_dir: Path | None = None  # versions texte brut

        # File d'attente : liste ordonnée des .docx à traiter dans _folder
        self._queue: list[Path] = []

        # Index du fichier courant dans _queue (-1 = aucun fichier chargé)
        self._queue_idx: int = -1

        # Mapping position_listbox → index dans self._flat_paras.
        # Nécessaire car la listbox n'affiche pas forcément tous les paragraphes
        # (filtrage par la recherche possible). Maintenu par _populate_listbox().
        self._para_indices: list[int] = []

        self._build_ui()
        self._refresh_log()

    # =========================================================================
    # Construction de l'interface
    # =========================================================================

    def _build_ui(self):
        """Crée le Notebook principal et ses deux onglets."""
        self._notebook = ttk.Notebook(self)
        self._notebook.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        tab_insert = ttk.Frame(self._notebook)
        self._notebook.add(tab_insert, text="  Insertion  ")
        self._build_insertion_tab(tab_insert)

        tab_clauses = ttk.Frame(self._notebook)
        self._notebook.add(tab_clauses, text="  Clauses  ")
        self._build_clauses_tab(tab_clauses)

    # -------------------------------------------------------------------------
    # Onglet 1 — Insertion
    # -------------------------------------------------------------------------

    def _build_insertion_tab(self, parent):
        """
        Construit l'onglet Insertion, organisé verticalement :
          1. Sélection du dossier de travail
          2. Navigation entre les fichiers (◀ / ▶)
          3. Sélection du code d'insertion et édition de la clause
          4. Barre de recherche de section
          5. Historique des insertions (log)        ← packés en BOTTOM en premier
          6. Boutons d'action (Insérer / Passer)    ← packés en BOTTOM en second
          7. PanedWindow listbox | prévisualisation  ← expand=True, remplit le reste

        Règle critique pack : les widgets side=BOTTOM doivent être packés AVANT
        tout widget avec expand=True. Sinon le PanedWindow consomme tout l'espace
        et les boutons/log disparaissent. On pack donc log et actions en premier,
        puis le PanedWindow en dernier.
        """
        pad = {"padx": 8, "pady": 4}

        # ── 1. Dossier de travail ─────────────────────────────────────────────
        f_folder = ttk.LabelFrame(parent, text="Dossier de travail")
        f_folder.pack(fill=tk.X, **pad)

        self._lbl_folder = ttk.Label(f_folder, text="Aucun dossier sélectionné", foreground="gray")
        self._lbl_folder.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6, pady=4)
        ttk.Button(f_folder, text="Choisir un dossier…", command=self._open_folder).pack(side=tk.RIGHT, padx=6, pady=4)

        # ── 2. Navigation fichier par fichier ─────────────────────────────────
        f_nav = ttk.LabelFrame(parent, text="Fichier en cours")
        f_nav.pack(fill=tk.X, **pad)

        nav_row = ttk.Frame(f_nav)
        nav_row.pack(fill=tk.X, padx=6, pady=4)

        self._btn_prev = ttk.Button(nav_row, text="◀ Précédent", command=self._prev_file, state=tk.DISABLED)
        self._btn_prev.pack(side=tk.LEFT)
        self._lbl_file = ttk.Label(nav_row, text="—", anchor=tk.CENTER, font=("", 10, "bold"))
        self._lbl_file.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        self._btn_next = ttk.Button(nav_row, text="Suivant ▶", command=self._next_file, state=tk.DISABLED)
        self._btn_next.pack(side=tk.RIGHT)

        # Compteur "3 / 47"
        self._lbl_progress = ttk.Label(f_nav, text="", foreground="gray")
        self._lbl_progress.pack(padx=6, pady=(0, 4))

        # ── 3. Clause à insérer ───────────────────────────────────────────────
        f_clause = ttk.LabelFrame(parent, text="Clause à insérer")
        f_clause.pack(fill=tk.X, **pad)

        row1 = ttk.Frame(f_clause)
        row1.pack(fill=tk.X, padx=6, pady=(4, 0))

        ttk.Label(row1, text="Code insertion :").pack(side=tk.LEFT)
        self._fund_var = tk.StringVar(value=list(self._clauses.keys())[0])
        self._fund_cb = ttk.Combobox(
            row1, textvariable=self._fund_var,
            values=list(self._clauses.keys()),
            state="readonly", width=14
        )
        self._fund_cb.pack(side=tk.LEFT, padx=6)
        # Pré-remplit sous-titre et texte à chaque changement de code
        self._fund_cb.bind("<<ComboboxSelected>>", self._on_code_change)

        ttk.Label(row1, text="Sous-titre :").pack(side=tk.LEFT, padx=(16, 0))
        self._subtitle_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self._subtitle_var, width=30).pack(side=tk.LEFT, padx=4)

        # ── Format du sous-titre ──────────────────────────────────────────────
        row_fmt = ttk.Frame(f_clause)
        row_fmt.pack(fill=tk.X, padx=6, pady=(2, 0))

        ttk.Label(row_fmt, text="Format s-titre :").pack(side=tk.LEFT)
        self._subtitle_type_var = tk.StringVar(value="bold")

        # Radio "Gras"
        ttk.Radiobutton(row_fmt, text="Gras", variable=self._subtitle_type_var,
                        value="bold", command=self._on_subtitle_type_change).pack(side=tk.LEFT, padx=(6, 2))

        # Radio "Souligné"
        ttk.Radiobutton(row_fmt, text="Souligné", variable=self._subtitle_type_var,
                        value="underline", command=self._on_subtitle_type_change).pack(side=tk.LEFT, padx=(8, 2))

        # Radio + combobox "Style Word"
        ttk.Radiobutton(row_fmt, text="Style Word", variable=self._subtitle_type_var,
                        value="style", command=self._on_subtitle_type_change).pack(side=tk.LEFT, padx=(8, 2))
        self._subtitle_style_var = tk.StringVar(value="Heading 3")
        self._cb_subtitle_style = ttk.Combobox(
            row_fmt, textvariable=self._subtitle_style_var, width=14,
            values=["Heading 1", "Heading 2", "Heading 3", "Heading 4",
                    "Titre 1", "Titre 2", "Titre 3", "Titre 4",
                    "Normal", "Corps de texte"])
        self._cb_subtitle_style.pack(side=tk.LEFT, padx=(0, 8))

        # Radio + options "À puce"
        ttk.Radiobutton(row_fmt, text="À puce", variable=self._subtitle_type_var,
                        value="puce", command=self._on_subtitle_type_change).pack(side=tk.LEFT, padx=(8, 2))
        self._subtitle_bullet_var = tk.StringVar(value="•")
        ttk.Combobox(row_fmt, textvariable=self._subtitle_bullet_var, width=3,
                     values=["•", "–", "-", "○", "▪", "◦", "→"]).pack(side=tk.LEFT)
        ttk.Label(row_fmt, text="Niv.").pack(side=tk.LEFT, padx=(6, 2))
        self._subtitle_indent_var = tk.IntVar(value=1)
        ttk.Spinbox(row_fmt, from_=1, to=3, textvariable=self._subtitle_indent_var,
                    width=3).pack(side=tk.LEFT)

        # ── Tailles de police ─────────────────────────────────────────────────
        row_sz = ttk.Frame(f_clause)
        row_sz.pack(fill=tk.X, padx=6, pady=(2, 0))
        ttk.Label(row_sz, text="Taille s-titre (pt) :").pack(side=tk.LEFT)
        self._subtitle_font_size_var = tk.IntVar(value=0)
        ttk.Spinbox(row_sz, from_=0, to=72, textvariable=self._subtitle_font_size_var,
                    width=4).pack(side=tk.LEFT, padx=(4, 2))
        ttk.Label(row_sz, text="(0 = auto)").pack(side=tk.LEFT, padx=(0, 16))
        ttk.Label(row_sz, text="Taille texte (pt) :").pack(side=tk.LEFT)
        self._text_font_size_var = tk.IntVar(value=0)
        ttk.Spinbox(row_sz, from_=0, to=72, textvariable=self._text_font_size_var,
                    width=4).pack(side=tk.LEFT, padx=(4, 2))
        ttk.Label(row_sz, text="(0 = auto)").pack(side=tk.LEFT)

        # ── Style du texte principal ──────────────────────────────────────────
        row_ts = ttk.Frame(f_clause)
        row_ts.pack(fill=tk.X, padx=6, pady=(2, 0))
        ttk.Label(row_ts, text="Style texte :").pack(side=tk.LEFT)
        self._text_style_var = tk.StringVar(value="auto")
        ttk.Combobox(row_ts, textvariable=self._text_style_var, width=20,
                     values=["auto", "Normal", "Corps de texte", "Body Text",
                             "Heading 4", "Titre 4"]).pack(side=tk.LEFT, padx=6)

        ttk.Label(f_clause, text="Texte de la clause :").pack(anchor=tk.W, padx=6, pady=(4, 0))
        self._txt_clause = tk.Text(f_clause, height=4, wrap=tk.WORD)
        self._txt_clause.pack(fill=tk.X, padx=6, pady=(0, 6))

        # Pré-remplissage initial avec le premier code d'insertion
        self._on_code_change()

        # ── 4. Recherche de section ───────────────────────────────────────────
        f_search = ttk.LabelFrame(parent, text="Recherche dans le document")
        f_search.pack(fill=tk.X, **pad)

        row2 = ttk.Frame(f_search)
        row2.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(row2, text="Mot-clé de section :").pack(side=tk.LEFT)
        self._search_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self._search_var, width=30).pack(side=tk.LEFT, padx=6)
        ttk.Button(row2, text="Rechercher", command=self._search).pack(side=tk.LEFT)
        # "Afficher tout" remet la listbox en mode complet sans filtre
        ttk.Button(row2, text="Afficher tout", command=self._show_all).pack(side=tk.LEFT, padx=4)

        # ── 5. Historique des insertions ──────────────────────────────────────
        # Packé en BOTTOM avant le PanedWindow (expand=True) pour qu'il ne soit
        # pas écrasé : avec pack, les widgets BOTTOM sont réservés en premier.
        f_log = ttk.LabelFrame(parent, text="Historique des insertions")
        f_log.pack(side=tk.BOTTOM, fill=tk.X, **pad)

        # Widget en lecture seule : on l'active brièvement pour écrire dedans
        self._txt_log = tk.Text(f_log, height=4, state=tk.DISABLED, font=("Courier", 8), foreground="#444")
        self._txt_log.pack(fill=tk.X, padx=4, pady=4)

        # ── 6. Boutons d'action ───────────────────────────────────────────────
        # Également packé en BOTTOM, avant le PanedWindow.
        f_actions = ttk.Frame(parent)
        f_actions.pack(side=tk.BOTTOM, fill=tk.X, **pad)

        ttk.Button(f_actions, text="Insérer la clause", command=self._insert).pack(side=tk.LEFT)
        # "Passer" avance au fichier suivant sans modifier le document courant
        ttk.Button(f_actions, text="Passer (sans insérer)", command=self._next_file).pack(side=tk.LEFT, padx=8)

        # Auteur affiché dans la bulle de révision Word (mode Track Changes)
        ttk.Label(f_actions, text="Auteur :").pack(side=tk.LEFT, padx=(16, 4))
        self._author_var = tk.StringVar(value="Juriste")
        ttk.Entry(f_actions, textvariable=self._author_var, width=16).pack(side=tk.LEFT)

        self._lbl_status = ttk.Label(f_actions, text="", foreground="gray")
        self._lbl_status.pack(side=tk.LEFT, padx=12)

        # ── 7. PanedWindow : listbox + prévisualisation ───────────────────────
        # Packé en dernier avec expand=True : il occupe tout l'espace restant
        # entre les sections fixes du haut et les boutons/log du bas.
        paned = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, **pad)

        # Pane gauche — liste des paragraphes du document
        f_para = ttk.LabelFrame(paned, text="Paragraphes  —  insérer APRÈS le paragraphe sélectionné")
        paned.add(f_para, weight=1)

        sb_para = ttk.Scrollbar(f_para, orient=tk.VERTICAL)
        self._listbox = tk.Listbox(
            f_para, yscrollcommand=sb_para.set,
            selectmode=tk.SINGLE, activestyle="dotbox",
            font=("Courier", 9)
        )
        sb_para.config(command=self._listbox.yview)
        sb_para.pack(side=tk.RIGHT, fill=tk.Y)
        self._listbox.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        # <<ListboxSelect>> = changement de sélection ; <ButtonRelease-1> = re-clic
        # sur le même item (repasse en vue fenêtrée après "Vue complète")
        self._listbox.bind("<<ListboxSelect>>",  self._on_para_select)
        self._listbox.bind("<ButtonRelease-1>",  self._on_para_select)

        # Pane droite — prévisualisation HTML du document
        # tkinterweb/HtmlFrame utilise le moteur tkhtml3 (pas de JavaScript).
        f_preview = ttk.LabelFrame(paned, text="Aperçu du document")
        paned.add(f_preview, weight=2)

        # Barre d'outils du panneau preview (bouton vue complète)
        f_prev_bar = ttk.Frame(f_preview)
        f_prev_bar.pack(fill=tk.X, padx=2, pady=(2, 0))
        ttk.Button(
            f_prev_bar, text="Vue complète", width=14,
            command=lambda: self._refresh_preview()
        ).pack(side=tk.LEFT, padx=2)

        self._html_frame = HtmlFrame(f_preview, messages_enabled=False)
        self._html_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self._html_frame.load_html("<body style='color:gray;padding:16px'>Aucun document chargé.</body>")

    # -------------------------------------------------------------------------
    # Onglet 2 — Éditeur de clauses
    # -------------------------------------------------------------------------

    def _build_clauses_tab(self, parent):
        """
        Construit l'onglet Clauses, divisé en deux panneaux :
          - Gauche : liste des codes d'insertion + boutons (Nouveau / Dupliquer / Supprimer)
          - Droite : formulaire d'édition (code, sous-titre, texte) + Enregistrer / Annuler

        Toute modification est immédiatement persistée dans clauses.json et
        synchronisée avec le combobox de l'onglet Insertion.
        """
        parent.columnconfigure(0, weight=1, minsize=180)
        parent.columnconfigure(1, weight=3)
        parent.rowconfigure(0, weight=1)

        # ── Panneau gauche : liste des codes ──────────────────────────────────
        f_left = ttk.LabelFrame(parent, text="Codes d'insertion")
        f_left.grid(row=0, column=0, sticky="nsew", padx=(8, 4), pady=8)
        f_left.rowconfigure(0, weight=1)
        f_left.columnconfigure(0, weight=1)

        sb_l = ttk.Scrollbar(f_left, orient=tk.VERTICAL)
        self._type_listbox = tk.Listbox(
            f_left, yscrollcommand=sb_l.set,
            selectmode=tk.SINGLE, activestyle="dotbox",
            font=("", 10)
        )
        sb_l.config(command=self._type_listbox.yview)
        sb_l.grid(row=0, column=1, sticky="ns", pady=4)
        self._type_listbox.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=4)
        self._type_listbox.bind("<<ListboxSelect>>", self._on_type_select)

        f_list_btns = ttk.Frame(f_left)
        f_list_btns.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4, pady=(0, 4))
        ttk.Button(f_list_btns, text="+ Nouveau",  command=self._clause_new).pack(side=tk.LEFT, expand=True, fill=tk.X)
        ttk.Button(f_list_btns, text="Dupliquer",  command=self._clause_duplicate).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(f_list_btns, text="Supprimer",  command=self._clause_delete).pack(side=tk.LEFT, expand=True, fill=tk.X)

        # ── Panneau droit : formulaire d'édition ──────────────────────────────
        f_right = ttk.LabelFrame(parent, text="Édition")
        f_right.grid(row=0, column=1, sticky="nsew", padx=(4, 8), pady=8)
        f_right.columnconfigure(1, weight=1)
        f_right.rowconfigure(6, weight=2)  # Éditeur de texte (plus de place)
        f_right.rowconfigure(9, weight=1)  # Aperçu de la clause

        ttk.Label(f_right, text="Code :").grid(row=0, column=0, sticky="w", padx=8, pady=(8, 4))
        self._edit_name_var = tk.StringVar()
        self._edit_name_entry = ttk.Entry(f_right, textvariable=self._edit_name_var, width=24)
        self._edit_name_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=(8, 4))

        ttk.Label(f_right, text="Sous-titre :").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        self._edit_subtitle_var = tk.StringVar()
        ttk.Entry(f_right, textvariable=self._edit_subtitle_var).grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=4)

        # ── Format du sous-titre ──────────────────────────────────────────────
        f_fmt = ttk.Frame(f_right)
        f_fmt.grid(row=2, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 4))

        ttk.Label(f_fmt, text="Format :").pack(side=tk.LEFT)
        self._edit_subtitle_type_var = tk.StringVar(value="bold")

        ttk.Radiobutton(f_fmt, text="Gras", variable=self._edit_subtitle_type_var,
                        value="bold", command=self._on_edit_subtitle_type_change).pack(side=tk.LEFT, padx=(6, 2))

        # Radio "Souligné"
        ttk.Radiobutton(f_fmt, text="Souligné", variable=self._edit_subtitle_type_var,
                        value="underline", command=self._on_edit_subtitle_type_change).pack(side=tk.LEFT, padx=(8, 2))

        ttk.Radiobutton(f_fmt, text="Style Word", variable=self._edit_subtitle_type_var,
                        value="style", command=self._on_edit_subtitle_type_change).pack(side=tk.LEFT, padx=(8, 2))
        self._edit_subtitle_style_var = tk.StringVar(value="Heading 3")
        self._edit_cb_style = ttk.Combobox(
            f_fmt, textvariable=self._edit_subtitle_style_var, width=14,
            values=["Heading 1", "Heading 2", "Heading 3", "Heading 4",
                    "Titre 1", "Titre 2", "Titre 3", "Titre 4",
                    "Normal", "Corps de texte"])
        self._edit_cb_style.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Radiobutton(f_fmt, text="À puce", variable=self._edit_subtitle_type_var,
                        value="puce", command=self._on_edit_subtitle_type_change).pack(side=tk.LEFT, padx=(8, 2))
        self._edit_subtitle_bullet_var = tk.StringVar(value="•")
        ttk.Combobox(f_fmt, textvariable=self._edit_subtitle_bullet_var, width=3,
                     values=["•", "–", "-", "○", "▪", "◦", "→"]).pack(side=tk.LEFT)
        ttk.Label(f_fmt, text="Niv.").pack(side=tk.LEFT, padx=(6, 2))
        self._edit_subtitle_indent_var = tk.IntVar(value=1)
        ttk.Spinbox(f_fmt, from_=1, to=3, textvariable=self._edit_subtitle_indent_var,
                    width=3).pack(side=tk.LEFT)

        # ── Style du texte principal ──────────────────────────────────────────
        f_ts = ttk.Frame(f_right)
        f_ts.grid(row=3, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 4))
        ttk.Label(f_ts, text="Style texte :").pack(side=tk.LEFT)
        self._edit_text_style_var = tk.StringVar(value="auto")
        ttk.Combobox(f_ts, textvariable=self._edit_text_style_var, width=20,
                     values=["auto", "Normal", "Corps de texte", "Body Text",
                             "Heading 4", "Titre 4"]).pack(side=tk.LEFT, padx=6)

        # ── Tailles de police ─────────────────────────────────────────────────
        f_sz = ttk.Frame(f_right)
        f_sz.grid(row=4, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 4))
        ttk.Label(f_sz, text="Taille s-titre (pt) :").pack(side=tk.LEFT)
        self._edit_subtitle_font_size_var = tk.IntVar(value=0)
        ttk.Spinbox(f_sz, from_=0, to=72, textvariable=self._edit_subtitle_font_size_var,
                    width=4).pack(side=tk.LEFT, padx=(4, 2))
        ttk.Label(f_sz, text="(0 = auto)").pack(side=tk.LEFT, padx=(0, 16))
        ttk.Label(f_sz, text="Taille texte (pt) :").pack(side=tk.LEFT)
        self._edit_text_font_size_var = tk.IntVar(value=0)
        ttk.Spinbox(f_sz, from_=0, to=72, textvariable=self._edit_text_font_size_var,
                    width=4).pack(side=tk.LEFT, padx=(4, 2))
        ttk.Label(f_sz, text="(0 = auto)").pack(side=tk.LEFT)

        ttk.Label(f_right, text="Texte de la clause :").grid(row=5, column=0, sticky="nw", padx=8, pady=4)

        self._edit_txt = tk.Text(f_right, wrap=tk.WORD, height=10)
        sb_r = ttk.Scrollbar(f_right, orient=tk.VERTICAL, command=self._edit_txt.yview)
        self._edit_txt.config(yscrollcommand=sb_r.set)
        self._edit_txt.grid(row=6, column=0, columnspan=2, sticky="nsew", padx=8, pady=(0, 4))
        sb_r.grid(row=6, column=2, sticky="ns", pady=(0, 4))

        f_form_btns = ttk.Frame(f_right)
        f_form_btns.grid(row=7, column=0, columnspan=3, sticky="ew", padx=8, pady=(0, 8))
        ttk.Button(f_form_btns, text="Enregistrer", command=self._clause_save).pack(side=tk.LEFT)
        ttk.Button(f_form_btns, text="Annuler",     command=self._clause_cancel).pack(side=tk.LEFT, padx=8)
        self._lbl_clause_status = ttk.Label(f_form_btns, text="", foreground="gray")
        self._lbl_clause_status.pack(side=tk.LEFT)

        # ── Aperçu de la clause ───────────────────────────────────────────────
        # Rendu HTML en temps réel : sous-titre dans son format + extrait du corps.
        # Ne nécessite pas de document chargé — utilise une approximation CSS.
        ttk.Label(f_right, text="Aperçu :").grid(
            row=8, column=0, sticky="w", padx=8, pady=(4, 0))
        self._clause_preview = HtmlFrame(f_right, messages_enabled=False)
        self._clause_preview.grid(
            row=9, column=0, columnspan=3, sticky="nsew", padx=8, pady=(0, 8))

        # Nom du code en cours d'édition (None si création d'un nouveau code)
        self._clause_editing_original_name: str | None = None

        self._refresh_type_listbox()

        # Traces pour mise à jour en temps réel de l'aperçu.
        # Les traces sont installées après _refresh_type_listbox() pour éviter
        # tout appel sur un widget non encore créé.
        for var in (
            self._edit_subtitle_var, self._edit_subtitle_type_var,
            self._edit_subtitle_style_var, self._edit_subtitle_bullet_var,
            self._edit_text_style_var,
        ):
            var.trace_add("write", lambda *_: self._update_clause_preview())
        for var in (
            self._edit_subtitle_font_size_var, self._edit_text_font_size_var,
            self._edit_subtitle_indent_var,
        ):
            var.trace_add("write", lambda *_: self._update_clause_preview())
        self._edit_txt.bind("<KeyRelease>", lambda _: self._update_clause_preview())
        self._update_clause_preview()

    # =========================================================================
    # Logique de l'éditeur de clauses (onglet 2)
    # =========================================================================

    def _refresh_type_listbox(self, select_name: str | None = None):
        """
        Reconstruit la listbox des codes d'insertion depuis self._clauses.

        Args:
            select_name: Si fourni, sélectionne automatiquement ce code après
                         le rechargement (utile après un enregistrement).
        """
        self._type_listbox.delete(0, tk.END)
        for name in self._clauses:
            self._type_listbox.insert(tk.END, name)
        if select_name and select_name in self._clauses:
            idx = list(self._clauses.keys()).index(select_name)
            self._type_listbox.selection_set(idx)
            self._type_listbox.see(idx)
            self._on_type_select()

    def _on_type_select(self, *_):
        """Remplit le formulaire d'édition avec les données du code sélectionné."""
        sel = self._type_listbox.curselection()
        if not sel:
            return
        name = self._type_listbox.get(sel[0])
        data = self._clauses.get(name, {})
        self._clause_editing_original_name = name
        self._edit_name_var.set(name)
        self._edit_subtitle_var.set(data.get("subtitle", ""))
        self._edit_subtitle_type_var.set(data.get("subtitle_type", "bold"))
        self._edit_subtitle_style_var.set(data.get("subtitle_style", "Heading 3"))
        self._edit_subtitle_bullet_var.set(data.get("subtitle_bullet", "•"))
        self._edit_subtitle_indent_var.set(data.get("subtitle_indent", 1))
        self._edit_subtitle_font_size_var.set(data.get("subtitle_font_size", 0))
        self._edit_text_style_var.set(data.get("text_style", "auto"))
        self._edit_text_font_size_var.set(data.get("text_font_size", 0))
        self._edit_txt.delete("1.0", tk.END)
        self._edit_txt.insert("1.0", data.get("text", ""))
        self._on_edit_subtitle_type_change()
        self._lbl_clause_status.config(text="")
        self._update_clause_preview()

    def _on_edit_subtitle_type_change(self, *_):
        """Active/désactive le combobox style Word dans l'éditeur de clauses."""
        t = self._edit_subtitle_type_var.get()
        self._edit_cb_style.config(state="normal" if t == "style" else "disabled")
        self._update_clause_preview()

    def _update_clause_preview(self):
        """Régénère l'aperçu HTML de la clause dans l'onglet Clauses."""
        if self._clause_preview is None:
            return
        self._clause_preview.load_html(self._build_clause_preview_html())

    def _build_clause_preview_html(self) -> str:
        """
        Génère un fragment HTML illustrant la clause en cours d'édition :
          - Sous-titre rendu selon son type (gras / souligné / heading / puce)
          - Corps de clause tronqué à 300 caractères
          - Note de style en gris italic (nom du style ou "style hérité" si auto)

        Utilise des approximations CSS car aucun document Word n'est chargé
        dans l'onglet Clauses. Pour les styles Word (Heading 3, etc.), le
        niveau de heading est déduit du chiffre présent dans le nom de style.
        """
        subtitle   = self._edit_subtitle_var.get().strip()
        sub_type   = self._edit_subtitle_type_var.get()
        sub_style  = self._edit_subtitle_style_var.get()
        sub_bullet = self._edit_subtitle_bullet_var.get()
        sub_size   = self._edit_subtitle_font_size_var.get()
        text       = self._edit_txt.get("1.0", tk.END).strip()
        text_size  = self._edit_text_font_size_var.get()
        text_style = self._edit_text_style_var.get()

        _CSS = (
            "body { font-family: Georgia, serif; font-size: 12px; margin: 8px;"
            "       color: #222; line-height: 1.5; }"
            "h1   { font-size: 1.4em; font-weight: bold; font-family: Arial, sans-serif;"
            "       color: #1a1a4e; margin: 2px 0; }"
            "h2   { font-size: 1.2em; font-weight: bold; font-family: Arial, sans-serif;"
            "       color: #1a1a4e; margin: 2px 0; }"
            "h3   { font-size: 1.05em; font-weight: bold; font-family: Arial, sans-serif;"
            "       color: #1a1a4e; margin: 2px 0; }"
            "h4   { font-size: 1.0em; font-weight: bold; font-family: Arial, sans-serif;"
            "       color: #1a1a4e; margin: 2px 0; }"
            "p    { margin: 4px 0; }"
            ".note { color: #999; font-style: italic; font-size: 0.82em; }"
            ".empty { color: #aaa; font-style: italic; }"
        )

        parts = []

        # ── Sous-titre ────────────────────────────────────────────────────────
        if subtitle:
            sz = f"font-size:{sub_size}pt;" if sub_size > 0 else ""
            esc = _html.escape(subtitle)

            if sub_type == "bold":
                parts.append(f'<p style="font-weight:bold;{sz}">{esc}</p>')

            elif sub_type == "underline":
                parts.append(f'<p style="text-decoration:underline;{sz}">{esc}</p>')

            elif sub_type == "style":
                # Déduit le niveau h1–h4 depuis le chiffre dans le nom de style
                lvl = 3
                for n in range(1, 5):
                    if str(n) in sub_style:
                        lvl = n
                        break
                tag = f"h{lvl}"
                note = f'<span class="note"> ({_html.escape(sub_style)})</span>'
                parts.append(f'<{tag} style="{sz}">{esc}{note}</{tag}>')

            elif sub_type == "puce":
                indent_px = self._edit_subtitle_indent_var.get() * 20
                parts.append(
                    f'<p style="margin-left:{indent_px}px;{sz}">'
                    f'{_html.escape(sub_bullet)} {esc}</p>'
                )

        # ── Corps de clause ───────────────────────────────────────────────────
        if text:
            sz = f"font-size:{text_size}pt;" if text_size > 0 else ""
            excerpt = _html.escape(text[:300]) + ("…" if len(text) > 300 else "")
            if text_style == "auto":
                note = '<span class="note"> [style hérité du paragraphe d\'ancrage]</span>'
            else:
                note = f'<span class="note"> [{_html.escape(text_style)}]</span>'
            parts.append(f'<p style="{sz}">{excerpt}{note}</p>')

        if not parts:
            body = '<p class="empty">Aperçu vide — saisissez un sous-titre ou un texte.</p>'
        else:
            body = "\n".join(parts)

        return (
            f'<!DOCTYPE html><html><head><meta charset="utf-8">'
            f'<style>{_CSS}</style></head><body>{body}</body></html>'
        )

    def _clause_new(self):
        """Réinitialise le formulaire pour la saisie d'un nouveau code d'insertion."""
        self._clause_editing_original_name = None
        self._type_listbox.selection_clear(0, tk.END)
        self._edit_name_var.set("")
        self._edit_subtitle_var.set("")
        self._edit_txt.delete("1.0", tk.END)
        self._edit_name_entry.focus_set()
        self._lbl_clause_status.config(text="Nouveau code — remplissez puis Enregistrer.")

    def _clause_duplicate(self):
        """Copie le code sélectionné dans le formulaire avec le suffixe '_copie'."""
        sel = self._type_listbox.curselection()
        if not sel:
            messagebox.showwarning("Attention", "Sélectionnez un code à dupliquer.")
            return
        name = self._type_listbox.get(sel[0])
        data = self._clauses[name]
        self._clause_editing_original_name = None
        self._edit_name_var.set(name + "_copie")
        self._edit_subtitle_var.set(data.get("subtitle", ""))
        self._edit_txt.delete("1.0", tk.END)
        self._edit_txt.insert("1.0", data.get("text", ""))
        self._lbl_clause_status.config(text="Duplicata — modifiez le nom puis Enregistrer.")

    def _clause_save(self):
        """
        Enregistre le code en cours d'édition dans self._clauses et clauses.json.

        Gère trois cas :
          - Création d'un nouveau code (self._clause_editing_original_name is None)
          - Modification sans renommage (le nom n'a pas changé)
          - Renommage (le nom a changé) : reconstruit le dict pour préserver l'ordre
            des clés (important pour l'affichage dans le combobox)
        """
        new_name = self._edit_name_var.get().strip()
        if not new_name:
            messagebox.showwarning("Attention", "Le code ne peut pas être vide.")
            return

        subtitle = self._edit_subtitle_var.get().strip()
        text = self._edit_txt.get("1.0", tk.END).strip()
        entry = {
            "subtitle": subtitle,
            "subtitle_type": self._edit_subtitle_type_var.get(),
            "subtitle_style": self._edit_subtitle_style_var.get().strip(),
            "subtitle_bullet": self._edit_subtitle_bullet_var.get(),
            "subtitle_indent": self._edit_subtitle_indent_var.get(),
            "subtitle_font_size": self._edit_subtitle_font_size_var.get(),
            "text": text,
            "text_style": self._edit_text_style_var.get().strip() or "auto",
            "text_font_size": self._edit_text_font_size_var.get(),
        }
        original = self._clause_editing_original_name

        if original and original != new_name:
            # Renommage : on reconstruit le dict en remplaçant la clé originale
            # pour conserver la position du code dans la liste
            if new_name in self._clauses:
                if not messagebox.askyesno("Confirmer", f"« {new_name} » existe déjà. Écraser ?"):
                    return
            self._clauses = {
                (new_name if k == original else k): (entry if k == original else v)
                for k, v in self._clauses.items()
            }
        else:
            # Création ou modification sans renommage
            if new_name in self._clauses and original is None:
                if not messagebox.askyesno("Confirmer", f"« {new_name} » existe déjà. Écraser ?"):
                    return
            self._clauses[new_name] = entry

        _save_clauses(self._clauses)
        self._clause_editing_original_name = new_name
        self._refresh_type_listbox(select_name=new_name)
        self._sync_code_combobox()
        self._lbl_clause_status.config(text=f"« {new_name} » enregistré.", foreground="green")

    def _clause_cancel(self):
        """Annule l'édition en cours et restaure les valeurs sauvegardées."""
        sel = self._type_listbox.curselection()
        if sel:
            # Si un code était sélectionné, on recharge ses valeurs depuis le dict
            self._on_type_select()
        else:
            self._edit_name_var.set("")
            self._edit_subtitle_var.set("")
            self._edit_txt.delete("1.0", tk.END)
        self._lbl_clause_status.config(text="")

    def _clause_delete(self):
        """Supprime le code sélectionné après confirmation, puis persiste la modification."""
        sel = self._type_listbox.curselection()
        if not sel:
            messagebox.showwarning("Attention", "Sélectionnez un code à supprimer.")
            return
        name = self._type_listbox.get(sel[0])
        if not messagebox.askyesno("Confirmer la suppression", f"Supprimer définitivement « {name} » ?"):
            return
        del self._clauses[name]
        _save_clauses(self._clauses)
        self._clause_editing_original_name = None
        self._edit_name_var.set("")
        self._edit_subtitle_var.set("")
        self._edit_txt.delete("1.0", tk.END)
        self._refresh_type_listbox()
        self._sync_code_combobox()
        self._lbl_clause_status.config(text=f"« {name} » supprimé.", foreground="gray")

    def _sync_code_combobox(self):
        """
        Met à jour le combobox de l'onglet Insertion pour refléter les codes
        actuels de clauses.json. Si le code sélectionné n'existe plus, on
        bascule sur le premier code disponible.
        """
        keys = list(self._clauses.keys())
        self._fund_cb.config(values=keys)
        if self._fund_var.get() not in keys:
            self._fund_var.set(keys[0] if keys else "")
            self._on_code_change()

    # =========================================================================
    # Logique de l'onglet Insertion — Dossier et navigation
    # =========================================================================

    def _open_folder(self):
        """
        Ouvre un sélecteur de dossier, scanne les .docx présents et
        charge le premier fichier automatiquement.
        """
        folder = filedialog.askdirectory(title="Sélectionner le dossier contenant les prospectus")
        if not folder:
            return
        self._folder = Path(folder)
        self._tc_dir    = self._folder / TC_FOLDER_NAME
        self._plain_dir = self._folder / PLAIN_FOLDER_NAME
        self._queue = sorted(self._folder.glob("*.docx"))

        if not self._queue:
            messagebox.showinfo("Dossier vide", "Aucun fichier .docx trouvé dans ce dossier.")
            return

        self._lbl_folder.config(
            text=f"{self._folder}  ({len(self._queue)} fichier(s))",
            foreground="black"
        )
        self._queue_idx = -1
        self._next_file()

    def _load_current_file(self):
        """
        Charge le fichier courant (self._queue[self._queue_idx]) en lecture
        seule, sans modifier l'original :
          1. Ouvre le document avec python-docx (original intact)
          2. Construit self._flat_paras via collect_paragraphs() — traversée
             XML complète qui capture les paragraphes dans les tableaux (w:td)
             et les contrôles de contenu (w:sdt), en plus des paragraphes
             de premier niveau. C'est la liste de référence utilisée partout.
          3. Met à jour les widgets de navigation et de statut
          4. Vide la listbox et rafraîchit la prévisualisation
        """
        path = self._queue[self._queue_idx]
        try:
            self._doc = docx_handler.open_document(str(path))
        except Exception as e:
            messagebox.showerror("Erreur", str(e))
            return

        total = len(self._queue)
        pos = self._queue_idx + 1
        self._lbl_file.config(text=path.name)
        self._lbl_progress.config(text=f"{pos} / {total}")
        self._btn_prev.config(state=tk.NORMAL if self._queue_idx > 0 else tk.DISABLED)
        self._btn_next.config(state=tk.NORMAL if self._queue_idx < total - 1 else tk.DISABLED)
        # Construire la liste plate des paragraphes (tableaux + SDT inclus)
        self._flat_paras = docx_handler.collect_paragraphs(self._doc)
        self._listbox.delete(0, tk.END)
        self._para_indices = []
        self._status(f"Fichier chargé — {len(self._flat_paras)} paragraphes.")
        self._refresh_preview()

    def _next_file(self):
        """Avance au fichier suivant dans la file, ou signale la fin de la campagne."""
        if not self._queue:
            return
        if self._queue_idx < len(self._queue) - 1:
            self._queue_idx += 1
            self._load_current_file()
        else:
            messagebox.showinfo("Terminé", "Tous les fichiers ont été traités.")

    def _prev_file(self):
        """Recule au fichier précédent dans la file."""
        if self._queue_idx > 0:
            self._queue_idx -= 1
            self._load_current_file()

    # =========================================================================
    # Logique de l'onglet Insertion — Clause
    # =========================================================================

    def _on_code_change(self, *_):
        """
        Pré-remplit tous les champs (sous-titre, format, texte, styles) à partir
        du code d'insertion sélectionné dans le combobox.
        """
        code = self._fund_var.get()
        data = self._clauses.get(code, {})
        self._subtitle_var.set(data.get("subtitle", ""))
        self._subtitle_type_var.set(data.get("subtitle_type", "bold"))
        self._subtitle_style_var.set(data.get("subtitle_style", "Heading 3"))
        self._subtitle_bullet_var.set(data.get("subtitle_bullet", "•"))
        self._subtitle_indent_var.set(data.get("subtitle_indent", 1))
        self._subtitle_font_size_var.set(data.get("subtitle_font_size", 0))
        self._text_style_var.set(data.get("text_style", "auto"))
        self._text_font_size_var.set(data.get("text_font_size", 0))
        self._txt_clause.delete("1.0", tk.END)
        self._txt_clause.insert("1.0", data.get("text", ""))
        self._on_subtitle_type_change()

    def _on_subtitle_type_change(self, *_):
        """Active/désactive le combobox style Word selon le type de sous-titre choisi."""
        t = self._subtitle_type_var.get()
        self._cb_subtitle_style.config(state="normal" if t == "style" else "disabled")

    # =========================================================================
    # Logique de l'onglet Insertion — Recherche et listbox
    # =========================================================================

    def _search(self):
        """
        Recherche le mot-clé saisi dans tous les paragraphes du document courant.
        Pour chaque occurrence trouvée, affiche les paragraphes environnants
        (contexte de 4 paragraphes de chaque côté) mis en évidence en jaune.
        """
        if not self._doc:
            messagebox.showwarning("Attention", "Ouvrez d'abord un dossier.")
            return
        keyword = self._search_var.get().strip()
        if not keyword:
            self._show_all()
            return

        results = docx_handler.search_paragraphs(self._doc, keyword, flat_paras=self._flat_paras)
        if not results:
            messagebox.showinfo("Aucun résultat", f"Mot-clé « {keyword} » introuvable.")
            return

        # Collecte les paragraphes de contexte autour de chaque occurrence,
        # en dédupliquant si plusieurs occurrences sont proches.
        displayed: list[tuple[int, str]] = []
        seen: set[int] = set()
        for (center_idx, _) in results:
            for (i, text) in docx_handler.get_paragraphs_around(self._doc, center_idx, flat_paras=self._flat_paras):
                if i not in seen:
                    seen.add(i)
                    displayed.append((i, text))
        displayed.sort(key=lambda x: x[0])

        self._populate_listbox(displayed, highlight_indices={r[0] for r in results})
        self._status(f"{len(results)} occurrence(s) trouvée(s) pour « {keyword} ».")

    def _show_all(self):
        """Affiche tous les paragraphes non vides du document dans la listbox."""
        if not self._doc:
            messagebox.showwarning("Attention", "Ouvrez d'abord un dossier.")
            return
        self._populate_listbox(docx_handler.get_all_paragraphs(self._doc, flat_paras=self._flat_paras))
        self._status(f"{len(self._para_indices)} paragraphes affichés.")

    def _on_para_select(self, *_):
        """
        Déclenché à chaque clic dans la listbox (y compris re-clic sur le même item).
        Régénère la prévisualisation en vue fenêtrée centrée sur le paragraphe sélectionné.
        La cible est toujours en haut du fragment HTML — aucun scroll nécessaire.
        """
        sel = self._listbox.curselection()
        if not sel:
            return
        para_idx = self._para_indices[sel[0]]
        self._refresh_preview(highlight_idx=para_idx)

    def _populate_listbox(self, paras: list[tuple[int, str]], highlight_indices: set | None = None):
        """
        Remplit la listbox avec la liste de paragraphes fournie.

        Les paragraphes dont l'index est dans highlight_indices sont affichés
        avec un fond jaune (occurrences de recherche).

        Args:
            paras: Liste de tuples (index_dans_doc, texte).
            highlight_indices: Ensemble d'index à mettre en évidence.
        """
        self._listbox.delete(0, tk.END)
        self._para_indices = []
        highlight_indices = highlight_indices or set()

        for (idx, text) in paras:
            preview = text[:120].replace("\n", " ")
            self._listbox.insert(tk.END, f"§{idx:>4}  {preview}")
            self._para_indices.append(idx)
            if idx in highlight_indices:
                self._listbox.itemconfig(tk.END, background="#fffacd", foreground="#8B4513")

    # =========================================================================
    # Logique de l'onglet Insertion — Prévisualisation
    # =========================================================================

    def _refresh_preview(self, highlight_idx: int | None = None):
        """
        Génère le HTML du document courant et l'injecte dans le widget HtmlFrame.

        Deux modes :
          - highlight_idx fourni : affiche une fenêtre de paragraphes centrée
            sur la cible (build_html_window). La cible est toujours en haut de
            la zone visible — aucun scroll nécessaire.
          - highlight_idx absent : affiche le document complet (build_html),
            utilisé au chargement initial du fichier.

        Note : tkinterweb/tkhtml3 ne supporte pas JavaScript, ce qui rend le
        scroll programmatique peu fiable (yview_moveto est basé sur une
        fraction du document total, impossible à calculer avec précision sans
        connaître les hauteurs rendues). La vue fenêtrée élimine ce problème.

        Args:
            highlight_idx: Index du paragraphe à mettre en évidence, ou None.
        """
        if not self._doc:
            return
        try:
            if highlight_idx is not None:
                # Vue fenêtrée : 4 paragraphes de contexte avant, 25 après.
                # La cible est en haut du fragment → toujours visible, sans scroll.
                html_body = docx_handler.build_html_window(
                    self._doc, highlight_idx, flat_paras=self._flat_paras
                )
            else:
                html_body = docx_handler.build_html(self._doc, flat_paras=self._flat_paras)
            html = (
                f"<!DOCTYPE html><html><head>"
                f'<meta charset="utf-8">'
                f"<style>{_PREVIEW_CSS}</style>"
                f"</head><body>{html_body}</body></html>"
            )
            self._html_frame.load_html(html)
        except Exception:
            self._html_frame.load_html(
                "<body style='color:gray;padding:16px'>Aperçu indisponible pour ce fichier.</body>"
            )

    # =========================================================================
    # Logique de l'onglet Insertion — Insertion et log
    # =========================================================================

    def _insert(self):
        """
        Insère la clause dans deux fichiers de sortie distincts, sans toucher
        l'original.

        Étapes :
          1. Validation : document ouvert, paragraphe sélectionné, clause non vide
          2. Pour la version Track Changes :
               - Ouvre une copie fraîche de l'original
               - Insère avec <w:ins> via insert_clause_after()
               - Écrit dans _track_changes/<fichier>.docx
          3. Pour la version texte brut :
               - Ouvre une nouvelle copie fraîche de l'original
               - Insère en paragraphes directs via insert_clause_plain_after()
               - Écrit dans _texte_brut/<fichier>.docx
          4. Journalisation dans logs/insertions.csv
          5. Rafraîchissement du log et passage au fichier suivant

        L'original n'est jamais modifié.
        """
        if not self._doc:
            messagebox.showwarning("Attention", "Ouvrez d'abord un dossier.")
            return
        selection = self._listbox.curselection()
        if not selection:
            messagebox.showwarning("Attention", "Sélectionnez le paragraphe après lequel insérer la clause.")
            return

        para_idx = self._para_indices[selection[0]]
        subtitle = self._subtitle_var.get().strip()
        clause_text = self._txt_clause.get("1.0", tk.END).strip()
        author = self._author_var.get().strip() or "Juriste"

        if not clause_text:
            messagebox.showwarning("Attention", "Le texte de la clause est vide.")
            return

        subtitle_config = {
            "type":   self._subtitle_type_var.get(),
            "style":  self._subtitle_style_var.get(),
            "bullet": self._subtitle_bullet_var.get(),
            "indent": self._subtitle_indent_var.get(),
        }
        text_style = self._text_style_var.get().strip() or None
        subtitle_font_size = self._subtitle_font_size_var.get()
        text_font_size = self._text_font_size_var.get()

        filepath = str(self._queue[self._queue_idx])
        filename = self._queue[self._queue_idx].name
        self._tc_dir.mkdir(parents=True, exist_ok=True)
        self._plain_dir.mkdir(parents=True, exist_ok=True)

        try:
            # ── Version Track Changes ─────────────────────────────────────────
            doc_tc = docx_handler.open_document(filepath)
            flat_tc = docx_handler.collect_paragraphs(doc_tc)
            docx_handler.insert_clause_after(
                doc_tc, para_idx, subtitle, clause_text, author,
                subtitle_config=subtitle_config, text_style=text_style,
                subtitle_font_size=subtitle_font_size, text_font_size=text_font_size,
                flat_paras=flat_tc,
            )
            date_result = docx_handler.update_dates(doc_tc, author, flat_paras=flat_tc)
            docx_handler.save_document(doc_tc, str(self._tc_dir / filename))

            # ── Version texte brut ────────────────────────────────────────────
            doc_plain = docx_handler.open_document(filepath)
            flat_plain = docx_handler.collect_paragraphs(doc_plain)
            docx_handler.insert_clause_plain_after(
                doc_plain, para_idx, subtitle, clause_text,
                subtitle_config=subtitle_config, text_style=text_style,
                subtitle_font_size=subtitle_font_size, text_font_size=text_font_size,
                flat_paras=flat_plain,
            )
            docx_handler.update_dates_plain(doc_plain, flat_paras=flat_plain)
            docx_handler.save_document(doc_plain, str(self._plain_dir / filename))

            # ── Avertissement si aucune date trouvée ──────────────────────────
            if not date_result["body"] and not date_result["footer"]:
                messagebox.showwarning(
                    "Date non trouvée",
                    f"Aucune date de publication n'a été trouvée dans « {filename} ».\n\n"
                    "Patterns recherchés :\n"
                    "  • Corps   : paragraphe contenant « Date de publication »\n"
                    "  • Footer  : paragraphe contenant « mise à jour le »\n\n"
                    "La clause a bien été insérée."
                )

            # ── Statut ────────────────────────────────────────────────────────
            date_parts = []
            if date_result["body"]:
                date_parts.append("corps")
            if date_result["footer"]:
                date_parts.append("footer")
            date_info = f" + date ({', '.join(date_parts)})" if date_parts else ""

            logger.log_insertion(filepath, self._fund_var.get(), para_idx, subtitle, clause_text)
            self._status(
                f"Clause insérée après §{para_idx}{date_info}"
                f" → {TC_FOLDER_NAME}/ et {PLAIN_FOLDER_NAME}/"
            )
            self._refresh_log()
            self._next_file()
        except Exception as e:
            messagebox.showerror("Erreur lors de l'insertion", str(e))

    def _status(self, msg: str):
        """Met à jour le label de statut dans la barre d'actions."""
        self._lbl_status.config(text=msg)

    def _refresh_log(self):
        """Recharge et affiche les 20 dernières lignes du journal CSV."""
        lines = logger.get_recent_logs(20)
        self._txt_log.config(state=tk.NORMAL)
        self._txt_log.delete("1.0", tk.END)
        self._txt_log.insert("1.0", "\n".join(lines) if lines else "(aucune insertion enregistrée)")
        self._txt_log.config(state=tk.DISABLED)
