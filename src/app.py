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
#   Onglet "Clauses" :
#     Éditeur de clauses. Permet de créer, modifier, renommer, dupliquer et
#     supprimer les codes d'insertion stockés dans clauses.json.
#
# Dépendances externes : tkinterweb (prévisualisation HTML), python-docx, lxml.
# =============================================================================

import json
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tkinterweb import HtmlFrame

from src import docx_handler, logger

# -----------------------------------------------------------------------------
# Constantes
# -----------------------------------------------------------------------------

# Chemin vers le fichier de configuration des clauses (à la racine du projet)
CLAUSES_PATH = Path(__file__).parent.parent / "clauses.json"

# Nom du sous-dossier créé dans le dossier de travail pour les sauvegardes _old
BACKUP_FOLDER_NAME = "_sauvegardes"

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

        # Dossier de travail sélectionné par l'utilisateur
        self._folder: Path | None = None

        # Sous-dossier de sauvegardes (_sauvegardes/) créé dans _folder
        self._backup_dir: Path | None = None

        # File d'attente : liste ordonnée des .docx à traiter dans _folder
        self._queue: list[Path] = []

        # Index du fichier courant dans _queue (-1 = aucun fichier chargé)
        self._queue_idx: int = -1

        # Mapping position_listbox → index_paragraphe dans le document courant.
        # Nécessaire car la listbox n'affiche pas forcément tous les paragraphes
        # (filtrage par recherche possible).
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
          5. PanedWindow : listbox des paragraphes | prévisualisation HTML
          6. Boutons d'action (Insérer / Passer)
          7. Historique des insertions (log)
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

        ttk.Label(f_clause, text="Texte de la clause :").pack(anchor=tk.W, padx=6)
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

        # ── 5. PanedWindow : listbox + prévisualisation ───────────────────────
        # Le PanedWindow est redimensionnable par l'utilisateur (poignée centrale).
        # Poids : listbox=1, preview=2 → la preview est deux fois plus large par défaut.
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
        # À chaque sélection, on met en évidence le paragraphe dans la preview
        self._listbox.bind("<<ListboxSelect>>", self._on_para_select)

        # Pane droite — prévisualisation HTML du document
        # tkinterweb/HtmlFrame utilise le moteur tkhtml3 (pas de JavaScript).
        f_preview = ttk.LabelFrame(paned, text="Aperçu du document")
        paned.add(f_preview, weight=2)

        self._html_frame = HtmlFrame(f_preview, messages_enabled=False)
        self._html_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self._html_frame.load_html("<body style='color:gray;padding:16px'>Aucun document chargé.</body>")

        # ── 6. Boutons d'action ───────────────────────────────────────────────
        f_actions = ttk.Frame(parent)
        f_actions.pack(fill=tk.X, **pad)

        ttk.Button(f_actions, text="Insérer la clause", command=self._insert).pack(side=tk.LEFT)
        # "Passer" avance au fichier suivant sans modifier le document courant
        ttk.Button(f_actions, text="Passer (sans insérer)", command=self._next_file).pack(side=tk.LEFT, padx=8)
        self._lbl_status = ttk.Label(f_actions, text="", foreground="gray")
        self._lbl_status.pack(side=tk.LEFT, padx=12)

        # ── 7. Historique des insertions ──────────────────────────────────────
        f_log = ttk.LabelFrame(parent, text="Historique des insertions")
        f_log.pack(fill=tk.X, **pad)

        # Widget en lecture seule : on l'active brièvement pour écrire dedans
        self._txt_log = tk.Text(f_log, height=4, state=tk.DISABLED, font=("Courier", 8), foreground="#444")
        self._txt_log.pack(fill=tk.X, padx=4, pady=4)

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
        f_right.rowconfigure(3, weight=1)  # La zone de texte occupe l'espace restant

        ttk.Label(f_right, text="Code :").grid(row=0, column=0, sticky="w", padx=8, pady=(8, 4))
        self._edit_name_var = tk.StringVar()
        self._edit_name_entry = ttk.Entry(f_right, textvariable=self._edit_name_var, width=24)
        self._edit_name_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=(8, 4))

        ttk.Label(f_right, text="Sous-titre :").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        self._edit_subtitle_var = tk.StringVar()
        ttk.Entry(f_right, textvariable=self._edit_subtitle_var).grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=4)

        ttk.Label(f_right, text="Texte de la clause :").grid(row=2, column=0, sticky="nw", padx=8, pady=4)

        self._edit_txt = tk.Text(f_right, wrap=tk.WORD, height=12)
        sb_r = ttk.Scrollbar(f_right, orient=tk.VERTICAL, command=self._edit_txt.yview)
        self._edit_txt.config(yscrollcommand=sb_r.set)
        self._edit_txt.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=8, pady=(0, 4))
        sb_r.grid(row=3, column=2, sticky="ns", pady=(0, 4))

        f_form_btns = ttk.Frame(f_right)
        f_form_btns.grid(row=4, column=0, columnspan=3, sticky="ew", padx=8, pady=(0, 8))
        ttk.Button(f_form_btns, text="Enregistrer", command=self._clause_save).pack(side=tk.LEFT)
        ttk.Button(f_form_btns, text="Annuler",     command=self._clause_cancel).pack(side=tk.LEFT, padx=8)
        self._lbl_clause_status = ttk.Label(f_form_btns, text="", foreground="gray")
        self._lbl_clause_status.pack(side=tk.LEFT)

        # Nom du code en cours d'édition (None si création d'un nouveau code)
        self._clause_editing_original_name: str | None = None

        self._refresh_type_listbox()

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
        self._edit_txt.delete("1.0", tk.END)
        self._edit_txt.insert("1.0", data.get("text", ""))
        self._lbl_clause_status.config(text="")

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
        original = self._clause_editing_original_name

        if original and original != new_name:
            # Renommage : on reconstruit le dict en remplaçant la clé originale
            # pour conserver la position du code dans la liste
            if new_name in self._clauses:
                if not messagebox.askyesno("Confirmer", f"« {new_name} » existe déjà. Écraser ?"):
                    return
            self._clauses = {
                (new_name if k == original else k): ({"subtitle": subtitle, "text": text} if k == original else v)
                for k, v in self._clauses.items()
            }
        else:
            # Création ou modification sans renommage
            if new_name in self._clauses and original is None:
                if not messagebox.askyesno("Confirmer", f"« {new_name} » existe déjà. Écraser ?"):
                    return
            self._clauses[new_name] = {"subtitle": subtitle, "text": text}

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
        self._backup_dir = self._folder / BACKUP_FOLDER_NAME
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
        Charge le fichier courant (self._queue[self._queue_idx]) :
          1. Crée une sauvegarde _old.docx dans _sauvegardes/
          2. Ouvre le document avec python-docx
          3. Met à jour les widgets de navigation et de statut
          4. Vide la listbox et rafraîchit la prévisualisation
        """
        path = self._queue[self._queue_idx]
        try:
            self._doc = docx_handler.backup_and_open(str(path), self._backup_dir)
        except Exception as e:
            messagebox.showerror("Erreur", str(e))
            return

        total = len(self._queue)
        pos = self._queue_idx + 1
        self._lbl_file.config(text=path.name)
        self._lbl_progress.config(text=f"{pos} / {total}")
        self._btn_prev.config(state=tk.NORMAL if self._queue_idx > 0 else tk.DISABLED)
        self._btn_next.config(state=tk.NORMAL if self._queue_idx < total - 1 else tk.DISABLED)
        self._listbox.delete(0, tk.END)
        self._para_indices = []
        self._status(f"Fichier chargé. Sauvegarde dans {BACKUP_FOLDER_NAME}/")
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
        Pré-remplit les champs sous-titre et texte de la clause
        à partir du code d'insertion sélectionné dans le combobox.
        """
        code = self._fund_var.get()
        data = self._clauses.get(code, {})
        self._subtitle_var.set(data.get("subtitle", ""))
        self._txt_clause.delete("1.0", tk.END)
        self._txt_clause.insert("1.0", data.get("text", ""))

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

        results = docx_handler.search_paragraphs(self._doc, keyword)
        if not results:
            messagebox.showinfo("Aucun résultat", f"Mot-clé « {keyword} » introuvable.")
            return

        # Collecte les paragraphes de contexte autour de chaque occurrence,
        # en dédupliquant si plusieurs occurrences sont proches.
        displayed: list[tuple[int, str]] = []
        seen: set[int] = set()
        for (center_idx, _) in results:
            for (i, text) in docx_handler.get_paragraphs_around(self._doc, center_idx):
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
        self._populate_listbox(docx_handler.get_all_paragraphs(self._doc))
        self._status(f"{len(self._para_indices)} paragraphes affichés.")

    def _on_para_select(self, *_):
        """
        Déclenché à chaque clic dans la listbox.
        Régénère la prévisualisation HTML avec le paragraphe sélectionné mis en
        évidence, puis scrolle jusqu'à lui via yview_moveto.
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

    def _scroll_fraction(self, highlight_idx: int) -> float:
        """
        Calcule la fraction de scroll (0.0 à 1.0) correspondant à la position
        visuelle du paragraphe highlight_idx dans la prévisualisation.

        On pondère chaque paragraphe par sa hauteur estimée plutôt que de
        diviser naïvement par le nombre total de paragraphes. Cela compense
        le fait que les paragraphes ont des longueurs très variables : un long
        paragraphe occupe beaucoup plus d'espace visuel qu'un titre court.

        Heuristique :
          - Poids de base = max(1, len(texte) / 80)  [~80 caractères par ligne]
          - Bonus heading  = +2.0 lignes             [marge CSS margin-top: 1.2em]

        Args:
            highlight_idx: Index du paragraphe cible dans doc.paragraphs.

        Returns:
            Fraction entre 0.0 et 0.99 à passer à yview_moveto().
        """
        _heading_keys = {"heading", "titre", "title"}
        weights: list[tuple[int, float]] = []

        for i, para in enumerate(self._doc.paragraphs):
            text = para.text.strip()
            if not text:
                continue
            style = (para.style.name or "").lower() if para.style else ""
            is_heading = any(k in style for k in _heading_keys)
            lines = max(1.0, len(text) / 80)
            weight = lines + (2.0 if is_heading else 0.0)
            weights.append((i, weight))

        total = sum(w for _, w in weights)
        if total == 0:
            return 0.0
        cumulative = sum(w for i, w in weights if i < highlight_idx)
        return max(0.0, min(0.99, cumulative / total))

    def _refresh_preview(self, highlight_idx: int | None = None):
        """
        Génère le HTML du document courant via docx_handler.build_html() et
        l'injecte dans le widget HtmlFrame.

        Si highlight_idx est fourni, le paragraphe correspondant reçoit la
        classe CSS "highlight". Après le chargement (délai de 120 ms pour
        laisser le temps au moteur HTML de rendre la page), on appelle
        yview_moveto() avec la fraction calculée par _scroll_fraction().

        Note : tkinterweb ne supporte pas JavaScript par défaut (moteur tkhtml3),
        donc on ne peut pas utiliser scrollIntoView(). On utilise à la place
        yview_moveto() qui est une méthode native du widget Tk.

        Args:
            highlight_idx: Index du paragraphe à mettre en évidence, ou None.
        """
        if not self._doc:
            return
        try:
            html_body = docx_handler.build_html(self._doc, highlight_idx=highlight_idx)
            html = (
                f"<!DOCTYPE html><html><head>"
                f'<meta charset="utf-8">'
                f"<style>{_PREVIEW_CSS}</style>"
                f"</head><body>{html_body}</body></html>"
            )
            self._html_frame.load_html(html)
            if highlight_idx is not None:
                fraction = self._scroll_fraction(highlight_idx)
                # Délai nécessaire : load_html() est asynchrone côté rendu tkhtml3
                self.after(120, lambda f=fraction: self._html_frame.yview_moveto(f))
        except Exception:
            self._html_frame.load_html(
                "<body style='color:gray;padding:16px'>Aperçu indisponible pour ce fichier.</body>"
            )

    # =========================================================================
    # Logique de l'onglet Insertion — Insertion et log
    # =========================================================================

    def _insert(self):
        """
        Insère la clause dans le document courant après le paragraphe sélectionné.

        Étapes :
          1. Validation : document ouvert, paragraphe sélectionné, clause non vide
          2. Insertion via docx_handler.insert_clause_after()
          3. Sauvegarde du document (écrase le .docx original)
          4. Journalisation dans logs/insertions.csv
          5. Rafraîchissement du log et passage au fichier suivant
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

        if not clause_text:
            messagebox.showwarning("Attention", "Le texte de la clause est vide.")
            return

        filepath = str(self._queue[self._queue_idx])
        try:
            docx_handler.insert_clause_after(self._doc, para_idx, subtitle, clause_text)
            docx_handler.save_document(self._doc, filepath)
            logger.log_insertion(filepath, self._fund_var.get(), para_idx, subtitle, clause_text)
            self._status(f"Clause insérée après §{para_idx}. Passage au fichier suivant…")
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
