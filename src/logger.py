# =============================================================================
# src/logger.py — Journalisation des insertions
#
# Enregistre chaque insertion dans un fichier CSV horodaté situé dans logs/.
# Ce journal permet de tracer toutes les modifications effectuées sur les
# prospectus : qui a inséré quoi, dans quel fichier, et à quel endroit.
#
# Format du CSV (séparateur point-virgule, encodage UTF-8) :
#   date ; heure ; fichier ; code_insertion ; paragraphe_index ; sous_titre ; extrait_clause
# =============================================================================

import csv
from datetime import datetime
from pathlib import Path

# Emplacement du fichier de log : <racine_projet>/logs/insertions.csv
LOG_DIR = Path(__file__).parent.parent / "logs"
LOG_FILE = LOG_DIR / "insertions.csv"

_HEADERS = ["date", "heure", "fichier", "code_insertion", "paragraphe_index", "sous_titre", "extrait_clause"]


def _ensure_log_file() -> None:
    """
    Crée le dossier logs/ et le fichier CSV avec ses en-têtes si nécessaire.
    Appelé avant chaque écriture pour garantir que le fichier existe.
    """
    LOG_DIR.mkdir(exist_ok=True)
    if not LOG_FILE.exists():
        with open(LOG_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(_HEADERS)


def log_insertion(filepath: str, code_insertion: str, para_idx: int, subtitle: str, clause_text: str) -> None:
    """
    Ajoute une ligne dans le journal CSV pour tracer une insertion effectuée.

    Args:
        filepath: Chemin complet du fichier .docx modifié.
        code_insertion: Code d'insertion utilisé (ex: "LMT_OPCVM").
        para_idx: Index du paragraphe après lequel la clause a été insérée.
        subtitle: Sous-titre inséré (peut être vide).
        clause_text: Texte complet de la clause insérée.
    """
    _ensure_log_file()
    now = datetime.now()
    # On tronque l'extrait à 80 caractères pour garder le CSV lisible
    extrait = clause_text[:80].replace("\n", " ") + ("..." if len(clause_text) > 80 else "")
    row = [
        now.strftime("%Y-%m-%d"),
        now.strftime("%H:%M:%S"),
        Path(filepath).name,
        code_insertion,
        para_idx,
        subtitle,
        extrait,
    ]
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(row)


def get_recent_logs(n: int = 20) -> list[str]:
    """
    Retourne les n dernières lignes du journal, formatées pour l'affichage
    dans le widget de log de l'interface.

    Format retourné : "YYYY-MM-DD HH:MM:SS  |  nom_fichier.docx  |  CODE_INSERTION"

    Args:
        n: Nombre maximum de lignes à retourner (les plus récentes).

    Returns:
        Liste de chaînes formatées, ou liste vide si le fichier n'existe pas.
    """
    if not LOG_FILE.exists():
        return []
    with open(LOG_FILE, "r", encoding="utf-8") as f:
        reader = csv.reader(f, delimiter=";")
        rows = list(reader)[1:]  # On saute la ligne d'en-tête
    return [
        f"{row[0]} {row[1]}  |  {row[2]}  |  {row[3]}"
        for row in rows[-n:]
        if len(row) >= 4
    ]