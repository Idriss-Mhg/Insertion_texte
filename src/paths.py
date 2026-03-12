import sys
from pathlib import Path


def get_runtime_dir() -> Path:
    """Retourne le dossier 'runtime' : dossier du .exe si frozen, sinon racine du projet."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent.parent


RUNTIME_DIR  = get_runtime_dir()
CLAUSES_PATH = RUNTIME_DIR / "clauses.json"
LOG_DIR      = RUNTIME_DIR / "logs"
LOG_FILE     = LOG_DIR / "insertions.csv"
BACKUP_DIR   = RUNTIME_DIR / "_sauvegardes"