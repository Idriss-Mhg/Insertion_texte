# =============================================================================
# main.py — Point d'entrée de l'application
#
# Lance la fenêtre principale Tkinter. À exécuter directement :
#   python main.py
# ou via le venv activé :
#   source .venv/bin/activate && python main.py
# =============================================================================

from src.app import App

if __name__ == "__main__":
    app = App()
    app.mainloop()
