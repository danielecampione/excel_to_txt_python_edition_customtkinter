"""
main.py
-------
Entry point: avvia la GUI.

Uso:
    python main.py

Dipendenze:
    pip install customtkinter tkinterdnd2 openpyxl pandas
"""

from gui import build_root


def main():
    root = build_root()
    root.mainloop()


if __name__ == "__main__":
    main()
