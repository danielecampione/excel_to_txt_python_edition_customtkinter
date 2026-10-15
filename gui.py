"""Interfaccia CustomTkinter completa con drag-and-drop, selezione file e join opzionale."""
from __future__ import annotations

import os
import threading
import customtkinter as ctk
from tkinter import filedialog

from logic import consume_warnings, process_excel, save_output

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    DND_FILES = None
    TkinterDnD = None

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

BG = "#f7f5f2"
SURFACE = "#ffffff"
CARD = "#ffffff"
ACCENT = "#e05c2a"
ACCENT_H = "#c94d1f"
DROP_BG = "#fff8f5"
DROP_ACT = "#fde8df"
SUCCESS = "#2d9e6b"
ERROR = "#d63b3b"
WARNING = "#d97706"
TEXT = "#1a1a1a"
MUTED = "#888077"
BORDER = "#ddd8d0"

DEFAULT_COLUMNS = "Numero\nDescrizione\nBreve descrizione"
DEFAULT_MAIN_KEYS = "Richiesto per\nReparto\nUbicazione"
DEFAULT_SECONDARY_KEYS = "Nome utente\nReparto\nUbicazione"
DEFAULT_JOIN_OUTPUT = "Nome utente\nNumero dipendente\nTitolo\nReparto"


class ExcelToTxtGUI(TkinterDnD.Tk if DND_AVAILABLE else ctk.CTk):
    def __init__(self):
        super().__init__()
        self.wm_attributes("-alpha", 0.0)
        self.configure(bg=BG)
        self.title("Excel → TXT")
        self.geometry(self._initial_geometry())
        self.minsize(700, 680)
        self.resizable(True, True)

        self.F_TITLE = ctk.CTkFont("Georgia", 26, "bold")
        self.F_SECTION = ctk.CTkFont("Georgia", 16, "bold")
        self.F_LABEL = ctk.CTkFont("Georgia", 14)
        self.F_BODY = ctk.CTkFont("Helvetica", 14)
        self.F_SMALL = ctk.CTkFont("Helvetica", 12)
        self.F_MONO = ctk.CTkFont("Courier", 13)
        self.F_BTN = ctk.CTkFont("Helvetica", 16, "bold")

        self._file_path: str | None = None
        self._secondary_path: str | None = None
        self._running = False
        self._build()
        self._anim_open()
        self.protocol("WM_DELETE_WINDOW", self._anim_close)

    def _initial_geometry(self):
        self.update_idletasks()
        width = 780
        height = min(940, max(700, self.winfo_screenheight() - 80))
        x = max(0, (self.winfo_screenwidth() - width) // 2)
        y = max(0, (self.winfo_screenheight() - height) // 2)
        return f"{width}x{height}+{x}+{y}"

    def _ease_out(self, t): return 1 - (1 - t) ** 3
    def _ease_in(self, t): return t ** 2

    def _anim_open(self, step=0):
        steps, interval = 40, 12
        if step > steps:
            self.wm_attributes("-alpha", 1.0)
            return
        self.wm_attributes("-alpha", self._ease_out(step / steps))
        self.after(interval, self._anim_open, step + 1)

    def _anim_close(self, step=0):
        steps, interval, slide_px = 13, 10, 14
        if step > steps:
            self.destroy()
            return
        ease = self._ease_in(step / steps)
        self.wm_attributes("-alpha", 1.0 - ease)
        self.geometry(f"+{self.winfo_x()}+{self.winfo_y() + int(slide_px * ease)}")
        self.after(interval, self._anim_close, step + 1)

    def _build(self):
        header = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0)
        header.pack(fill="x")
        ctk.CTkLabel(header, text="📊  Excel → TXT", font=self.F_TITLE,
                     text_color=ACCENT, anchor="w").pack(side="left", padx=28, pady=20)
        ctk.CTkLabel(header, text="Estrazione e join di documenti Excel",
                     font=self.F_SMALL, text_color=MUTED, anchor="w").pack(side="left", padx=4)
        ctk.CTkFrame(self, fg_color=BORDER, height=1, corner_radius=0).pack(fill="x")

        body = ctk.CTkScrollableFrame(self, fg_color=BG, corner_radius=0,
                                      scrollbar_button_color=BORDER,
                                      scrollbar_button_hover_color=MUTED)
        body.pack(fill="both", expand=True, padx=24, pady=20)

        self._section(body, "Documento principale")
        self.main_drop = self._drop_panel(
            body, "Trascina qui il file Excel principale (.xlsx / .xls)",
            "oppure clicca per sfogliare", True, 118)
        self.columns = self._area(body,
            "Colonne principali da estrarre (una per riga, nell'ordine desiderato)",
            DEFAULT_COLUMNS, 110)

        join = ctk.CTkFrame(body, fg_color=CARD, corner_radius=14,
                            border_width=1, border_color=BORDER)
        join.pack(fill="x", pady=(18, 10))
        ctk.CTkLabel(join, text="Join con documento secondario", font=self.F_SECTION,
                     text_color=TEXT, anchor="w").pack(fill="x", padx=16, pady=(14, 2))
        ctk.CTkLabel(join,
            text="Opzionale: senza documento secondario l'app continua a funzionare come prima.",
            font=self.F_SMALL, text_color=MUTED, anchor="w").pack(fill="x", padx=16, pady=(0, 10))

        self.secondary_drop = self._drop_panel(
            join, "Trascina qui il documento secondario per la join",
            "ad esempio: sys_user.xlsx", False, 100, 16)

        keys_card = ctk.CTkFrame(join, fg_color=DROP_BG, corner_radius=12,
                                 border_width=1, border_color=BORDER)
        keys_card.pack(fill="x", padx=16, pady=(12, 6))
        ctk.CTkLabel(keys_card, text="1. Chiavi di collegamento", font=self.F_LABEL,
                     text_color=ACCENT, anchor="w").pack(fill="x", padx=12, pady=(10, 0))
        self.main_keys = self._area(keys_card,
            "Chiavi del DOCUMENTO PRINCIPALE (una per riga)", DEFAULT_MAIN_KEYS, 92, 12)
        self.secondary_keys = self._area(keys_card,
            "Chiavi del DOCUMENTO SECONDARIO (una per riga)", DEFAULT_SECONDARY_KEYS, 92, 12)

        self.join_output = self._area(join,
            "2. Colonne del documento secondario da aggiungere all'output",
            DEFAULT_JOIN_OUTPUT, 110, 16)

        self.fallback = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(join, text="Abilita join con fallback sull'ultima chiave",
                        variable=self.fallback, font=self.F_SMALL,
                        text_color=TEXT, fg_color=ACCENT, hover_color=ACCENT_H,
                        border_color=BORDER).pack(anchor="w", padx=16, pady=(8, 4))
        ctk.CTkLabel(join,
            text="Le chiavi sono accoppiate riga per riga. Il fallback ignora soltanto l'ultima chiave\ne solo in assenza di ambiguità.",
            justify="left", font=self.F_SMALL, text_color=MUTED,
            anchor="w").pack(fill="x", padx=16)
        ctk.CTkButton(join, text="Rimuovi documento secondario", command=self._clear_secondary,
                      font=self.F_SMALL, height=32, corner_radius=8,
                      fg_color=SURFACE, hover_color=DROP_ACT, text_color=ACCENT,
                      border_width=1, border_color=BORDER).pack(anchor="w", padx=16, pady=(10, 16))

        self._status = ctk.CTkTextbox(body, font=self.F_SMALL, height=120,
                                      fg_color=SURFACE, text_color=MUTED,
                                      border_width=1, border_color=BORDER,
                                      corner_radius=12, wrap="word")
        self._status.pack(fill="x", pady=(6, 8))
        self._status.configure(state="disabled")
        self._set_status("In attesa del documento principale...", MUTED)
        if not DND_AVAILABLE:
            self._set_status("tkinterdnd2 non trovato: il clic per sfogliare resta disponibile.", WARNING)

        self._progress = ctk.CTkProgressBar(body, fg_color=BORDER,
                                            progress_color=ACCENT, height=6, corner_radius=3)
        self._progress.set(0)
        self._run_btn = ctk.CTkButton(body, text="Elabora  →", font=self.F_BTN,
                                      height=58, corner_radius=12, fg_color=ACCENT,
                                      hover_color=ACCENT_H, text_color="#fff",
                                      command=self._run, state="disabled")
        self._run_btn.pack(fill="x", pady=(8, 12))

    def _section(self, parent, text):
        ctk.CTkLabel(parent, text=text, font=self.F_SECTION,
                     text_color=TEXT, anchor="w").pack(fill="x", pady=(0, 6))

    def _area(self, parent, label, value, height, padx=0):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=padx, pady=(9, 2))
        ctk.CTkLabel(frame, text=label, font=self.F_LABEL,
                     text_color=TEXT, anchor="w").pack(fill="x", pady=(0, 4))
        area = ctk.CTkTextbox(frame, font=self.F_MONO, height=height,
                              fg_color=SURFACE, text_color=TEXT,
                              border_width=1, border_color=BORDER,
                              scrollbar_button_color=BORDER,
                              scrollbar_button_hover_color=MUTED,
                              corner_radius=10, wrap="none")
        area.pack(fill="x")
        area.insert("1.0", value)
        return area

    def _drop_panel(self, parent, first, second, main, height, padx=0):
        frame = ctk.CTkFrame(parent, fg_color=DROP_BG, corner_radius=14,
                             border_width=2, border_color=BORDER, height=height,
                             cursor="hand2")
        frame.pack(fill="x", padx=padx, pady=(0, 4))
        frame.pack_propagate(False)
        icon = ctk.CTkLabel(frame, text="FILE", font=self.F_LABEL, text_color=ACCENT)
        line1 = ctk.CTkLabel(frame, text=first, font=self.F_BODY, text_color=TEXT)
        line2 = ctk.CTkLabel(frame, text=second, font=self.F_SMALL, text_color=MUTED)
        icon.pack(expand=True, pady=(10, 0)); line1.pack(); line2.pack(pady=(0, 10))
        frame._labels = (line1, line2)
        frame._main = main
        for widget in (frame, icon, line1, line2):
            widget.bind("<Button-1>", lambda _e, m=main: self._choose(m))
            widget.bind("<Enter>", lambda _e, f=frame: f.configure(fg_color=DROP_ACT, border_color=ACCENT))
            widget.bind("<Leave>", lambda _e, f=frame: f.configure(fg_color=DROP_BG, border_color=BORDER))
            if DND_AVAILABLE:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", lambda e, m=main: self._on_drop(e, m))
                widget.dnd_bind("<<DragEnter>>", lambda _e, f=frame: f.configure(fg_color=DROP_ACT, border_color=ACCENT))
                widget.dnd_bind("<<DragLeave>>", lambda _e, f=frame: f.configure(fg_color=DROP_BG, border_color=BORDER))
        return frame

    @staticmethod
    def _lines(area):
        return [line.strip() for line in area.get("1.0", "end").splitlines() if line.strip()]

    @staticmethod
    def _valid(path):
        return path.lower().endswith((".xlsx", ".xls"))

    def _choose(self, main):
        path = filedialog.askopenfilename(title="Seleziona file Excel",
                                          filetypes=[("File Excel", "*.xlsx *.xls")])
        if path:
            self._handle_file(path, main)

    def _on_drop(self, event, main):
        paths = self.tk.splitlist(event.data.strip())
        if paths:
            self._handle_file(paths[0], main)

    def _handle_file(self, path, main):
        if not self._valid(path):
            if main:
                self._file_path = None
                self._run_btn.configure(state="disabled")
            else:
                self._secondary_path = None
            self._set_status("File non supportato. Usa .xlsx o .xls", ERROR)
            return
        panel = self.main_drop if main else self.secondary_drop
        panel._labels[0].configure(text=f"✓  {os.path.basename(path)}", text_color=SUCCESS)
        panel._labels[1].configure(text="Documento principale pronto" if main else "Join attiva")
        if main:
            self._file_path = path
            self._run_btn.configure(state="normal")
            self._set_status("Documento principale pronto.", MUTED)
        else:
            self._secondary_path = path
            self._set_status("Documento secondario pronto: la join verrà applicata.", MUTED)

    def _clear_secondary(self):
        self._secondary_path = None
        self.secondary_drop._labels[0].configure(
            text="Trascina qui il documento secondario per la join", text_color=TEXT)
        self.secondary_drop._labels[1].configure(text="ad esempio: sys_user.xlsx")
        self._set_status("Documento secondario rimosso. Elaborazione standard attiva.", MUTED)

    def _set_status(self, message, color=MUTED):
        self._status.configure(state="normal", text_color=color)
        self._status.delete("1.0", "end")
        self._status.insert("1.0", message)
        self._status.configure(state="disabled")
        self._status.see("1.0")

    def _run(self):
        if self._running or not self._file_path:
            return
        columns = self._lines(self.columns)
        if not columns:
            self._set_status("Inserisci almeno una colonna principale.", ERROR)
            return
        args = (self._file_path, columns, self._secondary_path,
                self._lines(self.main_keys), self._lines(self.secondary_keys),
                self._lines(self.join_output), self.fallback.get())
        self._running = True
        self._run_btn.configure(state="disabled", text="⏳  Un momento…")
        self._progress.pack(fill="x", pady=(0, 6), before=self._run_btn)
        self._progress.start()
        self._set_status("Elaborazione in corso…", MUTED)
        threading.Thread(target=self._worker, args=args, daemon=True).start()

    def _worker(self, *args):
        try:
            content = process_excel(*args)
            out = save_output(args[0], content)
            warning = consume_warnings()
            self.after(0, lambda: self._finish(out, warning))
        except Exception as exc:
            self.after(0, lambda e=exc: self._fail(e))

    def _finish(self, out, warning):
        self._progress.stop(); self._progress.set(1)
        if warning:
            self._set_status(f"File salvato correttamente:\n{out}\n\n{warning}", WARNING)
        else:
            self._set_status(f"✓  File salvato:\n{out}", SUCCESS)
        self._run_btn.configure(state="normal", text="Elabora  →")
        self._running = False

    def _fail(self, exc):
        self._progress.stop(); self._progress.set(0)
        self._set_status(f"✗  Errore:\n{exc}", ERROR)
        self._run_btn.configure(state="normal", text="Elabora  →")
        self._running = False


def build_root():
    return ExcelToTxtGUI()
