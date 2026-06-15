# gui.py
import os
import re
import threading
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
from logic import process_excel, save_output

# ── Tema ───────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

BG       = "#f7f5f2"
SURFACE  = "#ffffff"
CARD     = "#ffffff"
ACCENT   = "#e05c2a"
ACCENT_H = "#c94d1f"
DROP_BG  = "#fff8f5"
DROP_ACT = "#fde8df"
SUCCESS  = "#2d9e6b"
ERROR    = "#d63b3b"
TEXT     = "#1a1a1a"
MUTED    = "#888077"
BORDER   = "#ddd8d0"

DEFAULT_COLUMNS = "Numero\nDescrizione\nBreve descrizione"


def _parse_paths(data: str) -> list[str]:
    paths = []
    for match in re.finditer(r'\{([^}]+)\}|(\S+)', data):
        p = match.group(1) or match.group(2)
        if p:
            paths.append(p)
    return paths


# ── GUI ────────────────────────────────────────────────────────────────────────
class ExcelToTxtGUI(TkinterDnD.Tk):

    def __init__(self):
        super().__init__()
        self.wm_attributes("-alpha", 0.0)

        self.configure(bg=BG)
        ctk.set_appearance_mode("light")

        self.title("Excel → TXT")
        self.geometry("700x720")
        self.minsize(580, 600)
        self.resizable(True, True)

        self.F_TITLE = ctk.CTkFont("Georgia",   26, "bold")
        self.F_LABEL = ctk.CTkFont("Georgia",   15)
        self.F_BODY  = ctk.CTkFont("Helvetica", 14)
        self.F_SMALL = ctk.CTkFont("Helvetica", 12)
        self.F_MONO  = ctk.CTkFont("Courier",   13)
        self.F_BTN   = ctk.CTkFont("Helvetica", 16, "bold")

        self._file_path = None
        self._running   = False

        self._build()
        self._anim_open()
        self.protocol("WM_DELETE_WINDOW", self._anim_close)

    # ── Animazioni ─────────────────────────────────────────────────────────────
    def _ease_out(self, t): return 1 - (1 - t) ** 3
    def _ease_in(self, t):  return t ** 2

    def _anim_open(self, step=0):
        STEPS, INTERVAL = 40, 12
        if step > STEPS:
            self.wm_attributes("-alpha", 1.0)
            return
        self.wm_attributes("-alpha", self._ease_out(step / STEPS))
        self.after(INTERVAL, self._anim_open, step + 1)

    def _anim_close(self, step=0):
        STEPS, INTERVAL, SLIDE_PX = 13, 10, 14
        if step > STEPS:
            self.destroy()
            return
        ease = self._ease_in(step / STEPS)
        self.wm_attributes("-alpha", 1.0 - ease)
        x, y = self.winfo_x(), self.winfo_y()
        self.geometry(f"+{x}+{y + int(SLIDE_PX * ease)}")
        self.after(INTERVAL, self._anim_close, step + 1)

    # ── Layout ─────────────────────────────────────────────────────────────────
    def _build(self):
        # Header
        hdr = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0)
        hdr.pack(fill="x")
        ctk.CTkLabel(hdr, text="📊  Excel → TXT",
                     font=self.F_TITLE, text_color=ACCENT, anchor="w"
                     ).pack(side="left", padx=28, pady=20)
        ctk.CTkLabel(hdr, text="Estrae i record da un foglio Excel",
                     font=self.F_SMALL, text_color=MUTED, anchor="w"
                     ).pack(side="left", padx=4)
        ctk.CTkFrame(self, fg_color=BORDER, height=1, corner_radius=0).pack(fill="x")

        # Corpo scrollabile
        body = ctk.CTkScrollableFrame(
            self, fg_color=BG, corner_radius=0,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=MUTED,
        )
        body.pack(fill="both", expand=True, padx=24, pady=20)

        # ── 1. Zona drop ───────────────────────────────────────────────────────
        self._label(body, "Trascina qui il file Excel (.xlsx)")

        self._drop_frame = ctk.CTkFrame(
            body, fg_color=DROP_BG, corner_radius=14,
            border_width=2, border_color=BORDER
        )
        self._drop_frame.pack(fill="x", pady=(4, 4))

        self._drop_label = ctk.CTkLabel(
            self._drop_frame,
            text="⬇   Trascina qui\nil file .xlsx",
            font=self.F_BODY, text_color=MUTED,
            justify="center"
        )
        self._drop_label.pack(pady=22)

        for widget in (self._drop_frame, self._drop_label):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>",      self._on_drop)
            widget.dnd_bind("<<DragEnter>>", self._on_drag_enter)
            widget.dnd_bind("<<DragLeave>>", self._on_drag_leave)

        # ── 2. Colonne da estrarre ─────────────────────────────────────────────
        self._label(body, "Colonne da estrarre (una per riga, nell'ordine desiderato)")

        col_card = ctk.CTkFrame(
            body, fg_color=CARD, corner_radius=12,
            border_width=1, border_color=BORDER
        )
        col_card.pack(fill="x", pady=(0, 4))

        self._col_box = ctk.CTkTextbox(
            col_card,
            font=self.F_MONO,
            height=110,
            fg_color=CARD,
            text_color=TEXT,
            border_width=0,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=MUTED,
            wrap="none",
        )
        self._col_box.pack(fill="x", padx=12, pady=10)
        self._col_box.insert("1.0", DEFAULT_COLUMNS)

        # hint sotto la textarea
        ctk.CTkLabel(
            body,
            text="💡  I valori multi-riga vengono racchiusi automaticamente tra virgolette",
            font=ctk.CTkFont("Helvetica", 11),
            text_color=MUTED, anchor="w"
        ).pack(fill="x", pady=(0, 6))

        # ── Stato ──────────────────────────────────────────────────────────────
        self._status = ctk.CTkLabel(
            body, text="", font=self.F_SMALL,
            text_color=MUTED, anchor="w", wraplength=480
        )
        self._status.pack(fill="x", pady=(4, 4))

        # ── Barra avanzamento (nascosta finché non serve) ───────────────────────
        self._progress = ctk.CTkProgressBar(
            body, fg_color=BORDER, progress_color=ACCENT,
            height=6, corner_radius=3
        )
        self._progress.set(0)

        # ── Bottone Elabora ────────────────────────────────────────────────────
        self._run_btn = ctk.CTkButton(
            body, text="Elabora  →",
            font=self.F_BTN, height=58, corner_radius=12,
            fg_color=ACCENT, hover_color=ACCENT_H,
            text_color="#fff", command=self._run,
            state="disabled"
        )
        self._run_btn.pack(fill="x", pady=(8, 4))

    # ── Widget helpers ─────────────────────────────────────────────────────────
    def _label(self, parent, text):
        ctk.CTkLabel(parent, text=text, font=self.F_LABEL,
                     text_color=TEXT, anchor="w"
                     ).pack(fill="x", pady=(12, 2))

    def _get_columns(self) -> list[str]:
        """Legge la textarea e restituisce la lista di nomi colonna non vuoti."""
        raw = self._col_box.get("1.0", "end")
        return [line.strip() for line in raw.splitlines() if line.strip()]

    # ── Drag & drop ────────────────────────────────────────────────────────────
    def _on_drag_enter(self, _event):
        self._drop_frame.configure(fg_color=DROP_ACT, border_color=ACCENT)

    def _on_drag_leave(self, _event):
        self._drop_frame.configure(fg_color=DROP_BG, border_color=BORDER)

    def _on_drop(self, event):
        self._drop_frame.configure(fg_color=DROP_BG, border_color=BORDER)
        paths = _parse_paths(event.data)
        if not paths:
            return

        path = paths[0]
        if not path.lower().endswith((".xlsx", ".xls")):
            self._set_status("⚠  File non supportato: trascina un file .xlsx", ERROR)
            self._run_btn.configure(state="disabled")
            return

        self._file_path = path
        name = os.path.basename(path)
        self._drop_label.configure(text=f"✓  {name}", text_color=SUCCESS)
        self._set_status("File pronto. Premi «Elabora» per generare il TXT.", MUTED)
        self._run_btn.configure(state="normal")

    # ── Stato ──────────────────────────────────────────────────────────────────
    def _set_status(self, msg, color=MUTED):
        self._status.configure(text=msg, text_color=color)

    # ── Esecuzione ─────────────────────────────────────────────────────────────
    def _run(self):
        if self._running or not self._file_path:
            return

        columns = self._get_columns()
        if not columns:
            self._set_status("⚠  Inserisci almeno un nome colonna.", ERROR)
            return

        self._running = True
        self._run_btn.configure(state="disabled", text="⏳  Un momento…")
        self._progress.pack(fill="x", pady=(0, 6))
        self._progress.start()
        self._set_status("Elaborazione in corso…", MUTED)

        threading.Thread(
            target=self._worker,
            args=(self._file_path, columns),
            daemon=True
        ).start()

    def _worker(self, path, columns):
        try:
            content = process_excel(path, columns)
            out = save_output(path, content)
            self.after(0, self._on_success, out)
        except Exception as exc:
            self.after(0, self._on_error, str(exc))

    def _on_success(self, out_path):
        self._progress.stop()
        self._progress.set(1)
        size_kb = os.path.getsize(out_path) / 1024
        self._set_status(f"✓  Salvato: {out_path}  ({size_kb:.1f} KB)", SUCCESS)
        self._run_btn.configure(state="normal", text="Elabora  →")
        self._running = False

    def _on_error(self, msg):
        self._progress.stop()
        self._progress.set(0)
        self._set_status(f"✗  Errore: {msg}", ERROR)
        self._run_btn.configure(state="normal", text="Elabora  →")
        self._running = False


# ── factory ────────────────────────────────────────────────────────────────────
def build_root() -> TkinterDnD.Tk:
    return ExcelToTxtGUI()
