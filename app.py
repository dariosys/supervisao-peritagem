#!/usr/bin/env python3
"""
App de Preenchimento - Ficha de Supervisão de Peritagem
Empresa: Tranquilidade / Açoreana / Logo

Uso: python supervisao_app.py
Requisitos: pip install pypdf pandas openpyxl pillow reportlab
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, create_string_object
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.pagesizes import A4
from PIL import Image, ImageTk
import io
import os
import sys
import copy

# ── Paleta de cores ────────────────────────────────────────────────────────
BG         = "#F5F6FA"
CARD_BG    = "#FFFFFF"
ACCENT     = "#C8102E"        # vermelho Generali
ACCENT2    = "#1A1A2E"        # azul escuro
TEXT       = "#1A1A2E"
TEXT_LIGHT = "#6B7280"
BORDER     = "#E5E7EB"
SUCCESS    = "#059669"
HOVER      = "#F3F4F6"

FONT_TITLE  = ("Georgia", 18, "bold")
FONT_LABEL  = ("Helvetica", 10)
FONT_SMALL  = ("Helvetica", 9)
FONT_BTN    = ("Helvetica", 10, "bold")
FONT_HEADER = ("Helvetica", 11, "bold")

# ── Mapeamento campos Excel → PDF ──────────────────────────────────────────
FIELD_MAP = {
    "F[0].Page_1[0].TextField9[1]" : "Nº sinistro",
    "F[0].Page_1[0].TextField9[4]" : "Data/ Hora da Visita",
    "F[0].Page_1[0].TextField9[2]" : "Matrícula",
    "F[0].Page_1[0].TextField9[0]" : "Perito: Nome /Código",
    "F[0].Page_1[0].TextField9[6]" : "Nome da oficina",
    "F[0].Page_1[0].TextField9[7]" : "Nº prest ofic",
    "F[0].Page_1[0].TextField12[0]": "Observações",
    "F[0].Page_1[0].TextField12[1]": "Avalie o serviço do perito",
    "F[0].Page_1[0].TextField12[2]": "Avalie o serviço da oficina",
    "F[0].Page_1[0].TextField9[3]" : "supervisor",
    "F[0].Page_1[0].DateTimeField1[0]": "Data",
}

# Coordenadas PDF das 3 células de suporte fotográfico (x0, y0_bottom, x1, y1_top)
PHOTO_RECTS = [
    (29.602,  129.542, 206.331, 259.189),
    (209.165, 129.542, 385.894, 259.189),
    (388.728, 129.542, 565.457, 259.189),
]


# ── Utilitários PDF ────────────────────────────────────────────────────────

def create_image_overlay(page_width, page_height, photos):
    """Cria uma página PDF transparente com as fotos nas células corretas."""
    packet = io.BytesIO()
    c = rl_canvas.Canvas(packet, pagesize=(page_width, page_height))
    for i, (img_path, rect) in enumerate(zip(photos, PHOTO_RECTS)):
        if not img_path:
            continue
        x0, y0, x1, y1 = rect
        w = x1 - x0
        h = y1 - y0
        padding = 4
        try:
            img = Image.open(img_path)
            img.thumbnail((int(w - padding*2), int(h - padding*2)), Image.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            # centre the image in the cell
            iw, ih = img.size
            ix = x0 + padding + (w - padding*2 - iw) / 2
            iy = y0 + padding + (h - padding*2 - ih) / 2
            c.drawInlineImage(img, ix, iy, iw, ih)
        except Exception as e:
            print(f"Erro ao carregar foto {i+1}: {e}")
    c.save()
    packet.seek(0)
    return packet


def fill_pdf(template_path, row, photos, output_path):
    reader = PdfReader(template_path)
    writer = PdfWriter()
    writer.append(reader)
    writer.set_need_appearances_writer(True)

    # Preencher campos de texto
    values = {}
    for field_id, col in FIELD_MAP.items():
        val = row.get(col, "")
        if pd.isna(val):
            val = ""
        if hasattr(val, "strftime"):
            val = val.strftime("%d/%m/%Y")
        values[field_id] = str(val).strip()

    writer.update_page_form_field_values(writer.pages[0], values)

    # Guardar PDF intermédio (necessário para merge das fotos)
    intermediate = io.BytesIO()
    writer.write(intermediate)
    intermediate.seek(0)

    # Criar overlay com as fotos
    page = reader.pages[0]
    pw = float(page.mediabox.width)
    ph = float(page.mediabox.height)

    active_photos = [p for p in photos if p]
    if active_photos:
        overlay_buf = create_image_overlay(pw, ph, photos)
        overlay_reader = PdfReader(overlay_buf)
        overlay_page = overlay_reader.pages[0]

        # Merge overlay → página preenchida
        final_reader = PdfReader(intermediate)
        final_writer = PdfWriter()
        base_page = final_reader.pages[0]
        base_page.merge_page(overlay_page)
        final_writer.add_page(base_page)
        with open(output_path, "wb") as f:
            final_writer.write(f)
    else:
        with open(output_path, "wb") as f:
            f.write(intermediate.getvalue())


# ── Interface gráfica ──────────────────────────────────────────────────────

class SupervisaoApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Ficha de Supervisão de Peritagem")
        self.geometry("860x700")
        self.minsize(800, 640)
        self.configure(bg=BG)
        self.resizable(True, True)

        self.excel_path   = tk.StringVar()
        self.pdf_path     = tk.StringVar()
        self.output_dir   = tk.StringVar()
        self.df           = None
        self.selected_row = None
        self.photos       = [None, None, None]
        self.photo_labels = []

        self._build_ui()

    # ── Layout principal ───────────────────────────────────────────────────

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=ACCENT2, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Ficha de Supervisão de Peritagem",
                 font=FONT_TITLE, bg=ACCENT2, fg="white").pack(side="left", padx=24)
        tk.Label(hdr, text="Tranquilidade · Açoreana · Logo",
                 font=FONT_SMALL, bg=ACCENT2, fg="#9CA3AF").pack(side="right", padx=24)

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=16)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)

        # Coluna esquerda — ficheiros + sinistros
        left = tk.Frame(body, bg=BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self._section(left, "1 · Ficheiros").pack(fill="x", pady=(0, 12))
        self._file_row(left, "Excel (mapa):",   self.excel_path, self._load_excel, "xlsx").pack(fill="x", pady=3)
        self._file_row(left, "PDF (template):", self.pdf_path,   self._pick_pdf,   "pdf" ).pack(fill="x", pady=3)
        self._file_row(left, "Pasta de saída:", self.output_dir, self._pick_outdir, "dir").pack(fill="x", pady=3)

        self._section(left, "2 · Selecionar sinistro").pack(fill="x", pady=(12, 6))

        # search box
        sf = tk.Frame(left, bg=BG)
        sf.pack(fill="x", pady=(0, 4))
        tk.Label(sf, text="Pesquisar:", font=FONT_SMALL, bg=BG, fg=TEXT_LIGHT).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace("w", lambda *_: self._filter_list())
        tk.Entry(sf, textvariable=self.search_var, font=FONT_LABEL,
                 relief="flat", bd=1, highlightthickness=1,
                 highlightbackground=BORDER, highlightcolor=ACCENT).pack(side="left", fill="x", expand=True, padx=(6, 0))

        # listbox
        lf = tk.Frame(left, bg=BORDER, bd=1)
        lf.pack(fill="both", expand=True)
        self.listbox = tk.Listbox(lf, font=FONT_LABEL, bg=CARD_BG, fg=TEXT,
                                  selectbackground=ACCENT, selectforeground="white",
                                  relief="flat", activestyle="none",
                                  highlightthickness=0, borderwidth=0)
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.listbox.pack(side="left", fill="both", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

        self.info_label = tk.Label(left, text="Carregue um Excel para ver os sinistros.",
                                   font=FONT_SMALL, bg=BG, fg=TEXT_LIGHT, anchor="w")
        self.info_label.pack(fill="x", pady=(4, 0))

        # Coluna direita — pré-visualização + fotos
        right = tk.Frame(body, bg=BG)
        right.grid(row=0, column=1, sticky="nsew")

        self._section(right, "3 · Dados do sinistro").pack(fill="x", pady=(0, 8))
        self.preview_frame = tk.Frame(right, bg=CARD_BG, relief="flat",
                                      highlightthickness=1, highlightbackground=BORDER)
        self.preview_frame.pack(fill="x")
        self.preview_labels = {}
        preview_fields = [
            ("Nº Sinistro",   "Nº sinistro"),
            ("Data",          "Data/ Hora da Visita"),
            ("Matrícula",     "Matrícula"),
            ("Perito",        "Perito: Nome /Código"),
            ("Oficina",       "Nome da oficina"),
            ("Supervisor",    "supervisor"),
        ]
        for i, (lbl, col) in enumerate(preview_fields):
            row_f = tk.Frame(self.preview_frame, bg=CARD_BG if i % 2 == 0 else HOVER)
            row_f.pack(fill="x")
            tk.Label(row_f, text=lbl, font=FONT_SMALL, bg=row_f["bg"],
                     fg=TEXT_LIGHT, width=14, anchor="w").pack(side="left", padx=8, pady=4)
            var = tk.StringVar(value="—")
            self.preview_labels[col] = var
            tk.Label(row_f, textvariable=var, font=FONT_LABEL, bg=row_f["bg"],
                     fg=TEXT, anchor="w").pack(side="left", fill="x", expand=True, padx=4)

        self._section(right, "4 · Suporte fotográfico (máx. 3 fotos)").pack(fill="x", pady=(14, 6))
        photos_frame = tk.Frame(right, bg=BG)
        photos_frame.pack(fill="x")
        self.photo_labels = []
        for i in range(3):
            pf = tk.Frame(photos_frame, bg=CARD_BG, width=175, height=110,
                          relief="flat", highlightthickness=1, highlightbackground=BORDER,
                          cursor="hand2")
            pf.pack(side="left", padx=(0 if i == 0 else 8, 0))
            pf.pack_propagate(False)

            inner = tk.Frame(pf, bg=CARD_BG)
            inner.place(relx=0.5, rely=0.5, anchor="center")

            icon = tk.Label(inner, text="📷", font=("Helvetica", 22), bg=CARD_BG, fg=TEXT_LIGHT)
            icon.pack()
            lbl_txt = tk.Label(inner, text=f"Foto {i+1}", font=FONT_SMALL, bg=CARD_BG, fg=TEXT_LIGHT)
            lbl_txt.pack()

            self.photo_labels.append({"frame": pf, "icon": icon, "label": lbl_txt, "img_ref": None})

            idx = i
            pf.bind("<Button-1>", lambda e, n=idx: self._pick_photo(n))
            icon.bind("<Button-1>", lambda e, n=idx: self._pick_photo(n))
            lbl_txt.bind("<Button-1>", lambda e, n=idx: self._pick_photo(n))

            # botão remover
            rm = tk.Label(pf, text="✕", font=FONT_SMALL, bg=CARD_BG, fg=TEXT_LIGHT, cursor="hand2")
            rm.place(relx=1.0, rely=0.0, anchor="ne", x=-4, y=4)
            rm.bind("<Button-1>", lambda e, n=idx: self._remove_photo(n))

        # Botão gerar
        btn_frame = tk.Frame(self, bg=BG)
        btn_frame.pack(fill="x", padx=20, pady=(4, 16))

        self.generate_btn = tk.Button(
            btn_frame, text="⬇  Gerar PDF Preenchido",
            font=FONT_BTN, bg=ACCENT, fg="white", activebackground="#A00D24",
            activeforeground="white", relief="flat", padx=24, pady=10,
            cursor="hand2", command=self._generate
        )
        self.generate_btn.pack(side="right")

        self.status_label = tk.Label(btn_frame, text="", font=FONT_SMALL,
                                     bg=BG, fg=TEXT_LIGHT, anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True)

    # ── Helpers UI ─────────────────────────────────────────────────────────

    def _section(self, parent, title):
        f = tk.Frame(parent, bg=BG)
        tk.Label(f, text=title.upper(), font=("Helvetica", 9, "bold"),
                 bg=BG, fg=ACCENT).pack(side="left")
        tk.Frame(f, bg=BORDER, height=1).pack(side="left", fill="x", expand=True, padx=(8, 0), pady=6)
        return f

    def _file_row(self, parent, label, var, cmd, kind):
        f = tk.Frame(parent, bg=BG)
        tk.Label(f, text=label, font=FONT_SMALL, bg=BG, fg=TEXT, width=14, anchor="w").pack(side="left")
        e = tk.Entry(f, textvariable=var, font=FONT_SMALL, state="readonly",
                     relief="flat", bd=0, readonlybackground=CARD_BG,
                     fg=TEXT, highlightthickness=1,
                     highlightbackground=BORDER, highlightcolor=ACCENT)
        e.pack(side="left", fill="x", expand=True, ipady=4, padx=(0, 6))
        btn = tk.Button(f, text="...", font=FONT_SMALL, command=cmd,
                        bg=ACCENT2, fg="white", activebackground=ACCENT,
                        activeforeground="white", relief="flat", padx=8, pady=2, cursor="hand2")
        btn.pack(side="left")
        return f

    # ── Ações de ficheiros ─────────────────────────────────────────────────

    def _load_excel(self):
        path = filedialog.askopenfilename(
            title="Escolha o ficheiro Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]
        )
        if not path:
            return
        try:
            self.df = pd.read_excel(path)
            self.excel_path.set(path)
            self._populate_list()
            self._set_status(f"Excel carregado: {len(self.df)} sinistro(s)", SUCCESS)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o Excel:\n{e}")

    def _pick_pdf(self):
        path = filedialog.askopenfilename(
            title="Escolha o PDF template",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")]
        )
        if path:
            self.pdf_path.set(path)

    def _pick_outdir(self):
        path = filedialog.askdirectory(title="Escolha a pasta de destino")
        if path:
            self.output_dir.set(path)

    # ── Lista de sinistros ─────────────────────────────────────────────────

    def _populate_list(self):
        self.all_items = []
        for i, row in self.df.iterrows():
            sinistro = str(row.get("Nº sinistro", "")).strip()
            oficina  = str(row.get("Nome da oficina", "")).strip()
            mat      = str(row.get("Matrícula", "")).strip()
            label    = f"{sinistro}  ·  {mat}  ·  {oficina[:30]}"
            self.all_items.append((label, i))
        self._filter_list()

    def _filter_list(self):
        if self.df is None:
            return
        q = self.search_var.get().lower()
        self.listbox.delete(0, "end")
        self.filtered_indices = []
        for label, idx in self.all_items:
            if q in label.lower():
                self.listbox.insert("end", "  " + label)
                self.filtered_indices.append(idx)
        self.info_label.config(text=f"{self.listbox.size()} sinistro(s) encontrado(s).")

    def _on_select(self, event):
        sel = self.listbox.curselection()
        if not sel:
            return
        row_idx = self.filtered_indices[sel[0]]
        self.selected_row = self.df.iloc[row_idx]
        # Actualizar pré-visualização
        for col, var in self.preview_labels.items():
            val = self.selected_row.get(col, "")
            if pd.isna(val):
                val = "—"
            elif hasattr(val, "strftime"):
                val = val.strftime("%d/%m/%Y")
            var.set(str(val).strip() or "—")

    # ── Fotos ──────────────────────────────────────────────────────────────

    def _pick_photo(self, idx):
        path = filedialog.askopenfilename(
            title=f"Escolha a Foto {idx+1}",
            filetypes=[("Imagens", "*.jpg *.jpeg *.png *.bmp *.tiff *.webp"), ("Todos", "*.*")]
        )
        if not path:
            return
        self.photos[idx] = path
        self._refresh_photo(idx)

    def _remove_photo(self, idx):
        self.photos[idx] = None
        self._refresh_photo(idx)

    def _refresh_photo(self, idx):
        info = self.photo_labels[idx]
        if self.photos[idx]:
            try:
                img = Image.open(self.photos[idx])
                img.thumbnail((160, 95), Image.LANCZOS)
                tk_img = ImageTk.PhotoImage(img)
                info["icon"].configure(image=tk_img, text="")
                info["icon"].image = tk_img
                fname = os.path.basename(self.photos[idx])
                info["label"].configure(text=fname[:22], fg=TEXT)
            except Exception as e:
                info["icon"].configure(image="", text="⚠", fg=ACCENT)
                info["label"].configure(text="Erro na imagem", fg=ACCENT)
        else:
            info["icon"].configure(image="", text="📷", fg=TEXT_LIGHT)
            info["label"].configure(text=f"Foto {idx+1}", fg=TEXT_LIGHT)

    # ── Gerar PDF ──────────────────────────────────────────────────────────

    def _generate(self):
        if self.selected_row is None:
            messagebox.showwarning("Atenção", "Selecione um sinistro primeiro.")
            return
        if not self.pdf_path.get():
            messagebox.showwarning("Atenção", "Escolha o ficheiro PDF template.")
            return
        if not self.output_dir.get():
            messagebox.showwarning("Atenção", "Escolha a pasta de destino.")
            return

        sinistro = str(self.selected_row.get("Nº sinistro", "sinistro")).strip()
        mat      = str(self.selected_row.get("Matrícula", "")).strip().replace("-", "")
        out_name = f"supervisao_{sinistro}_{mat}.pdf"
        out_path = os.path.join(self.output_dir.get(), out_name)

        self._set_status("A gerar PDF...", TEXT_LIGHT)
        self.update_idletasks()

        try:
            fill_pdf(self.pdf_path.get(), self.selected_row, self.photos, out_path)
            self._set_status(f"✓ PDF gerado: {out_name}", SUCCESS)
            if messagebox.askyesno("Sucesso", f"PDF gerado com sucesso!\n\n{out_path}\n\nAbrir a pasta?"):
                self._open_folder(self.output_dir.get())
        except Exception as e:
            self._set_status(f"Erro: {e}", ACCENT)
            messagebox.showerror("Erro", f"Não foi possível gerar o PDF:\n{e}")

    def _set_status(self, msg, color=TEXT_LIGHT):
        self.status_label.config(text=msg, fg=color)

    def _open_folder(self, path):
        import subprocess, platform
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])


# ── Estilo ttk ────────────────────────────────────────────────────────────

def apply_style():
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Vertical.TScrollbar",
                    background=BG, troughcolor=BG, arrowcolor=TEXT_LIGHT,
                    borderwidth=0, relief="flat")


# ── Entrada ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = SupervisaoApp()
    apply_style()
    app.mainloop()
