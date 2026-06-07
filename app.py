"""
Ficha de Supervisão de Peritagem — App Web
Tranquilidade · Açoreana · Logo

Deploy gratuito em: https://streamlit.io/cloud
"""

import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rl_canvas
from PIL import Image
import io

# ── Config da página ────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Ficha de Supervisão de Peritagem",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS personalizado ───────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&family=DM+Serif+Display&display=swap');

  html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

  .header-bar {
    background: #1A1A2E;
    border-left: 6px solid #C8102E;
    padding: 18px 28px;
    border-radius: 8px;
    margin-bottom: 24px;
  }
  .header-bar h1 {
    font-family: 'DM Serif Display', serif;
    color: white;
    font-size: 1.7rem;
    margin: 0;
    padding: 0;
  }
  .header-bar p {
    color: #9CA3AF;
    font-size: 0.82rem;
    margin: 4px 0 0 0;
  }

  .step-label {
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    color: #C8102E;
    text-transform: uppercase;
    margin-bottom: 4px;
  }

  .info-card {
    background: #F9FAFB;
    border: 1px solid #E5E7EB;
    border-radius: 8px;
    padding: 14px 18px;
  }
  .info-row {
    display: flex;
    justify-content: space-between;
    padding: 5px 0;
    border-bottom: 1px solid #F3F4F6;
    font-size: 0.88rem;
  }
  .info-row:last-child { border-bottom: none; }
  .info-key { color: #6B7280; min-width: 110px; }
  .info-val { color: #1A1A2E; font-weight: 500; text-align: right; }

  .photo-placeholder {
    border: 2px dashed #D1D5DB;
    border-radius: 8px;
    background: #F9FAFB;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 20px 10px;
    text-align: center;
    color: #9CA3AF;
    font-size: 0.82rem;
    min-height: 120px;
  }

  div[data-testid="stButton"] > button {
    background-color: #C8102E !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
    padding: 0.5rem 2rem !important;
    font-size: 1rem !important;
    width: 100%;
  }
  div[data-testid="stButton"] > button:hover {
    background-color: #A00D24 !important;
  }

  .stDataFrame { border: 1px solid #E5E7EB; border-radius: 8px; }

  section[data-testid="stSidebar"] { display: none; }
  #MainMenu { visibility: hidden; }
  footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Mapeamento campos Excel → PDF ───────────────────────────────────────────
FIELD_MAP = {
    "F[0].Page_1[0].TextField9[1]"    : "Nº sinistro",
    "F[0].Page_1[0].TextField9[4]"    : "Data/ Hora da Visita",
    "F[0].Page_1[0].TextField9[2]"    : "Matrícula",
    "F[0].Page_1[0].TextField9[0]"    : "Perito: Nome /Código",
    "F[0].Page_1[0].TextField9[6]"    : "Nome da oficina",
    "F[0].Page_1[0].TextField9[7]"    : "Nº prest ofic",
    "F[0].Page_1[0].TextField12[0]"   : "Observações",
    "F[0].Page_1[0].TextField12[1]"   : "Avalie o serviço do perito",
    "F[0].Page_1[0].TextField12[2]"   : "Avalie o serviço da oficina",
    "F[0].Page_1[0].TextField9[3]"    : "supervisor",
    "F[0].Page_1[0].DateTimeField1[0]": "Data",
}

PHOTO_RECTS = [
    (29.602,  129.542, 206.331, 259.189),
    (209.165, 129.542, 385.894, 259.189),
    (388.728, 129.542, 565.457, 259.189),
]

PREVIEW_FIELDS = [
    ("Nº Sinistro",  "Nº sinistro"),
    ("Data Visita",  "Data/ Hora da Visita"),
    ("Matrícula",    "Matrícula"),
    ("Perito",       "Perito: Nome /Código"),
    ("Oficina",      "Nome da oficina"),
    ("Cód. Oficina", "Nº prest ofic"),
    ("Supervisor",   "supervisor"),
]

# ── Funções PDF ─────────────────────────────────────────────────────────────

def create_image_overlay(pw, ph, photo_files):
    packet = io.BytesIO()
    c = rl_canvas.Canvas(packet, pagesize=(pw, ph))
    for photo_file, rect in zip(photo_files, PHOTO_RECTS):
        if photo_file is None:
            continue
        x0, y0, x1, y1 = rect
        w, h = x1 - x0, y1 - y0
        pad = 4
        try:
            img = Image.open(photo_file)
            img.thumbnail((int(w - pad*2), int(h - pad*2)), Image.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            iw, ih = img.size
            ix = x0 + pad + (w - pad*2 - iw) / 2
            iy = y0 + pad + (h - pad*2 - ih) / 2
            c.drawInlineImage(img, ix, iy, iw, ih)
        except Exception as e:
            st.warning(f"Erro ao processar foto: {e}")
    c.save()
    packet.seek(0)
    return packet


def fill_pdf_bytes(template_bytes, row, photo_files):
    template_buf = io.BytesIO(template_bytes)
    reader = PdfReader(template_buf)
    writer = PdfWriter()
    writer.append(reader)
    writer.set_need_appearances_writer(True)

    values = {}
    for field_id, col in FIELD_MAP.items():
        val = row.get(col, "")
        if pd.isna(val):
            val = ""
        if hasattr(val, "strftime"):
            val = val.strftime("%d/%m/%Y")
        values[field_id] = str(val).strip()

    writer.update_page_form_field_values(writer.pages[0], values)

    intermediate = io.BytesIO()
    writer.write(intermediate)
    intermediate.seek(0)

    page = reader.pages[0]
    pw = float(page.mediabox.width)
    ph = float(page.mediabox.height)

    active = [p for p in photo_files if p is not None]
    if active:
        overlay_buf = create_image_overlay(pw, ph, photo_files)
        overlay_page = PdfReader(overlay_buf).pages[0]
        final_reader = PdfReader(intermediate)
        final_writer = PdfWriter()
        base_page = final_reader.pages[0]
        base_page.merge_page(overlay_page)
        final_writer.add_page(base_page)
        output = io.BytesIO()
        final_writer.write(output)
        return output.getvalue()
    else:
        return intermediate.getvalue()


# ── Header ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-bar">
  <h1>📋 Ficha de Supervisão de Peritagem</h1>
  <p>Tranquilidade · Açoreana · Logo &nbsp;·&nbsp; Preencha e descarregue o PDF automaticamente</p>
</div>
""", unsafe_allow_html=True)

# ── Layout principal ────────────────────────────────────────────────────────
col_left, col_right = st.columns([1.1, 1], gap="large")

with col_left:
    # ── PASSO 1: Excel ──────────────────────────────────────────────────────
    st.markdown('<p class="step-label">1 · Carregar Excel (mapa do dia)</p>', unsafe_allow_html=True)
    excel_file = st.file_uploader(
        "Ficheiro Excel", type=["xlsx", "xls"],
        label_visibility="collapsed", key="excel"
    )

    df = None
    selected_row = None

    if excel_file:
        try:
            df = pd.read_excel(excel_file)
            st.success(f"✓ {len(df)} sinistro(s) carregado(s)")
        except Exception as e:
            st.error(f"Erro ao ler Excel: {e}")

    # ── PASSO 2: Template PDF ───────────────────────────────────────────────
    st.markdown('<p class="step-label" style="margin-top:20px">2 · Template PDF</p>', unsafe_allow_html=True)
    pdf_file = st.file_uploader(
        "PDF template", type=["pdf"],
        label_visibility="collapsed", key="pdf"
    )
    if pdf_file:
        st.success("✓ Template PDF carregado")

    # ── PASSO 3: Selecionar sinistro ────────────────────────────────────────
    if df is not None:
        st.markdown('<p class="step-label" style="margin-top:20px">3 · Selecionar sinistro</p>', unsafe_allow_html=True)

        search = st.text_input("🔍 Pesquisar por nº sinistro, matrícula ou oficina",
                               placeholder="ex: 28765768 ou AH-79-JN",
                               label_visibility="collapsed")

        # Filtrar
        mask = pd.Series([True] * len(df))
        if search.strip():
            q = search.strip().lower()
            mask = df.apply(
                lambda r: any(q in str(r.get(c, "")).lower()
                              for c in ["Nº sinistro", "Matrícula", "Nome da oficina"]),
                axis=1
            )

        filtered_df = df[mask].reset_index(drop=False)

        if len(filtered_df) == 0:
            st.info("Nenhum sinistro encontrado.")
        else:
            # Criar opções para o selectbox
            options = []
            for _, r in filtered_df.iterrows():
                sin = str(r.get("Nº sinistro", "")).strip()
                mat = str(r.get("Matrícula", "")).strip()
                ofi = str(r.get("Nome da oficina", "")).strip()[:35]
                options.append(f"{sin}  ·  {mat}  ·  {ofi}")

            choice = st.selectbox(
                "Sinistro", options,
                label_visibility="collapsed"
            )

            chosen_idx = options.index(choice)
            orig_idx = filtered_df.iloc[chosen_idx]["index"]
            selected_row = df.iloc[orig_idx]

with col_right:
    # ── Pré-visualização dos dados ──────────────────────────────────────────
    st.markdown('<p class="step-label">Dados do sinistro selecionado</p>', unsafe_allow_html=True)

    if selected_row is not None:
        rows_html = ""
        for label, col in PREVIEW_FIELDS:
            val = selected_row.get(col, "")
            if pd.isna(val):
                val = "—"
            elif hasattr(val, "strftime"):
                val = val.strftime("%d/%m/%Y")
            else:
                val = str(val).strip() or "—"
            rows_html += f"""
            <div class="info-row">
              <span class="info-key">{label}</span>
              <span class="info-val">{val}</span>
            </div>"""
        st.markdown(f'<div class="info-card">{rows_html}</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="info-card" style="color:#9CA3AF;font-size:0.9rem;text-align:center;padding:32px">Selecione um sinistro à esquerda</div>', unsafe_allow_html=True)

    # ── PASSO 4: Fotografias ────────────────────────────────────────────────
    st.markdown('<p class="step-label" style="margin-top:24px">4 · Suporte fotográfico (até 3 fotos)</p>', unsafe_allow_html=True)

    photo_cols = st.columns(3)
    photo_files = [None, None, None]

    for i, pcol in enumerate(photo_cols):
        with pcol:
            uploaded = st.file_uploader(
                f"Foto {i+1}", type=["jpg", "jpeg", "png", "bmp", "webp"],
                key=f"photo_{i}", label_visibility="visible"
            )
            if uploaded:
                photo_files[i] = io.BytesIO(uploaded.read())
                try:
                    img = Image.open(photo_files[i])
                    photo_files[i].seek(0)
                    st.image(img, use_container_width=True)
                except:
                    pass

    # ── PASSO 5: Gerar PDF ──────────────────────────────────────────────────
    st.markdown('<p class="step-label" style="margin-top:24px">5 · Gerar PDF</p>', unsafe_allow_html=True)

    can_generate = (selected_row is not None) and (pdf_file is not None)

    if not can_generate:
        missing = []
        if pdf_file is None:
            missing.append("template PDF")
        if selected_row is None:
            missing.append("sinistro selecionado")
        st.info(f"Falta: {' e '.join(missing)}")

    if st.button("⬇ Gerar e Descarregar PDF", disabled=not can_generate):
        with st.spinner("A gerar PDF..."):
            try:
                pdf_file.seek(0)
                template_bytes = pdf_file.read()
                pdf_bytes = fill_pdf_bytes(template_bytes, selected_row, photo_files)

                sinistro = str(selected_row.get("Nº sinistro", "sinistro")).strip()
                mat      = str(selected_row.get("Matrícula", "")).strip().replace("-", "")
                filename = f"supervisao_{sinistro}_{mat}.pdf"

                st.download_button(
                    label="📥 Clique aqui para descarregar o PDF",
                    data=pdf_bytes,
                    file_name=filename,
                    mime="application/pdf",
                    use_container_width=True,
                )
                st.success(f"✓ PDF gerado: {filename}")
            except Exception as e:
                st.error(f"Erro ao gerar PDF: {e}")
