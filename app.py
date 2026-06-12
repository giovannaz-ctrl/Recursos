"""
Dashboard de Alocação de Consultores
Covers: US-001 (Dashboard), US-002 (Cockpit de Workshops), US-003 (Planejamento Semanal)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import re
from io import BytesIO

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Alocação de Consultores",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
# THEME / CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

  html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

  /* Sidebar */
  section[data-testid="stSidebar"] { background: #0f172a; }
  section[data-testid="stSidebar"] * { color: #cbd5e1 !important; }
  section[data-testid="stSidebar"] .stSelectbox label,
  section[data-testid="stSidebar"] .stMultiSelect label { color: #94a3b8 !important; font-size: 0.78rem; text-transform: uppercase; letter-spacing: .05em; }

  /* Main background */
  .main { background: #f8fafc; }

  /* KPI Cards */
  .kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; margin-bottom: 1.5rem; }
  .kpi-card {
    background: white;
    border-radius: 12px;
    padding: 1.2rem 1.4rem;
    border-left: 4px solid #6366f1;
    box-shadow: 0 1px 3px rgba(0,0,0,.06);
  }
  .kpi-card.green  { border-left-color: #10b981; }
  .kpi-card.amber  { border-left-color: #f59e0b; }
  .kpi-card.rose   { border-left-color: #f43f5e; }
  .kpi-value { font-size: 2rem; font-weight: 700; color: #1e293b; line-height: 1; }
  .kpi-label { font-size: 0.75rem; color: #64748b; margin-top: .25rem; text-transform: uppercase; letter-spacing: .05em; }

  /* Section titles */
  .section-title { font-size: 1rem; font-weight: 600; color: #1e293b; margin: 1.5rem 0 .75rem; border-bottom: 1px solid #e2e8f0; padding-bottom: .4rem; }

  /* Nav tabs override */
  div[data-testid="stTabs"] button { font-weight: 500; }

  /* Calendar cells */
  .cal-event {
    border-radius: 6px; padding: 4px 7px; font-size: .75rem;
    font-weight: 500; margin-bottom: 3px; line-height: 1.3;
    overflow: hidden; white-space: nowrap; text-overflow: ellipsis;
  }
  .conflict-badge {
    background: #fef2f2; color: #dc2626;
    border: 1px solid #fca5a5; border-radius: 4px;
    padding: 2px 6px; font-size: .7rem; font-weight: 600;
  }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
EXCEL_EPOCH = datetime(1899, 12, 30)

def excel_serial_to_date(serial):
    """Convert Excel serial number → Python date."""
    try:
        return EXCEL_EPOCH + timedelta(days=int(float(serial)))
    except Exception:
        return None

def excel_time_to_str(fraction):
    """Convert Excel time fraction → HH:MM string."""
    try:
        total_minutes = round(float(fraction) * 24 * 60)
        h, m = divmod(total_minutes, 60)
        return f"{h:02d}:{m:02d}"
    except Exception:
        return ""

def excel_time_to_float(fraction):
    """Convert Excel time fraction → decimal hours."""
    try:
        return float(fraction) * 24
    except Exception:
        return 0.0

def extract_name(raw: str) -> str:
    """'João Silva <j@email.com>' → 'João Silva'."""
    if pd.isna(raw) or str(raw).strip() == "":
        return ""
    return re.sub(r"\s*<[^>]+>", "", str(raw)).strip()

def extract_email(raw: str) -> str:
    """'João Silva <j@email.com>' → 'j@email.com'."""
    m = re.search(r"<([^>]+)>", str(raw))
    return m.group(1).strip().lower() if m else ""

def split_consultants(raw: str) -> list[str]:
    """Split comma-separated consultant cells into individual names."""
    if pd.isna(raw) or str(raw).strip() in ("", "nan"):
        return []
    parts = str(raw).split(",")
    return [n for n in (extract_name(p) for p in parts) if n and n.lower() != "nan"]

def split_emails(raw: str) -> list[str]:
    """Split comma-separated consultant cells into individual emails."""
    if pd.isna(raw) or str(raw).strip() in ("", "nan"):
        return []
    parts = str(raw).split(",")
    return [e for e in (extract_email(p) for p in parts) if e]

def get_client_from_project(project: str) -> str:
    """Best-effort: extract client name from project string like '3568 - Brasal - ...'."""
    parts = project.split(" - ")
    return parts[1].strip() if len(parts) >= 2 else project

PROJECT_COLORS = [
    "#6366f1","#10b981","#f59e0b","#3b82f6","#ec4899",
    "#14b8a6","#f97316","#8b5cf6","#06b6d4","#84cc16",
    "#e11d48","#0ea5e9","#a855f7","#22c55e","#fb923c",
]

@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes):
    wb = pd.ExcelFile(BytesIO(file_bytes))
    sheets = {s.lower().replace(" ", ""): s for s in wb.sheet_names}

    # ── Sheet 1: Cockpit ──────────────────────────────────────────
    s1_key = next((k for k in sheets if "cockpit" in k or "1-" in k), None)
    df_cockpit_raw = pd.read_excel(BytesIO(file_bytes), sheet_name=sheets[s1_key] if s1_key else 0)
    df_cockpit_raw.columns = [c.strip() for c in df_cockpit_raw.columns]

    rows = []
    vagas_rows = []
    for _, row in df_cockpit_raw.iterrows():
        projeto = str(row.get("Projeto", "")).strip()
        perfil  = str(row.get("Perfil", "")).strip()
        # Support both old "Consultor" and new "Consultor principal"/"Consultor secundário"
        raw_prin = str(row.get("Consultor principal", row.get("Consultor", "")) or "")
        raw_secu = str(row.get("Consultor secundário", "") or "")
        if not projeto or projeto == "nan":
            continue
        client = get_client_from_project(projeto)
        golive_raw  = row.get("Go line") or row.get("Go Live") or row.get("Go live")
        golive      = pd.Timestamp(golive_raw) if golive_raw is not None and pd.notna(golive_raw) else None
        senior_raw  = row.get("Senioridade", None)
        senioridade = str(senior_raw).strip() if senior_raw is not None and pd.notna(senior_raw) else "Sênior"

        # Principal consultant(s)
        for part in (raw_prin.split(",") if raw_prin.strip() not in ("","nan") else []):
            name  = extract_name(part.strip())
            email = extract_email(part.strip())
            if not name or name.lower() == "nan": continue
            rows.append({"Projeto": projeto, "Módulo": perfil, "Consultor": name,
                         "Email": email, "Cliente": client, "GoLive": golive,
                         "Senioridade": senioridade, "Papel": "Principal"})

        # Secondary / shadow consultant(s)
        for part in (raw_secu.split(",") if raw_secu.strip() not in ("","nan") else []):
            name  = extract_name(part.strip())
            email = extract_email(part.strip())
            if not name or name.lower() == "nan": continue
            rows.append({"Projeto": projeto, "Módulo": perfil, "Consultor": name,
                         "Email": email, "Cliente": client, "GoLive": golive,
                         "Senioridade": senioridade, "Papel": "Sombra"})

        # Vaga: no principal assigned
        if not split_consultants(raw_prin) and perfil not in ("", "nan"):
            vagas_rows.append({"Projeto": projeto, "Perfil": perfil, "Cliente": client})

    df1 = pd.DataFrame(rows).drop_duplicates()

    # ── Go Live conflict table ────────────────────────────────────
    _gl_cols = ["Consultor","Email","Cliente","Projeto","GoLive"]
    if "Senioridade" in df1.columns:
        _gl_cols.append("Senioridade")
    df_golive = df1[df1["GoLive"].notna()][_gl_cols].drop_duplicates().copy()
    df_vagas = pd.DataFrame(vagas_rows).drop_duplicates()

    # ── Sheet 2: Agenda Workshop ──────────────────────────────────
    s2_key = next((k for k in sheets if "agenda" in k or "2-" in k or "workshop" in k.replace("3-","")), None)
    df_ws_raw = pd.read_excel(BytesIO(file_bytes), sheet_name=sheets[s2_key] if s2_key else 1)
    df_ws_raw.columns = [c.strip() for c in df_ws_raw.columns]

    def _to_ts(val):
        if val is None: return None
        try:
            if pd.isna(val): return None
        except Exception: pass
        if isinstance(val, pd.Timestamp): return val
        if isinstance(val, datetime): return pd.Timestamp(val)
        import datetime as _dt
        if isinstance(val, _dt.date): return pd.Timestamp(val)
        try: return pd.Timestamp(EXCEL_EPOCH + timedelta(days=int(float(val))))
        except Exception: return None

    ws_rows = []
    for _, row in df_ws_raw.iterrows():
        projeto  = str(row.get("Projeto", "")).strip()
        workshop = str(row.get("Workshop", "")).strip()
        raw_con  = str(row.get("Consultor", ""))
        inicio  = row.get("Inicio", row.get("Início", None))
        termino = row.get("Término", row.get("Termino", None))
        if not projeto or projeto == "nan":
            continue
        dt_ini = _to_ts(inicio)
        dt_fim = _to_ts(termino)
        client = get_client_from_project(projeto)
        for part in (str(raw_con).split(",") if raw_con.strip() not in ("","nan") else [""]):
            name  = extract_name(part.strip()) or ""
            email = extract_email(part.strip())
            ws_rows.append({
                "Projeto": projeto, "Cliente": client, "Workshop": workshop,
                "Consultor": name, "Email": email,
                "DataInicio": dt_ini, "DataFim": dt_fim,
            })

    df2 = pd.DataFrame(ws_rows)

    # ── Sheet 3: Atividades da Semana ─────────────────────────────
    s3_key = next((k for k in sheets if "atividade" in k or "3-" in k), None)
    df_at_raw = pd.read_excel(BytesIO(file_bytes), sheet_name=sheets[s3_key] if s3_key else 2)
    df_at_raw.columns = [c.strip() for c in df_at_raw.columns]

    # Email column in sheet 3 (usually 'Unnamed: 1')
    email_col_3 = next((c for c in df_at_raw.columns if "unnamed" in c.lower() or "email" in c.lower()), None)

    at_rows = []
    for _, row in df_at_raw.iterrows():
        recurso   = str(row.get("Recurso", "")).strip()
        raw_rec   = str(row.get(email_col_3, "")) if email_col_3 else recurso
        projeto   = str(row.get("Projeto", "")).strip()
        fase      = str(row.get("Fase do Projeto", "")).strip()
        atividade = str(row.get("Atividade", "")).strip()
        data_raw  = row.get("Data")
        hi_raw    = row.get("Hora inicio") or row.get("Hora Inicio") or row.get("Hora Início")
        hf_raw    = row.get("Hora fim") or row.get("Hora Fim")

        # Extract proper name and email from 'Nome <email>' format
        email3 = extract_email(raw_rec) if raw_rec and "<" in raw_rec else ""
        name3  = extract_name(raw_rec)  if raw_rec and "<" in raw_rec else recurso.title()

        if not recurso or recurso == "nan":
            continue

        # Data: may be datetime64 or serial
        if pd.notna(data_raw):
            if isinstance(data_raw, (pd.Timestamp, datetime)):
                data = pd.Timestamp(data_raw)
            else:
                data = pd.Timestamp(excel_serial_to_date(data_raw))
        else:
            data = None

        def _time_to_float(t):
            """datetime.time or excel fraction → decimal hours."""
            if t is None or (isinstance(t, float) and np.isnan(t)):
                return None
            import datetime as _dt
            if isinstance(t, _dt.time):
                return t.hour + t.minute / 60
            try:
                return float(t) * 24
            except Exception:
                return None

        h_ini   = _time_to_float(hi_raw)
        h_fim   = _time_to_float(hf_raw)
        horas   = round(h_fim - h_ini, 2) if (h_ini is not None and h_fim is not None) else 0.0
        client  = get_client_from_project(projeto)
        semana  = f"W{data.isocalendar()[1]:02d}/{data.year}" if data else ""

        at_rows.append({
            "Consultor": name3, "Email": email3, "Projeto": projeto, "Cliente": client,
            "Fase": fase, "Atividade": atividade,
            "Data": data, "Semana": semana,
            "HorasIni": h_ini, "HorasFim": h_fim, "Horas": horas,
        })

    df3 = pd.DataFrame(at_rows)
    # ── Sheet 4: Recursos ────────────────────────────────────────
    s4_key = next((k for k in sheets if "recur" in k.lower() or "4" in k), None)
    df_rec_raw = pd.read_excel(BytesIO(file_bytes), sheet_name=sheets[s4_key] if s4_key else 3)
    df_rec_raw.columns = [str(c).strip() for c in df_rec_raw.columns]

    # Email column is "Unnamed: 1" when header row has no label
    email_col = next((c for c in df_rec_raw.columns if "unnamed" in c.lower() or "email" in c.lower()), None)
    modulos   = [c for c in df_rec_raw.columns if c not in ("Consultor", email_col or "")]

    rec_rows = []
    for _, row in df_rec_raw.iterrows():
        consultor = str(row.get("Consultor","")).strip()
        email_raw = str(row.get(email_col,"")).strip().lower() if email_col else ""
        if not consultor or consultor == "nan":
            continue
        specs = [m for m in modulos if str(row.get(m,"")).strip().lower() == "x"]
        senior_raw4  = row.get("Senioridade", None)
        senioridade4 = str(senior_raw4).strip() if senior_raw4 is not None and pd.notna(senior_raw4) else "Sênior"
        rec_rows.append({
            "Consultor":    consultor,
            "Email":        email_raw,
            "Especialidades": specs,
            "Modulos":      ", ".join(specs),
            "Senioridade":  senioridade4,
        })

    df_rec = pd.DataFrame(rec_rows)

    # ── Cross by EMAIL: who is allocated in the cockpit? ─────────
    allocated_emails = set()
    for _col in ["Consultor principal", "Consultor secundário", "Consultor"]:
        if _col in df_cockpit_raw.columns:
            for _raw in df_cockpit_raw[_col].dropna():
                for _em in split_emails(str(_raw)):
                    allocated_emails.add(_em)

    df_rec["Alocado"] = df_rec["Email"].apply(lambda e: e in allocated_emails)
    df_rec["Status"]  = df_rec["Alocado"].map({True: "Alocado", False: "Disponível"})

    return df1, df2, df3, df_vagas, df_rec, df_golive


def kpi_html(value, label, variant=""):
    return f"""
    <div class="kpi-card {variant}">
      <div class="kpi-value">{value}</div>
      <div class="kpi-label">{label}</div>
    </div>"""


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return out.getvalue()


# ─────────────────────────────────────────────
# SIDEBAR – file upload ou GitHub
# ─────────────────────────────────────────────
GITHUB_EXCEL_URL = "https://raw.githubusercontent.com/giovannaz-ctrl/Recursos/main/alocacao_controle_otimizado.xlsx"

with st.sidebar:
    st.markdown("## 📂 Fonte de Dados")
    uploaded = st.file_uploader("Carregar novo arquivo Excel", type=["xlsx","xls"],
                                 help="Deixe vazio para usar o arquivo padrão do repositório")
    st.markdown("---")

# ─────────────────────────────────────────────
# LOAD
# ─────────────────────────────────────────────
with st.spinner("Processando dados…"):
    if uploaded is not None:
        file_bytes = uploaded.read()
    else:
        import urllib.request
        try:
            with urllib.request.urlopen(GITHUB_EXCEL_URL) as resp:
                file_bytes = resp.read()
        except Exception as e:
            st.error(f"Não foi possível carregar o arquivo do repositório: {e}")
            st.stop()
    df1, df2, df3, df_vagas, df_rec, df_golive = load_data(file_bytes)

# Build project→color map (shared across tabs)
all_projects = sorted(set(df1["Projeto"]) | set(df2["Projeto"]) | set(df3["Projeto"]))
proj_color = {p: PROJECT_COLORS[i % len(PROJECT_COLORS)] for i, p in enumerate(all_projects)}


# ═══════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 Alocação de Consultores",
    "📅 Agenda de Workshops",
    "⏱️ Planejamento Semanal",
    "👥 Recursos",
    "⚠️ Go Lives",
])


# ───────────────────────────────────────────────────────────────
# TAB 1 – Dashboard de Alocação (US-001)
# ───────────────────────────────────────────────────────────────
with tab1:
    st.markdown("### Dashboard de Alocação de Consultores")

    # ── Filters ─────────────────────────────────────────────────
    with st.expander("🔍 Filtros", expanded=False):
        c1, c2, c3, c4 = st.columns(4)
        f_cons  = c1.multiselect("Consultor",  sorted(df1["Consultor"].dropna().unique()), key="t1_cons")
        f_proj  = c2.multiselect("Projeto",    sorted(df1["Projeto"].dropna().unique()),   key="t1_proj")
        f_cli   = c3.multiselect("Cliente",    sorted(df1["Cliente"].dropna().unique()),   key="t1_cli")
        f_mod   = c4.multiselect("Módulo",     sorted(df1["Módulo"].dropna().unique()),    key="t1_mod")

    dft = df1.copy()
    if f_cons: dft = dft[dft["Consultor"].isin(f_cons)]
    if f_proj: dft = dft[dft["Projeto"].isin(f_proj)]
    if f_cli:  dft = dft[dft["Cliente"].isin(f_cli)]
    if f_mod:  dft = dft[dft["Módulo"].isin(f_mod)]

    # Global max (unfiltered) keeps color scale stable across filters
    _global_max_proj = max(df1.groupby("Consultor")["Projeto"].nunique().max(), 2)

    # ── KPIs ─────────────────────────────────────────────────
    n_cons    = dft["Consultor"].nunique()
    n_proj    = dft["Projeto"].nunique()
    n_cli     = dft["Cliente"].nunique()
    n_mod     = dft["Módulo"].nunique()
    # Sobrealocação: consultores em > 1 projeto
    proj_per_cons = dft.groupby("Consultor")["Projeto"].nunique()
    n_sobre   = (proj_per_cons > 1).sum()

    st.markdown(f"""
    <div class="kpi-grid">
        {kpi_html(n_cons, "Consultores")}
        {kpi_html(n_proj, "Projetos", "green")}
        {kpi_html(n_sobre, "Multialocados", "rose")}
    </div>
    """, unsafe_allow_html=True)

    # ── Treemap ──────────────────────────────────────────────────
    st.markdown('<div class="section-title">Distribuição por Consultor (Treemap)</div>', unsafe_allow_html=True)

    # Include Papel in treemap path so sombra is visually distinct
    _tp_cols = ["Consultor","Projeto","Módulo"]
    if "Papel" in dft.columns:
        _tp_cols = ["Consultor","Papel","Projeto","Módulo"]
    treemap_df = (
        dft.groupby(["Consultor","Projeto"] + (["Papel"] if "Papel" in dft.columns else []))
           .agg(Atividades=("Módulo","count"))
           .reset_index()
    )
    if "Papel" in treemap_df.columns:
        # Label: append "(Sombra)" to project name for shadow rows
        treemap_df["ProjetoLabel"] = treemap_df.apply(
            lambda r: r["Projeto"] + (" 👥" if r["Papel"] == "Sombra" else " ⭐"),
            axis=1
        )
    else:
        treemap_df["ProjetoLabel"] = treemap_df["Projeto"]
    treemap_df["Projetos"] = treemap_df.groupby("Consultor")["Projeto"].transform("nunique")

    if not treemap_df.empty:
        fig_tree = px.treemap(
            treemap_df,
            path=["Consultor", "ProjetoLabel"],
            values="Atividades",
            color="Projetos",
            color_continuous_scale=[
                [0.0,  "#bfdbfe"],
                [0.35, "#60a5fa"],
                [0.65, "#2563eb"],
                [1.0,  "#1e3a8a"],
            ],
            range_color=[1, _global_max_proj],
            custom_data=["Projetos"],
        )
        fig_tree.update_traces(
            hovertemplate="<b>%{label}</b><br>Projetos: %{customdata[0]}<extra></extra>",
            textfont_size=13,
            marker_line_width=2,
            marker_line_color="white",
            opacity=0.9,
        )
        fig_tree.update_layout(
            margin=dict(l=0, r=0, t=10, b=0),
            height=420,
            coloraxis_colorbar=dict(
                title="Projetos",
                tickmode="linear",
                tick0=1,
                dtick=1,
                thickness=14,
                len=0.6,
            ),
        )
        st.plotly_chart(fig_tree, use_container_width=True)

        # Detail on click via selectbox
        sel_cons = st.selectbox("Selecione um consultor para detalhar:",
                                ["— todos —"] + sorted(dft["Consultor"].dropna().unique().tolist()),
                                key="t1_sel")
        if sel_cons != "— todos —":
            detail = dft[dft["Consultor"] == sel_cons]
            n_p = detail["Projeto"].nunique()
            n_m = detail["Módulo"].nunique()
            st.markdown(
                f"<div style='font-size:.85rem; color:#64748b; margin-bottom:.5rem;'>"
                f"<b style='color:#1e293b;'>{sel_cons}</b>"
                f" &nbsp;·&nbsp; {n_p} projeto{'s' if n_p>1 else ''}"
                f" &nbsp;·&nbsp; {n_m} módulo{'s' if n_m>1 else ''}"
                f"</div>",
                unsafe_allow_html=True,
            )
            st.dataframe(
                detail[["Projeto","Cliente","Módulo"]].drop_duplicates(),
                use_container_width=True, hide_index=True,
            )

    # ── Table ────────────────────────────────────────────────────
    st.markdown('<div class="section-title">Tabela Detalhada</div>', unsafe_allow_html=True)

    # Build display: pivot Principal + Sombra into same row
    # Use df1 (full unfiltered) as source for Sombra so filter on consultant doesn't hide shadows
    if "Papel" in dft.columns and "Papel" in df1.columns:
        _prin = (dft[dft["Papel"] == "Principal"]
                 [["Cliente","Projeto","Módulo","Consultor"]]
                 .rename(columns={"Consultor":"Consultor Principal"}))
        _somb = (df1[df1["Papel"] == "Sombra"]
                 [["Projeto","Módulo","Consultor"]]
                 .rename(columns={"Consultor":"Consultor Sombra"})
                 .drop_duplicates())
        display = _prin.merge(_somb, on=["Projeto","Módulo"], how="left")
        display = display[["Cliente","Projeto","Módulo","Consultor Principal","Consultor Sombra"]].drop_duplicates()
        display["Consultor Sombra"] = display["Consultor Sombra"].fillna("—")
    else:
        display = dft[["Consultor","Cliente","Projeto","Módulo"]].copy()

    st.dataframe(display, use_container_width=True, hide_index=True,
                 column_config={
                     "Cliente":              st.column_config.TextColumn("Cliente",              width="small"),
                     "Projeto":              st.column_config.TextColumn("Projeto",              width="large"),
                     "Módulo":               st.column_config.TextColumn("Módulo",               width="medium"),
                     "Consultor Principal":  st.column_config.TextColumn("Consultor Principal",  width="medium"),
                     "Consultor Sombra":     st.column_config.TextColumn("👥 Sombra",            width="medium"),
                 })

    st.download_button("⬇ Exportar Excel", to_excel_bytes(display),
                       file_name="alocacao_consultores.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── Vagas abertas (always-visible alert) ─────────────────────
    if not df_vagas.empty:
        n_vagas      = len(df_vagas)
        n_proj_vagas = df_vagas["Projeto"].nunique()

        proj_groups = df_vagas.groupby("Projeto")
        rows_html = ""
        for proj, grp in proj_groups:
            client = grp["Cliente"].iloc[0]
            perfis = grp["Perfil"].tolist()
            for i, perfil in enumerate(perfis):
                if i == 0:
                    proj_cell = (
                        f"<td rowspan='{len(perfis)}' style='vertical-align:top;"
                        f"padding:6px 12px; border-bottom:1px solid #fed7aa;'>"
                        f"<span style='font-weight:600;color:#1e293b;'>{proj}</span><br>"
                        f"<span style='font-size:.72rem;color:#92400e;'>{client}</span></td>"
                    )
                else:
                    proj_cell = ""
                rows_html += (
                    f"<tr>"
                    f"{proj_cell}"
                    f"<td style='padding:6px 12px;border-bottom:1px solid #fed7aa;color:#1e293b;'>{perfil}</td>"
                    f"<td style='padding:6px 12px;border-bottom:1px solid #fed7aa;'>"
                    f"<span style='background:#fee2e2;color:#dc2626;border-radius:4px;"
                    f"padding:2px 8px;font-size:.72rem;font-weight:600;'>Sem consultor</span>"
                    f"</td>"
                    f"</tr>"
                )

        st.markdown(f"""
        <div style="background:#fff7ed; border:1.5px solid #fb923c; border-radius:10px;
                    padding:1rem 1.2rem; margin-top:1.2rem;">
          <div style="display:flex; align-items:center; gap:.6rem; margin-bottom:.8rem;">
            <span style="font-size:1.3rem;">⚠️</span>
            <span style="font-weight:700; color:#c2410c; font-size:.95rem;">
              {n_vagas} vaga{'s' if n_vagas > 1 else ''} sem consultor alocado
              &nbsp;·&nbsp;
              {n_proj_vagas} projeto{'s' if n_proj_vagas > 1 else ''} afetado{'s' if n_proj_vagas > 1 else ''}
            </span>
          </div>
          <table style="width:100%; border-collapse:collapse; font-size:.83rem;">
            <thead>
              <tr style="background:#fed7aa;">
                <th style="padding:5px 12px; text-align:left; color:#7c2d12;
                           font-weight:600; border-radius:4px 0 0 0;">Projeto</th>
                <th style="padding:5px 12px; text-align:left; color:#7c2d12; font-weight:600;">Perfil / Módulo</th>
                <th style="padding:5px 12px; text-align:left; color:#7c2d12;
                           font-weight:600; border-radius:0 4px 0 0;">Status</th>
              </tr>
            </thead>
            <tbody>{rows_html}</tbody>
          </table>
        </div>
        """, unsafe_allow_html=True)


# ───────────────────────────────────────────────────────────────
# TAB 2 – Cockpit de Workshops (US-002)
# ───────────────────────────────────────────────────────────────
with tab2:
    st.markdown("### Agenda de Workshops")

    df2w = df2.copy()
    df2w = df2w[df2w["DataInicio"].notna()]

    if df2w.empty:
        st.warning("Nenhum workshop com data encontrado.")
    else:
        min_date = df2w["DataInicio"].min()
        max_date = df2w["DataFim"].fillna(df2w["DataInicio"]).max()

        # ── Nav: week selector ────────────────────────────────────
        all_mondays = pd.date_range(
            start=min_date - timedelta(days=min_date.weekday()),
            end=max_date,
            freq="W-MON",
        )
        if len(all_mondays) == 0:
            all_mondays = [min_date - timedelta(days=min_date.weekday())]

        week_labels = [f"{d.strftime('%d/%m/%Y')} – {(d+timedelta(6)).strftime('%d/%m/%Y')}" for d in all_mondays]

        # ── Filters ───────────────────────────────────────────────
        fw_cons = st.multiselect(
            "Filtrar por Consultor",
            sorted(df2w["Consultor"].dropna().unique()),
            key="t2_cons",
            placeholder="Todos os consultores…",
        )
        if fw_cons:
            df2w = df2w[df2w["Consultor"].isin(fw_cons)]

        # ── Week navigation – default to current week ────────────
        today = pd.Timestamp(datetime.today().date())
        current_monday = today - timedelta(days=today.weekday())
        # find index of current week (or nearest future)
        def _find_current_idx():
            for i, m in enumerate(all_mondays):
                if m >= current_monday:
                    return i
            return len(all_mondays) - 1

        if "t2_week_idx" not in st.session_state:
            st.session_state["t2_week_idx"] = _find_current_idx()

        nav1, nav2, nav3, nav4 = st.columns([1, 5, 1, 1])
        with nav1:
            if st.button("◀ Anterior", key="prev_w"):
                if st.session_state["t2_week_idx"] > 0:
                    st.session_state["t2_week_idx"] -= 1
                    st.rerun()
        with nav3:
            if st.button("Próxima ▶", key="next_w"):
                if st.session_state["t2_week_idx"] < len(all_mondays) - 1:
                    st.session_state["t2_week_idx"] += 1
                    st.rerun()
        with nav4:
            if st.button("Hoje", key="today_w"):
                st.session_state["t2_week_idx"] = _find_current_idx()
                st.rerun()

        sel_week_idx = max(0, min(st.session_state["t2_week_idx"], len(all_mondays) - 1))

        with nav2:
            label = week_labels[sel_week_idx]
            st.markdown(
                f"<div style='text-align:center; padding:.3rem 0; font-weight:400; "
                f"color:#94a3b8; font-size:.85rem;'>{label}</div>",
                unsafe_allow_html=True,
            )

        week_start = all_mondays[sel_week_idx]
        week_end   = week_start + timedelta(days=6)

        week_ws = df2w[
            (df2w["DataInicio"] <= week_end) &
            (df2w["DataFim"].fillna(df2w["DataInicio"]) >= week_start)
        ].copy()

        # ── Calendar grid ─────────────────────────────────────────
        st.markdown('<div class="section-title">Agenda da Semana</div>', unsafe_allow_html=True)

        days = [week_start + timedelta(d) for d in range(5)]
        day_names = ["Seg","Ter","Qua","Qui","Sex"]

        cols = st.columns(5)
        for col_idx, (day, dname) in enumerate(zip(days, day_names)):
            with cols[col_idx]:
                is_weekend = False
                bg = "white"
                st.markdown(f"""
                <div style="background:{bg}; border-radius:8px; padding:8px;
                     border: 1px solid #6366f1;
                     min-height:50px;">
                  <div style="font-weight:600; color:#6366f1;
                       font-size:.8rem; margin-bottom:6px;">
                    {dname} {day.strftime('%d/%m')}
                  </div>
                """, unsafe_allow_html=True)

                day_events = week_ws[week_ws["DataInicio"] <= day]
                day_events = day_events[day_events["DataFim"].fillna(day_events["DataInicio"]) >= day]

                if day_events.empty:
                    st.markdown(
                        "<div style='font-size:.68rem; color:#cbd5e1; padding:4px 0;'>—</div>",
                        unsafe_allow_html=True,
                    )
                for _, ev in day_events.iterrows():
                    color      = proj_color.get(ev["Projeto"], "#6366f1")
                    ws_label   = f'{ev["Workshop"][:26]}{"…" if len(ev["Workshop"])>26 else ""}'
                    proj_label = f'{ev["Projeto"][:24]}{"…" if len(ev["Projeto"])>24 else ""}'
                    cons       = ev["Consultor"] or "—"
                    # Full tooltip text shown on hover
                    tooltip_ws   = ev["Workshop"].replace('"', "&quot;")
                    tooltip_proj = ev["Projeto"].replace('"', "&quot;")
                    tooltip_cons = cons.replace('"', "&quot;")
                    d_ini = ev["DataInicio"].strftime("%d/%m/%Y") if pd.notna(ev["DataInicio"]) else "—"
                    d_fim = ev["DataFim"].strftime("%d/%m/%Y") if pd.notna(ev["DataFim"]) else d_ini
                    periodo = d_ini if d_ini == d_fim else f"{d_ini} – {d_fim}"
                    st.markdown(f"""
                    <style>
                    .cal-tooltip {{ position:relative; display:block; }}
                    .cal-tooltip .cal-tip {{
                        visibility:hidden; opacity:0;
                        background:#1e293b; color:#f8fafc;
                        font-size:.75rem; line-height:1.5;
                        border-radius:7px; padding:8px 11px;
                        position:absolute; z-index:999;
                        left:105%; top:0;
                        width:220px;
                        box-shadow: 0 4px 12px rgba(0,0,0,.25);
                        transition: opacity .15s ease;
                        white-space: normal;
                        pointer-events: none;
                    }}
                    .cal-tooltip:hover .cal-tip {{ visibility:visible; opacity:1; }}
                    </style>
                    <div class="cal-tooltip">
                      <div class="cal-event" style="background:{color}22; border-left:3px solid {color}; color:#1e293b; cursor:default;">
                        <div style="font-weight:600; font-size:.72rem;">{ws_label}</div>
                        <div style="font-size:.66rem; color:{color}; font-weight:500; margin-top:1px;">{proj_label}</div>
                        <div style="font-size:.64rem; color:#64748b; margin-top:1px;">{cons[:22]}</div>
                      </div>
                      <div class="cal-tip">
                        <div style="font-weight:700; margin-bottom:4px; color:#e2e8f0;">{tooltip_ws}</div>
                        <div>📁 {tooltip_proj}</div>
                        <div>👤 {tooltip_cons}</div>
                        <div>📅 {periodo}</div>
                      </div>
                    </div>
                    """, unsafe_allow_html=True)

                st.markdown("</div>", unsafe_allow_html=True)

        # ── Full list ─────────────────────────────────────────────
        if not week_ws.empty:
            st.markdown('<div class="section-title">Detalhes dos Workshops</div>', unsafe_allow_html=True)
            show = week_ws[["Projeto","Workshop","Consultor","DataInicio","DataFim"]].copy()
            show["DataInicio"] = show["DataInicio"].dt.strftime("%d/%m/%Y")
            show["DataFim"]    = show["DataFim"].dt.strftime("%d/%m/%Y").fillna("—")
            st.dataframe(show, use_container_width=True, hide_index=True)

            st.download_button("⬇ Exportar Excel", to_excel_bytes(show),
                               file_name="workshops_semana.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ───────────────────────────────────────────────────────────────
# TAB 3 – Planejamento Semanal de Atividades (US-003)
# ───────────────────────────────────────────────────────────────
with tab3:
    st.markdown("### Planejamento Semanal de Atividades")

    if df3.empty:
        st.warning("Nenhuma atividade encontrada na aba '3-Atividades da Semana'.")
    else:
        dfa_all = df3[df3["Data"].notna()].copy()

        # ── All consultants: activities + workshop-only ───────────
        # Build name→email from df2 for workshop consultants
        _ws_name_email = {}
        _ws_email_name = {}
        for _, _wr in df2[df2["Consultor"].str.strip() != ""].iterrows():
            _wn = str(_wr.get("Consultor","")).strip()
            _we = str(_wr.get("Email","")).strip().lower()
            if _wn and _we:
                _ws_name_email[_wn] = _we
                _ws_email_name[_we] = _wn

        # Consultants in df2 not in df3 → workshop-only
        _act_names = set(dfa_all["Consultor"].dropna().unique())
        _act_emails = set(dfa_all["Email"].str.lower().dropna().unique())
        _ws_only_names = {
            name for name, email in _ws_name_email.items()
            if email not in _act_emails
        }
        _all_cons_names = sorted(_act_names | _ws_only_names)

        # ── Filtro ───────────────────────────────────────────────
        fa_cons = st.multiselect(
            "Filtrar por Consultor",
            _all_cons_names,
            key="t3_cons",
            placeholder="Todos os consultores…",
        )
        if fa_cons:
            dfa_all = dfa_all[dfa_all["Consultor"].isin(fa_cons)]

        # Week dates
        _dates_act = dfa_all["Data"].dropna()
        if not _dates_act.empty:
            min_date3 = _dates_act.min()
            max_date3 = _dates_act.max()
        else:
            _ws_sel = df2[df2["DataInicio"].notna()].copy()
            if fa_cons:
                _sel_emails = {_ws_name_email.get(n, "") for n in fa_cons}
                _ws_sel = _ws_sel[_ws_sel["Email"].str.lower().isin(_sel_emails)]
            _today_fb = pd.Timestamp(datetime.today().date())
            min_date3 = _ws_sel["DataInicio"].min() if not _ws_sel.empty else _today_fb
            max_date3 = _ws_sel["DataFim"].fillna(_ws_sel["DataInicio"]).max() if not _ws_sel.empty else _today_fb

        all_mondays3 = pd.date_range(
            start=min_date3 - timedelta(days=int(min_date3.weekday())),
            end=max_date3, freq="W-MON",
        )
        if len(all_mondays3) == 0:
            all_mondays3 = pd.DatetimeIndex([min_date3 - timedelta(days=int(min_date3.weekday()))])

        week_labels3 = [
            f"{d.strftime('%d/%m/%Y')} – {(d + timedelta(6)).strftime('%d/%m/%Y')}"
            for d in all_mondays3
        ]

        today3 = pd.Timestamp(datetime.today().date())
        current_monday3 = today3 - timedelta(days=today3.weekday())

        def _find_current_idx3():
            for i, m in enumerate(all_mondays3):
                if m >= current_monday3:
                    return i
            return len(all_mondays3) - 1

        if "t3_week_idx" not in st.session_state:
            st.session_state["t3_week_idx"] = _find_current_idx3()

        sel_idx3    = max(0, min(st.session_state["t3_week_idx"], len(all_mondays3) - 1))
        week_start3 = all_mondays3[sel_idx3]
        week_end3   = week_start3 + timedelta(days=6)

        dfa = dfa_all[dfa_all["Data"].between(week_start3, week_end3)].copy()

        # ── KPIs ─────────────────────────────────────────────────
        total_h = dfa["Horas"].sum()
        st.markdown(f"""
        <div class="kpi-grid">
            {kpi_html(f"{total_h:.0f}h", "Total de Horas na Semana")}
            {kpi_html(dfa["Consultor"].nunique(), "Consultores", "green")}
            {kpi_html(dfa["Projeto"].nunique(),   "Projetos",    "amber")}
            {kpi_html(dfa["Atividade"].nunique(), "Atividades",  "rose")}
        </div>
        """, unsafe_allow_html=True)

        # ── Navegação ─────────────────────────────────────────────
        _g1, _g2, _g3, _g4 = st.columns([1, 5, 1, 1])
        with _g1:
            if st.button("◀ Anterior", key="t3_prev"):
                if st.session_state["t3_week_idx"] > 0:
                    st.session_state["t3_week_idx"] -= 1
                    st.rerun()
        with _g3:
            if st.button("Próxima ▶", key="t3_next"):
                if st.session_state["t3_week_idx"] < len(all_mondays3) - 1:
                    st.session_state["t3_week_idx"] += 1
                    st.rerun()
        with _g4:
            if st.button("Semana Atual", key="t3_today"):
                st.session_state["t3_week_idx"] = _find_current_idx3()
                st.rerun()
        with _g2:
            st.markdown(
                f"<div style='text-align:center; padding:.3rem 0; font-weight:400; "
                f"color:#94a3b8; font-size:.85rem;'>{week_labels3[sel_idx3]}</div>",
                unsafe_allow_html=True,
            )

        st.markdown('<div class="section-title">Gantt da Semana por Consultor</div>', unsafe_allow_html=True)

        # Check if there are workshops this week even if no activities
        _ws_check = df2[
            (df2["DataInicio"].notna()) &
            (df2["DataInicio"] <= week_end3) &
            (df2["DataFim"].fillna(df2["DataInicio"]) >= week_start3)
        ]
        # Filter by selected consultants if any
        if fa_cons:
            _fa_emails = {_ws_name_email.get(n,"") for n in fa_cons}
            _ws_check = _ws_check[_ws_check["Email"].str.lower().isin(_fa_emails)]

        if dfa.empty and _ws_check.empty:
            st.info("Nenhuma atividade ou workshop registrado para esta semana.")
        else:
            days_of_week = [week_start3 + timedelta(d) for d in range(5)]
            day_names3   = ["Seg","Ter","Qua","Qui","Sex"]

            # Build one Gantt bar per (Consultor, Projeto, Atividade, Data)
            # x-axis: days of the week; y-axis: consultants
            # Each bar spans from HorasIni to HorasFim on the given day.
            # We use a Plotly figure with shapes for precision.

            # Include consultants with workshops this week even if no activities
            _ws_week_all = df2[
                (df2["DataInicio"].notna()) &
                (df2["DataInicio"] <= week_end3) &
                (df2["DataFim"].fillna(df2["DataInicio"]) >= week_start3) &
                (df2["Consultor"].str.strip() != "")
            ].copy()

            # Build email↔name maps from df2 (email is the reliable key)
            _ws_email_map  = {}   # name  → email
            _ws_name_map   = {}   # email → name
            for _, _wr in _ws_week_all.iterrows():
                _wn = str(_wr.get("Consultor","")).strip()
                _we = str(_wr.get("Email","")).strip().lower()
                if _wn and _we:
                    _ws_email_map[_wn] = _we
                    _ws_name_map[_we]  = _wn

            # Build email map from activities
            _act_email_map = dfa.groupby("Consultor")["Email"].first().to_dict()
            # email → canonical name (prefer activity name, fallback to workshop name)
            _email_to_name = {v.lower(): k for k, v in _act_email_map.items() if v}
            for _we, _wn in _ws_name_map.items():
                if _we not in _email_to_name:
                    _email_to_name[_we] = _wn

            # Consultant list = union by email
            _all_emails = set(v.lower() for v in _act_email_map.values() if v)
            _all_emails |= set(_ws_name_map.keys())
            all_cons_set = {_email_to_name[e] for e in _all_emails if e in _email_to_name}

            # Apply fa_cons filter to gantt consultant list
            if fa_cons:
                _fa_emails_g = set()
                for _fn in fa_cons:
                    _fe = _ws_name_email.get(_fn, "")
                    if not _fe:
                        _rows = dfa_all[dfa_all["Consultor"] == _fn]["Email"]
                        _fe = _rows.iloc[0] if not _rows.empty else ""
                    _fa_emails_g.add(str(_fe).lower())
                all_cons_set = {
                    nm for nm in all_cons_set
                    if _ws_email_map.get(nm, "").lower() in _fa_emails_g
                    or _act_email_map.get(nm, "").lower() in _fa_emails_g
                    or nm in fa_cons
                }

            consultores = sorted(all_cons_set)
            # y positions: one row per consultant (bottom to top in plotly = reversed list)
            cons_y = {c: i for i, c in enumerate(reversed(consultores))}

            fig_gantt = go.Figure()

            # Background alternating bands per consultant row
            for c, yi in cons_y.items():
                fig_gantt.add_shape(
                    type="rect",
                    x0=-0.5, x1=4.5,
                    y0=yi - 0.45, y1=yi + 0.45,
                    fillcolor="#f8fafc" if yi % 2 == 0 else "white",
                    line_width=0, layer="below",
                )

            # Day separator lines
            for d in range(5):
                fig_gantt.add_shape(
                    type="line",
                    x0=d - 0.5, x1=d - 0.5,
                    y0=-0.5, y1=len(consultores) - 0.5,
                    line=dict(color="#e2e8f0", width=1), layer="below",
                )

            # Highlight today if in week
            today_in_week = (today3 >= week_start3) and (today3 <= week_end3)
            if today_in_week:
                today_x = (today3 - week_start3).days
                fig_gantt.add_shape(
                    type="rect",
                    x0=today_x - 0.48, x1=today_x + 0.48,
                    y0=-0.5, y1=len(consultores) - 0.5,
                    fillcolor="#eef2ff", line_width=0, layer="below",
                )

            # ── Activity bars (upper half of row) ────────────────
            bar_height  = 0.50
            bar_offset  = 0.13   # activities sit slightly above center
            ws_offset   = -0.13  # workshops sit slightly below center
            ws_height   = 0.34
            added_projects  = set()
            added_workshops = set()

            for _, row in dfa.iterrows():
                c     = row["Consultor"]
                proj  = row["Projeto"]
                atv   = row["Atividade"]
                data  = row["Data"]
                h_ini = row["HorasIni"] if pd.notna(row.get("HorasIni")) else 9.0
                h_fim = row["HorasFim"] if pd.notna(row.get("HorasFim")) else 18.0

                day_x = (pd.Timestamp(data) - week_start3).days
                if day_x < 0 or day_x > 4:
                    continue

                yi     = cons_y.get(c, 0)
                color  = proj_color.get(proj, "#6366f1")
                day_hours = 9.0
                x_start = day_x + (h_ini - 9.0) / day_hours * 0.9 - 0.45
                x_end   = day_x + (h_fim - 9.0) / day_hours * 0.9 - 0.45
                x_start = max(day_x - 0.45, min(x_start, day_x + 0.45))
                x_end   = max(x_start + 0.05, min(x_end, day_x + 0.45))

                show_legend = False
                added_projects.add(proj)
                atv_short = atv[:22] + "…" if len(atv) > 22 else atv

                fig_gantt.add_trace(go.Bar(
                    x=[x_end - x_start],
                    y=[yi + bar_offset],
                    base=[x_start],
                    orientation="h",
                    width=bar_height,
                    marker=dict(color=color, opacity=0.88, line=dict(color=color, width=1)),
                    name=proj,
                    legendgroup=proj,
                    showlegend=show_legend,
                    text=atv_short,
                    textposition="inside",
                    insidetextanchor="middle",
                    textfont=dict(size=9, color="white"),
                    hovertemplate=(
                        f"<b>{c}</b><br>"
                        f"Projeto: {proj}<br>"
                        f"Atividade: {atv}<br>"
                        f"Horário: {int(h_ini):02d}h – {int(h_fim):02d}h<br>"
                        "<extra></extra>"
                    ),
                ))

            # ── Workshop bars (lower half of row, slate color) ────
            # Unified email map: name → email
            email_map3 = {**{v: k for k, v in _email_to_name.items()}, **_act_email_map}
            # Also direct name→email from workshop map
            for _wc, _we in _ws_email_map.items():
                if _wc not in email_map3:
                    email_map3[_wc] = _we

            for c in consultores:
                c_email = email_map3.get(c, "").lower()
                if not c_email:
                    continue
                # Find workshops for this consultant in this week
                ws_week = df2[
                    (df2["Email"].str.lower() == c_email) &
                    (df2["DataInicio"].notna()) &
                    (df2["DataInicio"] <= week_end3) &
                    (df2["DataFim"].fillna(df2["DataInicio"]) >= week_start3)
                ]
                for _, wrow in ws_week.iterrows():
                    ws_name = wrow["Workshop"]
                    ws_proj = wrow["Projeto"]
                    # Span all days of this workshop within the week
                    d_ini = max(wrow["DataInicio"], week_start3)
                    d_fim = min(wrow["DataFim"] if pd.notna(wrow["DataFim"]) else wrow["DataInicio"], week_end3)
                    for d_cur in pd.date_range(d_ini, d_fim, freq="D"):
                        day_x = (d_cur - week_start3).days
                        if day_x < 0 or day_x > 4:
                            continue
                        yi = cons_y.get(c, 0)
                        ws_key = f"WS_{ws_name}"
                        show_ws_legend = ws_key not in added_workshops
                        added_workshops.add(ws_key)
                        ws_short = ws_name[:20] + "…" if len(ws_name) > 20 else ws_name
                        fig_gantt.add_trace(go.Bar(
                            x=[0.88],
                            y=[yi + ws_offset],
                            base=[day_x - 0.44],
                            orientation="h",
                            width=ws_height,
                            marker=dict(
                                color="#0ea5e9",
                                opacity=0.80,
                                line=dict(color="#0284c7", width=1),
                                pattern=dict(shape="/", size=4, fgcolor="white", fgopacity=0.3),
                            ),
                            name="Workshop",
                            legendgroup="workshops",
                            showlegend=False,
                            text=ws_short,
                            textposition="inside",
                            insidetextanchor="middle",
                            textfont=dict(size=8, color="white"),
                            hovertemplate=(
                                f"<b>{c}</b><br>"
                                f"📅 Workshop: {ws_name}<br>"
                                f"Projeto: {ws_proj}<br>"
                                f"Data: {d_cur.strftime('%d/%m/%Y')}<br>"
                                "<extra></extra>"
                            ),
                        ))

            # ── Legend entries: activity + workshop ──────────────
            fig_gantt.add_trace(go.Bar(
                x=[None], y=[None], orientation="h",
                name="▬  Atividade",
                marker=dict(color="#6366f1", opacity=0.88),
                showlegend=False,
                legendgroup="__legend_atv",
            ))
            fig_gantt.add_trace(go.Bar(
                x=[None], y=[None], orientation="h",
                name="▨  Workshop",
                marker=dict(
                    color="#0ea5e9", opacity=0.80,
                    pattern=dict(shape="/", size=4, fgcolor="white", fgopacity=0.3),
                ),
                showlegend=False,
                legendgroup="__legend_ws",
            ))

            # Axes config
            fig_gantt.update_layout(
                barmode="overlay",
                height=max(260, len(consultores) * 44 + 80),
                margin=dict(l=0, r=0, t=60, b=20),
                plot_bgcolor="white",
                paper_bgcolor="white",
                showlegend=True,
                legend=dict(
                    orientation="h",
                    x=0, xanchor="left",
                    y=1.12, yanchor="bottom",
                    font=dict(size=10),
                    bgcolor="rgba(0,0,0,0)",
                    title=dict(text="", font=dict(size=10)),
                ),
                xaxis=dict(
                    tickmode="array",
                    tickvals=list(range(5)),
                    ticktext=[
                        f"<b>{day_names3[i]}</b><br>{days_of_week[i].strftime('%d/%m')}"
                        for i in range(5)
                    ],
                    range=[-0.5, 4.5],
                    showgrid=False,
                    zeroline=False,
                    fixedrange=True,
                    side="top",
                ),
                yaxis=dict(
                    tickmode="array",
                    tickvals=list(cons_y.values()),
                    ticktext=list(cons_y.keys()),
                    showgrid=False,
                    zeroline=False,
                    fixedrange=True,
                    tickfont=dict(size=11),
                ),
            )

            st.plotly_chart(fig_gantt, use_container_width=True)

        # ── Detail table ──────────────────────────────────────────
        with st.expander("📄 Detalhe das atividades da semana"):
            if dfa.empty:
                st.info("Sem atividades nesta semana.")
            else:
                show3 = dfa[["Consultor","Projeto","Atividade","Fase","Data","Horas"]].copy()
                show3["Data"]  = show3["Data"].dt.strftime("%d/%m/%Y")
                show3["Horas"] = show3["Horas"].round(2)
                st.dataframe(show3.sort_values(["Consultor","Data"]),
                             use_container_width=True, hide_index=True)

        st.download_button("⬇ Exportar Excel", to_excel_bytes(dfa),
                           file_name="planejamento_semanal.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ───────────────────────────────────────────────────────────────
# TAB 4 – Recursos
# ───────────────────────────────────────────────────────────────
with tab4:
    st.markdown("### Recursos")

    if df_rec.empty:
        st.warning("Aba '4 - Recursos' não encontrada ou vazia.")
    else:
        all_modulos = sorted({m for specs in df_rec["Especialidades"] for m in specs})

        # ── Filters ──────────────────────────────────────────────
        fc1, fc2, fc3 = st.columns([2, 2, 2])
        f_status = fc1.multiselect("Status", ["Alocado", "Disponível"], key="r_status",
                                   placeholder="Todos…")
        f_mods   = fc2.multiselect("Especialidade / Módulo", all_modulos, key="r_mods",
                                   placeholder="Todas…")
        f_name   = fc3.multiselect("Consultor", sorted(df_rec["Consultor"].dropna().unique()),
                                   key="r_name", placeholder="Todos…")

        dfr = df_rec.copy()
        if f_status:
            dfr = dfr[dfr["Status"].isin(f_status)]
        if f_mods:
            dfr = dfr[dfr["Especialidades"].apply(lambda s: any(m in s for m in f_mods))]
        if f_name:
            dfr = dfr[dfr["Consultor"].isin(f_name)]

        # ── KPIs ─────────────────────────────────────────────────
        n_total     = len(dfr)
        n_alocados  = (dfr["Status"] == "Alocado").sum()
        n_disponiveis = (dfr["Status"] == "Disponível").sum()
        pct_disp    = f"{n_disponiveis/n_total*100:.0f}%" if n_total > 0 else "—"

        st.markdown(f"""
        <div class="kpi-grid">
            {kpi_html(n_total,       "Total de Recursos")}
            {kpi_html(n_alocados,    "Alocados em Projeto", "amber")}
            {kpi_html(n_disponiveis, "Disponíveis",         "green")}
            {kpi_html(pct_disp,      "% Disponíveis",       "rose")}
        </div>
        """, unsafe_allow_html=True)

        # ── Matrix view: consultores × módulos ───────────────────
        st.markdown('<div class="section-title">Matriz de Especialidades</div>',
                    unsafe_allow_html=True)

        # Build matrix HTML
        header_cells = "".join(
            f"<th style='padding:5px 8px; font-size:.72rem; font-weight:600; color:#475569;"
            f"writing-mode:vertical-rl; transform:rotate(180deg); white-space:nowrap;"
            f"min-width:32px;'>{m}</th>"
            for m in all_modulos
        )

        status_color = {"Alocado": "#f59e0b", "Disponível": "#10b981"}
        status_bg    = {"Alocado": "#fffbeb", "Disponível": "#f0fdf4"}

        body_rows = ""
        for _, row in dfr.sort_values(["Status","Consultor"]).iterrows():
            sc = status_color.get(row["Status"], "#6366f1")
            sb = status_bg.get(row["Status"], "#f8fafc")
            badge = (
                f"<span style='background:{sb}; color:{sc}; border:1px solid {sc};"
                f"border-radius:4px; padding:1px 7px; font-size:.7rem; font-weight:600;"
                f"white-space:nowrap;'>{row['Status']}</span>"
            )
            mod_cells = "".join(
                f"<td style='text-align:center; padding:4px;'>"
                f"{'<span style="color:#6366f1; font-size:1rem;">●</span>' if m in row['Especialidades'] else '<span style="color:#e2e8f0;">·</span>'}"
                f"</td>"
                for m in all_modulos
            )
            body_rows += (
                f"<tr style='border-bottom:1px solid #f1f5f9; background:white;'>"
                f"<td style='padding:6px 10px; font-size:.82rem; white-space:nowrap; color:#1e293b;"
                f"font-weight:500; min-width:180px;'>{row['Consultor']}</td>"
                f"<td style='padding:6px 10px;'>{badge}</td>"
                f"{mod_cells}"
                f"</tr>"
            )

        st.markdown(f"""
        <div style="overflow-x:auto;">
        <table style="border-collapse:collapse; width:100%; font-size:.82rem;">
          <thead>
            <tr style="background:#f8fafc; border-bottom:2px solid #e2e8f0;">
              <th style="padding:8px 10px; text-align:left; color:#475569; font-weight:600; min-width:180px;">Consultor</th>
              <th style="padding:8px 10px; text-align:left; color:#475569; font-weight:600; min-width:100px;">Status</th>
              {header_cells}
            </tr>
          </thead>
          <tbody>{body_rows}</tbody>
        </table>
        </div>
        """, unsafe_allow_html=True)

        # ── Bar chart: consultores por especialidade ──────────────
        st.markdown('<div class="section-title">Consultores por Especialidade</div>',
                    unsafe_allow_html=True)

        mod_counts = []
        for m in all_modulos:
            total_m  = dfr["Especialidades"].apply(lambda s: m in s).sum()
            aloc_m   = dfr[dfr["Status"]=="Alocado"]["Especialidades"].apply(lambda s: m in s).sum()
            disp_m   = total_m - aloc_m
            mod_counts.append({"Módulo": m, "Alocados": aloc_m, "Disponíveis": disp_m})

        df_chart = pd.DataFrame(mod_counts).sort_values("Módulo")

        fig_mod = go.Figure()
        fig_mod.add_trace(go.Bar(
            name="Alocados", x=df_chart["Módulo"], y=df_chart["Alocados"],
            marker_color="#f59e0b",
        ))
        fig_mod.add_trace(go.Bar(
            name="Disponíveis", x=df_chart["Módulo"], y=df_chart["Disponíveis"],
            marker_color="#10b981",
        ))
        fig_mod.update_layout(
            barmode="stack", height=300,
            margin=dict(l=0, r=0, t=10, b=0),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=dict(orientation="h", y=1.1, x=0),
            xaxis=dict(tickfont=dict(size=11)),
            yaxis=dict(title="Qtd", tickfont=dict(size=10)),
        )
        st.plotly_chart(fig_mod, use_container_width=True)

        st.download_button("⬇ Exportar Excel", to_excel_bytes(
            dfr[["Consultor","Status","Modulos"]].rename(columns={"Modulos":"Especialidades"})
        ), file_name="recursos.xlsx",
           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ───────────────────────────────────────────────────────────────
# TAB 5 – Go Lives
# ───────────────────────────────────────────────────────────────
with tab5:
    st.markdown("### ⚠️ Conflitos de Go Live")

    if df_golive.empty:
        st.warning("Nenhuma data de Go Live encontrada na aba Cockpit.")
    else:
        # ── Filters ──────────────────────────────────────────────
        fc1, fc2 = st.columns([2, 2])
        f_gl_cons = fc1.multiselect("Filtrar por Consultor",
                                    sorted(df_golive["Consultor"].dropna().unique()),
                                    key="gl_cons", placeholder="Todos…")
        f_gl_proj = fc2.multiselect("Filtrar por Projeto",
                                    sorted(df_golive["Projeto"].dropna().unique()),
                                    key="gl_proj", placeholder="Todos…")

        dfgl = df_golive.copy()
        if f_gl_cons: dfgl = dfgl[dfgl["Consultor"].isin(f_gl_cons)]
        if f_gl_proj: dfgl = dfgl[dfgl["Projeto"].isin(f_gl_proj)]

        # ── KPIs ─────────────────────────────────────────────────
        proj_per_cons_gl = dfgl.groupby(["Consultor","GoLive"])["Projeto"].nunique()
        n_conflitos = (proj_per_cons_gl > 1).sum()
        n_criticos  = (proj_per_cons_gl >= 3).sum()

        st.markdown(f"""
        <div class="kpi-grid">
            {kpi_html(dfgl["Projeto"].nunique(), "Projetos", "green")}
            {kpi_html(int(n_conflitos), "Consultores com conflito", "amber")}
            {kpi_html(int(n_criticos),  "Situações críticas (3+)", "rose")}
        </div>
        """, unsafe_allow_html=True)

        from collections import defaultdict as _dd

        # Exclude juniors using Senioridade column (more reliable than text parsing)
        _jr_emails = set(
            df_golive[df_golive["Senioridade"].str.lower() == "junior"]["Email"]
            .str.strip().str.lower()
        ) if "Senioridade" in df_golive.columns else set()

        # rec_map: email → {name, modules} — exclude juniors
        _rec_map = {}
        for _, _rr in df_rec.iterrows():
            _em  = str(_rr.get("Email","")).strip().lower()
            _nm  = str(_rr.get("Consultor","")).strip()
            _sen = str(_rr.get("Senioridade","")).strip().lower()
            if not _em or _sen == "junior": continue
            _rec_map[_em] = {"name": _nm, "modules": set(_rr.get("Especialidades",[]) or [])}

        # alloc_map: email → assignments — exclude juniors
        _alloc_map = _dd(lambda: {"name":"", "assignments":[]})
        _dfgl_sr = df_golive[df_golive.get("Senioridade", pd.Series("Sênior", index=df_golive.index)).str.lower() != "junior"]
        for _, _ar in _dfgl_sr.iterrows():
            _em2 = str(_ar.get("Email","")).strip().lower()
            if not _em2: continue
            _alloc_map[_em2]["name"] = _ar["Consultor"]
            _alloc_map[_em2]["assignments"].append({"projeto": _ar["Projeto"], "golive": _ar["GoLive"]})

        # ── Gráfico de bolhas de conflitos ───────────────────────
        st.markdown('<div class="section-title">Mapa de Conflitos por Consultor</div>',
                    unsafe_allow_html=True)

        # Build bubble data: consultant × month → n_projects + project list
        _bubble_rows = []
        for _em2, _data2 in _alloc_map.items():
            _by_m = _dd(list)
            for _a in _data2["assignments"]:
                _by_m[_a["golive"]].append(_a["projeto"])
            for _month2, _projs2 in _by_m.items():
                _pu = list(dict.fromkeys(_projs2))
                _nm2 = _data2["name"] or _rec_map.get(_em2, {}).get("name", "")
                if not _nm2: continue
                _bubble_rows.append({
                    "Consultor": f"{_nm2.split()[0]} {_nm2.split()[-1]}",
                    "ConsultorFull": _nm2,
                    "Mês": _month2,
                    "MêsStr": _month2.strftime("%b/%Y"),
                    "Projetos": len(_pu),
                    "Lista": "<br>".join(f"▸ {p.split(' - ')[1].strip() if ' - ' in p else p[:40]}" for p in _pu),
                })

        df_bubble = pd.DataFrame(_bubble_rows)

        if not df_bubble.empty:
            # Apply consultant filter
            if f_gl_cons:
                _cons_short = [f"{n.split()[0]} {n.split()[-1]}" for n in f_gl_cons]
                df_bubble = df_bubble[df_bubble["Consultor"].isin(_cons_short)]

            def _bubble_color(n):
                if n >= 3: return "#ef4444"
                if n == 2: return "#f59e0b"
                return "#10b981"

            df_bubble["Cor"]   = df_bubble["Projetos"].apply(_bubble_color)
            df_bubble["Label"] = df_bubble["Projetos"].apply(lambda n: "Crítico" if n >= 3 else ("Atenção" if n == 2 else "Ok"))

            # Sort consultants: most total conflicts first (bottom→top in plotly)
            _cons_order_all = (df_bubble.groupby("Consultor")["Projetos"]
                               .sum().sort_values(ascending=False).index.tolist())

            # ── Pagination state ─────────────────────────────────
            _page_size = 10
            _n_pages   = max(1, -(-len(_cons_order_all) // _page_size))

            if "bubble_page" not in st.session_state:
                st.session_state["bubble_page"] = 0
            _cur_page = st.session_state["bubble_page"]

            _start = _cur_page * _page_size
            _end   = _start + _page_size
            _cons_page = _cons_order_all[_start:_end]
            # For plotly y-axis bottom→top: reverse
            _cons_order = list(reversed(_cons_page))

            df_bubble_page = df_bubble[df_bubble["Consultor"].isin(_cons_page)]

            _month_order = sorted(df_bubble["MêsStr"].unique(),
                                  key=lambda s: pd.Timestamp("01 " + s))

            fig_bubble = go.Figure()

            # ── Alternating row bands ────────────────────────────
            for _ci, _cn in enumerate(_cons_order):
                fig_bubble.add_shape(
                    type="rect",
                    x0=-0.5, x1=len(_month_order) - 0.5,
                    y0=_ci - 0.45, y1=_ci + 0.45,
                    fillcolor="#f8fafc" if _ci % 2 == 0 else "white",
                    line_width=0, layer="below",
                )

            # ── Dot grid: empty circles for months without conflict
            for _cn in _cons_page:
                for _mo in _month_order:
                    _has = not df_bubble_page[
                        (df_bubble_page["Consultor"] == _cn) &
                        (df_bubble_page["MêsStr"] == _mo)
                    ].empty
                    if not _has:
                        fig_bubble.add_trace(go.Scatter(
                            x=[_mo], y=[_cn],
                            mode="markers",
                            marker=dict(size=36, color="#f8fafc",
                                        line=dict(color="#f1f5f9", width=1)),
                            showlegend=False,
                            hoverinfo="skip",
                        ))

            # ── Filled bubbles sized by project count
            for _label, _color, _border in [
                ("Crítico", "#ef4444", "#b91c1c"),
                ("Atenção", "#f59e0b", "#b45309"),
                ("Ok",      "#10b981", "#047857"),
            ]:
                _sub = df_bubble_page[df_bubble_page["Label"] == _label]
                if _sub.empty: continue
                fig_bubble.add_trace(go.Scatter(
                    x=_sub["MêsStr"],
                    y=_sub["Consultor"],
                    mode="markers+text",
                    name=_label,
                    marker=dict(
                        size=_sub["Projetos"] * 18 + 14,
                        color=_color,
                        opacity=0.45,
                        line=dict(color=_border, width=1.5),
                        symbol="circle",
                    ),
                    text=_sub["Projetos"].astype(str),
                    textposition="middle center",
                    textfont=dict(color="white", size=13, family="Inter"),
                    customdata=_sub[["ConsultorFull","Lista","Projetos"]].values,
                    hovertemplate=(
                        "<b>%{customdata[0]}</b><br>"
                        "Mês: %{x}<br>"
                        "<b>%{customdata[2]} projetos simultâneos</b><br><br>"
                        "%{customdata[1]}"
                        "<extra></extra>"
                    ),
                ))

            fig_bubble.update_layout(
                height=max(320, len(_cons_page) * 54 + 110),
                margin=dict(l=0, r=20, t=52, b=10),
                plot_bgcolor="white",
                paper_bgcolor="white",
                xaxis=dict(
                    categoryorder="array",
                    categoryarray=_month_order,
                    showgrid=False,
                    zeroline=False,
                    tickfont=dict(size=12, color="#475569", family="Inter"),
                    side="top",
                ),
                yaxis=dict(
                    categoryorder="array",
                    categoryarray=_cons_order,
                    showgrid=False,
                    zeroline=False,
                    tickfont=dict(size=11, color="#1e293b", family="Inter"),
                ),
                legend=dict(
                    orientation="h", x=0, y=-0.08, xanchor="left",
                    font=dict(size=11),
                    title=dict(text="Situação: ", font=dict(size=11, color="#64748b")),
                    bgcolor="rgba(0,0,0,0)",
                ),
            )

            st.plotly_chart(fig_bubble, use_container_width=True)

            # ── Pagination buttons (below chart) ─────────────────
            _pag1, _pag2, _pag3 = st.columns([1, 4, 1])
            with _pag1:
                if st.button("◀", key="bub_prev") and st.session_state["bubble_page"] > 0:
                    st.session_state["bubble_page"] -= 1
                    st.rerun()
            with _pag3:
                if st.button("▶", key="bub_next") and st.session_state["bubble_page"] < _n_pages - 1:
                    st.session_state["bubble_page"] += 1
                    st.rerun()
            with _pag2:
                st.markdown(
                    f"<div style='text-align:center; padding:.2rem 0; font-size:.85rem; color:#6366f1; font-weight:600;'>"
                    f"Página {_cur_page + 1} de {_n_pages} &nbsp;·&nbsp; {len(_cons_order_all)} consultores</div>",
                    unsafe_allow_html=True,
                )

        # ── Sugestões de substituição ─────────────────────────────
        st.markdown('<div class="section-title">💡 Sugestões de Substituição</div>',
                    unsafe_allow_html=True)

        # Detect conflicts and build suggestions
        _suggestions = []
        for _em, _data in _alloc_map.items():
            _by_month = _dd(list)
            for _a in _data["assignments"]:
                _by_month[_a["golive"]].append(_a["projeto"])
            for _month, _projs in _by_month.items():
                _projs_uniq = list(dict.fromkeys(_projs))
                if len(_projs_uniq) < 2: continue
                _my_mods = _rec_map.get(_em, {}).get("modules", set())
                _candidates = []
                # Busy consultants free in this month
                for _em2, _d2 in _alloc_map.items():
                    if _em2 == _em: continue
                    if _month in {_a["golive"] for _a in _d2["assignments"]}: continue
                    _overlap = _rec_map.get(_em2,{}).get("modules",set()) & _my_mods
                    if _overlap:
                        _candidates.append({"name": _d2["name"] or _rec_map.get(_em2,{}).get("name",""), "modules": sorted(_overlap)})
                # Fully free consultants (not in alloc_map at all)
                for _em3, _rd in _rec_map.items():
                    if _em3 in _alloc_map: continue
                    _overlap = _rd["modules"] & _my_mods
                    if _overlap:
                        _candidates.append({"name": _rd["name"], "modules": sorted(_overlap)})
                _suggestions.append({
                    "name": _data["name"], "month": _month,
                    "n_projects": len(_projs_uniq), "projects": _projs_uniq,
                    "my_modules": sorted(_my_mods), "candidates": _candidates[:3],
                })

        _suggestions.sort(key=lambda x: (-x["n_projects"], x["month"]))
        if f_gl_cons:
            _suggestions = [s for s in _suggestions if s["name"] in f_gl_cons]

        if not _suggestions:
            st.info("Nenhum conflito detectado nos perfis sênior.")
        else:
            for _i, _s in enumerate(_suggestions):
                _month_str = _s["month"].strftime("%b/%Y")
                _icon = "🔴" if _s["n_projects"] >= 3 else "🟡"
                _sc   = "#ef4444" if _s["n_projects"] >= 3 else "#f59e0b"
                with st.expander(
                    f"{_icon}  **{_s['name']}**  ·  {_s['n_projects']} projetos  ·  {_month_str}",
                    expanded=False,
                ):
                    _ec1, _ec2, _ec3 = st.columns([2, 1, 2])
                    with _ec1:
                        st.markdown("<div style='font-size:.75rem;color:#64748b;font-weight:600;"
                                    "text-transform:uppercase;letter-spacing:.05em;margin-bottom:.3rem;'>"
                                    "Projetos em conflito</div>", unsafe_allow_html=True)
                        for _p in _s["projects"]:
                            _pn = _p.split(" - ")[1].strip() if " - " in _p else _p[:40]
                            st.markdown(f"<div style='font-size:.82rem;color:#3730a3;padding:2px 0;'>▸ {_pn}</div>",
                                        unsafe_allow_html=True)
                    with _ec2:
                        st.markdown("<div style='font-size:.75rem;color:#64748b;font-weight:600;"
                                    "text-transform:uppercase;letter-spacing:.05em;margin-bottom:.3rem;'>"
                                    "Perfil</div>", unsafe_allow_html=True)
                        for _m in (_s["my_modules"] or ["—"]):
                            st.markdown(f"<span style='background:#e0e7ff;color:#3730a3;border-radius:3px;"
                                        f"padding:2px 8px;font-size:.78rem;display:inline-block;margin-bottom:3px;'>{_m}</span>",
                                        unsafe_allow_html=True)
                    with _ec3:
                        st.markdown("<div style='font-size:.75rem;color:#64748b;font-weight:600;"
                                    "text-transform:uppercase;letter-spacing:.05em;margin-bottom:.3rem;'>"
                                    "💡 Sugestões</div>", unsafe_allow_html=True)
                        if _s["candidates"]:
                            for _ci2, _cand in enumerate(_s["candidates"]):
                                _cn = f"{_cand['name'].split()[0]} {_cand['name'].split()[-1]}"
                                _cm = " · ".join(_cand["modules"])
                                st.markdown(
                                    f"<div style='display:flex;align-items:center;gap:.4rem;padding:2px 0;'>"
                                    f"<span style='background:#dcfce7;color:#166534;border-radius:50%;"
                                    f"width:20px;height:20px;display:inline-flex;align-items:center;"
                                    f"justify-content:center;font-size:.7rem;font-weight:700;'>{_ci2+1}</span>"
                                    f"<span style='font-size:.82rem;color:#1e293b;font-weight:500;'>{_cn}</span>"
                                    f"<span style='font-size:.72rem;color:#64748b;'>{_cm}</span></div>",
                                    unsafe_allow_html=True,
                                )
                        else:
                            st.markdown("<div style='font-size:.78rem;color:#94a3b8;'>Sem candidatos disponíveis</div>",
                                        unsafe_allow_html=True)

        # ── Detail table ──────────────────────────────────────────
        with st.expander("📄 Ver detalhe completo"):
            detail_gl = dfgl.copy()
            detail_gl["Go Live"] = detail_gl["GoLive"].dt.strftime("%b/%Y")
            if "Papel" in df1.columns:
                # Build direct map: (principal_name, projeto) → secondary_name
                # Each row in df1 already links principal to secondary via the same row
                # Reconstruct from original cockpit: join principal rows with secondary of SAME row
                _direct = (
                    df1[df1["Papel"]=="Principal"][["Consultor","Email","Projeto"]]
                    .drop_duplicates(subset=["Email","Projeto"])
                )
                _sombras_direct = (
                    df1[df1["Papel"]=="Sombra"][["Email","Projeto","Consultor"]]
                    .drop_duplicates()
                    .rename(columns={"Consultor":"NomeSombra","Email":"EmailSombra"})
                )
                # Match: same project, same "row origin" — use Email of principal
                # The sombra row has its OWN email, not principal email
                # We need the cockpit raw: for each row, principal→sombra are in same row
                # Rebuild from df1 by matching Projeto+Módulo
                _p = df1[df1["Papel"]=="Principal"][["Consultor","Email","Projeto","Módulo"]].drop_duplicates()
                _s = df1[df1["Papel"]=="Sombra"][["Projeto","Módulo","Consultor"]].drop_duplicates().rename(columns={"Consultor":"NomeSombra"})
                _somb_lk = (
                    _p.merge(_s, on=["Projeto","Módulo"], how="left")
                    .groupby(["Consultor","Projeto"])["NomeSombra"]
                    .apply(lambda x: ", ".join(x.dropna().unique()) or "—")
                    .reset_index()
                    .rename(columns={"NomeSombra": "👥 Sombra"})
                )
                detail_gl = detail_gl.merge(_somb_lk, on=["Consultor","Projeto"], how="left")
                detail_gl["👥 Sombra"] = detail_gl["👥 Sombra"].fillna("—")
                detail_gl = detail_gl[["Consultor","Cliente","Projeto","Go Live","👥 Sombra"]].sort_values(["Go Live","Consultor"])
            else:
                detail_gl = detail_gl[["Consultor","Cliente","Projeto","Go Live"]].sort_values(["Go Live","Consultor"])
            st.dataframe(detail_gl, use_container_width=True, hide_index=True)

        st.download_button("⬇ Exportar Excel", to_excel_bytes(detail_gl),
                           file_name="golives_conflitos.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
