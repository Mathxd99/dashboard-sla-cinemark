import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import date

# ─── CONFIGURAÇÃO DA PÁGINA ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Cinemark · SLA & Análise Operacional",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── PALETA OFICIAL CINEMARK CLEAN ────────────────────────────────────────────
BG_PAGE           = "#f1f3f8"
BG_CARD           = "#ffffff"
CINEMARK_RED      = "#ba1c1c"
CINEMARK_DARK_RED = "#7a1010"
BORDER_COLOR      = "#e2e8f0"
TEXT_MAIN         = "#0f172a"
TEXT_MUTED        = "#64748b"

SLA_ORDEM = ["até 2h", "até 24h", "até 48h", "fora", "Em Aberto"]
SLA_CORES = {
    "até 2h":    "#16a34a", # Verde Sucesso
    "até 24h":   "#2563eb", # Azul No Prazo
    "até 48h":   "#ea580c", # Laranja
    "fora":      "#ba1c1c", # Vermelho Violação
    "Em Aberto": "#94a3b8",
}

STATUS_MAP = {
    "TI_EncerraFaltaValidacao": "Encerrado s/ Validação",
    "TI_EmAtend":               "Em Atendimento",
    "TI_Encerrado":             "Encerrado",
    "TI_CanceladoPedSolicita":  "Cancelado (Sol.)",
    "TI_CanceladoSolucionador": "Cancelado (Soluc.)",
    "TI_Encerrar":              "A Encerrar",
    "Aguardando aprovacao":     "Aguard. Aprovação",
}
MESES = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",
         7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}

DIAS_MAP = {
    0: "Segunda-feira",
    1: "Terça-feira",
    2: "Quarta-feira",
    3: "Quinta-feira",
    4: "Sexta-feira",
    5: "Sábado",
    6: "Domingo"
}
DIAS_ORDEM = [
    "Segunda-feira",
    "Terça-feira",
    "Quarta-feira",
    "Quinta-feira",
    "Sexta-feira",
    "Sábado",
    "Domingo"
]

EXCEL_PATH = r"C:\Users\Matheus\.gemini\antigravity\brain\61c61a76-f4c1-4467-a187-845b4be89ff3\excel_extract2\Base Requestia 2024-2026.xlsx"
PARQUET_PATH = "base_chamados.parquet"
SHEET = "BASE CHAMADO"

def fn(v): return f"{int(v):,}".replace(",", ".")

# ─── ESTILIZAÇÃO CSS (CLONE 100% WHITE / CLEAN) ──────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Segoe+UI:wght@300;400;600;700;800;900&family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"] {{
    font-family: 'Segoe UI', 'Inter', -apple-system, sans-serif;
    color: {TEXT_MAIN};
}}

.stApp {{
    background-color: {BG_PAGE};
}}

/* Esconder barra superior do Streamlit */
header[data-testid="stHeader"], [data-testid="stToolbar"] {{
    display: none !important;
}}
.block-container {{
    padding-top: 0rem !important;
    padding-bottom: 2rem !important;
    padding-left: 1.2rem !important;
    padding-right: 1.2rem !important;
    max-width: 100% !important;
}}

/* ── Ribbon Superior Vinho/Bordô ── */
.top-ribbon {{
    background: linear-gradient(90deg, #8a1414 0%, #6e0c0c 60%, #450808 100%);
    margin-left: -1.2rem;
    margin-right: -1.2rem;
    padding: 10px 24px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    box-shadow: 0 3px 8px rgba(0,0,0,0.12);
    margin-bottom: 12px;
}}
.ribbon-left {{
    display: flex;
    align-items: center;
    gap: 12px;
}}
.ribbon-smile {{
    width: 32px;
    height: 32px;
    border: 2px solid #ffffff;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #ffffff;
    font-size: 19px;
    font-weight: bold;
    line-height: 1;
}}
.ribbon-divider {{
    width: 1px;
    height: 24px;
    background-color: rgba(255,255,255,0.3);
}}
.ribbon-title {{
    color: #ffffff;
    font-size: 1.05rem;
    font-weight: 700;
    letter-spacing: 0.5px;
}}
.ribbon-badges {{
    display: flex;
    gap: 8px;
}}
.ribbon-badge-btn {{
    background-color: #1d4ed8;
    color: #ffffff;
    padding: 6px 16px;
    border-radius: 4px;
    font-size: 0.78rem;
    font-weight: 600;
    border: none;
}}

/* ── Linha de Filtros Branca com Logo Cinemark ── */
.filter-row-container {{
    background: {BG_CARD};
    border: 1px solid {BORDER_COLOR};
    border-radius: 6px;
    padding: 10px 18px 6px;
    margin-bottom: 12px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.03);
}}
.cinemark-brand-txt {{
    color: #b91c1c;
    font-size: 1.7rem;
    font-weight: 900;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    line-height: 1;
    margin-top: 12px;
}}
.cinemark-brand-txt span {{
    font-size: 0.75rem;
    vertical-align: super;
    margin-left: 2px;
}}

/* ── Cards de Gauge / Indicadores Topo ── */
.gauge-card {{
    background: {BG_CARD};
    border: 1px solid {BORDER_COLOR};
    border-radius: 6px;
    padding: 10px 12px 6px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.03);
    position: relative;
    overflow: hidden;
    height: 100%;
}}
.gauge-card::before {{
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    bottom: 0;
    width: 5px;
}}
.gauge-card.g-red::before    {{ background-color: #b91c1c; }}
.gauge-card.g-orange::before {{ background-color: #ea580c; }}
.gauge-card.g-green::before  {{ background-color: #16a34a; }}
.gauge-card.g-blue::before   {{ background-color: #2563eb; }}

.gauge-header {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-size: 0.8rem;
    font-weight: 700;
    color: {TEXT_MAIN};
    padding-left: 4px;
}}
.gauge-meta {{
    font-size: 0.78rem;
    color: {TEXT_MUTED};
    font-weight: 600;
}}

/* ── Cards de Gráficos ── */
.chart-box {{
    background: {BG_CARD};
    border: 1px solid {BORDER_COLOR};
    border-radius: 6px;
    padding: 14px 16px 8px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.03);
    margin-bottom: 12px;
    height: 100%;
}}
.chart-box-title {{
    font-size: 0.88rem;
    font-weight: 700;
    color: {TEXT_MAIN};
    line-height: 1.2;
}}
.chart-box-sub {{
    font-size: 0.72rem;
    color: {TEXT_MUTED};
    margin-bottom: 6px;
}}

/* ── Customização de inputs 100% Brancos Clean ── */
div[data-baseweb="select"] > div {{
    background-color: #ffffff !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 5px !important;
    color: #0f172a !important;
    font-size: 0.84rem !important;
}}
div[data-baseweb="select"] span {{
    color: #0f172a !important;
}}
div[data-baseweb="input"] {{
    background-color: #ffffff !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 5px !important;
    color: #0f172a !important;
}}
div[data-baseweb="input"] input {{
    background-color: #ffffff !important;
    color: #0f172a !important;
}}
label {{
    font-size: 0.75rem !important;
    font-weight: 600 !important;
    color: #475569 !important;
    margin-bottom: 2px !important;
}}

/* ── Estilização de Tabelas 100% Brancas Clean ── */
[data-testid="stDataFrame"] {{
    background-color: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 6px !important;
}}
[data-testid="stDataFrame"] div[role="grid"] {{
    background-color: #ffffff !important;
    color: #0f172a !important;
}}
/* Cabeçalho da tabela */
[data-testid="stDataFrame"] div[role="columnheader"] {{
    background-color: #f8fafc !important;
    color: #334155 !important;
    font-weight: 700 !important;
    border-bottom: 2px solid #e2e8f0 !important;
}}
/* Células e linhas */
[data-testid="stDataFrame"] div[role="gridcell"] {{
    background-color: #ffffff !important;
    color: #0f172a !important;
    border-bottom: 1px solid #f1f5f9 !important;
}}
/* Expander branco */
[data-testid="stExpander"] {{
    background-color: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 6px !important;
}}
</style>
""", unsafe_allow_html=True)

# ─── CARREGAMENTO DE DADOS ────────────────────────────────────────────────────
@st.cache_data(show_spinner="Carregando Base Cinemark...")
def load_data(source):
    import os
    # Prioriza o arquivo parquet ultracompacto se existir
    if isinstance(source, str) and (source.endswith(".parquet") or os.path.exists("base_chamados.parquet")):
        target = "base_chamados.parquet" if os.path.exists("base_chamados.parquet") else source
        df = pd.read_parquet(target)
    elif isinstance(source, str) and source.endswith(".xlsx"):
        df = pd.read_excel(source, sheet_name=SHEET, engine="openpyxl")
    else:
        # Upload em tempo real
        try:
            df = pd.read_parquet(source)
        except Exception:
            df = pd.read_excel(source, engine="openpyxl")

    df.columns = df.columns.str.strip()
    df["Tempo"]        = df["Tempo"].fillna("Em Aberto").str.strip()
    df["STATUS_LABEL"] = df["STATUS"].map(STATUS_MAP).fillna(df["STATUS"].fillna("N/A"))
    df["DATA_ABERTURA"]   = pd.to_datetime(df["DATA_ABERTURA"],   dayfirst=True, errors="coerce")
    df["DATA_FECHAMENTO"] = pd.to_datetime(df["DATA_FECHAMENTO"], dayfirst=True, errors="coerce")
    df["Tempo_horas"]  = pd.to_numeric(df["Tempo minutos"], errors="coerce") / 60
    df["Mês_Num"]      = pd.to_numeric(df["Mês"], errors="coerce").astype("Int64")
    df["Ano"]          = pd.to_numeric(df["Ano"], errors="coerce").astype("Int64")
    df["Mês_Label"]    = df["Mês_Num"].map(MESES).fillna("?")
    df["Período"]      = df["Ano"].astype(str) + "-" + df["Mês_Num"].apply(lambda x: f"{x:02d}" if pd.notna(x) else "??")
    df["GOC"]          = df["GOC"].fillna("Sem GOC").str.strip().replace("#N/A","Sem GOC")
    df["ANALISTA"]     = df["ANALISTA"].fillna("N/A").str.strip()
    df["Cinema"]       = df["Cinema"].fillna("N/A").str.strip().str.upper()
    df["Dia_Semana_Num"] = df["DATA_ABERTURA"].dt.dayofweek
    df["Dia_Semana"]     = df["Dia_Semana_Num"].map(DIAS_MAP)
    return df

# ─── GAUGE HALF-DONUT COM PONTEIRO/META (ESTILO FOTO) ─────────────────────────
def draw_half_gauge(value_pct, target_pct, color_hex="#b91c1c"):
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value_pct,
        number={
            "suffix": "%",
            "font": {"size": 26, "color": TEXT_MAIN, "family": "Segoe UI, sans-serif", "weight": "bold"}
        },
        gauge={
            "axis": {"range": [0, 100], "visible": False},
            "bar": {"color": color_hex, "thickness": 0.45},
            "bgcolor": "#e2e8f0",
            "borderwidth": 0,
            "shape": "angular",
            "threshold": {
                "line": {"color": "#0f172a", "width": 2.5},
                "thickness": 0.75,
                "value": target_pct
            }
        }
    ))
    fig.update_layout(
        height=105,
        margin=dict(t=5, b=0, l=12, r=12),
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig

# ─── LAYOUT BASE PLOTLY CLEAN ─────────────────────────────────────────────────
def get_box_layout(h=250, margin=None):
    m = margin or dict(t=20, b=30, l=40, r=15)
    return dict(
        height=h,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#ffffff",
        font=dict(color=TEXT_MAIN, family="Segoe UI, Inter, sans-serif", size=10),
        margin=m,
        xaxis=dict(
            gridcolor="#f1f5f9",
            linecolor="#cbd5e1",
            tickfont=dict(color=TEXT_MUTED, size=9),
            zeroline=False
        ),
        yaxis=dict(
            gridcolor="#f1f5f9",
            linecolor="#cbd5e1",
            tickfont=dict(color=TEXT_MUTED, size=9),
            zeroline=False
        ),
    )

# ─── CONTROLE DE ACESSO / LOGIN DA EQUIPE ─────────────────────────────────────
TEAM_PASSWORD = "cinemark2026"

def check_password():
    if st.session_state.get("authenticated", False):
        return True

    # Estilo do Card Centralizado de Login
    st.markdown("""
    <div style="max-width: 440px; margin: 40px auto 10px auto; background: #ffffff; border: 1px solid #cbd5e1; border-top: 5px solid #ba1c1c; border-radius: 8px; padding: 32px 30px; box-shadow: 0 4px 16px rgba(0,0,0,0.06); text-align: center;">
        <div style="color: #ba1c1c; font-size: 2.1rem; font-weight: 900; letter-spacing: 2px; text-transform: uppercase;">
            CINEMARK<span style="font-size: 0.9rem; vertical-align: super;">™</span>
        </div>
        <div style="font-size: 1rem; font-weight: 700; color: #0f172a; margin-top: 10px;">
            Painel de SLA & Metas Operacionais
        </div>
        <div style="font-size: 0.8rem; color: #64748b; margin-top: 4px;">
            Acesso exclusivo à equipe e liderança
        </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.3, 1])
    with col2:
        with st.form("login_form"):
            senha_digitada = st.text_input("Senha de Acesso", type="password", placeholder="Digite a senha...")
            submit_btn = st.form_submit_button("Acessar Painel ➔", use_container_width=True)

            if submit_btn:
                if senha_digitada == TEAM_PASSWORD:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("Senha incorreta. Solicite a senha correta ao responsável.")

        st.caption("🔒 Ambiente seguro e restrito · Cinemark Brasil")

    return False

# ─── APLICAÇÃO PRINCIPAL ───────────────────────────────────────────────────────
def main():
    if not check_password():
        return
    source = st.session_state.get("uploaded_file", EXCEL_PATH)
    try:
        df_raw = load_data(source)
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return

    datas = df_raw["DATA_ABERTURA"].dropna()
    min_d, max_d = datas.min().date(), datas.max().date()

    # 1. TOP RIBBON VINHO/BORDÔ
    st.markdown(f"""
    <div class="top-ribbon">
        <div class="ribbon-left">
            <div class="ribbon-smile">☺</div>
            <div class="ribbon-divider"></div>
            <div class="ribbon-title">CINEMARK BRASIL · PAINEL DE METAS E CUMPRIMENTO DE SLA</div>
        </div>
        <div class="ribbon-badges">
            <button class="ribbon-badge-btn">Meta Oficial: 95%</button>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 2. BARRA DE FILTROS CLEAN (FUNDO BRANCO) COM LOGO CINEMARK™
    with st.container():
        st.markdown('<div class="filter-row-container">', unsafe_allow_html=True)
        col_logo, col_f1, col_f2, col_f3, col_f4 = st.columns([1.6, 2.3, 2, 2, 1.8])
        
        with col_logo:
            st.markdown('<div class="cinemark-brand-txt">CINEMARK<span>™</span></div>', unsafe_allow_html=True)
        
        with col_f1:
            dt_range = st.date_input(
                "Período de Atendimento",
                value=(date(max_d.year, 1, 1), max_d),
                min_value=min_d,
                max_value=max_d,
                format="DD/MM/YYYY"
            )
            if isinstance(dt_range, tuple) and len(dt_range) == 2:
                dt_ini, dt_fim = dt_range
            else:
                dt_ini, dt_fim = min_d, max_d

        with col_f2:
            cinemas_lista = ["Todos"] + sorted(df_raw["Cinema"].unique())
            cinema_sel = st.selectbox("Cinema", cinemas_lista)

        with col_f3:
            gocs_lista = ["Todos"] + sorted([g for g in df_raw["GOC"].unique() if g != "Sem GOC"])
            goc_sel = st.selectbox("GOC / Gerente", gocs_lista)

        with col_f4:
            sla_opts = ["Todos"] + SLA_ORDEM
            sla_sel = st.selectbox("Faixa de SLA", sla_opts)
            
        st.markdown('</div>', unsafe_allow_html=True)

    # 3. APLICAÇÃO DOS FILTROS
    m = df_raw["DATA_ABERTURA"].dt.date.between(dt_ini, dt_fim)
    if cinema_sel != "Todos":
        m &= (df_raw["Cinema"] == cinema_sel)
    if goc_sel != "Todos":
        m &= (df_raw["GOC"] == goc_sel)
    if sla_sel != "Todos":
        m &= (df_raw["Tempo"] == sla_sel)
    
    df = df_raw[m].copy()
    total_chamados = len(df)

    if total_chamados == 0:
        st.warning("Nenhum chamado encontrado para os filtros selecionados.")
        return

    # 4. CARDS SUPERIORES: 4 TACÔMETROS COM A META DE 90%
    dentro_2h    = (df["Tempo"] == "até 2h").sum()
    dentro_24h   = (df["Tempo"] == "até 24h").sum()
    fora_sla     = (df["Tempo"] == "fora").sum()
    total_sla_ok = dentro_2h + dentro_24h

    pct_fora    = (fora_sla / total_chamados * 100) if total_chamados > 0 else 0
    pct_24h     = (dentro_24h / total_chamados * 100) if total_chamados > 0 else 0
    pct_2h      = (dentro_2h / total_chamados * 100) if total_chamados > 0 else 0
    pct_sla_ok  = (total_sla_ok / total_chamados * 100) if total_chamados > 0 else 0

    col_g1, col_g2, col_g3, col_g4 = st.columns(4)

    with col_g1:
        st.markdown("""
        <div class="gauge-card g-red">
            <div class="gauge-header">
                <span>Fora do SLA %</span>
                <span class="gauge-meta">Meta: &lt; 5,0%</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.plotly_chart(draw_half_gauge(pct_fora, 5.0, "#b91c1c"), use_container_width=True, config={"displayModeBar": False})

    with col_g2:
        st.markdown("""
        <div class="gauge-card g-blue">
            <div class="gauge-header">
                <span>SLA até 24h %</span>
                <span class="gauge-meta">Meta: 70,0%</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.plotly_chart(draw_half_gauge(pct_24h, 70.0, "#2563eb"), use_container_width=True, config={"displayModeBar": False})

    with col_g3:
        st.markdown("""
        <div class="gauge-card g-green">
            <div class="gauge-header">
                <span>SLA Imediato (&le; 2h) %</span>
                <span class="gauge-meta">Meta: 25,0%</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.plotly_chart(draw_half_gauge(pct_2h, 25.0, "#16a34a"), use_container_width=True, config={"displayModeBar": False})

    with col_g4:
        # Cor do gauge muda se bateu a meta de 95%!
        cor_meta_global = "#16a34a" if pct_sla_ok >= 95.0 else "#b91c1c"
        st.markdown(f"""
        <div class="gauge-card {'g-green' if pct_sla_ok >= 95.0 else 'g-red'}">
            <div class="gauge-header">
                <span>SLA Global Atendido %</span>
                <span class="gauge-meta">Meta: &ge; 95,0%</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.plotly_chart(draw_half_gauge(pct_sla_ok, 95.0, cor_meta_global), use_container_width=True, config={"displayModeBar": False})

    # 5. QUADRO DE GRÁFICOS: GRID 2x2 TOTALMENTE DISTINTOS E SEM REPETIÇÃO
    c_l1, c_r1 = st.columns([1.3, 1])

    # ── Gráfico 1 (Superior Esquerdo): Histórico de SLA % vs Meta Oficial (95%) ──
    with c_l1:
        st.markdown("""
        <div class="chart-box">
            <div class="chart-box-title">Histórico de SLA Geral (%) vs Meta de 95%</div>
            <div class="chart-box-sub">Identificação clara de meses que bateram (Verde) ou não bateram a meta (Vermelho)</div>
        </div>
        """, unsafe_allow_html=True)

        df_m = df.groupby("Período").agg(
            Total=("CHAMADO", "count"),
            Ok=("Tempo", lambda x: ((x == "até 2h") | (x == "até 24h")).sum())
        ).reset_index().sort_values("Período")
        df_m["Taxa_SLA"] = (df_m["Ok"] / df_m["Total"] * 100).round(1)
        df_m["Meta"] = 95.0
        df_m["Status_Meta"] = df_m["Taxa_SLA"].apply(lambda v: "Bateu Meta" if v >= 95.0 else "Abaixo da Meta")
        
        # Cores dos marcadores: Verde se >= 95%, Vermelho se < 95%
        marker_colors = ["#16a34a" if v >= 95.0 else "#dc2626" for v in df_m["Taxa_SLA"]]

        fig_meta = go.Figure()
        # Linha Curva de Realizado
        fig_meta.add_trace(go.Scatter(
            x=df_m["Período"], y=df_m["Taxa_SLA"],
            mode="lines+markers+text",
            name="SLA Realizado %",
            line=dict(color="#7f1d1d", width=3, shape="spline"),
            marker=dict(size=8, color=marker_colors, line=dict(width=1.5, color="#ffffff")),
            text=df_m["Taxa_SLA"].apply(lambda v: f"{v:.1f}%"),
            textposition="top center",
            textfont=dict(size=9.5, color="#0f172a", family="Segoe UI", weight="bold")
        ))
        # Linha de Meta 95% Pontilhada de Referência
        fig_meta.add_trace(go.Scatter(
            x=df_m["Período"], y=df_m["Meta"],
            mode="lines",
            name="Meta Mínima (95%)",
            line=dict(color="#2563eb", width=2, dash="dash")
        ))
        m_cfg = get_box_layout(h=250)
        m_cfg["legend"] = dict(orientation="h", y=-0.25, x=0.2, font=dict(size=9.5))
        m_cfg["yaxis"]["ticksuffix"] = "%"
        m_cfg["yaxis"]["range"] = [min(df_m["Taxa_SLA"].min() - 8, 70), 105]
        fig_meta.update_layout(**m_cfg)
        st.plotly_chart(fig_meta, use_container_width=True, config={"displayModeBar": False})

    # ── Gráfico 2 (Superior Direito): Distribuição por Faixa de SLA (Donut) ──
    with c_r1:
        st.markdown("""
        <div class="chart-box">
            <div class="chart-box-title">Distribuição Geral de Faixas de SLA</div>
            <div class="chart-box-sub">Volume consolidado no período selecionado</div>
        </div>
        """, unsafe_allow_html=True)

        cnt_sla = df["Tempo"].value_counts().reindex(SLA_ORDEM).dropna().reset_index()
        cnt_sla.columns = ["Faixa", "Qtd"]
        
        # Mapeamento com nomes explicativos para a legenda
        legenda_faixa = {
            "até 2h": "Até 2 horas (Imediato)",
            "até 24h": "Até 24 horas",
            "até 48h": "Até 48 horas",
            "fora": "Fora do SLA (>24h)",
            "Em Aberto": "Em Aberto"
        }
        cnt_sla["Legenda"] = cnt_sla["Faixa"].map(legenda_faixa).fillna(cnt_sla["Faixa"])

        cores_map = {legenda_faixa.get(k, k): v for k, v in SLA_CORES.items()}

        fig_pie = px.pie(
            cnt_sla, values="Qtd", names="Legenda", hole=0.62,
            color="Legenda", color_discrete_map=cores_map
        )
        fig_pie.update_traces(
            textposition="inside",
            textinfo="percent",
            textfont=dict(size=10.5, color="#ffffff", family="Segoe UI", weight="bold"),
            marker=dict(line=dict(color="#ffffff", width=2))
        )
        fig_pie.update_layout(
            height=250,
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(t=10, b=10, l=10, r=10),
            showlegend=True,
            legend=dict(
                orientation="v",
                yanchor="middle",
                y=0.5,
                xanchor="left",
                x=1.02,
                font=dict(size=10, color=TEXT_MAIN)
            ),
            annotations=[dict(
                text=f"<b>{fn(total_chamados)}</b><br><span style='font-size:10px;color:#64748b'>Chamados</span>",
                x=0.38, y=0.5, showarrow=False, font=dict(size=13, color=TEXT_MAIN)
            )]
        )
        st.plotly_chart(fig_pie, use_container_width=True, config={"displayModeBar": False})

    # ── Linha Inferior: 2 Gráficos ÚNICOS (Analistas à esquerda, Cinemas à direita) ──
    c_l2, c_r2 = st.columns([1.1, 1.2])

    # ── Gráfico 3 (Inferior Esquerdo): Top Analistas (Excluindo vjesus e tmatsuo) ──
    with c_l2:
        st.markdown("""
        <div class="chart-box">
            <div class="chart-box-title">Top Analistas por Volume</div>
            <div class="chart-box-sub">Equipe com maior quantidade de chamados atendidos</div>
        </div>
        """, unsafe_allow_html=True)

        top_a = df[(df["ANALISTA"] != "N/A") & (~df["ANALISTA"].isin(["vjesus", "tmatsuo"]))].groupby("ANALISTA").size().sort_values(ascending=False).head(10).reset_index(name="Total")
        top_a = top_a.sort_values("Total", ascending=True)

        fig_an = px.bar(
            top_a, x="Total", y="ANALISTA", orientation="h",
            color_discrete_sequence=["#2563eb"],
            text="Total"
        )
        fig_an.update_traces(
            texttemplate="%{x:,}",
            textposition="outside",
            textfont=dict(size=9, color=TEXT_MUTED)
        )
        l2_cfg = get_box_layout(h=250, margin=dict(t=10, b=20, l=95, r=35))
        l2_cfg["xaxis"]["title"] = dict(text="")
        l2_cfg["yaxis"]["title"] = dict(text="")
        fig_an.update_layout(**l2_cfg)
        st.plotly_chart(fig_an, use_container_width=True, config={"displayModeBar": False})

    # ── Gráfico 4 (Inferior Direito): Top 10 Cinemas ──
    with c_r2:
        st.markdown("""
        <div class="chart-box">
            <div class="chart-box-title">Top 10 Cinemas com Maior Demanda</div>
            <div class="chart-box-sub">Complexos mais acionados da rede</div>
        </div>
        """, unsafe_allow_html=True)

        top_cin = df.groupby("Cinema").size().sort_values(ascending=False).head(10).reset_index(name="Total")
        top_cin = top_cin.sort_values("Total", ascending=True)

        fig_cin = px.bar(
            top_cin, x="Total", y="Cinema", orientation="h",
            color_discrete_sequence=["#ba1c1c"],
            text="Total"
        )
        fig_cin.update_traces(
            texttemplate="%{x:,}",
            textposition="outside",
            textfont=dict(size=9, color=TEXT_MUTED)
        )
        r2_cfg = get_box_layout(h=250, margin=dict(t=10, b=20, l=130, r=40))
        r2_cfg["xaxis"]["title"] = dict(text="")
        r2_cfg["yaxis"]["title"] = dict(text="")
        fig_cin.update_layout(**r2_cfg)
        st.plotly_chart(fig_cin, use_container_width=True, config={"displayModeBar": False})

    # 6. SEÇÃO DEDICADA: ANÁLISE DE VOLUMETRIA E SLA POR DIA DA SEMANA
    st.markdown("""
    <div style="margin-top: 14px; margin-bottom: 8px;">
        <div style="font-size: 0.95rem; font-weight: 800; color: #0f172a; text-transform: uppercase; letter-spacing: 0.5px;">
            📅 Análise de Chamados e Demanda por Dia da Semana
        </div>
        <div style="font-size: 0.75rem; color: #64748b;">
            Identificação de dias com maior pico de chamados e conformidade operacional de Segunda a Domingo
        </div>
    </div>
    """, unsafe_allow_html=True)

    df_dias = df.dropna(subset=["Dia_Semana_Num"]).groupby(["Dia_Semana_Num", "Dia_Semana"]).agg(
        Total=("CHAMADO", "count"),
        Ok=("Tempo", lambda x: ((x == "até 2h") | (x == "até 24h")).sum()),
        Fora=("Tempo", lambda x: (x == "fora").sum())
    ).reset_index().sort_values("Dia_Semana_Num")

    if not df_dias.empty:
        total_periodo = df_dias["Total"].sum()
        df_dias["Pct_Vol"] = (df_dias["Total"] / total_periodo * 100).round(1)
        df_dias["Taxa_SLA"] = (df_dias["Ok"] / df_dias["Total"] * 100).round(1)

        dia_pico_idx = df_dias["Total"].idxmax()
        dia_pico_nome = df_dias.loc[dia_pico_idx, "Dia_Semana"]
        dia_pico_total = df_dias.loc[dia_pico_idx, "Total"]
        dia_pico_pct = df_dias.loc[dia_pico_idx, "Pct_Vol"]

        c_sem1, c_sem2 = st.columns([1.2, 1.2])

        # ── Gráfico 1: Volumetria com Destaque no Dia de Pico ──
        with c_sem1:
            st.markdown(f"""
            <div class="chart-box">
                <div class="chart-box-title">Volumetria de Chamados por Dia da Semana</div>
                <div class="chart-box-sub">Dia de Maior Pico: <b>{dia_pico_nome}</b> com <b>{fn(dia_pico_total)}</b> chamados ({dia_pico_pct:.1f}% da demanda)</div>
            </div>
            """, unsafe_allow_html=True)

            bar_colors = ["#ba1c1c" if d == dia_pico_nome else "#2563eb" for d in df_dias["Dia_Semana"]]

            fig_vol_dia = go.Figure()
            fig_vol_dia.add_trace(go.Bar(
                x=df_dias["Dia_Semana"],
                y=df_dias["Total"],
                marker=dict(color=bar_colors, line=dict(color="#ffffff", width=1.5)),
                text=[f"{fn(t)}<br><span style='font-size:9.5px;color:#64748b'>({p:.1f}%)</span>" for t, p in zip(df_dias["Total"], df_dias["Pct_Vol"])],
                textposition="outside",
                textfont=dict(size=9.5, color=TEXT_MAIN, weight="bold"),
                hovertemplate="<b>%{x}</b><br>Total de Chamados: %{y:,}<extra></extra>"
            ))
            
            cfg_vd = get_box_layout(h=265, margin=dict(t=25, b=25, l=45, r=20))
            cfg_vd["yaxis"]["range"] = [0, df_dias["Total"].max() * 1.22]
            fig_vol_dia.update_layout(**cfg_vd)
            st.plotly_chart(fig_vol_dia, use_container_width=True, config={"displayModeBar": False})

        # ── Gráfico 2: Composição de Faixas de SLA por Dia da Semana ──
        with c_sem2:
            st.markdown("""
            <div class="chart-box">
                <div class="chart-box-title">Composição de Faixas de SLA por Dia da Semana</div>
                <div class="chart-box-sub">Volume de atendimentos imediatos (≤2h), no prazo (≤24h) e fora do SLA por dia</div>
            </div>
            """, unsafe_allow_html=True)

            cross_sla = df.dropna(subset=["Dia_Semana_Num"]).groupby(["Dia_Semana_Num", "Dia_Semana", "Tempo"]).size().unstack(fill_value=0).reset_index().sort_values("Dia_Semana_Num")
            
            fig_sla_comp = go.Figure()
            ordem_exibicao = [
                ("até 2h", "Até 2h (Imediato)", "#16a34a"),
                ("até 24h", "Até 24h", "#2563eb"),
                ("até 48h", "Até 48h", "#ea580c"),
                ("fora", "Fora do SLA (>24h)", "#ba1c1c")
            ]

            for col_faixa, label_nome, cor in ordem_exibicao:
                if col_faixa in cross_sla.columns:
                    fig_sla_comp.add_trace(go.Bar(
                        name=label_nome,
                        x=cross_sla["Dia_Semana"],
                        y=cross_sla[col_faixa],
                        marker_color=cor
                    ))

            cfg_sc = get_box_layout(h=265, margin=dict(t=15, b=30, l=45, r=20))
            cfg_sc["barmode"] = "stack"
            cfg_sc["legend"] = dict(
                orientation="h",
                y=-0.22,
                x=0.05,
                font=dict(size=9.5, color=TEXT_MAIN)
            )
            fig_sla_comp.update_layout(**cfg_sc)
            st.plotly_chart(fig_sla_comp, use_container_width=True, config={"displayModeBar": False})

    # 7. SEÇÃO DEDICADA: ANÁLISE DETALHADA POR ANALISTA (PERFORMANCE INDIVIDUAL)
    st.markdown("""
    <div style="margin-top: 14px; margin-bottom: 8px;">
        <div style="font-size: 0.95rem; font-weight: 800; color: #0f172a; text-transform: uppercase; letter-spacing: 0.5px;">
            👨‍💼 Desempenho e SLA por Analista
        </div>
        <div style="font-size: 0.75rem; color: #64748b;">
            Taxa de cumprimento da meta de 95%, tempo de resposta e volume atendido por profissional
        </div>
    </div>
    """, unsafe_allow_html=True)

    col_an_g1, col_an_g2 = st.columns([1.3, 1])

    # Analistas a ignorar/excluir da análise
    ANALISTAS_EXCLUIR = ["vjesus", "tmatsuo"]

    # Agregação completa por Analista (excluindo os solicitados e N/A)
    df_an_filtrado = df[(df["ANALISTA"] != "N/A") & (~df["ANALISTA"].isin(ANALISTAS_EXCLUIR))]

    df_an_full = df_an_filtrado.groupby("ANALISTA").agg(
        Total=("CHAMADO", "count"),
        No_Prazo=("Tempo", lambda x: ((x == "até 2h") | (x == "até 24h")).sum()),
        Imediato_2h=("Tempo", lambda x: (x == "até 2h").sum()),
        Fora_SLA=("Tempo", lambda x: (x == "fora").sum()),
        Mediana_h=("Tempo_horas", "median")
    ).reset_index()

    df_an_full["% SLA Geral"] = (df_an_full["No_Prazo"] / df_an_full["Total"] * 100).round(1)
    df_an_full["% Até 2h"] = (df_an_full["Imediato_2h"] / df_an_full["Total"] * 100).round(1)
    df_an_full["Mediana (h)"] = df_an_full["Mediana_h"].round(1)
    
    # Filtra os com maior volume para manter os gráficos legíveis
    top_12_an = df_an_full.sort_values("Total", ascending=False).head(10).sort_values("% SLA Geral", ascending=True)

    with col_an_g1:
        st.markdown("""
        <div class="chart-box">
            <div class="chart-box-title">% de Cumprimento de SLA por Analista (Meta 95%)</div>
            <div class="chart-box-sub">Verde: Atingiu meta &ge;95% | Vermelho: Abaixo de 95%</div>
        </div>
        """, unsafe_allow_html=True)

        cores_sla_an = ["#16a34a" if v >= 95.0 else "#dc2626" for v in top_12_an["% SLA Geral"]]

        fig_an_sla = px.bar(
            top_12_an, x="% SLA Geral", y="ANALISTA", orientation="h",
            color="% SLA Geral",
            color_continuous_scale=["#dc2626", "#ea580c", "#16a34a"],
            range_color=[80, 100],
            text=top_12_an["% SLA Geral"].apply(lambda v: f"{v:.1f}%")
        )
        fig_an_sla.update_traces(
            textposition="outside",
            textfont=dict(size=9.5, color=TEXT_MAIN, weight="bold")
        )
        an_sla_cfg = get_box_layout(h=290, margin=dict(t=10, b=20, l=100, r=40))
        an_sla_cfg["xaxis"]["ticksuffix"] = "%"
        an_sla_cfg["xaxis"]["range"] = [70, 105]
        an_sla_cfg["xaxis"]["title"] = dict(text="")
        an_sla_cfg["yaxis"]["title"] = dict(text="")
        an_sla_cfg["coloraxis_showscale"] = False
        fig_an_sla.update_layout(**an_sla_cfg)
        st.plotly_chart(fig_an_sla, use_container_width=True, config={"displayModeBar": False})

    with col_an_g2:
        st.markdown("""
        <div class="chart-box">
            <div class="chart-box-title">Tempo Mediano de Atendimento (Horas)</div>
            <div class="chart-box-sub">Horas necessárias por analista até a resolução</div>
        </div>
        """, unsafe_allow_html=True)

        top_12_tempo = df_an_full.sort_values("Total", ascending=False).head(12).sort_values("Mediana (h)", ascending=False)
        
        fig_an_time = px.bar(
            top_12_tempo, x="Mediana (h)", y="ANALISTA", orientation="h",
            color_discrete_sequence=["#ea580c"],
            text=top_12_tempo["Mediana (h)"].apply(lambda v: f"{v:.1f}h" if pd.notna(v) else "—")
        )
        fig_an_time.update_traces(
            textposition="outside",
            textfont=dict(size=9, color=TEXT_MUTED)
        )
        an_time_cfg = get_box_layout(h=290, margin=dict(t=10, b=20, l=100, r=40))
        an_time_cfg["xaxis"]["title"] = dict(text="")
        an_time_cfg["yaxis"]["title"] = dict(text="")
        fig_an_time.update_layout(**an_time_cfg)
        st.plotly_chart(fig_an_time, use_container_width=True, config={"displayModeBar": False})

    # Tabela Executiva Consolidada por Analista
    st.markdown("""
    <div class="chart-box" style="margin-bottom: 14px;">
        <div class="chart-box-title">Quadro Comparativo de Analistas</div>
        <div class="chart-box-sub">Tabela com métricas de volume, taxa de cumprimento da meta de 95% e resoluções imediatas</div>
    </div>
    """, unsafe_allow_html=True)

    df_an_table = df_an_full.rename(columns={
        "ANALISTA": "Analista",
        "Total": "Total Chamados",
        "No_Prazo": "Atendidos no Prazo",
        "Imediato_2h": "Até 2h (Imediato)",
        "Fora_SLA": "Fora do SLA",
        "% SLA Geral": "% SLA Geral",
        "% Até 2h": "% Imediato (≤2h)",
        "Mediana (h)": "Tempo Mediano (h)"
    }).sort_values("Total Chamados", ascending=False)

    def style_analista(val):
        if isinstance(val, (int, float)):
            if val >= 95.0:
                return "color: #16a34a; font-weight: bold; background-color: #f0fdf4;"
            elif val < 90.0:
                return "color: #dc2626; font-weight: bold; background-color: #fef2f2;"
        return ""

    st.dataframe(
        df_an_table[["Analista", "Total Chamados", "Atendidos no Prazo", "Fora do SLA", "% SLA Geral", "% Imediato (≤2h)", "Tempo Mediano (h)"]].style.map(
            style_analista, subset=["% SLA Geral"]
        ),
        use_container_width=True,
        hide_index=True
    )

    # 7. TABELA EXPANSÍVEL NO RODAPÉ COM EXPORTAÇÃO
    with st.expander("📋 Visualizar Base Completa de Chamados Filtrada"):
        st.dataframe(
            df[["CHAMADO", "DATA_ABERTURA", "DATA_FECHAMENTO", "Tempo", "ANALISTA", "GOC", "Cinema", "STATUS_LABEL"]].head(500),
            use_container_width=True,
            hide_index=True
        )
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="Chamados")
        st.download_button(
            "⬇️ Exportar Dados para Excel (.xlsx)",
            buf.getvalue(),
            "chamados_cinemark.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
