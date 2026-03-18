import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from openpyxl import load_workbook
import io, os

# ── Config pagina ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Portfolio AM — Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS personalizzato ────────────────────────────────────────────────────────
st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #0f1117; }
  [data-testid="stSidebar"] { background: #1a1f2e; }
  .metric-card {
    background: linear-gradient(135deg, #1e2d5f 0%, #2e3a5c 100%);
    border-radius: 12px; padding: 20px; margin: 6px 0;
    border-left: 4px solid #c8a227;
  }
  .metric-card.green  { border-left-color: #2d7a4f; }
  .metric-card.blue   { border-left-color: #2e75b6; }
  .metric-card.purple { border-left-color: #7030a0; }
  .metric-card.gold   { border-left-color: #c8a227; }
  .metric-val  { font-size: 28px; font-weight: 700; color: #ffffff; margin: 4px 0; }
  .metric-lbl  { font-size: 11px; color: #8a9ab5; text-transform: uppercase; letter-spacing: 1px; }
  .metric-sub  { font-size: 12px; color: #c8a227; margin-top: 2px; }
  .section-title {
    font-size: 14px; font-weight: 700; color: #c8a227;
    text-transform: uppercase; letter-spacing: 2px;
    border-bottom: 1px solid #2e3a5c; padding-bottom: 8px; margin: 24px 0 16px 0;
  }
  div[data-testid="metric-container"] { background: #1a1f2e; border-radius: 8px; padding: 12px; }
</style>
""", unsafe_allow_html=True)

# ── Caricamento dati ──────────────────────────────────────────────────────────
@st.cache_data
def load_data(file_content):
    wb = load_workbook(io.BytesIO(file_content), data_only=True)

    # ── Fonte_Dati ──────────────────────────────────────────────────────────
    wsFD = wb['Fonte_Dati']
    fd = []
    for row in range(8, wsFD.max_row+1):
        mese  = wsFD.cell(row=row, column=3).value
        imp   = wsFD.cell(row=row, column=4).value
        en    = wsFD.cell(row=row, column=5).value
        magg  = wsFD.cell(row=row, column=6).value
        fl    = wsFD.cell(row=row, column=7).value
        fn    = wsFD.cell(row=row, column=9).value
        tipo  = wsFD.cell(row=row, column=11).value
        if imp:
            fd.append({'Mese':mese,'Impianto':imp,
                       'En_Mis_kWh': en or 0,'En_Magg_kWh': magg or 0,
                       'Fat_Lordo': fl or 0,'Fat_Netto': fn or 0,
                       'Tipo': tipo or ''})
    df_fd = pd.DataFrame(fd)

    # ── DB Impianti ─────────────────────────────────────────────────────────
    wsDB = wb['DB_Impianti']
    db = []
    for row in range(8, 24):
        nome = wsDB.cell(row=row, column=3).value
        tipo = wsDB.cell(row=row, column=4).value
        kwp  = wsDB.cell(row=row, column=11).value
        fn_q1= wsDB.cell(row=row, column=18).value
        en_q1= wsDB.cell(row=row, column=14).value
        if nome:
            db.append({'Impianto':nome,'Tipo':tipo or '','kWp':kwp,
                       'FatNetto_Q1': fn_q1 or 0, 'EnMis_Q1': en_q1 or 0})
    df_db = pd.DataFrame(db)

    # ── Incentivi Acconto ───────────────────────────────────────────────────
    wsI = wb['💹 Incentivi']
    acc, con = [], []
    for row in range(8, 24):
        site = wsI.cell(row=row, column=4).value
        tot  = wsI.cell(row=row, column=18).value
        g    = wsI.cell(row=row, column=6).value
        f    = wsI.cell(row=row, column=7).value
        m    = wsI.cell(row=row, column=8).value
        if site:
            acc.append({'Site':site,'Gen':g or 0,'Feb':f or 0,'Mar':m or 0,'Tot':tot or 0})
    for row in range(28, 44):
        site = wsI.cell(row=row, column=4).value
        tot  = wsI.cell(row=row, column=18).value
        g    = wsI.cell(row=row, column=6).value
        f    = wsI.cell(row=row, column=7).value
        m    = wsI.cell(row=row, column=8).value
        if site:
            con.append({'Site':site,'Gen':g or 0,'Feb':f or 0,'Mar':m or 0,'Tot':tot or 0})
    df_acc = pd.DataFrame(acc)
    df_con = pd.DataFrame(con)

    # ── Analisi Perdite ─────────────────────────────────────────────────────
    wsAP = wb['Analisi_Perdite']
    ap = []
    for row in range(41, wsAP.max_row+1):
        nome  = wsAP.cell(row=row, column=3).value
        tipo  = wsAP.cell(row=row, column=4).value
        mese  = wsAP.cell(row=row, column=5).value
        input_= wsAP.cell(row=row, column=7).value
        et    = wsAP.cell(row=row, column=9).value
        er    = wsAP.cell(row=row, column=10).value
        de    = wsAP.cell(row=row, column=11).value
        dep   = wsAP.cell(row=row, column=12).value
        if nome and mese and tipo:
            ap.append({'Impianto':nome,'Tipo':tipo,'Mese':mese,
                       'Input':input_ or 0,'E_Teorica':et or 0,
                       'E_Reale':er or 0,'Delta_MWh':de or 0,'Delta_pct':dep or 0})
    df_ap = pd.DataFrame(ap)

    return df_fd, df_db, df_acc, df_con, df_ap

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚡ Portfolio AM")
    st.markdown("**Moncada Energy Group**")
    st.markdown("*Agrigento — Sicilia*")
    st.divider()

    uploaded = st.file_uploader("📂 Carica Excel aggiornato", type=['xlsx','xlsm'])
    if uploaded:
        file_bytes = uploaded.read()
        st.success(f"✅ {uploaded.name}")
    else:
        with open('/home/claude/portfolio_integrato_Q1_2026.xlsx','rb') as f:
            file_bytes = f.read()
        st.info("📊 Dati Q1 2026 (file demo)")

    st.divider()
    st.markdown("**Filtri**")
    show_tipo = st.multiselect("Tipo impianto",["Fotovoltaico","Eolico"],
                                default=["Fotovoltaico","Eolico"])
    show_mese = st.multiselect("Mese", ["Gennaio","Febbraio","Marzo"],
                                default=["Gennaio","Febbraio","Marzo"])

# ── Carica dati ───────────────────────────────────────────────────────────────
df_fd, df_db, df_acc, df_con, df_ap = load_data(file_bytes)

# Filtri applicati
df_filt = df_fd[df_fd['Tipo'].isin(show_tipo) & df_fd['Mese'].isin(show_mese)]

# KPI aggregati
tot_en   = df_filt['En_Mis_kWh'].sum() / 1_000_000  # TWh
tot_fn   = df_filt['Fat_Netto'].sum()
tot_fl   = df_filt['Fat_Lordo'].sum()
tot_magg = df_filt['En_Magg_kWh'].sum()
eur_mwh  = (tot_fn / (tot_magg/1000)) if tot_magg > 0 else 0
tot_acc  = df_acc['Tot'].sum() if not df_acc.empty else 0
tot_con  = df_con['Tot'].sum() if not df_con.empty else 0

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("# ⚡ Portfolio Rinnovabili — Dashboard Q1 2026")
st.markdown("**Gruppo Moncada** | 13 Impianti FV + 3 Eolici | Provincia di Agrigento, Sicilia")
st.divider()

# ── KPI Cards ─────────────────────────────────────────────────────────────────
def kpi(label, value, sub="", color="gold"):
    return f"""<div class="metric-card {color}">
        <div class="metric-lbl">{label}</div>
        <div class="metric-val">{value}</div>
        <div class="metric-sub">{sub}</div>
    </div>"""

c1,c2,c3,c4,c5,c6 = st.columns(6)
with c1: st.markdown(kpi("⚡ Energia Misurata",f"{df_filt['En_Mis_kWh'].sum()/1000:,.0f} MWh","Q1 2026","blue"), unsafe_allow_html=True)
with c2: st.markdown(kpi("💰 Fatturato Netto",f"€ {tot_fn:,.0f}","Q1 2026","green"), unsafe_allow_html=True)
with c3: st.markdown(kpi("📈 EUR/MWh",f"€ {eur_mwh:,.2f}","Prezzo medio","gold"), unsafe_allow_html=True)
with c4: st.markdown(kpi("💹 Acconto GSE",f"€ {tot_acc:,.0f}","ADVANCE-GSE × Tariffa","purple"), unsafe_allow_html=True)
with c5: st.markdown(kpi("✅ Consuntivo GSE",f"€ {tot_con:,.0f}","OWNER × Tariffa","green"), unsafe_allow_html=True)
with c6: st.markdown(kpi("🏭 Impianti",f"{df_db['Tipo'].value_counts().get('Fotovoltaico',0)} FV + {df_db['Tipo'].value_counts().get('Eolico',0)} Eo.","Portfolio attivo","blue"), unsafe_allow_html=True)

st.divider()

# ── Tab navigation ────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Produzione", "💰 Finanziario", "💹 Incentivi", "⚡ Analisi Perdite", "🏭 Impianti"
])

COLORS = {"Fotovoltaico":"#2d7a4f","Eolico":"#065a82"}
MESI_ORDER = ["Gennaio","Febbraio","Marzo"]

# ═══ TAB 1: PRODUZIONE ═══════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-title">Produzione Energetica Mensile</div>', unsafe_allow_html=True)

    # Stacked bar mensile
    df_mese = df_fd[df_fd['Mese'].isin(show_mese)].groupby(['Mese','Tipo'])['En_Mis_kWh'].sum().reset_index()
    df_mese['MWh'] = df_mese['En_Mis_kWh'] / 1000
    df_mese['Mese'] = pd.Categorical(df_mese['Mese'], categories=MESI_ORDER, ordered=True)
    df_mese = df_mese.sort_values('Mese')

    fig_bar = px.bar(df_mese, x='Mese', y='MWh', color='Tipo',
                     color_discrete_map=COLORS, barmode='stack',
                     labels={'MWh':'Energia (MWh)'},
                     text_auto='.0f')
    fig_bar.update_layout(
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        font_color='#c8d3e8', legend_title_text='',
        xaxis=dict(gridcolor='#2e3a5c'), yaxis=dict(gridcolor='#2e3a5c'),
        height=350
    )
    fig_bar.update_traces(textfont_color='white', textposition='inside')

    col1, col2 = st.columns([2,1])
    with col1:
        st.plotly_chart(fig_bar, use_container_width=True)
    with col2:
        # Torta FV vs Eolico
        df_pie = df_fd[df_fd['Tipo'].isin(show_tipo)].groupby('Tipo')['En_Mis_kWh'].sum().reset_index()
        fig_pie = px.pie(df_pie, names='Tipo', values='En_Mis_kWh',
                         color='Tipo', color_discrete_map=COLORS,
                         hole=0.5)
        fig_pie.update_layout(
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            font_color='#c8d3e8', showlegend=True, height=350,
            legend=dict(orientation='h', yanchor='bottom', y=-0.2)
        )
        fig_pie.update_traces(textinfo='percent+label', textfont_color='white')
        st.plotly_chart(fig_pie, use_container_width=True)

    # Top impianti
    st.markdown('<div class="section-title">Top Impianti per Produzione Q1</div>', unsafe_allow_html=True)
    df_top = df_fd[df_fd['Tipo'].isin(show_tipo)].groupby(['Impianto','Tipo'])['En_Mis_kWh'].sum().reset_index()
    df_top['MWh'] = df_top['En_Mis_kWh']/1000
    df_top = df_top.sort_values('MWh', ascending=True)
    fig_horiz = px.bar(df_top, y='Impianto', x='MWh', color='Tipo',
                       color_discrete_map=COLORS, orientation='h',
                       labels={'MWh':'Energia (MWh)'}, text_auto='.0f')
    fig_horiz.update_layout(
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        font_color='#c8d3e8', showlegend=False, height=500,
        xaxis=dict(gridcolor='#2e3a5c'), yaxis=dict(gridcolor='#2e3a5c')
    )
    fig_horiz.update_traces(textfont_color='white')
    st.plotly_chart(fig_horiz, use_container_width=True)

# ═══ TAB 2: FINANZIARIO ══════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">Performance Finanziaria</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        # Fatturato mensile
        df_fin_m = df_fd[df_fd['Mese'].isin(show_mese)].groupby(['Mese','Tipo'])['Fat_Netto'].sum().reset_index()
        df_fin_m['Mese'] = pd.Categorical(df_fin_m['Mese'], categories=MESI_ORDER, ordered=True)
        df_fin_m = df_fin_m.sort_values('Mese')
        fig_fin = px.bar(df_fin_m, x='Mese', y='Fat_Netto', color='Tipo',
                         color_discrete_map=COLORS, barmode='stack',
                         labels={'Fat_Netto':'Fatturato Netto (€)'}, text_auto=',.0f')
        fig_fin.update_layout(
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            font_color='#c8d3e8', legend_title_text='', height=350,
            xaxis=dict(gridcolor='#2e3a5c'), yaxis=dict(gridcolor='#2e3a5c')
        )
        fig_fin.update_traces(textfont_color='white', textposition='inside')
        st.plotly_chart(fig_fin, use_container_width=True)

    with col2:
        # EUR/MWh per impianto
        df_eur = df_fd[df_fd['Tipo'].isin(show_tipo)].groupby('Impianto').agg(
            Fat_Netto=('Fat_Netto','sum'), En_Magg=('En_Magg_kWh','sum')).reset_index()
        df_eur['EUR_MWh'] = df_eur['Fat_Netto'] / (df_eur['En_Magg']/1000)
        df_eur = df_eur.sort_values('EUR_MWh', ascending=False).head(10)
        fig_eur = px.bar(df_eur, x='Impianto', y='EUR_MWh',
                         color='EUR_MWh', color_continuous_scale='Viridis',
                         labels={'EUR_MWh':'€/MWh'}, text_auto='.1f')
        fig_eur.update_layout(
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            font_color='#c8d3e8', height=350, showlegend=False,
            xaxis=dict(gridcolor='#2e3a5c', tickangle=45),
            yaxis=dict(gridcolor='#2e3a5c')
        )
        fig_eur.update_traces(textfont_color='white')
        st.plotly_chart(fig_eur, use_container_width=True)

    # Tabella riepilogativa
    st.markdown('<div class="section-title">Dettaglio Finanziario per Impianto Q1</div>', unsafe_allow_html=True)
    df_tbl = df_fd[df_fd['Tipo'].isin(show_tipo)].groupby(['Impianto','Tipo']).agg(
        En_MWh=('En_Mis_kWh', lambda x: x.sum()/1000),
        Fat_Lordo=('Fat_Lordo','sum'),
        Fat_Netto=('Fat_Netto','sum')
    ).reset_index()
    df_tbl['EUR_MWh'] = df_tbl['Fat_Netto'] / df_tbl['En_MWh']
    df_tbl = df_tbl.sort_values('Fat_Netto', ascending=False)
    df_tbl.columns = ['Impianto','Tipo','MWh','Fat. Lordo (€)','Fat. Netto (€)','€/MWh']
    st.dataframe(
        df_tbl.style
            .format({'MWh':'{:,.1f}','Fat. Lordo (€)':'{:,.0f}','Fat. Netto (€)':'{:,.0f}','€/MWh':'{:,.2f}'}),
        use_container_width=True, hide_index=True
    )

# ═══ TAB 3: INCENTIVI ════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-title">KPI Incentivi GSE — Q1 2026</div>', unsafe_allow_html=True)

    c1,c2,c3 = st.columns(3)
    with c1:
        st.metric("📊 Acconto Totale (ADVANCE-GSE)", f"€ {tot_acc:,.0f}")
    with c2:
        st.metric("✅ Consuntivo Totale (OWNER)", f"€ {tot_con:,.0f}")
    with c3:
        delta = tot_con - tot_acc
        st.metric("📐 Delta Cons.−Acc.", f"€ {delta:+,.0f}", delta=f"{delta/tot_acc*100:+.1f}%" if tot_acc else "—")

    if not df_acc.empty and not df_con.empty:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown('<div class="section-title">Acconto per Impianto (Gen/Feb/Mar)</div>', unsafe_allow_html=True)
            df_acc_m = df_acc.melt(id_vars='Site', value_vars=['Gen','Feb','Mar'],
                                   var_name='Mese', value_name='€')
            df_acc_m = df_acc_m[df_acc_m['€'] > 0]
            fig_acc = px.bar(df_acc_m, x='Site', y='€', color='Mese',
                             color_discrete_map={'Gen':'#2e75b6','Feb':'#375623','Mar':'#7030a0'},
                             barmode='group', labels={'€':'€','Site':''})
            fig_acc.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#c8d3e8', height=380,
                xaxis=dict(tickangle=45, gridcolor='#2e3a5c'),
                yaxis=dict(gridcolor='#2e3a5c')
            )
            st.plotly_chart(fig_acc, use_container_width=True)

        with col2:
            st.markdown('<div class="section-title">Acconto vs Consuntivo Q1</div>', unsafe_allow_html=True)
            df_cmp = pd.merge(
                df_acc[['Site','Tot']].rename(columns={'Tot':'Acconto'}),
                df_con[['Site','Tot']].rename(columns={'Tot':'Consuntivo'}),
                on='Site', how='outer'
            ).fillna(0)
            df_cmp = df_cmp[df_cmp['Acconto'] + df_cmp['Consuntivo'] > 0]
            df_cmp_m = df_cmp.melt(id_vars='Site', var_name='Tipo', value_name='€')
            fig_cmp = px.bar(df_cmp_m, x='Site', y='€', color='Tipo',
                             color_discrete_map={'Acconto':'#c8a227','Consuntivo':'#2d7a4f'},
                             barmode='group', labels={'Site':''})
            fig_cmp.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#c8d3e8', height=380,
                xaxis=dict(tickangle=45, gridcolor='#2e3a5c'),
                yaxis=dict(gridcolor='#2e3a5c')
            )
            st.plotly_chart(fig_cmp, use_container_width=True)

# ═══ TAB 4: ANALISI PERDITE ══════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-title">Analisi Perdite Energetiche — Teorica vs Reale</div>', unsafe_allow_html=True)

    if not df_ap.empty:
        df_ap_f = df_ap[df_ap['Tipo'].isin(show_tipo) & df_ap['Mese'].isin(show_mese)]

        c1,c2,c3,c4 = st.columns(4)
        tot_teo  = df_ap_f['E_Teorica'].sum()
        tot_real = df_ap_f['E_Reale'].sum()
        tot_delt = tot_teo - tot_real
        with c1: st.metric("E. Teorica totale", f"{tot_teo:,.0f} MWh")
        with c2: st.metric("E. Reale totale",   f"{tot_real:,.0f} MWh")
        with c3: st.metric("ΔE perdita",        f"{tot_delt:+,.0f} MWh")
        with c4: st.metric("ΔE %",              f"{tot_delt/tot_teo*100:+.1f}%" if tot_teo else "—")

        col1, col2 = st.columns(2)
        with col1:
            # Teorica vs Reale per impianto
            df_cmp_ap = df_ap_f.groupby(['Impianto','Tipo']).agg(
                E_Teorica=('E_Teorica','sum'), E_Reale=('E_Reale','sum')).reset_index()
            df_cmp_ap = df_cmp_ap[df_cmp_ap['E_Teorica'] > 0]
            df_cmp_m = df_cmp_ap.melt(id_vars=['Impianto','Tipo'],
                                       value_vars=['E_Teorica','E_Reale'],
                                       var_name='Categoria', value_name='MWh')
            fig_loss = px.bar(df_cmp_m, x='Impianto', y='MWh', color='Categoria',
                              color_discrete_map={'E_Teorica':'#2e75b6','E_Reale':'#2d7a4f'},
                              barmode='group', labels={'Impianto':''})
            fig_loss.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#c8d3e8', height=400,
                xaxis=dict(tickangle=45, gridcolor='#2e3a5c'),
                yaxis=dict(gridcolor='#2e3a5c', title='MWh')
            )
            st.plotly_chart(fig_loss, use_container_width=True)

        with col2:
            # % perdita per impianto
            df_pct = df_ap_f.groupby(['Impianto','Tipo']).agg(
                E_Teorica=('E_Teorica','sum'), E_Reale=('E_Reale','sum')).reset_index()
            df_pct = df_pct[df_pct['E_Teorica'] > 0]
            df_pct['Delta_pct'] = (df_pct['E_Teorica'] - df_pct['E_Reale']) / df_pct['E_Teorica'] * 100
            df_pct = df_pct.sort_values('Delta_pct', ascending=False)
            fig_pct = px.bar(df_pct, x='Impianto', y='Delta_pct',
                             color='Tipo', color_discrete_map=COLORS,
                             labels={'Delta_pct':'Perdita %','Impianto':''},
                             text_auto='.1f')
            fig_pct.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#c8d3e8', height=400, showlegend=False,
                xaxis=dict(tickangle=45, gridcolor='#2e3a5c'),
                yaxis=dict(gridcolor='#2e3a5c', title='Perdita (%)')
            )
            fig_pct.update_traces(textfont_color='white')
            fig_pct.add_hline(y=20, line_dash="dash", line_color="#c00000",
                              annotation_text="⚠ Soglia 20%")
            fig_pct.add_hline(y=8, line_dash="dash", line_color="#c8a227",
                              annotation_text="⚡ Soglia 8%")
            st.plotly_chart(fig_pct, use_container_width=True)

        # Tabella dettaglio
        st.markdown('<div class="section-title">Dettaglio Mensile</div>', unsafe_allow_html=True)
        df_tbl_ap = df_ap_f[['Impianto','Tipo','Mese','E_Teorica','E_Reale','Delta_MWh','Delta_pct']].copy()
        df_tbl_ap = df_tbl_ap[df_tbl_ap['E_Teorica'] > 0]
        df_tbl_ap.columns = ['Impianto','Tipo','Mese','E Teorica (MWh)','E Reale (MWh)','ΔE (MWh)','ΔE (%)']
        st.dataframe(
            df_tbl_ap.style
                .format({'E Teorica (MWh)':'{:,.1f}','E Reale (MWh)':'{:,.1f}',
                         'ΔE (MWh)':'{:+,.1f}','ΔE (%)':'{:+.1%}'}),
            use_container_width=True, hide_index=True
        )
    else:
        st.info("Dati analisi perdite non disponibili — inserire kWp DC in DB_Impianti")

# ═══ TAB 5: IMPIANTI ═════════════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="section-title">Anagrafica & Status Impianti</div>', unsafe_allow_html=True)

    df_imp = df_db[df_db['Tipo'].isin(show_tipo)].copy()
    df_imp_fn = df_fd[df_fd['Tipo'].isin(show_tipo)].groupby('Impianto').agg(
        En_MWh=('En_Mis_kWh', lambda x: x.sum()/1000),
        Fat_Netto=('Fat_Netto','sum')
    ).reset_index()
    df_imp = df_imp.merge(df_imp_fn, on='Impianto', how='left')
    df_imp['EUR_MWh'] = df_imp['Fat_Netto'] / df_imp['En_MWh']
    df_imp['Status'] = df_imp['En_MWh'].apply(lambda x: '✅ Attivo' if x and x > 0 else '⚠ No dati')

    # Scatter plot: En vs Fatturato
    fig_sc = px.scatter(df_imp[df_imp['En_MWh']>0], x='En_MWh', y='Fat_Netto',
                        color='Tipo', size='En_MWh',
                        color_discrete_map=COLORS,
                        hover_name='Impianto',
                        labels={'En_MWh':'Energia Q1 (MWh)','Fat_Netto':'Fatturato Netto Q1 (€)'},
                        text='Impianto')
    fig_sc.update_traces(textposition='top center', textfont_size=9)
    fig_sc.update_layout(
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        font_color='#c8d3e8', height=450,
        xaxis=dict(gridcolor='#2e3a5c'), yaxis=dict(gridcolor='#2e3a5c')
    )
    st.plotly_chart(fig_sc, use_container_width=True)

    # Tabella
    df_show = df_imp[['Impianto','Tipo','kWp','En_MWh','Fat_Netto','EUR_MWh','Status']].copy()
    df_show.columns = ['Impianto','Tipo','kWp DC','MWh Q1','Fat. Netto Q1 (€)','€/MWh','Status']
    st.dataframe(
        df_show.style
            .format({'kWp DC':lambda x: f'{x:,.0f}' if pd.notna(x) and x else '—',
                     'MWh Q1':'{:,.1f}','Fat. Netto Q1 (€)':'{:,.0f}','€/MWh':'{:,.2f}'}),
        use_container_width=True, hide_index=True
    )

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.markdown("<center><small style='color:#4a5a75'>Portfolio AM Dashboard | Q1 2026 | Moncada Energy Group</small></center>", unsafe_allow_html=True)
