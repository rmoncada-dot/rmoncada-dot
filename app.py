import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook
import io, os

st.set_page_config(
    page_title="Portfolio AM — Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #f4f6fa; }
  [data-testid="stSidebar"] { background: #ffffff; border-right: 1px solid #e0e5f0; }
  [data-testid="stSidebar"] * { color: #1e2d5f !important; }

  .metric-card {
    background: #ffffff;
    border-radius: 12px;
    padding: 18px 20px;
    margin: 6px 0;
    border-left: 4px solid #2e75b6;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
  }
  .metric-card.green  { border-left-color: #2d7a4f; }
  .metric-card.blue   { border-left-color: #2e75b6; }
  .metric-card.purple { border-left-color: #7030a0; }
  .metric-card.gold   { border-left-color: #c8a227; }
  .metric-card.red    { border-left-color: #c00000; }

  .metric-val  { font-size: 26px; font-weight: 700; color: #1e2d5f; margin: 4px 0; }
  .metric-lbl  { font-size: 10px; color: #6b7a99; text-transform: uppercase; letter-spacing: 1px; }
  .metric-sub  { font-size: 12px; color: #2e75b6; margin-top: 2px; }

  .section-title {
    font-size: 13px; font-weight: 700; color: #1e2d5f;
    text-transform: uppercase; letter-spacing: 2px;
    border-bottom: 2px solid #2e75b6;
    padding-bottom: 6px; margin: 24px 0 16px 0;
  }
  h1, h2, h3 { color: #1e2d5f !important; }
  [data-testid="stMarkdownContainer"] p { color: #3a4a65; }
  div[data-testid="stTabs"] button { color: #3a4a65; font-weight: 600; }
  div[data-testid="stTabs"] button[aria-selected="true"] { color: #1e2d5f; border-bottom: 2px solid #2e75b6; }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data(file_content):
    wb = load_workbook(io.BytesIO(file_content), data_only=True)
    wsFD = wb['Fonte_Dati']
    fd = []
    for row in range(8, wsFD.max_row+1):
        mese=wsFD.cell(row=row,column=3).value; imp=wsFD.cell(row=row,column=4).value
        en=wsFD.cell(row=row,column=5).value;   magg=wsFD.cell(row=row,column=6).value
        fl=wsFD.cell(row=row,column=7).value;   fn=wsFD.cell(row=row,column=9).value
        tipo=wsFD.cell(row=row,column=11).value
        if imp:
            fd.append({'Mese':mese,'Impianto':imp,'En_Mis_kWh':en or 0,'En_Magg_kWh':magg or 0,
                       'Fat_Lordo':fl or 0,'Fat_Netto':fn or 0,'Tipo':tipo or ''})
    df_fd = pd.DataFrame(fd)

    wsDB = wb['DB_Impianti']
    db = []
    for row in range(8, 24):
        nome=wsDB.cell(row=row,column=3).value; tipo=wsDB.cell(row=row,column=4).value
        kwp=wsDB.cell(row=row,column=11).value; fn_q1=wsDB.cell(row=row,column=18).value
        en_q1=wsDB.cell(row=row,column=14).value
        if nome:
            db.append({'Impianto':nome,'Tipo':tipo or '','kWp':kwp,
                       'FatNetto_Q1':fn_q1 or 0,'EnMis_Q1':en_q1 or 0})
    df_db = pd.DataFrame(db)

    wsI = wb['💹 Incentivi']
    acc, con = [], []
    for row in range(8, 24):
        site=wsI.cell(row=row,column=4).value; tot=wsI.cell(row=row,column=18).value
        g=wsI.cell(row=row,column=6).value; f=wsI.cell(row=row,column=7).value; m=wsI.cell(row=row,column=8).value
        if site: acc.append({'Site':site,'Gen':g or 0,'Feb':f or 0,'Mar':m or 0,'Tot':tot or 0})
    for row in range(28, 44):
        site=wsI.cell(row=row,column=4).value; tot=wsI.cell(row=row,column=18).value
        g=wsI.cell(row=row,column=6).value; f=wsI.cell(row=row,column=7).value; m=wsI.cell(row=row,column=8).value
        if site: con.append({'Site':site,'Gen':g or 0,'Feb':f or 0,'Mar':m or 0,'Tot':tot or 0})
    df_acc = pd.DataFrame(acc); df_con = pd.DataFrame(con)

    wsAP = wb['Analisi_Perdite']
    ap = []
    for row in range(41, wsAP.max_row+1):
        nome=wsAP.cell(row=row,column=3).value; tipo=wsAP.cell(row=row,column=4).value
        mese=wsAP.cell(row=row,column=5).value; et=wsAP.cell(row=row,column=9).value
        er=wsAP.cell(row=row,column=10).value;  de=wsAP.cell(row=row,column=11).value
        dep=wsAP.cell(row=row,column=12).value
        if nome and mese and tipo:
            ap.append({'Impianto':nome,'Tipo':tipo,'Mese':mese,'E_Teorica':et or 0,
                       'E_Reale':er or 0,'Delta_MWh':de or 0,'Delta_pct':dep or 0})
    df_ap = pd.DataFrame(ap)
    return df_fd, df_db, df_acc, df_con, df_ap

# Sidebar
with st.sidebar:
    # Logo
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logo.jpg')
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
    else:
        st.markdown("### ⚡ Portfolio AM")
    st.divider()
    uploaded = st.file_uploader("📂 Carica Excel aggiornato", type=['xlsx','xlsm'])
    if uploaded:
        file_bytes = uploaded.read()
        st.success(f"✅ {uploaded.name}")
    else:
        # Percorso assoluto relativo alla posizione di app.py
        base_dir = os.path.dirname(os.path.abspath(__file__))
        default_file = os.path.join(base_dir, 'portfolio_integrato_Q1_2026.xlsx')
        if os.path.exists(default_file):
            with open(default_file, 'rb') as f:
                file_bytes = f.read()
            st.info("📊 Dati Q1 2026")
        else:
            st.error("⚠ File Excel non trovato. Carica il file dalla sidebar.")
            st.stop()
    st.divider()
    st.markdown("**Filtri**")
    show_tipo = st.multiselect("Tipo impianto",["Fotovoltaico","Eolico"],default=["Fotovoltaico","Eolico"])
    show_mese = st.multiselect("Mese",["Gennaio","Febbraio","Marzo"],default=["Gennaio","Febbraio","Marzo"])

df_fd, df_db, df_acc, df_con, df_ap = load_data(file_bytes)
df_filt = df_fd[df_fd['Tipo'].isin(show_tipo) & df_fd['Mese'].isin(show_mese)]

tot_fn   = df_filt['Fat_Netto'].sum()
tot_magg = df_filt['En_Magg_kWh'].sum()
eur_mwh  = (tot_fn/(tot_magg/1000)) if tot_magg > 0 else 0
tot_acc  = df_acc['Tot'].sum() if not df_acc.empty else 0
tot_con  = df_con['Tot'].sum() if not df_con.empty else 0
n_fv     = df_db['Tipo'].value_counts().get('Fotovoltaico',0)
n_eo     = df_db['Tipo'].value_counts().get('Eolico',0)

# Header
col_logo, col_title = st.columns([1, 4])
with col_logo:
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logo.jpg')
    if os.path.exists(logo_path):
        st.image(logo_path, width=180)
with col_title:
    st.markdown("# ⚡ Portfolio Rinnovabili — Dashboard Q1 2026")
    st.markdown("13 Impianti FV + 3 Eolici &nbsp;|&nbsp; Q1 2026")
st.divider()

def kpi(label, value, sub="", color="blue"):
    return f"""<div class="metric-card {color}">
        <div class="metric-lbl">{label}</div>
        <div class="metric-val">{value}</div>
        <div class="metric-sub">{sub}</div>
    </div>"""

c1,c2,c3,c4,c5,c6 = st.columns(6)
with c1: st.markdown(kpi("⚡ Energia Misurata", f"{df_filt['En_Mis_kWh'].sum()/1000:,.0f} MWh","Q1 2026","blue"), unsafe_allow_html=True)
with c2: st.markdown(kpi("💰 Fatturato Netto",  f"€ {tot_fn:,.0f}","Q1 2026","green"), unsafe_allow_html=True)
with c3: st.markdown(kpi("📈 EUR/MWh",          f"€ {eur_mwh:,.2f}","Prezzo medio","gold"), unsafe_allow_html=True)
with c4: st.markdown(kpi("💹 Acconto GSE",       f"€ {tot_acc:,.0f}","ADVANCE-GSE × Tariffa","purple"), unsafe_allow_html=True)
with c5: st.markdown(kpi("✅ Consuntivo GSE",    f"€ {tot_con:,.0f}","OWNER × Tariffa","green"), unsafe_allow_html=True)
with c6: st.markdown(kpi("🏭 Impianti",          f"{n_fv} FV + {n_eo} Eolici","Portfolio attivo","blue"), unsafe_allow_html=True)

st.divider()

COLORS     = {"Fotovoltaico":"#2d7a4f","Eolico":"#065a82"}
MESI_ORDER = ["Gennaio","Febbraio","Marzo"]
PLOT_BG    = "rgba(244,246,250,0)"   # trasparente su sfondo chiaro
GRID_COLOR = "#dde3ef"
FONT_COLOR = "#3a4a65"

def chart_layout(fig, height=360):
    fig.update_layout(
        plot_bgcolor=PLOT_BG, paper_bgcolor=PLOT_BG,
        font_color=FONT_COLOR, height=height,
        xaxis=dict(gridcolor=GRID_COLOR, linecolor=GRID_COLOR),
        yaxis=dict(gridcolor=GRID_COLOR, linecolor=GRID_COLOR),
        legend=dict(bgcolor="rgba(255,255,255,0.8)", bordercolor=GRID_COLOR, borderwidth=1),
    )
    return fig

tab1,tab2,tab3,tab4,tab5 = st.tabs(["📊 Produzione","💰 Finanziario","💹 Incentivi","⚡ Analisi Perdite","🏭 Impianti"])

# ═══ TAB 1 ═══════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-title">Produzione Energetica Mensile</div>', unsafe_allow_html=True)
    df_mese = df_fd[df_fd['Mese'].isin(show_mese)].groupby(['Mese','Tipo'])['En_Mis_kWh'].sum().reset_index()
    df_mese['MWh'] = df_mese['En_Mis_kWh']/1000
    df_mese['Mese'] = pd.Categorical(df_mese['Mese'],categories=MESI_ORDER,ordered=True)
    df_mese = df_mese.sort_values('Mese')

    col1,col2 = st.columns([2,1])
    with col1:
        fig = px.bar(df_mese,x='Mese',y='MWh',color='Tipo',color_discrete_map=COLORS,
                     barmode='stack',text_auto='.0f',labels={'MWh':'Energia (MWh)'})
        fig.update_traces(textfont_color='white',textposition='inside')
        st.plotly_chart(chart_layout(fig), use_container_width=True)
    with col2:
        df_pie = df_fd[df_fd['Tipo'].isin(show_tipo)].groupby('Tipo')['En_Mis_kWh'].sum().reset_index()
        fig2 = px.pie(df_pie,names='Tipo',values='En_Mis_kWh',color='Tipo',
                      color_discrete_map=COLORS,hole=0.5)
        fig2.update_layout(plot_bgcolor=PLOT_BG,paper_bgcolor=PLOT_BG,font_color=FONT_COLOR,height=360,
                           legend=dict(orientation='h',yanchor='bottom',y=-0.25))
        fig2.update_traces(textinfo='percent+label',textfont_size=12)
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown('<div class="section-title">Top Impianti per Produzione Q1</div>', unsafe_allow_html=True)
    df_top = df_fd[df_fd['Tipo'].isin(show_tipo)].groupby(['Impianto','Tipo'])['En_Mis_kWh'].sum().reset_index()
    df_top['MWh'] = df_top['En_Mis_kWh']/1000
    df_top = df_top.sort_values('MWh',ascending=True)
    fig3 = px.bar(df_top,y='Impianto',x='MWh',color='Tipo',color_discrete_map=COLORS,
                  orientation='h',text_auto='.0f',labels={'MWh':'MWh','Impianto':''})
    fig3.update_traces(textfont_color='white')
    st.plotly_chart(chart_layout(fig3,height=500), use_container_width=True)

# ═══ TAB 2 ═══════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">Performance Finanziaria</div>', unsafe_allow_html=True)
    col1,col2 = st.columns(2)
    with col1:
        df_fin = df_fd[df_fd['Mese'].isin(show_mese)].groupby(['Mese','Tipo'])['Fat_Netto'].sum().reset_index()
        df_fin['Mese'] = pd.Categorical(df_fin['Mese'],categories=MESI_ORDER,ordered=True)
        df_fin = df_fin.sort_values('Mese')
        fig = px.bar(df_fin,x='Mese',y='Fat_Netto',color='Tipo',color_discrete_map=COLORS,
                     barmode='stack',text_auto=',.0f',labels={'Fat_Netto':'Fatturato Netto (€)'})
        fig.update_traces(textfont_color='white',textposition='inside')
        st.plotly_chart(chart_layout(fig), use_container_width=True)
    with col2:
        df_eur = df_fd[df_fd['Tipo'].isin(show_tipo)].groupby('Impianto').agg(
            Fat_Netto=('Fat_Netto','sum'),En_Magg=('En_Magg_kWh','sum')).reset_index()
        df_eur['EUR_MWh'] = df_eur['Fat_Netto']/(df_eur['En_Magg']/1000)
        df_eur = df_eur[df_eur['EUR_MWh']>0].sort_values('EUR_MWh',ascending=False).head(12)
        fig2 = px.bar(df_eur,x='Impianto',y='EUR_MWh',
                      color='EUR_MWh',color_continuous_scale=[[0,'#d4e9ff'],[1,'#1e2d5f']],
                      text_auto='.1f',labels={'EUR_MWh':'€/MWh'})
        fig2.update_layout(xaxis_tickangle=45,coloraxis_showscale=False)
        fig2.update_traces(textfont_color='white')
        st.plotly_chart(chart_layout(fig2), use_container_width=True)

    st.markdown('<div class="section-title">Dettaglio per Impianto Q1</div>', unsafe_allow_html=True)
    df_tbl = df_fd[df_fd['Tipo'].isin(show_tipo)].groupby(['Impianto','Tipo']).agg(
        En_MWh=('En_Mis_kWh',lambda x: x.sum()/1000),
        Fat_Lordo=('Fat_Lordo','sum'), Fat_Netto=('Fat_Netto','sum')).reset_index()
    df_tbl['EUR_MWh'] = df_tbl['Fat_Netto']/df_tbl['En_MWh']
    df_tbl = df_tbl.sort_values('Fat_Netto',ascending=False)
    df_tbl.columns = ['Impianto','Tipo','MWh','Fat. Lordo (€)','Fat. Netto (€)','€/MWh']
    st.dataframe(
        df_tbl.style.format({'MWh':'{:,.1f}','Fat. Lordo (€)':'{:,.0f}','Fat. Netto (€)':'{:,.0f}','€/MWh':'{:,.2f}'}),
        use_container_width=True, hide_index=True)

# ═══ TAB 3 ═══════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-title">KPI Incentivi GSE — Q1 2026</div>', unsafe_allow_html=True)
    c1,c2,c3 = st.columns(3)
    delta = tot_con - tot_acc
    with c1: st.metric("📊 Acconto Totale",   f"€ {tot_acc:,.0f}")
    with c2: st.metric("✅ Consuntivo Totale", f"€ {tot_con:,.0f}")
    with c3: st.metric("📐 Delta Cons.−Acc.",  f"€ {delta:+,.0f}", delta=f"{delta/tot_acc*100:+.1f}%" if tot_acc else "—")

    if not df_acc.empty and not df_con.empty:
        col1,col2 = st.columns(2)
        with col1:
            st.markdown('<div class="section-title">Acconto Mensile per Impianto</div>', unsafe_allow_html=True)
            df_am = df_acc.melt(id_vars='Site',value_vars=['Gen','Feb','Mar'],var_name='Mese',value_name='€')
            df_am = df_am[df_am['€']>0]
            fig = px.bar(df_am,x='Site',y='€',color='Mese',barmode='group',
                         color_discrete_map={'Gen':'#2e75b6','Feb':'#375623','Mar':'#7030a0'},
                         labels={'€':'€','Site':''})
            fig.update_layout(xaxis_tickangle=45)
            st.plotly_chart(chart_layout(fig,380), use_container_width=True)
        with col2:
            st.markdown('<div class="section-title">Acconto vs Consuntivo Q1</div>', unsafe_allow_html=True)
            df_cmp = pd.merge(df_acc[['Site','Tot']].rename(columns={'Tot':'Acconto'}),
                              df_con[['Site','Tot']].rename(columns={'Tot':'Consuntivo'}),
                              on='Site',how='outer').fillna(0)
            df_cmp = df_cmp[df_cmp['Acconto']+df_cmp['Consuntivo']>0]
            df_cm = df_cmp.melt(id_vars='Site',var_name='Tipo',value_name='€')
            fig2 = px.bar(df_cm,x='Site',y='€',color='Tipo',barmode='group',
                          color_discrete_map={'Acconto':'#c8a227','Consuntivo':'#2d7a4f'},
                          labels={'Site':''})
            fig2.update_layout(xaxis_tickangle=45)
            st.plotly_chart(chart_layout(fig2,380), use_container_width=True)

# ═══ TAB 4 ═══════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-title">Analisi Perdite — Teorica vs Reale</div>', unsafe_allow_html=True)
    if not df_ap.empty:
        df_apf = df_ap[df_ap['Tipo'].isin(show_tipo) & df_ap['Mese'].isin(show_mese)]
        tot_teo=df_apf['E_Teorica'].sum(); tot_real=df_apf['E_Reale'].sum()
        tot_delt=tot_teo-tot_real
        c1,c2,c3,c4 = st.columns(4)
        with c1: st.metric("E. Teorica",  f"{tot_teo:,.0f} MWh")
        with c2: st.metric("E. Reale",    f"{tot_real:,.0f} MWh")
        with c3: st.metric("ΔE perdita",  f"{tot_delt:+,.0f} MWh")
        with c4: st.metric("ΔE %", f"{tot_delt/tot_teo*100:+.1f}%" if tot_teo else "—")

        col1,col2 = st.columns(2)
        with col1:
            df_ca = df_apf.groupby(['Impianto','Tipo']).agg(
                E_Teorica=('E_Teorica','sum'),E_Reale=('E_Reale','sum')).reset_index()
            df_ca = df_ca[df_ca['E_Teorica']>0]
            df_cm = df_ca.melt(id_vars=['Impianto','Tipo'],value_vars=['E_Teorica','E_Reale'],
                               var_name='Cat',value_name='MWh')
            fig = px.bar(df_cm,x='Impianto',y='MWh',color='Cat',barmode='group',
                         color_discrete_map={'E_Teorica':'#2e75b6','E_Reale':'#2d7a4f'},
                         labels={'Impianto':'','MWh':'MWh'})
            fig.update_layout(xaxis_tickangle=45)
            st.plotly_chart(chart_layout(fig,400), use_container_width=True)
        with col2:
            df_pct = df_apf.groupby(['Impianto','Tipo']).agg(
                E_Teorica=('E_Teorica','sum'),E_Reale=('E_Reale','sum')).reset_index()
            df_pct = df_pct[df_pct['E_Teorica']>0]
            df_pct['Perdita_%'] = (df_pct['E_Teorica']-df_pct['E_Reale'])/df_pct['E_Teorica']*100
            df_pct = df_pct.sort_values('Perdita_%',ascending=False)
            fig2 = px.bar(df_pct,x='Impianto',y='Perdita_%',color='Tipo',
                          color_discrete_map=COLORS,text_auto='.1f',labels={'Impianto':'','Perdita_%':'Perdita (%)'})
            fig2.update_traces(textfont_color='white')
            fig2.update_layout(xaxis_tickangle=45)
            fig2.add_hline(y=20,line_dash="dash",line_color="#c00000",annotation_text="⚠ 20%")
            fig2.add_hline(y=8, line_dash="dash",line_color="#c8a227",annotation_text="⚡ 8%")
            st.plotly_chart(chart_layout(fig2,400), use_container_width=True)

        st.markdown('<div class="section-title">Dettaglio Mensile</div>', unsafe_allow_html=True)
        df_ta = df_apf[['Impianto','Tipo','Mese','E_Teorica','E_Reale','Delta_MWh','Delta_pct']].copy()
        df_ta = df_ta[df_ta['E_Teorica']>0]
        df_ta.columns = ['Impianto','Tipo','Mese','E Teorica (MWh)','E Reale (MWh)','ΔE (MWh)','ΔE (%)']
        st.dataframe(
            df_ta.style.format({'E Teorica (MWh)':'{:,.1f}','E Reale (MWh)':'{:,.1f}',
                                'ΔE (MWh)':'{:+,.1f}','ΔE (%)':'{:+.1%}'}),
            use_container_width=True, hide_index=True)
    else:
        st.info("Dati analisi perdite non disponibili — inserire kWp DC in DB_Impianti")

# ═══ TAB 5 ═══════════════════════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="section-title">Anagrafica & Status Impianti</div>', unsafe_allow_html=True)
    df_imp = df_db[df_db['Tipo'].isin(show_tipo)].copy()
    df_fn = df_fd[df_fd['Tipo'].isin(show_tipo)].groupby('Impianto').agg(
        En_MWh=('En_Mis_kWh',lambda x: x.sum()/1000),Fat_Netto=('Fat_Netto','sum')).reset_index()
    df_imp = df_imp.merge(df_fn,on='Impianto',how='left')
    df_imp['EUR_MWh'] = df_imp['Fat_Netto']/df_imp['En_MWh']

    fig = px.scatter(df_imp[df_imp['En_MWh']>0],x='En_MWh',y='Fat_Netto',
                     color='Tipo',size='En_MWh',color_discrete_map=COLORS,
                     hover_name='Impianto',text='Impianto',
                     labels={'En_MWh':'Energia Q1 (MWh)','Fat_Netto':'Fatturato Netto Q1 (€)'})
    fig.update_traces(textposition='top center',textfont_size=9,textfont_color=FONT_COLOR)
    st.plotly_chart(chart_layout(fig,450), use_container_width=True)

    df_show = df_imp[['Impianto','Tipo','kWp','En_MWh','Fat_Netto','EUR_MWh']].copy()
    df_show.columns = ['Impianto','Tipo','kWp DC','MWh Q1','Fat. Netto Q1 (€)','€/MWh']
    st.dataframe(
        df_show.style.format({
            'kWp DC':lambda x: f'{x:,.0f}' if pd.notna(x) and x else '—',
            'MWh Q1':'{:,.1f}','Fat. Netto Q1 (€)':'{:,.0f}','€/MWh':'{:,.2f}'}),
        use_container_width=True, hide_index=True)

st.divider()
st.markdown("<center><small style='color:#8a9ab5'>Portfolio AM Dashboard &nbsp;|&nbsp; Q1 2026</small></center>", unsafe_allow_html=True)
