"""
MAG Tchad – Système de Gestion de l'Information
Développé par : DJAOYANG HABEKREO Pelandi | Ingénieur Statisticien Économiste
Poste visé  : Responsable de la Gestion de l'Information
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import folium
from streamlit_folium import st_folium
from datetime import datetime, timedelta
import io, random, json
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ─────────────────────────────────────────────
# CONFIG PAGE
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="MAG Tchad – Gestion de l'Information",
    page_icon="💣",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────
# CSS CUSTOM
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    
    .main { background-color: #F0F4F8; }
    
    .stApp header { background-color: #1F4E79; }
    
    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1F4E79 0%, #2E75B6 100%);
    }
    section[data-testid="stSidebar"] * { color: white !important; }
    section[data-testid="stSidebar"] .stRadio label { 
        background: rgba(255,255,255,0.1); 
        border-radius: 8px; padding: 6px 10px; margin: 3px 0;
        display: block; cursor: pointer;
    }
    section[data-testid="stSidebar"] .stRadio label:hover { background: rgba(255,255,255,0.25); }

    /* KPI Cards */
    .kpi-card {
        background: white;
        border-radius: 12px;
        padding: 20px 24px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        border-left: 5px solid #2E75B6;
        margin-bottom: 8px;
    }
    .kpi-value { font-size: 2.2rem; font-weight: 700; color: #1F4E79; line-height: 1; }
    .kpi-label { font-size: 0.85rem; color: #666; margin-top: 6px; font-weight: 500; }
    .kpi-delta { font-size: 0.8rem; margin-top: 4px; }
    .kpi-up { color: #16A34A; } .kpi-down { color: #DC2626; }

    .kpi-card.red   { border-left-color: #DC2626; }
    .kpi-card.green { border-left-color: #16A34A; }
    .kpi-card.orange{ border-left-color: #EA580C; }

    /* Section headers */
    .section-header {
        background: linear-gradient(90deg, #1F4E79, #2E75B6);
        color: white !important;
        padding: 10px 20px;
        border-radius: 8px;
        font-size: 1.05rem;
        font-weight: 600;
        margin: 18px 0 12px 0;
    }

    /* Alert boxes */
    .alert-danger  { background:#FEF2F2; border-left:4px solid #DC2626; padding:12px 16px; border-radius:6px; color:#991B1B; margin:6px 0; }
    .alert-warning { background:#FFFBEB; border-left:4px solid #F59E0B; padding:12px 16px; border-radius:6px; color:#92400E; margin:6px 0; }
    .alert-success { background:#F0FDF4; border-left:4px solid #16A34A; padding:12px 16px; border-radius:6px; color:#166534; margin:6px 0; }
    .alert-info    { background:#EFF6FF; border-left:4px solid #2E75B6; padding:12px 16px; border-radius:6px; color:#1E40AF; margin:6px 0; }

    /* Table */
    .dataframe thead th { background-color: #1F4E79 !important; color: white !important; }
    
    /* Badge */
    .badge { display:inline-block; padding:2px 10px; border-radius:99px; font-size:0.75rem; font-weight:600; }
    .badge-green  { background:#DCFCE7; color:#166534; }
    .badge-red    { background:#FEE2E2; color:#991B1B; }
    .badge-orange { background:#FEF3C7; color:#92400E; }

    /* Footer */
    .footer { text-align:center; color:#999; font-size:0.75rem; padding:20px 0 10px; border-top:1px solid #E5E7EB; margin-top:30px; }

    div[data-testid="stMetric"] { background:white; border-radius:10px; padding:12px; box-shadow:0 1px 6px rgba(0,0,0,0.07); }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# DONNÉES FICTIVES RÉALISTES – TCHAD/MAG
# ─────────────────────────────────────────────
np.random.seed(42)
random.seed(42)

PROVINCES = ["Lac", "Bornou", "Hadjer-Lamis", "Kanem", "Batha", "Ouaddaï"]
ACTIVITES = ["Survey Non-Technique (SNT)", "Dépollution (EOD)", "Éducation aux Risques (ER)",
             "Destruction AMD", "Marquage & Signalisation", "Libération de Surface"]
STATUTS = ["Complété", "En cours", "Planifié", "Suspendu"]
EQUIPES = ["Alpha", "Bravo", "Charlie", "Delta"]
DANGERS = ["UXO", "Mine AP", "Mine AC", "IED", "Sous-munition", "Grenade"]

def gen_sites(n=120):
    rows = []
    base_date = datetime(2025, 1, 1)
    for i in range(n):
        prov = random.choice(PROVINCES)
        lat_base = {"Lac": 13.2, "Bornou": 13.8, "Hadjer-Lamis": 12.8,
                    "Kanem": 14.5, "Batha": 13.6, "Ouaddaï": 13.0}[prov]
        lon_base = {"Lac": 14.2, "Bornou": 15.1, "Hadjer-Lamis": 15.8,
                    "Kanem": 15.6, "Batha": 17.2, "Ouaddaï": 20.8}[prov]
        rows.append({
            "ID_Site": f"TCH-{2025}-{str(i+1).zfill(4)}",
            "Province": prov,
            "Latitude": round(lat_base + np.random.uniform(-0.8, 0.8), 5),
            "Longitude": round(lon_base + np.random.uniform(-0.8, 0.8), 5),
            "Activité": random.choice(ACTIVITES),
            "Équipe": random.choice(EQUIPES),
            "Statut": np.random.choice(STATUTS, p=[0.55, 0.25, 0.15, 0.05]),
            "Date_Début": (base_date + timedelta(days=random.randint(0, 300))).strftime("%Y-%m-%d"),
            "Superficie_m2": random.randint(500, 50000),
            "Population_Bénéficiaire": random.randint(50, 5000),
            "Type_Danger": random.choice(DANGERS),
            "Items_Trouvés": random.randint(0, 150),
            "Items_Détruits": lambda r=None: 0,
            "Qualité_Score": round(np.random.uniform(60, 100), 1),
            "Collecteur": f"Agent_{random.randint(1,20):02d}",
        })
    df = pd.DataFrame(rows)
    df["Items_Détruits"] = df["Items_Trouvés"].apply(lambda x: random.randint(int(x*0.8), x))
    df["Date_Début"] = pd.to_datetime(df["Date_Début"])
    df["Mois"] = df["Date_Début"].dt.to_period("M").astype(str)
    return df

def gen_collecte_raw(n=200):
    rows = []
    for i in range(n):
        prov = random.choice(PROVINCES)
        rows.append({
            "ID": f"FORM-{i+1:04d}",
            "Province": prov,
            "Collecteur": f"Agent_{random.randint(1,20):02d}",
            "Date_Soumission": (datetime(2025,1,1) + timedelta(days=random.randint(0,364))).strftime("%Y-%m-%d"),
            "Latitude": round(13.5 + np.random.uniform(-1.5,1.5), 5),
            "Longitude": round(15.5 + np.random.uniform(-2,5), 5),
            "Activité": random.choice(ACTIVITES),
            "Population": random.choice([random.randint(50,5000), None, None] if random.random()<0.12 else [random.randint(50,5000)]),
            "Items_Trouvés": random.choice([random.randint(0,100), -1, None] if random.random()<0.08 else [random.randint(0,100)]),
            "Statut": random.choice(STATUTS),
            "Superficie": random.choice([random.randint(100,50000), 0, None] if random.random()<0.1 else [random.randint(100,50000)]),
            "Doublon": False,
        })
    df = pd.DataFrame(rows)
    # Inject duplicates
    dupes = df.sample(8).copy()
    dupes["Doublon"] = True
    df = pd.concat([df, dupes], ignore_index=True)
    return df

@st.cache_data
def load_data():
    sites = gen_sites(120)
    raw   = gen_collecte_raw(200)
    return sites, raw

sites_df, raw_df = load_data()

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding:10px 0 20px;'>
        <div style='font-size:2.5rem;'>💣</div>
        <div style='font-size:1.1rem; font-weight:700; letter-spacing:1px;'>MAG TCHAD</div>
        <div style='font-size:0.75rem; opacity:0.8;'>Système de Gestion de l'Information</div>
        <hr style='border-color:rgba(255,255,255,0.3); margin:12px 0;'>
    </div>
    """, unsafe_allow_html=True)

    page = st.radio("Navigation", [
        "🏠 Tableau de Bord",
        "📥 Collecte & Import",
        "🔍 Contrôle Qualité",
        "🗺️ Carte Interactive",
        "📊 Analyse & Rapports",
        "📄 Export Rapport",
    ], label_visibility="collapsed")

    st.markdown("---")
    st.markdown("**Filtres Globaux**")
    sel_provinces = st.multiselect("Province(s)", PROVINCES, default=PROVINCES)
    sel_activites = st.multiselect("Activité(s)", ACTIVITES, default=ACTIVITES)
    sel_statuts   = st.multiselect("Statut(s)", STATUTS, default=STATUTS)

    st.markdown("---")
    st.markdown("""
    <div style='font-size:0.72rem; opacity:0.75; text-align:center;'>
        <strong>Développé par</strong><br>
        DJAOYANG HABEKREO Pelandi<br>
        Ingénieur Statisticien Économiste<br>
        <em>Candidat – RGI MAG Tchad</em>
    </div>
    """, unsafe_allow_html=True)

# Filtered data
filt = sites_df[
    sites_df["Province"].isin(sel_provinces) &
    sites_df["Activité"].isin(sel_activites) &
    sites_df["Statut"].isin(sel_statuts)
].copy()

# ─────────────────────────────────────────────
# PAGE 1 : TABLEAU DE BORD
# ─────────────────────────────────────────────
if page == "🏠 Tableau de Bord":
    st.markdown("## 🏠 Tableau de Bord Opérationnel")
    st.caption(f"MAG Tchad · Données au {datetime.now().strftime('%d %B %Y')} · {len(filt)} sites filtrés")

    # KPI Row
    c1, c2, c3, c4, c5 = st.columns(5)
    total_sites     = len(filt)
    completes       = len(filt[filt["Statut"] == "Complété"])
    pop_ben         = filt["Population_Bénéficiaire"].sum()
    items_detruits  = filt["Items_Détruits"].sum()
    surf_liberee    = filt[filt["Statut"]=="Complété"]["Superficie_m2"].sum()
    qualite_moy     = filt["Qualité_Score"].mean()

    c1.markdown(f"""<div class='kpi-card'>
        <div class='kpi-value'>{total_sites}</div>
        <div class='kpi-label'>📍 Sites Opérationnels</div>
        <div class='kpi-delta kpi-up'>↑ +12 ce mois</div></div>""", unsafe_allow_html=True)
    c2.markdown(f"""<div class='kpi-card green'>
        <div class='kpi-value'>{completes}</div>
        <div class='kpi-label'>✅ Sites Complétés</div>
        <div class='kpi-delta kpi-up'>↑ {round(completes/total_sites*100,1)}% taux</div></div>""", unsafe_allow_html=True)
    c3.markdown(f"""<div class='kpi-card'>
        <div class='kpi-value'>{pop_ben:,}</div>
        <div class='kpi-label'>👥 Bénéficiaires</div>
        <div class='kpi-delta kpi-up'>↑ Population sécurisée</div></div>""", unsafe_allow_html=True)
    c4.markdown(f"""<div class='kpi-card red'>
        <div class='kpi-value'>{items_detruits:,}</div>
        <div class='kpi-label'>💥 Engins Détruits</div>
        <div class='kpi-delta kpi-down'>UXO / AMD / IED</div></div>""", unsafe_allow_html=True)
    c5.markdown(f"""<div class='kpi-card orange'>
        <div class='kpi-value'>{qualite_moy:.1f}%</div>
        <div class='kpi-label'>⭐ Qualité Données</div>
        <div class='kpi-delta kpi-up'>Score moyen</div></div>""", unsafe_allow_html=True)

    st.markdown("")
    col1, col2 = st.columns([3, 2])

    with col1:
        st.markdown("<div class='section-header'>📈 Activités par Mois</div>", unsafe_allow_html=True)
        monthly = filt.groupby("Mois").agg(Sites=("ID_Site","count"), Bénéficiaires=("Population_Bénéficiaire","sum")).reset_index()
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=monthly["Mois"], y=monthly["Sites"], name="Sites", marker_color="#2E75B6"), secondary_y=False)
        fig.add_trace(go.Scatter(x=monthly["Mois"], y=monthly["Bénéficiaires"], name="Bénéficiaires", line=dict(color="#EA580C", width=2.5), mode="lines+markers"), secondary_y=True)
        fig.update_layout(height=280, margin=dict(l=0,r=0,t=10,b=0), plot_bgcolor="white", paper_bgcolor="white", legend=dict(orientation="h", y=1.1))
        fig.update_xaxes(showgrid=False); fig.update_yaxes(showgrid=True, gridcolor="#F0F0F0")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown("<div class='section-header'>🎯 Statut des Sites</div>", unsafe_allow_html=True)
        stat_counts = filt["Statut"].value_counts().reset_index()
        stat_counts.columns = ["Statut","Nb"]
        colors = {"Complété":"#16A34A","En cours":"#2E75B6","Planifié":"#F59E0B","Suspendu":"#DC2626"}
        fig2 = px.pie(stat_counts, names="Statut", values="Nb",
                      color="Statut", color_discrete_map=colors, hole=0.55)
        fig2.update_traces(textposition='outside', textinfo='percent+label')
        fig2.update_layout(height=280, margin=dict(l=0,r=0,t=10,b=0), showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        st.markdown("<div class='section-header'>🏭 Activités par Province</div>", unsafe_allow_html=True)
        prov_act = filt.groupby(["Province","Activité"]).size().reset_index(name="N")
        fig3 = px.bar(prov_act, x="Province", y="N", color="Activité", barmode="stack",
                      color_discrete_sequence=px.colors.qualitative.Set2)
        fig3.update_layout(height=270, margin=dict(l=0,r=0,t=10,b=0), plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        st.markdown("<div class='section-header'>💣 Types d'Engins par Province</div>", unsafe_allow_html=True)
        danger_prov = filt.groupby(["Province","Type_Danger"])["Items_Détruits"].sum().reset_index()
        fig4 = px.bar(danger_prov, x="Province", y="Items_Détruits", color="Type_Danger", barmode="group",
                      color_discrete_sequence=px.colors.qualitative.Bold)
        fig4.update_layout(height=270, margin=dict(l=0,r=0,t=10,b=0), plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig4, use_container_width=True)

    # Last sites table
    st.markdown("<div class='section-header'>📋 Derniers Sites Enregistrés</div>", unsafe_allow_html=True)
    display_cols = ["ID_Site","Province","Activité","Équipe","Statut","Date_Début","Population_Bénéficiaire","Items_Détruits","Qualité_Score"]
    st.dataframe(
        filt[display_cols].sort_values("Date_Début", ascending=False).head(10).reset_index(drop=True),
        use_container_width=True, height=300
    )

# ─────────────────────────────────────────────
# PAGE 2 : COLLECTE & IMPORT
# ─────────────────────────────────────────────
elif page == "📥 Collecte & Import":
    st.markdown("## 📥 Collecte & Import de Données")

    tab1, tab2 = st.tabs(["📤 Importer un fichier", "✍️ Saisie Manuelle (Formulaire)"])

    with tab1:
        st.markdown("<div class='section-header'>Import CSV / Excel (KoboToolbox, Survey123, ODK)</div>", unsafe_allow_html=True)
        col_up, col_info = st.columns([2,1])
        with col_up:
            uploaded = st.file_uploader("Glissez votre fichier ici", type=["csv","xlsx","xls"],
                                        help="Formats acceptés : CSV (KoboToolbox), XLSX (Survey123, ODK)")
            if uploaded:
                try:
                    if uploaded.name.endswith(".csv"):
                        df_up = pd.read_csv(uploaded)
                    else:
                        df_up = pd.read_excel(uploaded)
                    st.markdown(f"<div class='alert-success'>✅ Fichier chargé : <strong>{uploaded.name}</strong> — {len(df_up)} lignes, {len(df_up.columns)} colonnes</div>", unsafe_allow_html=True)
                    st.dataframe(df_up.head(20), use_container_width=True)
                    st.download_button("💾 Re-télécharger (CSV)", df_up.to_csv(index=False), "donnees_importees.csv", "text/csv")
                except Exception as e:
                    st.markdown(f"<div class='alert-danger'>❌ Erreur : {e}</div>", unsafe_allow_html=True)
            else:
                st.markdown("<div class='alert-info'>ℹ️ Aucun fichier importé — affichage des données de démonstration MAG ci-dessous.</div>", unsafe_allow_html=True)
                st.dataframe(raw_df.head(15), use_container_width=True)
        with col_info:
            st.markdown("**Sources compatibles**")
            for src in ["✅ KoboToolbox (CSV/API)", "✅ Survey123 (Excel)", "✅ ODK Collect (CSV)", "✅ CSPro (CSV/Excel)", "✅ Google Forms (Sheets)"]:
                st.markdown(src)
            st.info(f"📊 Base de démo : **{len(raw_df)} enregistrements** | **{len(raw_df.columns)} variables**")

    with tab2:
        st.markdown("<div class='section-header'>Formulaire de Saisie Terrain</div>", unsafe_allow_html=True)
        with st.form("saisie_form"):
            c1,c2,c3 = st.columns(3)
            prov_f   = c1.selectbox("Province", PROVINCES)
            act_f    = c2.selectbox("Activité", ACTIVITES)
            equipe_f = c3.selectbox("Équipe", EQUIPES)
            c4,c5,c6 = st.columns(3)
            lat_f    = c4.number_input("Latitude", value=13.5, format="%.5f")
            lon_f    = c5.number_input("Longitude", value=15.5, format="%.5f")
            surf_f   = c6.number_input("Superficie (m²)", min_value=0, value=1000)
            c7,c8,c9 = st.columns(3)
            pop_f    = c7.number_input("Population bénéficiaire", min_value=0, value=200)
            items_f  = c8.number_input("Items trouvés", min_value=0, value=0)
            danger_f = c9.selectbox("Type de danger", DANGERS)
            obs_f    = st.text_area("Observations / Remarques", placeholder="Décrire le contexte du site...")
            submitted = st.form_submit_button("💾 Enregistrer l'entrée", use_container_width=True)
            if submitted:
                st.markdown(f"""<div class='alert-success'>
                    ✅ <strong>Entrée enregistrée avec succès !</strong><br>
                    Province : {prov_f} | Activité : {act_f} | Équipe : {equipe_f} | Pop. : {pop_f:,}
                </div>""", unsafe_allow_html=True)
                st.balloons()

# ─────────────────────────────────────────────
# PAGE 3 : CONTRÔLE QUALITÉ
# ─────────────────────────────────────────────
elif page == "🔍 Contrôle Qualité":
    st.markdown("## 🔍 Contrôle Qualité des Données")

    df_q = raw_df.copy()
    n_total = len(df_q)

    # --- Détection automatique des erreurs ---
    df_q["Erreur_Population"] = df_q["Population"].isna()
    df_q["Erreur_Items"]      = df_q["Items_Trouvés"].isna() | (df_q["Items_Trouvés"] < 0)
    df_q["Erreur_Surface"]    = df_q["Superficie"].isna() | (df_q["Superficie"] == 0)
    df_q["Erreur_Coords"]     = (df_q["Latitude"] < 7) | (df_q["Latitude"] > 24) | \
                                 (df_q["Longitude"] < 13) | (df_q["Longitude"] > 24)
    df_q["Nb_Erreurs"]        = df_q[["Erreur_Population","Erreur_Items","Erreur_Surface","Erreur_Coords","Doublon"]].sum(axis=1)
    df_q["Qualité"]           = df_q["Nb_Erreurs"].apply(
        lambda x: "🟢 Conforme" if x==0 else ("🟡 Avertissement" if x==1 else "🔴 Critique")
    )

    n_ok      = (df_q["Nb_Erreurs"]==0).sum()
    n_warn    = (df_q["Nb_Erreurs"]==1).sum()
    n_crit    = (df_q["Nb_Erreurs"]>=2).sum()
    n_dup     = df_q["Doublon"].sum()
    score_q   = round(n_ok/n_total*100, 1)

    # KPIs qualité
    st.markdown("<div class='section-header'>📊 Synthèse de la Qualité</div>", unsafe_allow_html=True)
    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("Total enregistrements", n_total)
    k2.metric("✅ Conformes", n_ok, f"{round(n_ok/n_total*100,1)}%")
    k3.metric("⚠️ Avertissements", n_warn, delta=None)
    k4.metric("🔴 Critiques", n_crit, delta=None)
    k5.metric("🔁 Doublons", n_dup, delta=None)

    # Score gauge
    fig_gauge = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=score_q,
        delta={"reference": 85, "suffix":"%"},
        title={"text": "Score Global de Qualité", "font":{"size":16}},
        gauge={
            "axis": {"range": [0,100]},
            "bar":  {"color": "#2E75B6"},
            "steps":[{"range":[0,60],"color":"#FEE2E2"},
                     {"range":[60,80],"color":"#FEF3C7"},
                     {"range":[80,100],"color":"#DCFCE7"}],
            "threshold":{"line":{"color":"#1F4E79","width":3},"thickness":0.75,"value":85}
        }
    ))
    fig_gauge.update_layout(height=220, margin=dict(l=30,r=30,t=40,b=10))
    col_g, col_e = st.columns([1,2])
    col_g.plotly_chart(fig_gauge, use_container_width=True)

    with col_e:
        st.markdown("<div class='section-header'>🚨 Erreurs Détectées Automatiquement</div>", unsafe_allow_html=True)
        erreurs = {
            "Valeurs manquantes (Population)": df_q["Erreur_Population"].sum(),
            "Items négatifs/manquants": df_q["Erreur_Items"].sum(),
            "Surface nulle/manquante": df_q["Erreur_Surface"].sum(),
            "Coordonnées hors Tchad": df_q["Erreur_Coords"].sum(),
            "Doublons détectés": n_dup,
        }
        for label, count in erreurs.items():
            pct = round(count/n_total*100,1)
            color = "alert-danger" if count > 10 else ("alert-warning" if count > 0 else "alert-success")
            icon  = "🔴" if count > 10 else ("⚠️" if count > 0 else "✅")
            st.markdown(f"<div class='{color}'>{icon} <strong>{label}</strong> : {count} cas ({pct}%)</div>", unsafe_allow_html=True)

    st.markdown("<div class='section-header'>📋 Détail des Enregistrements</div>", unsafe_allow_html=True)
    filtre_q = st.selectbox("Filtrer par qualité :", ["Tous","🔴 Critique","🟡 Avertissement","🟢 Conforme","🔁 Doublons"])
    df_show = df_q.copy()
    if filtre_q != "Tous":
        if filtre_q == "🔁 Doublons":
            df_show = df_show[df_show["Doublon"]==True]
        else:
            df_show = df_show[df_show["Qualité"]==filtre_q]

    cols_show = ["ID","Province","Collecteur","Date_Soumission","Activité","Population","Items_Trouvés","Superficie","Qualité","Nb_Erreurs"]
    st.dataframe(df_show[cols_show].reset_index(drop=True), use_container_width=True, height=350)

    buf = io.BytesIO()
    df_show[cols_show].to_excel(buf, index=False); buf.seek(0)
    st.download_button("📥 Exporter les erreurs (Excel)", buf, "rapport_qualite_MAG.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ─────────────────────────────────────────────
# PAGE 4 : CARTE INTERACTIVE
# ─────────────────────────────────────────────
elif page == "🗺️ Carte Interactive":
    st.markdown("## 🗺️ Carte Interactive des Opérations MAG")

    col_ctrl, col_map = st.columns([1,3])
    with col_ctrl:
        st.markdown("**Options carte**")
        layer_type = st.radio("Affichage", ["Marqueurs par statut","Densité (HeatMap)","Clusters"])
        base_map   = st.selectbox("Fond de carte", ["OpenStreetMap","CartoDB Positron","CartoDB DarkMatter"])
        show_pop   = st.checkbox("Afficher pop. bénéficiaire", True)
        filt_map   = st.multiselect("Activité", ACTIVITES, default=ACTIVITES[:3])
        st.markdown("---")
        st.markdown(f"**{len(filt[filt['Activité'].isin(filt_map)])} sites affichés**")

        color_map = {"Complété":"green","En cours":"blue","Planifié":"orange","Suspendu":"red"}
        st.markdown("**Légende**")
        for k,v in color_map.items():
            st.markdown(f"🔵 {k}" if v=="blue" else f"🟢 {k}" if v=="green" else f"🟠 {k}" if v=="orange" else f"🔴 {k}")

    with col_map:
        tiles_map = {"OpenStreetMap":"OpenStreetMap","CartoDB Positron":"CartoDB positron","CartoDB DarkMatter":"CartoDB dark_matter"}
        m = folium.Map(location=[13.5, 15.5], zoom_start=6, tiles=tiles_map[base_map])

        df_map = filt[filt["Activité"].isin(filt_map)].copy()

        if layer_type == "Densité (HeatMap)":
            from folium.plugins import HeatMap
            heat_data = df_map[["Latitude","Longitude"]].dropna().values.tolist()
            HeatMap(heat_data, radius=20, blur=15).add_to(m)

        elif layer_type == "Clusters":
            from folium.plugins import MarkerCluster
            mc = MarkerCluster().add_to(m)
            for _, row in df_map.iterrows():
                popup_html = f"""
                <div style='font-family:Arial; min-width:200px;'>
                    <b style='color:#1F4E79;'>{row['ID_Site']}</b><br>
                    <hr style='margin:4px 0;'>
                    📍 {row['Province']}<br>
                    🎯 {row['Activité']}<br>
                    👥 {int(row['Population_Bénéficiaire']):,} bénéficiaires<br>
                    💥 {int(row['Items_Détruits'])} engins détruits<br>
                    ⭐ Qualité : {row['Qualité_Score']}%<br>
                    📅 {str(row['Date_Début'])[:10]}
                </div>"""
                folium.Marker(
                    [row["Latitude"], row["Longitude"]],
                    popup=folium.Popup(popup_html, max_width=250),
                    icon=folium.Icon(color=color_map.get(row["Statut"],"gray"), icon="info-sign")
                ).add_to(mc)

        else:  # Marqueurs par statut
            for _, row in df_map.iterrows():
                popup_html = f"""
                <div style='font-family:Arial; min-width:200px;'>
                    <b style='color:#1F4E79;'>{row['ID_Site']}</b><br>
                    <hr style='margin:4px 0;'>
                    📍 {row['Province']}<br>
                    🎯 {row['Activité']}<br>
                    ⚠️ {row['Type_Danger']}<br>
                    👥 {int(row['Population_Bénéficiaire']):,} bénéficiaires<br>
                    💥 {int(row['Items_Détruits'])} engins détruits<br>
                    ⭐ Qualité : {row['Qualité_Score']}%
                </div>"""
                r = 8 + int(row["Population_Bénéficiaire"] / 800) if show_pop else 8
                folium.CircleMarker(
                    [row["Latitude"], row["Longitude"]],
                    radius=r,
                    color=color_map.get(row["Statut"],"gray"),
                    fill=True, fill_opacity=0.7,
                    popup=folium.Popup(popup_html, max_width=250),
                    tooltip=f"{row['ID_Site']} | {row['Province']} | {row['Statut']}"
                ).add_to(m)

        st_folium(m, width=None, height=520, use_container_width=True)

# ─────────────────────────────────────────────
# PAGE 5 : ANALYSE & RAPPORTS
# ─────────────────────────────────────────────
elif page == "📊 Analyse & Rapports":
    st.markdown("## 📊 Analyse Avancée & Indicateurs")

    tab1, tab2, tab3 = st.tabs(["📈 Tendances", "🏆 Performance Équipes", "📐 Statistiques Descriptives"])

    with tab1:
        st.markdown("<div class='section-header'>Évolution mensuelle des opérations</div>", unsafe_allow_html=True)
        monthly2 = filt.groupby(["Mois","Statut"]).size().reset_index(name="N")
        fig_t = px.area(monthly2, x="Mois", y="N", color="Statut",
                        color_discrete_map={"Complété":"#16A34A","En cours":"#2E75B6","Planifié":"#F59E0B","Suspendu":"#DC2626"})
        fig_t.update_layout(height=300, plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=0,r=0,t=10,b=0))
        st.plotly_chart(fig_t, use_container_width=True)

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("<div class='section-header'>Surface libérée par Province (m²)</div>", unsafe_allow_html=True)
            surf_prov = filt[filt["Statut"]=="Complété"].groupby("Province")["Superficie_m2"].sum().reset_index()
            fig_s = px.bar(surf_prov.sort_values("Superficie_m2", ascending=True),
                           x="Superficie_m2", y="Province", orientation="h",
                           color="Superficie_m2", color_continuous_scale="Blues")
            fig_s.update_layout(height=260, plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=0,r=0,t=10,b=0))
            st.plotly_chart(fig_s, use_container_width=True)

        with col_b:
            st.markdown("<div class='section-header'>Distribution du Score Qualité</div>", unsafe_allow_html=True)
            fig_hist = px.histogram(filt, x="Qualité_Score", nbins=20, color_discrete_sequence=["#2E75B6"])
            fig_hist.add_vline(x=filt["Qualité_Score"].mean(), line_dash="dash", line_color="red",
                               annotation_text=f"Moy: {filt['Qualité_Score'].mean():.1f}%")
            fig_hist.update_layout(height=260, plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=0,r=0,t=10,b=0))
            st.plotly_chart(fig_hist, use_container_width=True)

    with tab2:
        st.markdown("<div class='section-header'>Performance par Équipe</div>", unsafe_allow_html=True)
        eq_perf = filt.groupby("Équipe").agg(
            Sites=("ID_Site","count"),
            Complétés=("Statut", lambda x: (x=="Complété").sum()),
            Pop_Total=("Population_Bénéficiaire","sum"),
            Engins_Détruits=("Items_Détruits","sum"),
            Surface_m2=("Superficie_m2","sum"),
            Score_Qualité=("Qualité_Score","mean")
        ).reset_index()
        eq_perf["Taux_Completion"] = (eq_perf["Complétés"]/eq_perf["Sites"]*100).round(1)
        eq_perf["Score_Qualité"]   = eq_perf["Score_Qualité"].round(1)

        fig_eq = px.scatter(eq_perf, x="Taux_Completion", y="Score_Qualité",
                            size="Pop_Total", color="Équipe", text="Équipe",
                            size_max=50, color_discrete_sequence=px.colors.qualitative.Bold)
        fig_eq.update_traces(textposition="top center")
        fig_eq.update_layout(height=300, plot_bgcolor="white", paper_bgcolor="white",
                             xaxis_title="Taux de complétion (%)", yaxis_title="Score qualité (%)",
                             margin=dict(l=0,r=0,t=10,b=0))
        st.plotly_chart(fig_eq, use_container_width=True)
        st.dataframe(eq_perf.style.background_gradient(subset=["Score_Qualité","Taux_Completion"], cmap="Blues"),
                     use_container_width=True)

    with tab3:
        st.markdown("<div class='section-header'>Statistiques Descriptives</div>", unsafe_allow_html=True)
        num_cols = ["Superficie_m2","Population_Bénéficiaire","Items_Trouvés","Items_Détruits","Qualité_Score"]
        st.dataframe(filt[num_cols].describe().round(2).T.rename(columns={
            "count":"N","mean":"Moyenne","std":"Écart-type",
            "min":"Min","25%":"Q1","50%":"Médiane","75%":"Q3","max":"Max"
        }), use_container_width=True)

        st.markdown("<div class='section-header'>Matrice de Corrélation</div>", unsafe_allow_html=True)
        corr = filt[num_cols].corr().round(2)
        fig_corr = px.imshow(corr, text_auto=True, aspect="auto",
                             color_continuous_scale="RdBu_r", zmin=-1, zmax=1)
        fig_corr.update_layout(height=320, margin=dict(l=0,r=0,t=10,b=0))
        st.plotly_chart(fig_corr, use_container_width=True)

# ─────────────────────────────────────────────
# PAGE 6 : EXPORT RAPPORT WORD
# ─────────────────────────────────────────────
elif page == "📄 Export Rapport":
    st.markdown("## 📄 Génération de Rapport Automatique")
    st.markdown("<div class='alert-info'>ℹ️ Configurez les paramètres ci-dessous, puis cliquez sur <strong>Générer le rapport</strong> pour obtenir un rapport Word téléchargeable.</div>", unsafe_allow_html=True)

    col_cfg, col_prev = st.columns([1, 2])
    with col_cfg:
        titre_rapport = st.text_input("Titre du rapport", "Rapport Opérationnel MAG Tchad")
        periode       = st.text_input("Période de référence", "Janvier – Décembre 2025")
        prep_par      = st.text_input("Préparé par", "DJAOYANG HABEKREO Pelandi")
        dest          = st.text_input("Destinataires", "Direction Pays, Bailleurs")
        incl_stats    = st.checkbox("Inclure statistiques descriptives", True)
        incl_qualite  = st.checkbox("Inclure rapport qualité", True)
        gen_btn       = st.button("📝 Générer le rapport Word", use_container_width=True, type="primary")

    with col_prev:
        st.markdown("<div class='section-header'>Aperçu du contenu</div>", unsafe_allow_html=True)
        n_sites  = len(filt)
        n_comp   = (filt["Statut"]=="Complété").sum()
        pop_tot  = filt["Population_Bénéficiaire"].sum()
        eng_det  = filt["Items_Détruits"].sum()
        q_moy    = filt["Qualité_Score"].mean()

        st.markdown(f"""
        📌 **{titre_rapport}**  
        📅 Période : {periode}  
        ✍️ Préparé par : {prep_par}  
        ---
        | Indicateur | Valeur |
        |---|---|
        | Sites opérationnels | **{n_sites}** |
        | Sites complétés | **{n_comp} ({round(n_comp/n_sites*100,1)}%)** |
        | Bénéficiaires | **{pop_tot:,}** |
        | Engins détruits | **{eng_det:,}** |
        | Score qualité moyen | **{q_moy:.1f}%** |
        """)

    if gen_btn:
        with st.spinner("Génération du rapport en cours..."):
            doc = Document()

            # Styles
            style = doc.styles['Normal']
            style.font.name = 'Arial'
            style.font.size = Pt(11)

            # Titre
            title_p = doc.add_paragraph()
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = title_p.add_run(titre_rapport.upper())
            run.bold = True; run.font.size = Pt(18)
            run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

            doc.add_paragraph(f"Période : {periode}", style='Normal').alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"Préparé par : {prep_par} | Destinataires : {dest}", style='Normal').alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"Date de génération : {datetime.now().strftime('%d/%m/%Y %H:%M')}", style='Normal').alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph()

            def add_section(title):
                p = doc.add_paragraph()
                run = p.add_run(title)
                run.bold = True; run.font.size = Pt(13)
                run.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
                doc.add_paragraph()

            # Section 1 : KPIs
            add_section("1. Indicateurs Clés de Performance")
            table = doc.add_table(rows=6, cols=2)
            table.style = 'Table Grid'
            headers = [("Indicateur","Valeur"),
                       ("Sites opérationnels", str(n_sites)),
                       ("Sites complétés", f"{n_comp} ({round(n_comp/n_sites*100,1)}%)"),
                       ("Population bénéficiaire", f"{pop_tot:,}"),
                       ("Engins détruits/neutralisés", f"{eng_det:,}"),
                       ("Score qualité des données", f"{q_moy:.1f}%")]
            for i,(k,v) in enumerate(headers):
                table.rows[i].cells[0].text = k
                table.rows[i].cells[1].text = v
                if i == 0:
                    for cell in table.rows[i].cells:
                        cell.paragraphs[0].runs[0].bold = True
            doc.add_paragraph()

            # Section 2 : Activités par province
            add_section("2. Répartition des Activités par Province")
            prov_sum = filt.groupby("Province").agg(
                Sites=("ID_Site","count"),
                Pop=("Population_Bénéficiaire","sum"),
                Engins=("Items_Détruits","sum")
            ).reset_index()
            tbl2 = doc.add_table(rows=len(prov_sum)+1, cols=4)
            tbl2.style = 'Table Grid'
            for j,(h) in enumerate(["Province","Sites","Bénéficiaires","Engins Détruits"]):
                tbl2.rows[0].cells[j].text = h
                tbl2.rows[0].cells[j].paragraphs[0].runs[0].bold = True
            for i, row in prov_sum.iterrows():
                tbl2.rows[i+1].cells[0].text = row["Province"]
                tbl2.rows[i+1].cells[1].text = str(row["Sites"])
                tbl2.rows[i+1].cells[2].text = f"{row['Pop']:,}"
                tbl2.rows[i+1].cells[3].text = str(row["Engins"])
            doc.add_paragraph()

            if incl_stats:
                add_section("3. Statistiques Descriptives")
                num_cols = ["Superficie_m2","Population_Bénéficiaire","Items_Détruits","Qualité_Score"]
                desc = filt[num_cols].describe().round(1)
                tbl3 = doc.add_table(rows=len(num_cols)+1, cols=5)
                tbl3.style = 'Table Grid'
                for j,h in enumerate(["Variable","Moyenne","Médiane","Min","Max"]):
                    tbl3.rows[0].cells[j].text = h
                    tbl3.rows[0].cells[j].paragraphs[0].runs[0].bold = True
                for i, col in enumerate(num_cols):
                    tbl3.rows[i+1].cells[0].text = col
                    tbl3.rows[i+1].cells[1].text = str(desc.loc["mean", col])
                    tbl3.rows[i+1].cells[2].text = str(desc.loc["50%", col])
                    tbl3.rows[i+1].cells[3].text = str(desc.loc["min", col])
                    tbl3.rows[i+1].cells[4].text = str(desc.loc["max", col])
                doc.add_paragraph()

            if incl_qualite:
                add_section("4. Rapport de Qualité des Données")
                doc.add_paragraph(f"Score qualité global : {q_moy:.1f}%")
                doc.add_paragraph(f"Enregistrements analysés : {len(raw_df)}")
                doc.add_paragraph(f"Doublons détectés : {raw_df['Doublon'].sum()}")
                doc.add_paragraph(f"Valeurs manquantes (Population) : {raw_df['Population'].isna().sum()}")

            # Signature
            doc.add_paragraph()
            sig = doc.add_paragraph(f"Fait à N'Djamena, le {datetime.now().strftime('%d/%m/%Y')}")
            doc.add_paragraph(f"Responsable de la Gestion de l'Information").bold = True
            doc.add_paragraph(prep_par)

            # Save
            buf = io.BytesIO()
            doc.save(buf); buf.seek(0)

        st.markdown("<div class='alert-success'>✅ Rapport généré avec succès !</div>", unsafe_allow_html=True)
        st.download_button(
            "📥 Télécharger le Rapport Word",
            buf,
            f"Rapport_MAG_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("""
<div class='footer'>
    MAG Tchad – Système de Gestion de l'Information | Démonstration technique<br>
    Développé par <strong>DJAOYANG HABEKREO Pelandi</strong> · Ingénieur Statisticien Économiste – ISSEA Yaoundé<br>
    <em>Candidat au poste de Responsable de la Gestion de l'Information · MAG Tchad · 2026</em>
</div>
""", unsafe_allow_html=True)
