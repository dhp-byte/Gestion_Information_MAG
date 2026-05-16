"""
MAG Tchad – Système de Gestion de l'Information
Développé par : DJAOYANG HABEKREO Pelandi | Ingénieur Statisticien Économiste
Poste visé  : Responsable de la Gestion de l'Information
Source       : https://www.maginternational.org/what-we-do/where-we-work/chad/
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
import io, random, time
from docx import Document
from docx.shared import Pt, RGBColor
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
# VRAIES DONNÉES DU SITE MAG TCHAD
# ─────────────────────────────────────────────
MAG_INFO = {
    "present_depuis": 2004,
    "resultats_2023": {
        "armes_legeres_detruites": 771,
        "munitions_detruites": 25101,
        "route_liberee_km": 167,
    },
    "zones_intervention": ["Lac Tchad", "Borkou", "Ennedi", "Fada", "Kalait", "Tibesti"],
    "activites": [
        "Déminage Humanitaire (HMA)",
        "Destruction d'Armes & Munitions (AMD)",
        "Éducation aux Risques (ER)",
        "Enquête Non-Technique (SNT)",
        "Gestion & Stockage d'Armements",
        "Libération de Surface",
    ],
    "bailleurs": [
        "GFFO (Allemagne)",
        "Union Européenne",
        "US Department of State",
        "EU Trust Fund for Africa",
        "Ministère belge des AE",
        "Ministère français des AE",
    ],
    "citation": {
        "texte": "La sécurité et la sûreté avec les armes sont très importantes et nous avons formé plus de 2 000 recrues au cours des neuf derniers mois.",
        "auteur": "Danemadji Toubard, Tchad"
    },
    "logo": "https://www.maginternational.org/static/img/mag-logo.a63f5ba03eb7.png",
}

MAG_IMAGES = [
    {
        "url": "https://www.maginternational.org/media/filer_public_thumbnails/filer_public/6d/ca/6dca76da-4ec7-471c-ae6f-875b9153d2a9/explosion.jpg__160x160_q50_crop_progressive_subsampling-2_upscale.jpg",
        "caption": "🔥 Destruction d'armes légères – Tchad 2023",
        "stat": "771 armes légères détruites en 2023"
    },
    {
        "url": "https://www.maginternational.org/media/filer_public_thumbnails/filer_public/00/f9/00f9b603-c086-415b-b1ba-80dfd1185acb/bullets.jpg__160x160_q50_crop_progressive_subsampling-2_upscale.jpg",
        "caption": "💥 Munitions neutralisées – Opérations AMD Tchad",
        "stat": "25 101 munitions détruites en 2023"
    },
    {
        "url": "https://www.maginternational.org/static/img/share-image.543598c7cbbf.jpg",
        "caption": "🌍 MAG – Sauve des vies, bâtit l'avenir",
        "stat": "Présent au Tchad depuis 2004 · 70 pays · 20M+ bénéficiaires"
    },
]

# ─────────────────────────────────────────────
# DONNÉES FICTIVES RÉALISTES
# ─────────────────────────────────────────────
np.random.seed(42)
random.seed(42)

PROVINCES = ["Lac", "Borkou", "Hadjer-Lamis", "Kanem", "Batha", "Ouaddaï", "Ennedi", "Tibesti"]
ACTIVITES = MAG_INFO["activites"]
STATUTS   = ["Complété", "En cours", "Planifié", "Suspendu"]
EQUIPES   = ["Alpha", "Bravo", "Charlie", "Delta"]
DANGERS   = ["UXO", "Mine AP", "Mine AC", "IED", "Sous-munition", "Grenade", "ALPC"]

LAT_MAP = {"Lac":13.2,"Borkou":18.5,"Hadjer-Lamis":12.8,"Kanem":14.5,
           "Batha":13.6,"Ouaddaï":13.0,"Ennedi":16.8,"Tibesti":21.0}
LON_MAP = {"Lac":14.2,"Borkou":18.9,"Hadjer-Lamis":15.8,"Kanem":15.6,
           "Batha":17.2,"Ouaddaï":20.8,"Ennedi":22.0,"Tibesti":17.5}

def gen_sites(n=140):
    rows = []
    base = datetime(2024, 1, 1)
    for i in range(n):
        prov = random.choice(PROVINCES)
        itms = random.randint(0, 200)
        rows.append({
            "ID_Site"                : f"TCH-2024-{str(i+1).zfill(4)}",
            "Province"               : prov,
            "Latitude"               : round(LAT_MAP[prov] + np.random.uniform(-0.6,0.6), 5),
            "Longitude"              : round(LON_MAP[prov] + np.random.uniform(-0.6,0.6), 5),
            "Activité"               : random.choice(ACTIVITES),
            "Équipe"                 : random.choice(EQUIPES),
            "Statut"                 : np.random.choice(STATUTS, p=[0.55,0.25,0.15,0.05]),
            "Date_Début"             : (base + timedelta(days=random.randint(0,450))).strftime("%Y-%m-%d"),
            "Superficie_m2"          : random.randint(500, 60000),
            "Population_Bénéficiaire": random.randint(50, 8000),
            "Type_Danger"            : random.choice(DANGERS),
            "Items_Trouvés"          : itms,
            "Items_Détruits"         : random.randint(int(itms*0.8), itms),
            "Qualité_Score"          : round(np.random.uniform(60, 100), 1),
            "Collecteur"             : f"Agent_{random.randint(1,25):02d}",
        })
    df = pd.DataFrame(rows)
    df["Date_Début"] = pd.to_datetime(df["Date_Début"])
    df["Mois"] = df["Date_Début"].dt.to_period("M").astype(str)
    return df

def gen_collecte_raw(n=220):
    rows = []
    for i in range(n):
        has_err = random.random() < 0.15
        rows.append({
            "ID"             : f"FORM-{i+1:04d}",
            "Province"       : random.choice(PROVINCES),
            "Collecteur"     : f"Agent_{random.randint(1,25):02d}",
            "Date_Soumission": (datetime(2024,1,1)+timedelta(days=random.randint(0,450))).strftime("%Y-%m-%d"),
            "Latitude"       : round(15.0 + np.random.uniform(-3,3), 5),
            "Longitude"      : round(17.0 + np.random.uniform(-4,7), 5),
            "Activité"       : random.choice(ACTIVITES),
            "Population"     : None if (has_err and random.random()<0.5) else random.randint(50,8000),
            "Items_Trouvés"  : (None if (has_err and random.random()<0.4) else
                                (-1 if (has_err and random.random()<0.3) else random.randint(0,150))),
            "Statut"         : random.choice(STATUTS),
            "Superficie"     : (None if (has_err and random.random()<0.4) else
                                (0 if (has_err and random.random()<0.3) else random.randint(100,60000))),
            "Doublon"        : False,
        })
    df = pd.DataFrame(rows)
    dupes = df.sample(10).copy()
    dupes["Doublon"] = True
    return pd.concat([df, dupes], ignore_index=True)

@st.cache_data
def load_data():
    return gen_sites(140), gen_collecte_raw(220)

sites_df, raw_df = load_data()

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
.main{background:#F0F4F8;}

section[data-testid="stSidebar"]{background:linear-gradient(180deg,#1F4E79 0%,#2E75B6 100%);}
section[data-testid="stSidebar"] *{color:white !important;}
section[data-testid="stSidebar"] .stRadio label{
    background:rgba(255,255,255,0.12);border-radius:8px;
    padding:7px 12px;margin:3px 0;display:block;cursor:pointer;transition:background 0.2s;}
section[data-testid="stSidebar"] .stRadio label:hover{background:rgba(255,255,255,0.28);}

.kpi-card{background:white;border-radius:12px;padding:18px 20px;
    box-shadow:0 2px 12px rgba(0,0,0,0.08);border-left:5px solid #2E75B6;margin-bottom:8px;}
.kpi-value{font-size:2rem;font-weight:700;color:#1F4E79;line-height:1.1;}
.kpi-label{font-size:0.82rem;color:#555;margin-top:5px;font-weight:500;}
.kpi-delta{font-size:0.78rem;margin-top:3px;}
.kpi-up{color:#16A34A;} .kpi-down{color:#DC2626;}
.kpi-card.red{border-left-color:#DC2626;}
.kpi-card.green{border-left-color:#16A34A;}
.kpi-card.orange{border-left-color:#EA580C;}

.section-header{background:linear-gradient(90deg,#1F4E79,#2E75B6);
    color:white !important;padding:9px 18px;border-radius:8px;
    font-size:1rem;font-weight:600;margin:16px 0 10px 0;}

.alert-danger{background:#FEF2F2;border-left:4px solid #DC2626;padding:11px 15px;border-radius:6px;color:#991B1B;margin:5px 0;}
.alert-warning{background:#FFFBEB;border-left:4px solid #F59E0B;padding:11px 15px;border-radius:6px;color:#92400E;margin:5px 0;}
.alert-success{background:#F0FDF4;border-left:4px solid #16A34A;padding:11px 15px;border-radius:6px;color:#166534;margin:5px 0;}
.alert-info{background:#EFF6FF;border-left:4px solid #2E75B6;padding:11px 15px;border-radius:6px;color:#1E40AF;margin:5px 0;}

.mag-quote{background:#1F4E79;color:white;border-radius:10px;padding:20px 28px;
    font-style:italic;font-size:1.05rem;margin:12px 0;}
.mag-quote .author{font-size:0.85rem;opacity:0.8;margin-top:8px;font-style:normal;}

.bailleur-card{background:white;border-radius:8px;padding:10px 14px;
    box-shadow:0 1px 6px rgba(0,0,0,0.07);text-align:center;
    font-size:0.82rem;font-weight:500;color:#1F4E79;border-top:3px solid #2E75B6;}

.result-card{background:white;border-radius:12px;padding:24px;
    text-align:center;box-shadow:0 2px 12px rgba(0,0,0,0.08);border-top:4px solid #DC2626;}
.result-number{font-size:2.8rem;font-weight:800;color:#1F4E79;}
.result-label{font-size:0.88rem;color:#555;font-weight:500;margin-top:6px;}

.carousel-caption{background:rgba(31,78,121,0.88);color:white;padding:10px 16px;
    font-size:0.9rem;font-weight:500;border-radius:0 0 10px 10px;}
.carousel-stat{background:#DC2626;color:white;padding:4px 12px;
    font-size:0.8rem;font-weight:600;border-radius:4px;display:inline-block;margin-top:4px;}

.footer{text-align:center;color:#999;font-size:0.72rem;
    padding:20px 0 8px;border-top:1px solid #E5E7EB;margin-top:28px;}

div[data-testid="stMetric"]{background:white;border-radius:10px;padding:12px;
    box-shadow:0 1px 6px rgba(0,0,0,0.07);}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style='text-align:center;padding:6px 0 18px;'>
        <img src='{MAG_INFO["logo"]}' width='90'
             style='filter:brightness(0) invert(1);margin-bottom:8px;'
             onerror="this.style.display='none'">
        <div style='font-size:1.05rem;font-weight:700;letter-spacing:1px;'>MAG TCHAD</div>
        <div style='font-size:0.72rem;opacity:0.85;'>Système de Gestion de l'Information</div>
        <div style='font-size:0.68rem;opacity:0.7;margin-top:2px;'>Présent depuis {MAG_INFO["present_depuis"]}</div>
        <hr style='border-color:rgba(255,255,255,0.3);margin:12px 0;'>
    </div>
    """, unsafe_allow_html=True)

    page = st.radio("Navigation", [
        "🏠 Accueil MAG Tchad",
        "📊 Tableau de Bord",
        "📥 Collecte & Import",
        "🔍 Contrôle Qualité",
        "🗺️ Carte Opérationnelle",
        "📈 Analyse & Rapports",
        "📄 Export Rapport Word",
    ], label_visibility="collapsed")

    st.markdown("---")
    st.markdown("**Filtres Globaux**")
    sel_provinces = st.multiselect("Province(s)", PROVINCES, default=PROVINCES)
    sel_activites = st.multiselect("Activité(s)", ACTIVITES, default=ACTIVITES)
    sel_statuts   = st.multiselect("Statut(s)", STATUTS, default=STATUTS)

    st.markdown("---")
    st.markdown(f"""
    <div style='font-size:0.7rem;opacity:0.78;text-align:center;'>
        <strong>Développé par</strong><br>
        DJAOYANG HABEKREO Pelandi<br>
        Ingénieur Statisticien Économiste<br>
        <em>Candidat – RGI MAG Tchad 2026</em>
    </div>
    """, unsafe_allow_html=True)

# Données filtrées
filt = sites_df[
    sites_df["Province"].isin(sel_provinces) &
    sites_df["Activité"].isin(sel_activites) &
    sites_df["Statut"].isin(sel_statuts)
].copy()

# ═══════════════════════════════════════════════════════
# PAGE 0 : ACCUEIL MAG TCHAD
# ═══════════════════════════════════════════════════════
if page == "🏠 Accueil MAG Tchad":

    # Hero banner
    st.markdown("""
    <div style='background:linear-gradient(135deg,#1F4E79 0%,#DC2626 100%);
        border-radius:14px;padding:36px 40px;margin-bottom:24px;'>
        <div style='color:white;font-size:2.2rem;font-weight:800;'>🇹🇩 MAG au Tchad</div>
        <div style='color:rgba(255,255,255,0.88);font-size:1.05rem;margin-top:8px;max-width:680px;'>
            Depuis 2004, MAG réduit les risques quotidiens de mort ou de blessures causés par les mines
            terrestres, engins non explosés et armes légères, pour créer des conditions sûres et durables.
        </div>
        <div style='margin-top:16px;'>
            <span style='background:rgba(255,255,255,0.2);color:white;padding:5px 14px;
                border-radius:20px;font-size:0.82rem;font-weight:600;margin-right:8px;'>
                ✅ Présent depuis 2004
            </span>
            <span style='background:rgba(255,255,255,0.2);color:white;padding:5px 14px;
                border-radius:20px;font-size:0.82rem;font-weight:600;margin-right:8px;'>
                🌍 Afrique centrale
            </span>
            <span style='background:rgba(220,38,38,0.7);color:white;padding:5px 14px;
                border-radius:20px;font-size:0.82rem;font-weight:600;'>
                💣 Déminage & AMD
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── CARROUSEL ──
    st.markdown("<div class='section-header'>📸 Galerie – Opérations MAG Tchad</div>",
                unsafe_allow_html=True)

    if "carousel_idx" not in st.session_state:
        st.session_state.carousel_idx = 0

    n_imgs = len(MAG_IMAGES)
    img    = MAG_IMAGES[st.session_state.carousel_idx]

    col_prev, col_img, col_next = st.columns([1, 10, 1])
    with col_prev:
        st.markdown("<br><br>", unsafe_allow_html=True)
        if st.button("◀", key="prev_btn", use_container_width=True):
            st.session_state.carousel_idx = (st.session_state.carousel_idx - 1) % n_imgs
            st.rerun()
    with col_img:
        st.markdown(f"""
        <div style='border-radius:12px;overflow:hidden;'>
            <img src='{img["url"]}'
                 style='width:100%;max-height:360px;object-fit:cover;display:block;'
                 onerror="this.style.background='#1F4E79';this.alt='MAG Tchad';">
            <div class='carousel-caption'>
                {img["caption"]}<br>
                <span class='carousel-stat'>{img["stat"]}</span>
            </div>
        </div>
        <div style='text-align:center;margin-top:8px;font-size:1.1rem;'>
            {''.join(['🔵' if i==st.session_state.carousel_idx else '⚪' for i in range(n_imgs)])}
            &nbsp;&nbsp;<span style='font-size:0.78rem;color:#888;'>
                {st.session_state.carousel_idx+1}/{n_imgs}</span>
        </div>
        """, unsafe_allow_html=True)
    with col_next:
        st.markdown("<br><br>", unsafe_allow_html=True)
        if st.button("▶", key="next_btn", use_container_width=True):
            st.session_state.carousel_idx = (st.session_state.carousel_idx + 1) % n_imgs
            st.rerun()

    col_auto, _ = st.columns([1,4])
    with col_auto:
        auto_play = st.checkbox("▶️ Auto", value=False, key="autoplay")
    if auto_play:
        time.sleep(3.5)
        st.session_state.carousel_idx = (st.session_state.carousel_idx + 1) % n_imgs
        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # ── RÉSULTATS 2023 (vrais chiffres site MAG) ──
    st.markdown("<div class='section-header'>🏆 Résultats Officiels MAG Tchad – 2023</div>",
                unsafe_allow_html=True)
    r1, r2, r3 = st.columns(3)
    r1.markdown(f"""<div class='result-card'>
        <div class='result-number'>{MAG_INFO["resultats_2023"]["armes_legeres_detruites"]:,}</div>
        <div class='result-label'>🔫 Armes légères & armes légères détruites</div>
    </div>""", unsafe_allow_html=True)
    r2.markdown(f"""<div class='result-card'>
        <div class='result-number'>{MAG_INFO["resultats_2023"]["munitions_detruites"]:,}</div>
        <div class='result-label'>💥 Munitions de petit calibre détruites</div>
    </div>""", unsafe_allow_html=True)
    r3.markdown(f"""<div class='result-card'>
        <div class='result-number'>{MAG_INFO["resultats_2023"]["route_liberee_km"]} km</div>
        <div class='result-label'>🛣️ Route libérée (Kouba Olanga – Faya, Borkou)</div>
    </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── POURQUOI / COMMENT ──
    col_why, col_how = st.columns(2)
    with col_why:
        st.markdown("<div class='section-header'>❓ Pourquoi nous intervenons</div>",
                    unsafe_allow_html=True)
        st.markdown("""
        <div style='background:white;border-radius:10px;padding:18px;box-shadow:0 1px 8px rgba(0,0,0,0.07);line-height:1.7;'>
            <p>Le <strong>nord du Tchad</strong> est fortement contaminé par des mines terrestres et engins
            non explosés issus du conflit <strong>Libye-Tchad (1979-1987)</strong>.</p>
            <p>Ces engins affectent particulièrement les <strong>communautés nomades</strong> qui ont dû
            abandonner leurs routes traditionnelles de commerce camélidé pour des itinéraires plus longs
            et dangereux.</p>
            <p>La stabilité régionale est menacée par la <strong>prolifération des armes légères</strong>
            et les explosifs non sécurisés qui alimentent les marchés illicites et les groupes armés non étatiques.</p>
        </div>
        """, unsafe_allow_html=True)

    with col_how:
        st.markdown("<div class='section-header'>🛠️ Nos activités</div>", unsafe_allow_html=True)
        items = "".join([f"<li style='margin:7px 0;'>✅ {a}</li>" for a in MAG_INFO["activites"]])
        st.markdown(f"""
        <div style='background:white;border-radius:10px;padding:18px;box-shadow:0 1px 8px rgba(0,0,0,0.07);'>
            <ul style='padding-left:16px;margin:0;'>{items}</ul>
            <p style='margin-top:14px;font-size:0.84rem;color:#555;'>
                <strong>Zones :</strong> {", ".join(MAG_INFO["zones_intervention"])}
            </p>
        </div>
        """, unsafe_allow_html=True)

    # ── CITATION ──
    st.markdown(f"""
    <div class='mag-quote'>
        « {MAG_INFO["citation"]["texte"]} »
        <div class='author'>— {MAG_INFO["citation"]["auteur"]}</div>
    </div>
    """, unsafe_allow_html=True)

    # ── BAILLEURS ──
    st.markdown("<div class='section-header'>🤝 Nos Bailleurs de Fonds</div>", unsafe_allow_html=True)
    b_cols = st.columns(len(MAG_INFO["bailleurs"]))
    for i, b in enumerate(MAG_INFO["bailleurs"]):
        b_cols[i].markdown(f"<div class='bailleur-card'>{b}</div>", unsafe_allow_html=True)

    # ── CARTE ZONES ──
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-header'>📍 Zones d'Intervention MAG Tchad</div>",
                unsafe_allow_html=True)
    zones_coords = {
        "Lac Tchad":(13.0,14.0),"Borkou":(18.0,18.5),"Ennedi":(16.5,22.0),
        "Fada":(17.2,21.6),"Kalait":(15.2,20.0),"Tibesti":(21.0,17.5),
    }
    m_acc = folium.Map(location=[16.0,18.0], zoom_start=5, tiles="CartoDB positron")
    for zone,(lat,lon) in zones_coords.items():
        folium.CircleMarker(
            [lat,lon], radius=14, color="#DC2626", fill=True, fill_opacity=0.75,
            popup=folium.Popup(f"<b style='color:#1F4E79;'>MAG – {zone}</b>", max_width=200),
            tooltip=f"🔴 {zone}"
        ).add_to(m_acc)
    st_folium(m_acc, height=320, use_container_width=True)


# ═══════════════════════════════════════════════════════
# PAGE 1 : TABLEAU DE BORD
# ═══════════════════════════════════════════════════════
elif page == "📊 Tableau de Bord":
    st.markdown("## 📊 Tableau de Bord Opérationnel")
    st.caption(f"Données au {datetime.now().strftime('%d %B %Y')} · {len(filt)} sites filtrés · Source : maginternational.org/chad")

    if len(filt) == 0:
        st.warning("Aucun site ne correspond aux filtres sélectionnés.")
        st.stop()

    total  = len(filt)
    comp   = (filt["Statut"]=="Complété").sum()
    pop    = filt["Population_Bénéficiaire"].sum()
    eng    = filt["Items_Détruits"].sum()
    qual   = filt["Qualité_Score"].mean()

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.markdown(f"""<div class='kpi-card'><div class='kpi-value'>{total}</div>
        <div class='kpi-label'>📍 Sites Opérationnels</div>
        <div class='kpi-delta kpi-up'>↑ +14 ce trimestre</div></div>""", unsafe_allow_html=True)
    c2.markdown(f"""<div class='kpi-card green'><div class='kpi-value'>{comp}</div>
        <div class='kpi-label'>✅ Sites Complétés</div>
        <div class='kpi-delta kpi-up'>↑ {round(comp/total*100,1)}% taux</div></div>""", unsafe_allow_html=True)
    c3.markdown(f"""<div class='kpi-card'><div class='kpi-value'>{pop:,}</div>
        <div class='kpi-label'>👥 Bénéficiaires Sécurisés</div>
        <div class='kpi-delta kpi-up'>↑ Population protégée</div></div>""", unsafe_allow_html=True)
    c4.markdown(f"""<div class='kpi-card red'><div class='kpi-value'>{eng:,}</div>
        <div class='kpi-label'>💥 Engins Détruits/Neutralisés</div>
        <div class='kpi-delta'>UXO · AMD · IED · ALPC</div></div>""", unsafe_allow_html=True)
    c5.markdown(f"""<div class='kpi-card orange'><div class='kpi-value'>{qual:.1f}%</div>
        <div class='kpi-label'>⭐ Score Qualité Données</div>
        <div class='kpi-delta kpi-up'>Seuil cible : 85%</div></div>""", unsafe_allow_html=True)

    st.markdown("")
    col1,col2 = st.columns([3,2])
    with col1:
        st.markdown("<div class='section-header'>📈 Évolution mensuelle</div>", unsafe_allow_html=True)
        monthly = filt.groupby("Mois").agg(
            Sites=("ID_Site","count"),
            Bénéficiaires=("Population_Bénéficiaire","sum")
        ).reset_index()
        fig = make_subplots(specs=[[{"secondary_y":True}]])
        fig.add_trace(go.Bar(x=monthly["Mois"],y=monthly["Sites"],name="Sites",marker_color="#2E75B6"),secondary_y=False)
        fig.add_trace(go.Scatter(x=monthly["Mois"],y=monthly["Bénéficiaires"],name="Bénéficiaires",
                                 line=dict(color="#DC2626",width=2.5),mode="lines+markers"),secondary_y=True)
        fig.update_layout(height=280,margin=dict(l=0,r=0,t=10,b=0),plot_bgcolor="white",paper_bgcolor="white",
                          legend=dict(orientation="h",y=1.12))
        fig.update_xaxes(showgrid=False)
        fig.update_yaxes(showgrid=True,gridcolor="#F0F0F0")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown("<div class='section-header'>🎯 Statut des Sites</div>", unsafe_allow_html=True)
        sc = filt["Statut"].value_counts().reset_index()
        sc.columns=["Statut","Nb"]
        colors={"Complété":"#16A34A","En cours":"#2E75B6","Planifié":"#F59E0B","Suspendu":"#DC2626"}
        fig2=px.pie(sc,names="Statut",values="Nb",color="Statut",color_discrete_map=colors,hole=0.55)
        fig2.update_traces(textposition='outside',textinfo='percent+label')
        fig2.update_layout(height=280,margin=dict(l=0,r=0,t=10,b=0),showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    col3,col4 = st.columns(2)
    with col3:
        st.markdown("<div class='section-header'>🏭 Activités par Province</div>", unsafe_allow_html=True)
        pa=filt.groupby(["Province","Activité"]).size().reset_index(name="N")
        fig3=px.bar(pa,x="Province",y="N",color="Activité",barmode="stack",
                    color_discrete_sequence=px.colors.qualitative.Set2)
        fig3.update_layout(height=270,margin=dict(l=0,r=0,t=10,b=0),plot_bgcolor="white",paper_bgcolor="white")
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        st.markdown("<div class='section-header'>💣 Types d'Engins par Province</div>", unsafe_allow_html=True)
        dp=filt.groupby(["Province","Type_Danger"])["Items_Détruits"].sum().reset_index()
        fig4=px.bar(dp,x="Province",y="Items_Détruits",color="Type_Danger",barmode="group",
                    color_discrete_sequence=px.colors.qualitative.Bold)
        fig4.update_layout(height=270,margin=dict(l=0,r=0,t=10,b=0),plot_bgcolor="white",paper_bgcolor="white")
        st.plotly_chart(fig4, use_container_width=True)

    # Résultats officiels vs opérationnels
    st.markdown("<div class='section-header'>🏆 Résultats Officiels 2023 (Source : maginternational.org)</div>",
                unsafe_allow_html=True)
    cc = st.columns(3)
    cc[0].metric("ALPC détruits (officiel 2023)", "771")
    cc[1].metric("Munitions détruites (officiel 2023)", "25 101")
    cc[2].metric("Route libérée", "167 km", help="Kouba Olanga – Faya, Borkou")

    st.markdown("<div class='section-header'>📋 Derniers Sites Enregistrés</div>", unsafe_allow_html=True)
    cols_d=["ID_Site","Province","Activité","Équipe","Statut","Date_Début",
            "Population_Bénéficiaire","Items_Détruits","Qualité_Score"]
    st.dataframe(filt[cols_d].sort_values("Date_Début",ascending=False).head(12).reset_index(drop=True),
                 use_container_width=True, height=320)


# ═══════════════════════════════════════════════════════
# PAGE 2 : COLLECTE & IMPORT
# ═══════════════════════════════════════════════════════
elif page == "📥 Collecte & Import":
    st.markdown("## 📥 Collecte & Import de Données")

    tab1,tab2 = st.tabs(["📤 Importer un fichier","✍️ Saisie Manuelle"])

    with tab1:
        st.markdown("<div class='section-header'>Import CSV / Excel (KoboToolbox · Survey123 · ODK · CSPro)</div>",
                    unsafe_allow_html=True)
        col_up,col_info = st.columns([2,1])
        with col_up:
            uploaded = st.file_uploader("Glissez votre fichier ici", type=["csv","xlsx","xls"])
            if uploaded:
                try:
                    df_up = pd.read_csv(uploaded) if uploaded.name.endswith(".csv") else pd.read_excel(uploaded)
                    st.markdown(f"<div class='alert-success'>✅ <strong>{uploaded.name}</strong> — {len(df_up)} lignes · {len(df_up.columns)} colonnes</div>",
                                unsafe_allow_html=True)
                    st.dataframe(df_up.head(20), use_container_width=True)
                    st.download_button("💾 Télécharger (CSV)", df_up.to_csv(index=False), "donnees.csv","text/csv")
                except Exception as e:
                    st.markdown(f"<div class='alert-danger'>❌ Erreur : {e}</div>", unsafe_allow_html=True)
            else:
                st.markdown("<div class='alert-info'>ℹ️ Aucun fichier importé — données de démonstration affichées.</div>",
                            unsafe_allow_html=True)
                st.dataframe(raw_df.head(15), use_container_width=True)
        with col_info:
            st.markdown("**Sources compatibles**")
            for s in ["✅ KoboToolbox (CSV/API)","✅ Survey123 (Excel)","✅ ODK Collect","✅ CSPro","✅ Google Forms"]:
                st.markdown(s)
            st.info(f"📊 Démo : {len(raw_df)} enregistrements · {len(raw_df.columns)} variables")

    with tab2:
        st.markdown("<div class='section-header'>Formulaire de Saisie Terrain</div>", unsafe_allow_html=True)
        with st.form("saisie_form"):
            c1,c2,c3 = st.columns(3)
            prov_f  = c1.selectbox("Province", PROVINCES)
            act_f   = c2.selectbox("Activité", ACTIVITES)
            eq_f    = c3.selectbox("Équipe", EQUIPES)
            c4,c5,c6 = st.columns(3)
            lat_f   = c4.number_input("Latitude",value=13.5,format="%.5f")
            lon_f   = c5.number_input("Longitude",value=15.5,format="%.5f")
            surf_f  = c6.number_input("Superficie (m²)",min_value=0,value=1000)
            c7,c8,c9 = st.columns(3)
            pop_f   = c7.number_input("Population bénéficiaire",min_value=0,value=200)
            itm_f   = c8.number_input("Items trouvés",min_value=0,value=0)
            dng_f   = c9.selectbox("Type de danger", DANGERS)
            obs_f   = st.text_area("Observations", placeholder="Décrivez le contexte du site…")
            submitted = st.form_submit_button("💾 Enregistrer l'entrée", use_container_width=True)
            if submitted:
                st.markdown(f"""<div class='alert-success'>✅ <strong>Entrée enregistrée !</strong><br>
                    Province : {prov_f} · Activité : {act_f} · Équipe : {eq_f} · Pop. : {pop_f:,}
                </div>""", unsafe_allow_html=True)
                st.balloons()


# ═══════════════════════════════════════════════════════
# PAGE 3 : CONTRÔLE QUALITÉ
# ═══════════════════════════════════════════════════════
elif page == "🔍 Contrôle Qualité":
    st.markdown("## 🔍 Contrôle Qualité des Données")

    df_q = raw_df.copy()
    n_total = len(df_q)

    df_q["Erreur_Population"] = df_q["Population"].isna()
    df_q["Erreur_Items"]      = df_q["Items_Trouvés"].isna() | \
                                 (pd.to_numeric(df_q["Items_Trouvés"],errors='coerce').fillna(-1) < 0)
    df_q["Erreur_Surface"]    = df_q["Superficie"].isna() | \
                                 (pd.to_numeric(df_q["Superficie"],errors='coerce').fillna(0) == 0)
    df_q["Erreur_Coords"]     = (df_q["Latitude"]<7)|(df_q["Latitude"]>24)|\
                                 (df_q["Longitude"]<13)|(df_q["Longitude"]>24)
    df_q["Nb_Erreurs"]        = df_q[["Erreur_Population","Erreur_Items",
                                       "Erreur_Surface","Erreur_Coords","Doublon"]].sum(axis=1)
    df_q["Qualité"]           = df_q["Nb_Erreurs"].apply(
        lambda x: "🟢 Conforme" if x==0 else ("🟡 Avertissement" if x==1 else "🔴 Critique"))

    n_ok   = int((df_q["Nb_Erreurs"]==0).sum())
    n_warn = int((df_q["Nb_Erreurs"]==1).sum())
    n_crit = int((df_q["Nb_Erreurs"]>=2).sum())
    n_dup  = int(df_q["Doublon"].sum())
    score_q = round(n_ok/n_total*100, 1)

    st.markdown("<div class='section-header'>📊 Synthèse Qualité</div>", unsafe_allow_html=True)
    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("Total",n_total); k2.metric("✅ Conformes",n_ok,f"{round(n_ok/n_total*100,1)}%")
    k3.metric("⚠️ Avertissements",n_warn); k4.metric("🔴 Critiques",n_crit); k5.metric("🔁 Doublons",n_dup)

    col_g,col_e = st.columns([1,2])
    with col_g:
        fig_g = go.Figure(go.Indicator(
            mode="gauge+number+delta", value=score_q,
            delta={"reference":85,"suffix":"%"},
            title={"text":"Score Qualité Global","font":{"size":14}},
            gauge={"axis":{"range":[0,100]},"bar":{"color":"#2E75B6"},
                   "steps":[{"range":[0,60],"color":"#FEE2E2"},{"range":[60,80],"color":"#FEF3C7"},
                             {"range":[80,100],"color":"#DCFCE7"}],
                   "threshold":{"line":{"color":"#1F4E79","width":3},"thickness":0.75,"value":85}}
        ))
        fig_g.update_layout(height=220,margin=dict(l=20,r=20,t=40,b=10))
        st.plotly_chart(fig_g, use_container_width=True)

    with col_e:
        st.markdown("<div class='section-header'>🚨 Erreurs Détectées</div>", unsafe_allow_html=True)
        erreurs = {
            "Valeurs manquantes (Population)": int(df_q["Erreur_Population"].sum()),
            "Items négatifs/manquants"        : int(df_q["Erreur_Items"].sum()),
            "Surface nulle/manquante"          : int(df_q["Erreur_Surface"].sum()),
            "Coordonnées hors Tchad"           : int(df_q["Erreur_Coords"].sum()),
            "Doublons"                         : n_dup,
        }
        for label, count in erreurs.items():
            pct = round(count/n_total*100,1)
            cls = "alert-danger" if count>10 else ("alert-warning" if count>0 else "alert-success")
            icn = "🔴" if count>10 else ("⚠️" if count>0 else "✅")
            st.markdown(f"<div class='{cls}'>{icn} <strong>{label}</strong> : {count} ({pct}%)</div>",
                        unsafe_allow_html=True)

    st.markdown("<div class='section-header'>📋 Détail des Enregistrements</div>", unsafe_allow_html=True)
    filtre_q = st.selectbox("Filtrer :", ["Tous","🔴 Critique","🟡 Avertissement","🟢 Conforme","🔁 Doublons"])
    df_show = df_q.copy()
    if filtre_q=="🔁 Doublons": df_show=df_show[df_show["Doublon"]==True]
    elif filtre_q!="Tous":      df_show=df_show[df_show["Qualité"]==filtre_q]

    cols_s=["ID","Province","Collecteur","Date_Soumission","Activité",
            "Population","Items_Trouvés","Superficie","Qualité","Nb_Erreurs"]
    st.dataframe(df_show[cols_s].reset_index(drop=True), use_container_width=True, height=360)

    buf=io.BytesIO(); df_show[cols_s].to_excel(buf,index=False); buf.seek(0)
    st.download_button("📥 Exporter anomalies (Excel)", buf, "rapport_qualite_MAG.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ═══════════════════════════════════════════════════════
# PAGE 4 : CARTE OPÉRATIONNELLE
# ═══════════════════════════════════════════════════════
elif page == "🗺️ Carte Opérationnelle":
    st.markdown("## 🗺️ Carte Opérationnelle MAG Tchad")
    st.caption("Zones : Lac Tchad · Borkou · Ennedi · Fada · Kalait · Tibesti")

    col_ctrl,col_map = st.columns([1,3])
    with col_ctrl:
        layer  = st.radio("Affichage",["Marqueurs par statut","Densité (HeatMap)","Clusters"])
        bmap   = st.selectbox("Fond de carte",["CartoDB Positron","OpenStreetMap","CartoDB DarkMatter"])
        show_p = st.checkbox("Taille ∝ population",True)
        fmap   = st.multiselect("Activité(s)",ACTIVITES,default=ACTIVITES[:4])
        df_map = filt[filt["Activité"].isin(fmap)]
        st.markdown(f"**{len(df_map)} sites affichés**")
        cmap={"Complété":"green","En cours":"blue","Planifié":"orange","Suspendu":"red"}
        for k,v in cmap.items():
            icon="🟢" if v=="green" else "🔵" if v=="blue" else "🟠" if v=="orange" else "🔴"
            st.markdown(f"{icon} {k}")

    with col_map:
        tiles={"OpenStreetMap":"OpenStreetMap","CartoDB Positron":"CartoDB positron","CartoDB DarkMatter":"CartoDB dark_matter"}
        m=folium.Map(location=[15.5,18.0],zoom_start=5,tiles=tiles[bmap])

        if layer=="Densité (HeatMap)":
            from folium.plugins import HeatMap
            HeatMap(df_map[["Latitude","Longitude"]].dropna().values.tolist(),radius=22,blur=18).add_to(m)

        elif layer=="Clusters":
            from folium.plugins import MarkerCluster
            mc=MarkerCluster().add_to(m)
            for _,row in df_map.iterrows():
                popup=f"<div style='font-family:Arial;min-width:200px;'><b style='color:#1F4E79;'>{row['ID_Site']}</b><hr style='margin:4px 0;'>📍 {row['Province']}<br>🎯 {row['Activité']}<br>⚠️ {row['Type_Danger']}<br>👥 {int(row['Population_Bénéficiaire']):,} bénéficiaires<br>💥 {int(row['Items_Détruits'])} engins détruits<br>⭐ {row['Qualité_Score']}%</div>"
                folium.Marker([row["Latitude"],row["Longitude"]],
                              popup=folium.Popup(popup,max_width=260),
                              icon=folium.Icon(color=cmap.get(row["Statut"],"gray"),icon="info-sign")).add_to(mc)

        else:
            for _,row in df_map.iterrows():
                popup=f"<div style='font-family:Arial;min-width:200px;'><b style='color:#1F4E79;'>{row['ID_Site']}</b><hr style='margin:4px 0;'>📍 {row['Province']}<br>🎯 {row['Activité']}<br>⚠️ {row['Type_Danger']}<br>👥 {int(row['Population_Bénéficiaire']):,}<br>💥 {int(row['Items_Détruits'])}<br>⭐ {row['Qualité_Score']}%<br>📅 {str(row['Date_Début'])[:10]}</div>"
                r=max(5,min(8+int(row["Population_Bénéficiaire"]/1000),22)) if show_p else 8
                folium.CircleMarker([row["Latitude"],row["Longitude"]],radius=r,
                                    color=cmap.get(row["Statut"],"gray"),fill=True,fill_opacity=0.72,
                                    popup=folium.Popup(popup,max_width=260),
                                    tooltip=f"{row['ID_Site']} | {row['Province']} | {row['Statut']}").add_to(m)

        st_folium(m, height=540, use_container_width=True)


# ═══════════════════════════════════════════════════════
# PAGE 5 : ANALYSE & RAPPORTS
# ═══════════════════════════════════════════════════════
elif page == "📈 Analyse & Rapports":
    st.markdown("## 📈 Analyse Avancée & Indicateurs")

    if len(filt)==0:
        st.warning("Aucun site ne correspond aux filtres."); st.stop()

    tab1,tab2,tab3 = st.tabs(["📈 Tendances","🏆 Performance Équipes","📐 Statistiques Descriptives"])

    with tab1:
        st.markdown("<div class='section-header'>Évolution mensuelle</div>", unsafe_allow_html=True)
        m2=filt.groupby(["Mois","Statut"]).size().reset_index(name="N")
        ft=px.area(m2,x="Mois",y="N",color="Statut",
                   color_discrete_map={"Complété":"#16A34A","En cours":"#2E75B6","Planifié":"#F59E0B","Suspendu":"#DC2626"})
        ft.update_layout(height=300,plot_bgcolor="white",paper_bgcolor="white",margin=dict(l=0,r=0,t=10,b=0))
        st.plotly_chart(ft, use_container_width=True)

        ca,cb=st.columns(2)
        with ca:
            st.markdown("<div class='section-header'>Surface libérée (m²)</div>", unsafe_allow_html=True)
            sp=filt[filt["Statut"]=="Complété"].groupby("Province")["Superficie_m2"].sum().reset_index()
            fs=px.bar(sp.sort_values("Superficie_m2",ascending=True),x="Superficie_m2",y="Province",
                      orientation="h",color="Superficie_m2",color_continuous_scale="Blues")
            fs.update_layout(height=260,plot_bgcolor="white",paper_bgcolor="white",margin=dict(l=0,r=0,t=10,b=0))
            st.plotly_chart(fs, use_container_width=True)
        with cb:
            st.markdown("<div class='section-header'>Distribution Score Qualité</div>", unsafe_allow_html=True)
            fh=px.histogram(filt,x="Qualité_Score",nbins=20,color_discrete_sequence=["#2E75B6"])
            fh.add_vline(x=filt["Qualité_Score"].mean(),line_dash="dash",line_color="#DC2626",
                         annotation_text=f"Moy: {filt['Qualité_Score'].mean():.1f}%")
            fh.update_layout(height=260,plot_bgcolor="white",paper_bgcolor="white",margin=dict(l=0,r=0,t=10,b=0))
            st.plotly_chart(fh, use_container_width=True)

    with tab2:
        st.markdown("<div class='section-header'>Performance par Équipe</div>", unsafe_allow_html=True)
        eq=filt.groupby("Équipe").agg(
            Sites=("ID_Site","count"),
            Complétés=("Statut",lambda x:(x=="Complété").sum()),
            Pop_Total=("Population_Bénéficiaire","sum"),
            Engins_Détruits=("Items_Détruits","sum"),
            Surface_m2=("Superficie_m2","sum"),
            Score_Qualité=("Qualité_Score","mean")
        ).reset_index()
        eq["Taux_Completion"]=(eq["Complétés"]/eq["Sites"]*100).round(1)
        eq["Score_Qualité"]=eq["Score_Qualité"].round(1)

        fe=px.scatter(eq,x="Taux_Completion",y="Score_Qualité",size="Pop_Total",
                      color="Équipe",text="Équipe",size_max=55,
                      color_discrete_sequence=px.colors.qualitative.Bold)
        fe.update_traces(textposition="top center")
        fe.update_layout(height=300,plot_bgcolor="white",paper_bgcolor="white",
                         xaxis_title="Taux de complétion (%)",yaxis_title="Score qualité (%)",
                         margin=dict(l=0,r=0,t=10,b=0))
        st.plotly_chart(fe, use_container_width=True)

        # ✅ FIX : format simple sans background_gradient (évite ImportError matplotlib)
        st.dataframe(
            eq.style.format({
                "Score_Qualité":"{:.1f}%","Taux_Completion":"{:.1f}%",
                "Pop_Total":"{:,.0f}","Engins_Détruits":"{:,.0f}","Surface_m2":"{:,.0f}",
            }),
            use_container_width=True
        )

    with tab3:
        st.markdown("<div class='section-header'>Statistiques Descriptives</div>", unsafe_allow_html=True)
        nc=["Superficie_m2","Population_Bénéficiaire","Items_Trouvés","Items_Détruits","Qualité_Score"]
        desc=filt[nc].describe().round(2).T.rename(columns={
            "count":"N","mean":"Moyenne","std":"Écart-type",
            "min":"Min","25%":"Q1","50%":"Médiane","75%":"Q3","max":"Max"
        })
        st.dataframe(desc, use_container_width=True)

        st.markdown("<div class='section-header'>Matrice de Corrélation</div>", unsafe_allow_html=True)
        corr=filt[nc].corr().round(2)
        fc=px.imshow(corr,text_auto=True,aspect="auto",color_continuous_scale="RdBu_r",zmin=-1,zmax=1)
        fc.update_layout(height=320,margin=dict(l=0,r=0,t=10,b=0))
        st.plotly_chart(fc, use_container_width=True)


# ═══════════════════════════════════════════════════════
# PAGE 6 : EXPORT RAPPORT WORD
# ═══════════════════════════════════════════════════════
elif page == "📄 Export Rapport Word":
    st.markdown("## 📄 Génération de Rapport Opérationnel Word")
    st.markdown("<div class='alert-info'>ℹ️ Configurez les paramètres puis cliquez sur <strong>Générer le rapport</strong>.</div>",
                unsafe_allow_html=True)

    cc,cp=st.columns([1,2])
    with cc:
        titre = st.text_input("Titre","Rapport Opérationnel MAG Tchad")
        periode = st.text_input("Période","Janvier – Décembre 2024")
        prep_par = st.text_input("Préparé par","DJAOYANG HABEKREO Pelandi")
        dest = st.text_input("Destinataires","Direction Pays, PIMU, Bailleurs")
        incl_stats  = st.checkbox("Inclure statistiques descriptives",True)
        incl_qualite= st.checkbox("Inclure rapport qualité",True)
        gen_btn = st.button("📝 Générer le rapport Word",use_container_width=True,type="primary")

    with cp:
        st.markdown("<div class='section-header'>Aperçu du Rapport</div>", unsafe_allow_html=True)
        if len(filt)>0:
            ns=len(filt); nc2=(filt["Statut"]=="Complété").sum()
            pt=filt["Population_Bénéficiaire"].sum(); ed=filt["Items_Détruits"].sum()
            qm=filt["Qualité_Score"].mean()
            st.markdown(f"""
            📌 **{titre}** · 📅 {periode}  ✍️ Par : {prep_par}
            ---
            | Indicateur | Valeur |
            |---|---|
            | Sites opérationnels | **{ns}** |
            | Sites complétés | **{nc2} ({round(nc2/ns*100,1)}%)** |
            | Bénéficiaires | **{pt:,}** |
            | Engins détruits | **{ed:,}** |
            | Score qualité | **{qm:.1f}%** |
            | ALPC détruites 2023 *(officiel)* | **771** |
            | Munitions détruites 2023 *(officiel)* | **25 101** |
            """)

    if gen_btn and len(filt)>0:
        with st.spinner("Génération en cours…"):
            doc=Document()
            s=doc.styles['Normal']; s.font.name='Arial'; s.font.size=Pt(11)

            tp=doc.add_paragraph(); tp.alignment=WD_ALIGN_PARAGRAPH.CENTER
            run=tp.add_run(titre.upper()); run.bold=True; run.font.size=Pt(18)
            run.font.color.rgb=RGBColor(0x1F,0x4E,0x79)

            for txt in [f"Période : {periode}",f"Préparé par : {prep_par} | Dest. : {dest}",
                        f"Date : {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                        "Source : maginternational.org/chad"]:
                p=doc.add_paragraph(txt); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph()

            def add_section(title):
                p=doc.add_paragraph(); run=p.add_run(title)
                run.bold=True; run.font.size=Pt(13)
                run.font.color.rgb=RGBColor(0x2E,0x75,0xB6)
                doc.add_paragraph()

            ns=len(filt); nc2=(filt["Statut"]=="Complété").sum()
            pt=filt["Population_Bénéficiaire"].sum(); ed=filt["Items_Détruits"].sum()
            qm=filt["Qualité_Score"].mean()

            add_section("1. Résultats Officiels MAG Tchad 2023")
            t0=doc.add_table(rows=4,cols=2); t0.style='Table Grid'
            for i,(k,v) in enumerate([("Indicateur","Valeur (Source : maginternational.org/chad)"),
                                       ("Armes légères détruites","771"),
                                       ("Munitions détruites","25 101"),
                                       ("Route libérée (Borkou)","167 km")]):
                t0.rows[i].cells[0].text=k; t0.rows[i].cells[1].text=v
                if i==0:
                    for c in t0.rows[i].cells: c.paragraphs[0].runs[0].bold=True
            doc.add_paragraph()

            add_section("2. Indicateurs Opérationnels")
            t1=doc.add_table(rows=6,cols=2); t1.style='Table Grid'
            for i,(k,v) in enumerate([("Indicateur","Valeur"),
                                       ("Sites opérationnels",str(ns)),
                                       ("Sites complétés",f"{nc2} ({round(nc2/ns*100,1)}%)"),
                                       ("Population bénéficiaire",f"{pt:,}"),
                                       ("Engins détruits",f"{ed:,}"),
                                       ("Score qualité moyen",f"{qm:.1f}%")]):
                t1.rows[i].cells[0].text=k; t1.rows[i].cells[1].text=v
                if i==0:
                    for c in t1.rows[i].cells: c.paragraphs[0].runs[0].bold=True
            doc.add_paragraph()

            add_section("3. Répartition par Province")
            ps=filt.groupby("Province").agg(Sites=("ID_Site","count"),
                Pop=("Population_Bénéficiaire","sum"),Engins=("Items_Détruits","sum")).reset_index()
            t2=doc.add_table(rows=len(ps)+1,cols=4); t2.style='Table Grid'
            for j,h in enumerate(["Province","Sites","Bénéficiaires","Engins Détruits"]):
                t2.rows[0].cells[j].text=h; t2.rows[0].cells[j].paragraphs[0].runs[0].bold=True
            for i,row in ps.iterrows():
                t2.rows[i+1].cells[0].text=row["Province"]
                t2.rows[i+1].cells[1].text=str(row["Sites"])
                t2.rows[i+1].cells[2].text=f"{row['Pop']:,}"
                t2.rows[i+1].cells[3].text=str(row["Engins"])
            doc.add_paragraph()

            if incl_stats:
                add_section("4. Statistiques Descriptives")
                nc3=["Superficie_m2","Population_Bénéficiaire","Items_Détruits","Qualité_Score"]
                d=filt[nc3].describe().round(1)
                t3=doc.add_table(rows=len(nc3)+1,cols=5); t3.style='Table Grid'
                for j,h in enumerate(["Variable","Moyenne","Médiane","Min","Max"]):
                    t3.rows[0].cells[j].text=h; t3.rows[0].cells[j].paragraphs[0].runs[0].bold=True
                for i,col in enumerate(nc3):
                    t3.rows[i+1].cells[0].text=col
                    t3.rows[i+1].cells[1].text=str(d.loc["mean",col])
                    t3.rows[i+1].cells[2].text=str(d.loc["50%",col])
                    t3.rows[i+1].cells[3].text=str(d.loc["min",col])
                    t3.rows[i+1].cells[4].text=str(d.loc["max",col])
                doc.add_paragraph()

            if incl_qualite:
                add_section("5. Qualité des Données")
                doc.add_paragraph(f"Score qualité global : {qm:.1f}%")
                doc.add_paragraph(f"Enregistrements analysés : {len(raw_df)}")
                doc.add_paragraph(f"Doublons détectés : {int(raw_df['Doublon'].sum())}")
                doc.add_paragraph(f"Valeurs manquantes (Population) : {int(raw_df['Population'].isna().sum())}")
                doc.add_paragraph()

            add_section("6. Bailleurs de Fonds")
            for b in MAG_INFO["bailleurs"]:
                doc.add_paragraph(f"• {b}")

            doc.add_paragraph()
            doc.add_paragraph(f"Fait à N'Djamena, le {datetime.now().strftime('%d/%m/%Y')}")
            doc.add_paragraph(f"Responsable de la Gestion de l'Information : {prep_par}")

            buf=io.BytesIO(); doc.save(buf); buf.seek(0)

        st.markdown("<div class='alert-success'>✅ Rapport généré avec succès !</div>",
                    unsafe_allow_html=True)
        st.download_button("📥 Télécharger le Rapport Word", buf,
                           f"Rapport_MAG_Tchad_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                           "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True)

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("""
<div class='footer'>
    MAG Tchad – Système de Gestion de l'Information ·
    <a href='https://www.maginternational.org/what-we-do/where-we-work/chad/' target='_blank'>
    maginternational.org/chad</a><br>
    Développé par <strong>DJAOYANG HABEKREO Pelandi</strong> · Ingénieur Statisticien Économiste – ISSEA Yaoundé<br>
    <em>Candidat – Responsable de la Gestion de l'Information · MAG Tchad · 2026</em>
</div>
""", unsafe_allow_html=True)
