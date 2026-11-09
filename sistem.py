# ============================================
# EDUCLUSTER SIANTAR - FINAL PERFECT REVISION
# ============================================

import streamlit as st
import pandas as pd
import numpy as np
from sklearn.cluster import DBSCAN
from sklearn.preprocessing import StandardScaler
import plotly.express as px
from datetime import datetime
import io

# ============================================
# 1. KONFIGURASI HALAMAN
# ============================================
st.set_page_config(
    page_title="EduCluster Siantar",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================
# 2. CSS CUSTOM (PERBAIKAN SLIDER & TOMBOL)
# ============================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    /* Global Font */
    .stApp {
        font-family: 'Inter', sans-serif;
        background-color: #f8fafc;
    }
    
    /* PAKSA FONT JUDUL GEDE (30PX) */
    h1, h2, h3, h4, h5 {
        font-size: 30px !important;
        font-weight: 800 !important;
        color: #0f172a;
        margin-bottom: 15px;
    }
    p, label, div {
        font-size: 18px !important;
    }

    /* PAKSA FONT TABEL JADI GEDE */
    .stDataFrame {
        font-size: 24px !important; 
    }
    .stDataFrame th, .stDataFrame td {
        font-size: 20px !important; 
        padding: 12px !important;
    }

    /* SIDEBAR STYLING */
    [data-testid="stSidebar"] {
        background-color: #0f172a;
        border-right: 1px solid #1e293b;
        padding-top: 20px;
    }
    
    /* Paksa semua teks di sidebar jadi putih dan bold */
    [data-testid="stSidebar"] * {
        color: white !important;
        font-weight: bold !important;
    }

    /* Kelas Font Khusus untuk Sidebar */
    .sidebar-title-large {
        font-size: 20px !important;
        font-weight: 800 !important;
        color: white !important;
        margin-top: 25px;
        margin-bottom: 15px;
    }

    /* PERBAIKAN SLIDER (BIAR GA NEMPEL) */
    [data-testid="stSidebar"] .stSlider label {
        font-size: 18px !important;
        margin-bottom: 12px !important; /* Jarak antara judul dan angka */
        display: block !important;
    }
    [data-testid="stSidebar"] .stSlider [data-baseweb="slider"] {
        margin-top: 5px !important; /* Jarak antara angka dan garis slider */
    }
    /* Kecilkan sedikit angka di slider biar ga tabrakan */
    [data-testid="stSidebar"] .stSlider label span {
        font-size: 16px !important; 
    }

    /* PERBAIKAN TOMBOL MENU (RATA KIRI & ADA BACKGROUND) */
    [data-testid="stSidebar"] .stButton > button {
        text-align: left !important; /* RATA KIRI */
        background-color: #1e293b !important; /* BACKGROUND SEBELUM DIKLIK */
        color: white !important;
        border: 1px solid #334155 !important;
        width: 100% !important;
        padding: 12px 15px !important;
        margin-bottom: 8px !important; /* Jarak antar tombol */
        border-radius: 6px !important;
        font-size: 16px !important;
        font-weight: 600 !important;
        transition: all 0.2s !important;
        display: flex !important;
        align-items: center !important;
    }
    [data-testid="stSidebar"] .stButton > button:hover {
        background-color: #334155 !important; /* Warna saat di-hover */
        transform: translateX(3px) !important; /* Efek geser dikit */
        border-color: #475569 !important;
    }
    
    /* Tombol Analisis Utama (Beda gaya, tetap full width tapi center text) */
    [data-testid="stSidebar"] .stButton > button[kind="primary"] {
        background-color: #1e3a8a !important;
        color: white !important;
        border: none !important;
        text-align: center !important;
        margin-top: 25px !important;
        justify-content: center !important;
    }
    [data-testid="stSidebar"] .stButton > button[kind="primary"]:hover {
        background-color: #1e40af !important;
        transform: scale(1.02) !important;
    }

    /* File Uploader (Background Putih, Teks Hitam) */
    [data-testid="stFileUploader"] {
        background-color: white !important;
        border-radius: 8px;
        padding: 10px;
        margin-bottom: 15px;
        border: 1px solid #e2e8f0 !important;
    }
    [data-testid="stFileUploader"] * {
        color: black !important;
        font-weight: normal !important;
    }
    [data-testid="stFileUploader"] button {
        background-color: #f1f5f9 !important;
        color: #0f172a !important;
        font-weight: bold !important;
    }

    /* Expander Hak Cipta */
    .stExpander {
        background-color: #1e293b !important;
        border: 1px solid #334155 !important;
        border-radius: 8px;
    }
    .stExpander * {
        color: white !important;
        font-weight: normal !important;
    }

    /* METRIC CARDS */
    .metric-box {
        background: white;
        padding: 20px;
        border-radius: 10px;
        border-top: 5px solid;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        text-align: center;
    }
    .metric-value { font-size: 40px !important; font-weight: 900; margin: 10px 0; }
    .metric-label { font-size: 16px; font-weight: 700; color: #64748b; text-transform: uppercase; }
    .metric-sub { font-size: 14px; color: #94a3b8; }
    </style>
""", unsafe_allow_html=True)

# ============================================
# 3. FUNGSI HELPER
# ============================================
def create_watermarked_excel(df):
    watermark_info = pd.DataFrame({
        'Kecamatan': ['=== DILINDUNGI HAK CIPTA ==='],
        'Jumlah_Sekolah': ['© 2026 - Analisis oleh [KIP Analytics]'],
        'Fasilitas_Skor': ['Dilarang mengklaim tanpa izin'],
        'Partisipasi_Sekolah': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
        'Cluster': ['']
    })
    df_final = pd.concat([watermark_info, df], ignore_index=True)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Hasil Klaster')
    return output.getvalue()

def load_dummy_data():
    return pd.DataFrame({
        'Kecamatan': ['Siantar Barat', 'Siantar Timur', 'Siantar Utara', 'Siantar Selatan', 
                      'Siantar Martoba', 'Siantar Marihat', 'Siantar Marimbun', 'Siantar Sitalasari'],
        'Jumlah_Sekolah': [25, 30, 20, 28, 22, 15, 10, 18],
        'Fasilitas_Skor': [8.5, 9.0, 7.5, 8.0, 7.0, 5.5, 4.0, 6.5],
        'Partisipasi_Sekolah': [95, 97, 90, 93, 88, 75, 65, 82]
    })

# ============================================
# 4. INISIALISASI STATE
# ============================================
if 'page' not in st.session_state:
    st.session_state['page'] = 'dashboard'
if 'df' not in st.session_state:
    st.session_state['df'] = None

# ============================================
# 5. SIDEBAR (RAPAT, FONT DIBALIK, RAPI)
# ============================================
with st.sidebar:
    # Header Profil
    st.markdown("""
    <div style="background: #1e293b; padding: 15px; border-radius: 10px; margin-bottom: 20px; display: flex; align-items: center; gap: 15px; border: 1px solid #334155;">
        <div style="width: 45px; height: 45px; background: #3b82f6; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: white; font-weight: bold; font-size: 20px;">A</div>
        <div>
            <div style="font-size: 18px;">Administrator</div>
            <div style="font-size: 14px; color: #4ade80;">● Online</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 1. MENU NAVIGASI (Tombol Rata Kiri dengan Background)
    if st.button("🏠 Dashboard", use_container_width=True):
        st.session_state['page'] = 'dashboard'

    if st.button("📊 Tampilkan Database", use_container_width=True):
        st.session_state['df'] = load_dummy_data()
        st.session_state['page'] = 'data'
        st.success("Database dimuat!", icon="✅")

    # Upload CSV (Background putih, teks hitam via CSS)
    st.markdown('<p style="font-size: 14px; margin-bottom: 5px; margin-top: 10px;">Atau upload file CSV:</p>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type=['csv'], label_visibility="collapsed")
    if uploaded_file is not None:
        st.session_state['df'] = pd.read_csv(uploaded_file)
        st.session_state['page'] = 'data'
        st.success(f"{len(st.session_state['df'])} baris dimuat!", icon="✅")

    # 2. PARAMETER DBSCAN (Font Dibalik: Judul Gede, Label Kecil)
    st.markdown('<p class="sidebar-title-large">🧠 Parameter DBSCAN</p>', unsafe_allow_html=True)
    
    # Slider dengan CSS yang sudah diperbaiki (ada jarak)
    eps = st.slider("Epsilon (eps)", 0.1, 3.0, 0.8, 0.1, help="Jarak maksimal antar titik")
    min_samples = st.slider("Min Samples", 2, 10, 2, help="Jumlah minimal anggota cluster")

    # 3. TOMBOL ANALISIS (Dikasih jarak)
    run_analysis = st.button("🚀 Jalankan Analisis", type="primary", use_container_width=True)

    # 4. HAK CIPTA (Paling Bawah, Teks Putih)
    st.markdown('<div style="margin-top: 20px;"></div>', unsafe_allow_html=True)
    with st.expander("ℹ️ Informasi Hak Cipta"):
        st.markdown("""
        <p style="color: white; font-size: 14px;">
        <b>Sistem ini dilindungi Hak Cipta.</b><br><br>
        Seluruh data, algoritma, dan hasil analisis merupakan kekayaan intelektual.<br><br>
        Dilarang memperbanyak atau mengklaim karya ini tanpa izin.<br><br>
        © 2026 EduCluster Team.
        </p>
        """, unsafe_allow_html=True)

# ============================================
# 6. KONTEN UTAMA (DASHBOARD & HASIL)
# ============================================

# --- HALAMAN DASHBOARD (WELCOME) ---
if st.session_state['page'] == 'dashboard':
    st.title("🎓 Dashboard Analisis Pendidikan")
    st.caption("Kota Pematangsiantar • Sistem Pendukung Keputusan Berbasis Data")
    
    st.markdown("---")
    st.subheader("Selamat Datang di EduCluster Siantar")
    st.info("Silakan pilih menu di sidebar sebelah kiri untuk memulai analisis.", icon="ℹ️")
    
    st.markdown("""
    ### 📌 Tentang Sistem Ini
    Sistem ini menggunakan algoritma **DBSCAN** untuk mengelompokkan kecamatan berdasarkan:
    1. Jumlah Sekolah
    2. Skor Fasilitas
    3. Partisipasi Sekolah
    
    Hasil analisis akan membantu pemerintah menentukan prioritas anggaran pendidikan secara akurat dan berbasis data.
    """)

# --- HALAMAN DATA & ANALISIS ---
else:
    if st.session_state['df'] is not None:
        # Tampilkan Data Mentah
        st.subheader("📋 Database Pendidikan")
        st.dataframe(st.session_state['df'], use_container_width=True, hide_index=True)

        # --- HASIL ANALISIS ---
        if run_analysis:
            st.subheader("🔬 Hasil Analisis Klaster DBSCAN")
            
            df = st.session_state['df']
            X = df[['Jumlah_Sekolah', 'Fasilitas_Skor', 'Partisipasi_Sekolah']]
            scaler = StandardScaler()
            X_scaled = scaler.fit_transform(X)
            dbscan = DBSCAN(eps=eps, min_samples=min_samples)
            df['Cluster'] = dbscan.fit_predict(X_scaled)
            st.session_state['df'] = df 
            
            # Metrik
            c0 = len(df[df['Cluster'] == 0])
            c1 = len(df[df['Cluster'] == 1])
            noise = len(df[df['Cluster'] == -1])
            total = len(df)
            
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.markdown(f'<div class="metric-box" style="border-top-color: #10b981;"><div class="metric-label">Zona Maju</div><div class="metric-value" style="color: #10b981;">{c0}</div><div class="metric-sub">Kecamatan</div></div>', unsafe_allow_html=True)
            with m2:
                st.markdown(f'<div class="metric-box" style="border-top-color: #ef4444;"><div class="metric-label">Zona Prioritas</div><div class="metric-value" style="color: #ef4444;">{c1}</div><div class="metric-sub">Kecamatan</div></div>', unsafe_allow_html=True)
            with m3:
                st.markdown(f'<div class="metric-box" style="border-top-color: #f59e0b;"><div class="metric-label">Anomali</div><div class="metric-value" style="color: #f59e0b;">{noise}</div><div class="metric-sub">Kecamatan</div></div>', unsafe_allow_html=True)
            with m4:
                st.markdown(f'<div class="metric-box" style="border-top-color: #1e3a8a;"><div class="metric-label">Total Data</div><div class="metric-value" style="color: #1e3a8a;">{total}</div><div class="metric-sub">Kecamatan</div></div>', unsafe_allow_html=True)

            # Visualisasi (Layout Vertikal)
            st.subheader("📊 Visualisasi Data")
            
            # Grafik 1: Scatter 2D (PERBAIKAN WARNA)
            st.markdown("##### 🗺️ Peta Sebaran: Fasilitas vs Partisipasi")
            
            # Copy dataframe dan ubah Cluster jadi STRING biar warnanya diskrit (Hijau/Merah/Kuning)
            df_2d = df.copy()
            df_2d['Cluster'] = df_2d['Cluster'].astype(str)

            fig1 = px.scatter(
                df_2d, 
                x='Fasilitas_Skor', 
                y='Partisipasi_Sekolah',
                color='Cluster', 
                hover_data=['Kecamatan', 'Jumlah_Sekolah'],
                # Pakai kunci String ('0', '1', '-1') biar warnanya pas
                color_discrete_map={'0': '#10b981', '1': '#ef4444', '-1': '#f59e0b'},
                labels={'Fasilitas_Skor': 'Skor Fasilitas', 'Partisipasi_Sekolah': 'Partisipasi (%)'}
            )
            fig1.update_layout(plot_bgcolor='white', paper_bgcolor='white', font=dict(family="Inter", size=16), height=500)
            fig1.update_traces(marker=dict(size=15, line=dict(width=2, color='white')))
            st.plotly_chart(fig1, use_container_width=True)

            # Grafik 2: 3D (PERBAIKAN WARNA & SPASI)
            st.markdown("##### 📊 Distribusi 3 Dimensi")
            
            # 1. Kita copy dataframe dan ubah kolom Cluster jadi STRING
            # Ini biar Plotly nganggap dia kategori (Hijau/Merah), bukan angka gradasi biru
            df_3d = df.copy()
            df_3d['Cluster'] = df_3d['Cluster'].astype(str)
            
            fig2 = px.scatter_3d(
                df_3d, 
                x='Jumlah_Sekolah', 
                y='Fasilitas_Skor', 
                z='Partisipasi_Sekolah',
                color='Cluster', 
                hover_data=['Kecamatan'],
                # Kunci warna sekarang pakai String ('0', '1', '-1')
                color_discrete_map={'0': '#10b981', '1': '#ef4444', '-1': '#f59e0b'}
            )
            
            fig2.update_layout(
                plot_bgcolor='white', 
                paper_bgcolor='white', 
                font=dict(family="Inter", size=14), 
                height=600, 
                # INI KUNCINYA: aspectmode='cube' bikin grafik jadi kubus, titik menyebar lega!
                scene=dict(
                    xaxis_title='Jumlah Sekolah', 
                    yaxis_title='Skor Fasilitas', 
                    zaxis_title='Partisipasi (%)',
                    aspectmode='cube' 
                )
            )
            # Ukuran titik dibesarkan dikit dan opacity penuh biar warna keluar
            fig2.update_traces(marker=dict(size=12, opacity=1.0, line=dict(width=1, color='white')))
            
            st.plotly_chart(fig2, use_container_width=True)

            # Export
            st.subheader("💾 Export Laporan & Hak Cipta")
            st.info("🔒 File Excel yang diunduh otomatis dilindungi Digital Watermark.", icon="🛡️")
            
            excel_data = create_watermarked_excel(df)
            st.download_button(
                label="⬇️ Download Laporan Excel (Terproteksi)",
                data=excel_data,
                file_name=f"Laporan_Klaster_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
    else:
        st.warning("Silakan pilih menu di sidebar untuk memuat data.")

# Footer
st.divider()
st.caption("© 2026 EduCluster Siantar. Sistem Analisis Klaster Pendidikan Berbasis DBSCAN.")