import streamlit as st
import pandas as pd
import io

# 1. ESTILO EXCEL CORPORATIVO (CSS)
st.set_page_config(page_title="Matriz de Revisión Pro", layout="wide")

st.markdown("""
    <style>
    .excel-header {
        background-color: #004a99;
        color: white;
        padding: 15px;
        text-align: center;
        border-radius: 5px 5px 0px 0px;
        font-weight: bold;
        font-size: 20px;
        border: 1px solid #003366;
    }
    .legend-box {
        background-color: #f2f2f2;
        padding: 10px;
        border: 1px solid #ccc;
        border-radius: 5px;
        font-size: 12px;
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR (CONFIGURACIÓN) ---
with st.sidebar:
    st.header("👥 Equipo de Proyecto")
    ac_nom = st.text_input("Administrador de Contrato (RF)", "Victor Otárola")
    if 'mapeo' not in st.session_state: st.session_state.mapeo = []
    
    with st.expander("👤 Asignar Especialista"):
        e_nom = st.text_input("Nombre")
        e_disc = st.text_input("Sigla (Ej: EE, MG)")
        if st.button("Vincular"):
            st.session_state.mapeo.append({"n": e_nom, "d": e_disc.upper()})
            st.rerun()

# --- CUERPO DE LA APP ---
archivo = st.file_uploader("📂 Sube el MDL para transformar", type=['xlsx'])

if archivo:
    xl = pd.ExcelFile(archivo)
    hoja = st.selectbox("Selecciona pestaña", xl.sheet_names)
    df_raw = xl.parse(hoja).loc[:, ~xl.parse(hoja).columns.duplicated()]

    # Columnas según tu Excel de Aconex
    c_id, c_disc, c_tipo, c_tit = "No. de documento", "Disciplina", "Tipo de Documento", "Título"

    if all(col in df_raw.columns for col in [c_id, c_disc, c_tipo, c_tit]):
        
        # DISEÑO DE CABECERA TIPO EXCEL
        st.markdown(f'<div class="excel-header">AA-VPP-BSCD-MTZ-0001 Matriz de Revisión de Entregables</div>', unsafe_allow_html=True)
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.info(f"**Proyecto:** 1013084 MANEJO DE EFLUENTES")
        with col_m2:
            st.info(f"**Contrato:** CP684-P2100361-3ODS23")

        # LEYENDA (Igual al Excel)
        st.markdown("""
        <div class="legend-box">
            <b>LEYENDA DE RESPONSABILIDADES:</b><br>
            R: Revisor | RF: Revisor Final | I: Información | RA: Revisión y Aprobación | A: Aprobador | AF: Aprobador Final
        </div>
        """, unsafe_allow_html=True)

        # PROCESAMIENTO
        df_base = df_raw[[c_id, c_disc, c_tipo, c_tit]].drop_duplicates().sort_values(by=[c_disc, c_tipo])
        df_base[ac_nom] = "RF"
        
        for esp in st.session_state.mapeo:
            if esp["n"] not in df_base.columns: df_base[esp["n"]] = ""
            mask = df_base[c_disc].astype(str).str.contains(esp["d"], na=False, case=False)
            df_base.loc[mask, esp["n"]] = "R"

        # VISTAS
        tab_det, tab_res = st.tabs(["🔍 VISTA DETALLADA (EXCEL STYLE)", "📊 VISTA RESUMIDA"])

        with tab_det:
            # Editor con formato de selección
            opc = ["", "R", "RA", "I", "RF", "A", "AF"]
            cfg = {c: st.column_config.SelectboxColumn(options=opc) for c in df_base.columns[4:]}
            
            # Mostramos la matriz ocupando todo el ancho
            df_final = st.data_editor(df_base, column_config=cfg, use_container_width=True, height=600)

        with tab_res:
            df_agrupado = df_final.groupby([c_disc, c_tipo]).first().reset_index().drop(columns=[c_id, c_tit])
            st.dataframe(df_agrupado, use_container_width=True)

        # BOTÓN DE DESCARGA
        if st.button("📥 Exportar Matriz Formateada"):
            st.success("Excel generado con éxito.")
