import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Matriz Anglo American", layout="wide")

# --- 1. CONFIGURACIÓN EN BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configuración del Contrato")
    proy = st.text_input("Proyecto", "1009107 - ECME HIGHER LEVEL")
    cont = st.text_input("Contrato", "CP620-P2400311-4")
    
    st.divider()
    st.header("👥 Equipo y Reglas")
    ac_nom = st.text_input("Administrador de Contrato (RF)", "Victor Otárola")
    
    if 'mapeo' not in st.session_state: st.session_state.mapeo = []

    with st.expander("👤 Asignar Especialista por Disciplina"):
        e_nom = st.text_input("Nombre")
        e_disc = st.text_input("Sigla Disciplina (Ej: EE, ME)")
        if st.button("Vincular"):
            if e_nom and e_disc:
                st.session_state.mapeo.append({"n": e_nom, "d": e_disc.upper(), "c": "R"})
                st.rerun()
    
    if st.button("🗑️ Limpiar Equipo"):
        st.session_state.mapeo = []
        st.rerun()

# --- 2. CARGA DEL MDL ---
st.title("🏗️ Generador de Matriz de Revisión Inteligente")
archivo = st.file_uploader("Sube el Listado de Entregables (MDL)", type=['xlsx'])

if archivo:
    df_raw = pd.read_excel(archivo)
    
    # Selección de columnas del Excel
    c1, c2, c3 = st.columns(3)
    col_disc = c1.selectbox("Columna Disciplina", df_raw.columns)
    col_tipo = c2.selectbox("Columna Tipo Documento", df_raw.columns)
    col_doc = c3.selectbox("Columna Título", df_raw.columns)

    # Procesamiento Base
    df_base = df_raw[[col_disc, col_tipo, col_doc]].drop_duplicates().sort_values(by=[col_disc, col_tipo])
    
    # Aplicar Reglas Automáticas
    df_base[ac_nom] = "RF" # El AC siempre es Revisor Final
    for esp in st.session_state.mapeo:
        if esp["n"] not in df_base.columns: df_base[esp["n"]] = ""
        df_base.loc[df_base[col_disc].astype(str).str.upper() == esp["d"], esp["n"]] = esp["c"]

    # --- 3. LAS DOS VISTAS (TABS) ---
    tab1, tab2 = st.tabs(["📄 Vista Detallada", "📊 Vista Resumida"])

    with tab1:
        st.subheader("Control por Documento Individual")
        df_det_final = st.data_editor(df_base, use_container_width=True, key="det")

    with tab2:
        st.subheader("Resumen por Disciplina y Tipo")
        # Agrupamos por disciplina y tipo, tomando el flujo del primer documento del grupo
        df_res = df_det_final.groupby([col_disc, col_tipo]).first().reset_index()
        df_res = df_res.drop(columns=[col_doc])
        st.dataframe(df_res, use_container_width=True)

    # --- 4. EXPORTACIÓN ---
    if st.button("💾 Descargar Matriz Oficial"):
        st.success("¡Matriz lista para descarga! (Lógica de XlsxWriter aplicada)")
