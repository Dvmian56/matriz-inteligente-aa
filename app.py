import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Matriz Pro Visual", layout="wide")

# ESTILOS CSS PARA QUE PAREZCA EXCEL REAL
st.markdown("""
<style>
    .stApp { background-color: #f0f2f6; }
    div[data-testid="stExpander"] { border: 1px solid #ccc; border-radius: 5px; background: white; }
    h1 { color: #003366; }
    
    /* Estilo de Tabla Personalizado */
    table { width: 100%; border-collapse: collapse; font-family: Calibri, sans-serif; font-size: 13px; }
    th { background-color: #004a99 !important; color: white !important; text-align: center !important; border: 1px solid #ccc !important; padding: 8px !important; }
    td { border: 1px solid #ddd !important; padding: 5px !important; color: #333; }
    tr:nth-child(even) { background-color: #f9f9f9; }
</style>
""", unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Configuración")
    ac_nom = st.text_input("Administrador (RF)", "Victor Otárola")
    
    if 'mapeo' not in st.session_state: st.session_state.mapeo = []
    with st.expander("👤 Vincular Especialista"):
        e_nom = st.text_input("Nombre")
        e_disc = st.text_input("Sigla (Ej: EE)")
        if st.button("Agregar"):
            st.session_state.mapeo.append({"n": e_nom, "d": e_disc.upper()})
            st.rerun()
            
    if st.button("🗑️ Reset Equipo"):
        st.session_state.mapeo = []
        st.rerun()

# --- APP ---
st.title("🛡️ Matriz de Revisión: Vista Ejecutiva")

archivo = st.file_uploader("Sube el MDL (Excel)", type=['xlsx'])

if archivo:
    xl = pd.ExcelFile(archivo)
    hoja = st.selectbox("Selecciona Pestaña", xl.sheet_names)
    df = xl.parse(hoja).loc[:, ~xl.parse(hoja).columns.duplicated()]

    c_req = ["No. de documento", "Disciplina", "Tipo de Documento", "Título"]
    
    if all(c in df.columns for c in c_req):
        # Procesamiento
        df_base = df[c_req].drop_duplicates().sort_values(by=["Disciplina", "Tipo de Documento"])
        
        # Reglas
        df_base[ac_nom] = "RF"
        for esp in st.session_state.mapeo:
            if esp["n"] not in df_base.columns: df_base[esp["n"]] = ""
            df_base.loc[df_base["Disciplina"].astype(str).str.contains(esp["d"], na=False), esp["n"]] = "R"

        # --- PESTAÑAS VISUALES ---
        t1, t2, t3 = st.tabs(["👁️ VISTA EXCEL (Solo Lectura)", "✏️ EDITOR (Modo Trabajo)", "📥 EXPORTAR"])

        with t1:
            st.info("Vista renderizada con formato corporativo. (No editable, solo visualización)")
            
            # MOTOR DE ESTILO (PANDAS STYLER)
            def color_disciplina(val):
                return 'background-color: #e6f3ff' if isinstance(val, str) else ''

            styler = df_base.style.set_properties(**{'text-align': 'center'})\
                .set_table_styles([
                    {'selector': 'th', 'props': [('background-color', '#004a99'), ('color', 'white'), ('font-weight', 'bold')]},
                    {'selector': 'tr:hover', 'props': [('background-color', '#ffff99')]}
                ])\
                .hide(axis="index")
            
            # Renderizamos HTML puro
            st.write(styler.to_html(), unsafe_allow_html=True)

        with t2:
            st.warning("Aquí puedes cambiar los datos manualmente.")
            # Aquí dejamos el editor funcional
            cols_edit = [c for c in df_base.columns if c not in c_req]
            cfg = {c: st.column_config.SelectboxColumn(options=["", "R", "RA", "I", "RF"]) for c in cols_edit}
            df_final = st.data_editor(df_base, column_config=cfg, use_container_width=True, height=600)

        with t3:
            st.success("Descarga el archivo final para firmar.")
            if st.button("Generar Excel Oficial"):
                # (Lógica de exportación xlsxwriter igual que antes)
                pass # Aquí iría el bloque de descarga que ya tienes
    else:
        st.error("Faltan columnas clave.")
