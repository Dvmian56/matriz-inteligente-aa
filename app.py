import streamlit as st
import pandas as pd
import io

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Matriz Pro - Anglo American", layout="wide")

# --- SIDEBAR: EQUIPO Y REGLAS ---
with st.sidebar:
    st.header("📋 Datos del Contrato")
    proy_input = st.text_input("Proyecto", "1013084 MANEJO DE EFLUENTES")
    cont_input = st.text_input("Contrato", "CP684-P2100361-3ODS23")
    
    st.divider()
    st.header("👥 Equipo de Revisión")
    ac_nom = st.text_input("Administrador de Contrato (RF)", "Victor Otárola")
    
    if 'mapeo' not in st.session_state: st.session_state.mapeo = []

    with st.expander("👤 Asignar Especialista por Disciplina"):
        e_nom = st.text_input("Nombre del Revisor")
        e_disc = st.text_input("Sigla Disciplina (Ej: EE, MG, PE)")
        if st.button("Vincular"):
            if e_nom and e_disc:
                st.session_state.mapeo.append({"n": e_nom, "d": e_disc.upper()})
                st.rerun()

    if st.button("🗑️ Limpiar Equipo"):
        st.session_state.mapeo = []
        st.rerun()

# --- CUERPO PRINCIPAL ---
st.title("🛡️ Generador de Matriz Multicapa")
st.write("Selecciona una pestaña del MDL para generar su matriz de responsabilidades.")

archivo = st.file_uploader("Sube el MDL (Excel)", type=['xlsx'])

if archivo:
    # 1. LEER TODAS LAS HOJAS
    xl = pd.ExcelFile(archivo)
    hoja_sel = st.selectbox("📂 Selecciona la pestaña a procesar", xl.sheet_names)
    
    # Carga y limpieza de nombres duplicados en columnas
    df_raw = xl.parse(hoja_sel)
    df_raw = df_raw.loc[:, ~df_raw.columns.duplicated()]

    # Columnas sagradas de tu archivo Aconex
    col_id = "No. de documento"
    col_disc = "Disciplina"
    col_tipo = "Tipo de Documento"
    col_tit = "Título"

    columnas_necesarias = [col_id, col_disc, col_tipo, col_tit]
    faltantes = [c for c in columnas_necesarias if c not in df_raw.columns]

    if not faltantes:
        # 2. PROCESAMIENTO
        df_base = df_raw[columnas_necesarias].drop_duplicates().copy()
        df_base = df_base.sort_values(by=[col_disc, col_tipo])

        # REGLAS AUTOMÁTICAS
        df_base[ac_nom] = "RF"
        for esp in st.session_state.mapeo:
            if esp["n"] not in df_base.columns: df_base[esp["n"]] = ""
            mask = df_base[col_disc].astype(str).str.contains(esp["d"], na=False, case=False)
            df_base.loc[mask, esp["n"]] = "R"

        # --- VISTAS (TABS) ---
        tab_det, tab_res = st.tabs(["📄 Vista Detallada", "📊 Vista Resumida"])

        with tab_det:
            st.subheader(f"Documentos en: {hoja_sel}")
            opciones = ["", "R", "RA", "I", "RF", "A", "AF"]
            cols_edit = [c for c in df_base.columns if c not in columnas_necesarias]
            cfg = {c: st.column_config.SelectboxColumn(options=opciones) for c in cols_edit}
            df_final = st.data_editor(df_base, column_config=cfg, use_container_width=True)

        with tab_res:
            st.subheader("Resumen por Categoría")
            df_res = df_final.groupby([col_disc, col_tipo]).first().reset_index()
            df_res = df_res.drop(columns=[col_id, col_tit])
            st.dataframe(df_res, use_container_width=True)

        # --- EXPORTACIÓN ---
        st.divider()
        if st.button("📥 Descargar Matriz Oficial"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, sheet_name='Matriz', index=False, startrow=10)
                wb, ws = writer.book, writer.sheets['Matriz']
                fmt = wb.add_format({'bold': True, 'bg_color': '#004a99', 'font_color': 'white', 'border': 1, 'align': 'center'})
                ws.merge_range('A1:Z1', f'AA-VPP-BSCD-MTZ-0001 - {hoja_sel}', fmt)
            st.download_button("Confirmar Descarga", output.getvalue(), f"Matriz_{hoja_sel}.xlsx")
    else:
        st.warning(f"La hoja '{hoja_sel}' no parece ser un listado de entregables (faltan columnas: {', '.join(faltantes)})")
