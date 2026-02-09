import streamlit as st
import pandas as pd
import io

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Matriz Pro - Anglo American", layout="wide")

# --- SIDEBAR: EQUIPO Y REGLAS ---
with st.sidebar:
    st.header("📋 Datos del Contrato")
    proy = st.text_input("Proyecto", "1013084 MANEJO DE EFLUENTES")
    cont = st.text_input("Contrato", "CP684-P2100361-3ODS23")
    
    st.divider()
    st.header("👥 Configuración de Equipo")
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
st.title("🛡️ Matriz de Revisión Inteligente")

archivo = st.file_uploader("Sube el MDL (Excel)", type=['xlsx'])

if archivo:
    xl = pd.ExcelFile(archivo)
    hoja_sel = st.selectbox("Selecciona la pestaña (Ej: Ingeniería)", xl.sheet_names)
    df_raw = xl.parse(hoja_sel)

    # Columnas específicas del archivo Aconex
    col_id = "No. de documento"
    col_disc = "Disciplina"
    col_tipo = "Tipo de Documento"
    col_tit = "Título"

    if all(c in df_raw.columns for c in [col_id, col_disc, col_tipo, col_tit]):
        
        # 1. Preparación y Ordenamiento
        df_base = df_raw[[col_id, col_disc, col_tipo, col_tit]].drop_duplicates()
        df_base = df_base.sort_values(by=[col_disc, col_tipo])

        # 2. Aplicar Reglas Automáticas
        # Administrador es RF en todo el listado
        df_base[ac_nom] = "RF"
        
        # Especialistas por disciplina
        for esp in st.session_state.mapeo:
            if esp["n"] not in df_base.columns: df_base[esp["n"]] = ""
            mask = df_base[col_disc].astype(str).str.contains(esp["d"], na=False, case=False)
            df_base.loc[mask, esp["n"]] = "R"

        # --- 3. LAS DOS VISTAS (TABS) ---
        tab_det, tab_res = st.tabs(["📄 Vista Detallada", "📊 Vista Resumida"])

        with tab_det:
            st.subheader("Control Operativo por Documento")
            opciones = ["", "R", "RA", "I", "RF", "A", "AF"]
            config_cols = {col: st.column_config.SelectboxColumn(options=opciones) 
                          for col in df_base.columns if col not in [col_id, col_disc, col_tipo, col_tit]}
            df_det_final = st.data_editor(df_base, column_config=config_cols, use_container_width=True)

        with tab_res:
            st.subheader("Resumen de Responsabilidades (Agrupado)")
            # Agrupación por Disciplina y Tipo
            df_res_final = df_det_final.groupby([col_disc, col_tipo]).first().reset_index()
            df_res_final = df_res_final.drop(columns=[col_id, col_tit])
            st.dataframe(df_res_final, use_container_width=True)

        # --- 4. EXPORTACIÓN A EXCEL OFICIAL ---
        st.divider()
        if st.button("📥 Descargar Matriz Oficial (Excel)"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_det_final.to_excel(writer, sheet_name='Matriz', index=False, startrow=10)
                wb = writer.book
                ws = writer.sheets['Matriz']
                
                # Formatos Corporativos
                fmt_azul = wb.add_format({'bold': True, 'bg_color': '#004a99', 'font_color': 'white', 'border': 1, 'align': 'center'})
                fmt_gris = wb.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center'})
                
                # Cabecera Anglo American
                ws.merge_range('A1:Z1', 'AA-VPP-BSCD-MTZ-0001 Matriz de Revisión de Entregables', fmt_azul)
                ws.write('B3', f'Proyecto: {proy}')
                ws.write('B4', f'Contrato: {cont}')
                
                # Pintar cargos de los especialistas
                col_idx = 4
                for r in st.session_state.mapeo:
                    ws.write(9, col_idx, "ESPECIALISTA", fmt_gris)
                    col_idx += 1

            st.download_button("Confirmar Descarga", output.getvalue(), f"Matriz_{cont}.xlsx")
    else:
        st.error(f"Faltan columnas obligatorias en la hoja '{hoja_sel}'.")
