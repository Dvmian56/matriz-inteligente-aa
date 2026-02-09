import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Matriz Anglo Real", layout="wide")

# --- ESTILOS CSS (PARA QUE SE VEA COMO EXCEL) ---
st.markdown("""
<style>
    .tabla-matriz { width: 100%; border-collapse: collapse; font-family: Calibri, sans-serif; font-size: 11px; }
    .header-azul { background-color: #004a99; color: white; padding: 10px; text-align: center; font-weight: bold; border: 1px solid black; }
    .fila-disciplina { background-color: #d9d9d9; font-weight: bold; text-align: left; padding: 5px; border: 1px solid #666; }
    .th-col { background-color: #004a99; color: white; border: 1px solid black; padding: 4px; text-align: center; font-size: 10px; writing-mode: vertical-rl; transform: rotate(180deg); height: 100px; }
    .td-celda { border: 1px solid #ccc; padding: 4px; color: black; }
    .td-dato { border: 1px solid #ccc; padding: 4px; text-align: center; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- SIDEBAR: DATOS Y EQUIPO ---
with st.sidebar:
    st.header("1. Configuración")
    ac_nom = st.text_input("Administrador (RF)", "Victor Otárola")
    
    st.divider()
    st.header("2. Especialistas")
    if 'equipo' not in st.session_state: st.session_state.equipo = []
    
    nombre = st.text_input("Nombre Revisor")
    sigla = st.text_input("Sigla Disciplina (Ej: EE, ME)")
    if st.button("Agregar a la Matriz"):
        st.session_state.equipo.append({"n": nombre, "d": sigla.upper()})
        st.rerun()
        
    # Mostrar equipo actual
    if st.session_state.equipo:
        st.write("---")
        for e in st.session_state.equipo:
            st.caption(f"👤 {e['n']} (Revisa: {e['d']})")
            
    if st.button("Borrar Todo"):
        st.session_state.equipo = []
        st.rerun()

# --- APP PRINCIPAL ---
st.title("🛡️ Visualizador de Matriz AA-VPP-BSCD-MTZ-0001")

archivo = st.file_uploader("Sube el MDL (Excel)", type=['xlsx'])

if archivo:
    xl = pd.ExcelFile(archivo)
    hoja = st.selectbox("Selecciona Pestaña", xl.sheet_names)
    df = xl.parse(hoja)
    
    # Limpieza de columnas duplicadas
    df = df.loc[:, ~df.columns.duplicated()]

    # Columnas Clave
    c_id = "No. de documento"
    c_disc = "Disciplina"
    c_tipo = "Tipo de Documento"
    c_tit = "Título"

    if all(c in df.columns for c in [c_id, c_disc, c_tipo, c_tit]):
        
        # 1. Preparar Datos
        df_base = df[[c_id, c_disc, c_tipo, c_tit]].drop_duplicates().sort_values(by=[c_disc, c_tipo])
        
        # 2. Asignar Roles (Lógica interna)
        # AC siempre es RF
        df_base[ac_nom] = "RF"
        # Especialistas
        for integrante in st.session_state.equipo:
            if integrante['n'] not in df_base.columns: df_base[integrante['n']] = ""
            mask = df_base[c_disc].astype(str).str.contains(integrante['d'], na=False, case=False)
            df_base.loc[mask, integrante['n']] = "R"

        # --- 3. CONSTRUCCIÓN VISUAL (HTML) ---
        # Aquí es donde ocurre la magia para que se vea agrupado
        
        html = f"""
        <table class="tabla-matriz">
            <tr>
                <td colspan="{4 + len(st.session_state.equipo) + 1}" class="header-azul">
                    MATRIZ DE REVISIÓN DE ENTREGABLES - {hoja.upper()}
                </td>
            </tr>
            <tr style="background-color: #f2f2f2;">
                <td class="td-celda" style="font-weight:bold; width:150px;">N° Documento</td>
                <td class="td-celda" style="font-weight:bold;">Título</td>
                <td class="td-celda" style="font-weight:bold;">Tipo</td>
                <td class="td-celda" style="font-weight:bold;">Disciplina</td>
                <td class="th-col">{ac_nom}</td>
        """
        
        # Columnas dinámicas de especialistas
        revisores_cols = [e['n'] for e in st.session_state.equipo]
        for rev in revisores_cols:
            html += f'<td class="th-col">{rev}</td>'
        
        html += "</tr>"

        # --- ITERACIÓN POR GRUPOS (LO QUE PEDÍAS) ---
        disciplinas_unicas = df_base[c_disc].unique()
        
        for disciplina in disciplinas_unicas:
            # 1. FILA SEPARADORA DE DISCIPLINA (LA BARRA GRIS)
            html += f"""
            <tr>
                <td colspan="{4 + len(st.session_state.equipo) + 1}" class="fila-disciplina">
                    ▼ {disciplina}
                </td>
            </tr>
            """
            
            # 2. DOCUMENTOS DE ESA DISCIPLINA
            docs_disciplina = df_base[df_base[c_disc] == disciplina]
            
            for index, row in docs_disciplina.iterrows():
                html += f"""
                <tr>
                    <td class="td-celda">{row[c_id]}</td>
                    <td class="td-celda" style="font-size:10px;">{row[c_tit]}</td>
                    <td class="td-celda">{row[c_tipo]}</td>
                    <td class="td-celda" style="text-align:center;">{row[c_disc]}</td>
                    
                    <td class="td-dato" style="background-color: #e6f7ff;">RF</td>
                """
                
                # Roles de Especialistas
                for rev in revisores_cols:
                    rol = row.get(rev, "")
                    color_bg = "#fff"
                    if rol == "R": color_bg = "#ffffcc" # Amarillo suave si revisa
                    html += f'<td class="td-dato" style="background-color:{color_bg};">{rol}</td>'
                
                html += "</tr>"

        html += "</table>"
        
        # RENDERIZAR LA TABLA
        st.markdown(html, unsafe_allow_html=True)
        
        # --- BOTÓN DE DESCARGA ---
        st.divider()
        if st.button("📥 Descargar este Excel"):
            # Lógica de xlsxwriter (simplificada para no alargar el código visual)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_base.to_excel(writer, index=False)
            st.download_button("Bajar Archivo", output.getvalue(), f"Matriz_{hoja}.xlsx")

    else:
        st.error(f"El archivo no tiene las columnas requeridas: {c_id}, {c_disc}, {c_tipo}, {c_tit}")
