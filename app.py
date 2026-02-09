import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Matriz Anglo Real", layout="wide")

# --- ESTILOS CSS (EL MAQUILLAJE PARA QUE SE VEA COMO EXCEL) ---
st.markdown("""
<style>
    .tabla-matriz { width: 100%; border-collapse: collapse; font-family: 'Calibri', sans-serif; font-size: 11px; }
    
    /* Cabecera Azul Principal */
    .header-azul { background-color: #004a99; color: white; padding: 15px; text-align: center; font-weight: bold; font-size: 16px; border: 1px solid #000; }
    
    /* Fila Separadora de Disciplina (LA BARRA GRIS) */
    .fila-disciplina { background-color: #bfbfbf; font-weight: bold; text-align: left; padding: 6px; border: 1px solid #666; color: black; }
    
    /* Encabezados de Columna (Verticales para los revisores) */
    .th-col { background-color: #004a99; color: white; border: 1px solid black; padding: 5px; text-align: center; vertical-align: bottom; }
    .th-vertical { writing-mode: vertical-rl; transform: rotate(180deg); height: 120px; padding-bottom: 5px; }
    
    /* Celdas de Datos */
    .td-celda { border: 1px solid #ccc; padding: 4px; color: black; vertical-align: middle; }
    .td-dato { border: 1px solid #ccc; padding: 4px; text-align: center; font-weight: bold; vertical-align: middle; }
    
    /* Colores de Responsabilidad */
    .rol-rf { background-color: #e6f7ff; color: #004a99; } /* Azulito para RF */
    .rol-r { background-color: #ffffcc; } /* Amarillito para R */
</style>
""", unsafe_allow_html=True)

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("1. Configuración")
    ac_nom = st.text_input("Administrador (RF)", "Victor Otárola")
    
    st.divider()
    st.header("2. Especialistas")
    if 'equipo' not in st.session_state: st.session_state.equipo = []
    
    c1, c2 = st.columns(2)
    nombre = c1.text_input("Nombre")
    sigla = c2.text_input("Sigla (Ej: EE)")
    
    if st.button("➕ Agregar"):
        if nombre and sigla:
            st.session_state.equipo.append({"n": nombre, "d": sigla.upper()})
            st.rerun()
            
    # Lista de equipo
    if st.session_state.equipo:
        st.write("---")
        st.markdown("**Equipo Actual:**")
        for e in st.session_state.equipo:
            st.caption(f"👤 {e['n']} ➡️ {e['d']}")
    
    if st.button("🗑️ Borrar Todo"):
        st.session_state.equipo = []
        st.rerun()

# --- APP PRINCIPAL ---
st.title("🛡️ Visualizador de Matriz AA-VPP-BSCD-MTZ-0001")

archivo = st.file_uploader("Sube el MDL (Excel)", type=['xlsx'])

if archivo:
    xl = pd.ExcelFile(archivo)
    hoja = st.selectbox("Selecciona Pestaña", xl.sheet_names)
    
    # Carga segura
    df = xl.parse(hoja)
    df = df.loc[:, ~df.columns.duplicated()]

    # Columnas Clave
    c_id, c_disc, c_tipo, c_tit = "No. de documento", "Disciplina", "Tipo de Documento", "Título"

    if all(c in df.columns for c in [c_id, c_disc, c_tipo, c_tit]):
        
        # 1. Preparar Datos
        df_base = df[[c_id, c_disc, c_tipo, c_tit]].drop_duplicates().sort_values(by=[c_disc, c_tipo])
        
        # 2. Lógica de Asignación (En memoria)
        # AC siempre es RF
        df_base[ac_nom] = "RF"
        # Especialistas
        cols_revisores = [ac_nom] + [e['n'] for e in st.session_state.equipo]
        
        for integrante in st.session_state.equipo:
            # Creamos columna si no existe
            if integrante['n'] not in df_base.columns: df_base[integrante['n']] = ""
            # Aplicamos regla: Si disciplina contiene sigla, es R
            mask = df_base[c_disc].astype(str).str.contains(integrante['d'], na=False, case=False)
            df_base.loc[mask, integrante['n']] = "R"

        # --- 3. DIBUJAR LA MATRIZ (HTML) ---
        # Aquí construimos la tabla fila por fila
        
        html = f"""
        <table class="tabla-matriz">
            <tr>
                <td colspan="{4 + len(cols_revisores)}" class="header-azul">
                    MATRIZ DE REVISIÓN DE ENTREGABLES - {hoja.upper()}
                </td>
            </tr>
            <tr style="background-color: #f2f2f2;">
                <td class="th-col">N° Documento</td>
                <td class="th-col">Título</td>
                <td class="th-col">Tipo</td>
                <td class="th-col">Disciplina</td>
        """
        
        # Encabezados de Revisores (VERTICALES)
        for rev in cols_revisores:
            html += f'<td class="th-col"><div class="th-vertical">{rev}</div></td>'
        
        html += "</tr>"

        # ITERACIÓN INTELIGENTE (AGRUPAR POR DISCIPLINA)
        disciplinas_unicas = df_base[c_disc].unique()
        
        for disciplina in disciplinas_unicas:
            # A) LA BARRA GRIS (Separador)
            html += f"""
            <tr>
                <td colspan="{4 + len(cols_revisores)}" class="fila-disciplina">
                    ▼ {disciplina}
                </td>
            </tr>
            """
            
            # B) LOS DOCUMENTOS DE ESA DISCIPLINA
            docs_disciplina = df_base[df_base[c_disc] == disciplina]
            
            for index, row in docs_disciplina.iterrows():
                html += f"""
                <tr>
                    <td class="td-celda" style="white-space:nowrap;">{row[c_id]}</td>
                    <td class="td-celda" style="font-size:10px;">{row[c_tit]}</td>
                    <td class="td-celda">{row[c_tipo]}</td>
                    <td class="td-celda" style="text-align:center;">{row[c_disc]}</td>
                """
                
                # Celdas de Revisores
                for rev in cols_revisores:
                    rol = row.get(rev, "")
                    # Estilo según el rol
                    clase_rol = ""
                    if rol == "RF": clase_rol = "rol-rf"
                    elif rol == "R": clase_rol = "rol-r"
                    
                    html += f'<td class="td-dato {clase_rol}">{rol}</td>'
                
                html += "</tr>"

        html += "</table>"
        
        # PINTAR EN PANTALLA
        st.markdown(html, unsafe_allow_html=True)
        
        # --- DESCARGA ---
        st.divider()
        st.caption("Nota: Esta vista es una representación visual. El archivo descargable será un Excel editable.")
        if st.button("📥 Descargar Excel Oficial"):
            # Aquí va tu lógica de XlsxWriter (la mantenemos simple para que funcione el botón)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_base.to_excel(writer, index=False)
            st.download_button("Confirmar Descarga", output.getvalue(), f"Matriz_{hoja}.xlsx")

    else:
        st.error(f"Faltan columnas clave. Asegúrate de tener: {c_id}, {c_disc}, {c_tipo}, {c_tit}")
