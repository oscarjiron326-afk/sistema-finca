import streamlit as st
import pandas as pd
import mysql.connector
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 1. Configuración de la página
st.set_page_config(page_title="Sistema Agrícola", layout="wide")
st.title("🌱 Panel de Control Central - Finca (300 Hectáreas)")
st.markdown("---")

def obtener_conexion():
    return mysql.connector.connect(
        host="bpvhrmazb58ojyt1ynth-mysql.services.clever-cloud.com",
        user="uunwlo7t4x1ihcw9",
        password="3OEBQ5ERzoCLY6u9SvXG", 
        database="bpvhrmazb58ojyt1ynth"
    )

# --- EL CEREBRO AUTOMÁTICO ---
def calcular_estado_automatico(cultivo, humedad):
    cultivo_limpio = str(cultivo).lower()
    if "arroz" in cultivo_limpio:
        return "Óptimo" if humedad >= 75 else "Requiere Riego (Crítico)"
    elif "frijol" in cultivo_limpio:
        return "Óptimo" if humedad >= 40 else "Requiere Riego"
    elif "maíz" in cultivo_limpio or "maiz" in cultivo_limpio:
        return "Óptimo" if humedad >= 50 else "Requiere Riego"
    else:
        return "Óptimo" if humedad >= 45 else "Requiere Riego"

# --- PANEL LATERAL (Registro Manual Automatizado) ---
st.sidebar.header("➕ Registrar Nueva Parcela")

with st.sidebar.form("formulario_parcela", clear_on_submit=True):
    id_parcela = st.text_input("ID de la Parcela (Ej: PAR-006)").strip()
    sector = st.selectbox("Sector", ["Norte", "Sur", "Este", "Oeste", "Central"])
    hectareas = st.number_input("Hectáreas", min_value=1, max_value=300, value=10)
    cultivo = st.text_input("Tipo de Cultivo (Ej: Arroz, Caña)").strip()
    humedad = st.slider("Humedad del Suelo (%)", 0, 100, 50)
    # ELIMINAMOS la pregunta del estado. ¡El sistema lo decide ahora!
    
    boton_guardar = st.form_submit_button("Guardar en Base de Datos")

if boton_guardar:
    if id_parcela and cultivo:
        try:
            # Calculamos el estado matemáticamente antes de guardar
            estado_calculado = calcular_estado_automatico(cultivo, humedad)
            
            conn = obtener_conexion()
            cursor = conn.cursor()
            cursor.execute("INSERT INTO registro_parcelas (id_parcela, sector, hectareas, cultivo, humedad_suelo_pct, estado) VALUES (%s, %s, %s, %s, %s, %s)", 
                           (id_parcela, sector, hectareas, cultivo, humedad, estado_calculado))
            conn.commit()
            cursor.close()
            conn.close()
            st.sidebar.success(f"¡{id_parcela} guardada como: {estado_calculado}! 🎉")
        except Exception as e:
            st.sidebar.error(f"Error al guardar: {e}")
    else:
        st.sidebar.warning("⚠️ Rellena todos los campos.")

# --- LECTURA DE DATOS Y DISEÑO CENTRAL POR PESTAÑAS ---
try:
    conn = obtener_conexion()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT id_parcela, sector, hectareas, cultivo, humedad_suelo_pct, estado FROM registro_parcelas")
    resultados = cursor.fetchall()
    cursor.close()
    conn.close()
    
    if resultados:
        df = pd.DataFrame(resultados)
        
        # CREAMOS LAS PESTAÑAS (TABS)
        pestana1, pestana2 = st.tabs(["📈 Dashboard Gráfico", "🗄️ Explorador de Datos y Reportes"])
        
        # --- CONTENIDO DE LA PESTAÑA 1 (Gráficos) ---
        with pestana1:
            st.subheader("⏱️ Resumen del Sistema en Tiempo Real")
            total_parcelas = len(df)
            total_hectareas = int(df['hectareas'].sum())
            promedio_humedad = round(df['humedad_suelo_pct'].mean(), 1)
            
            metrica1, metrica2, metrica3 = st.columns(3)
            metrica1.metric(label="📦 Total Parcelas Registradas", value=f"{total_parcelas}")
            metrica2.metric(label="🗺️ Hectáreas Administradas", value=f"{total_hectareas} ha")
            metrica3.metric(label="💧 Promedio de Humedad", value=f"{promedio_humedad} %")
            
            st.markdown("---")
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("💧 Humedad por Parcela (%)")
                st.bar_chart(data=df, x='id_parcela', y='humedad_suelo_pct', color="#1f77b4")
            with col2:
                st.subheader("📏 Distribución de Hectáreas")
                st.bar_chart(data=df, x='id_parcela', y='hectareas', color="#2ca02c")
                
        # --- CONTENIDO DE LA PESTAÑA 2 (Tabla y Descarga) ---
        with pestana2:
            st.subheader("🔍 Explorador y Filtro de Datos")
            
            busqueda = st.text_input("🔎 Escribe el Sector o ID de la parcela para filtrar (Ej: 'Norte' o 'PAR-002'):")
            
            if busqueda:
                df_filtrado = df[df['id_parcela'].str.contains(busqueda, case=False, na=False) | 
                                 df['sector'].str.contains(busqueda, case=False, na=False)]
            else:
                df_filtrado = df 
            
            st.dataframe(df_filtrado, use_container_width=True)
            
            # Exportar a Excel con formato Profesional
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_filtrado.to_excel(writer, index=False, sheet_name='Reporte_Finca')
                
                workbook = writer.book
                worksheet = writer.sheets['Reporte_Finca']
                
                color_enc = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") 
                fuente_enc = Font(color="FFFFFF", bold=True) 
                alineacion_centro = Alignment(horizontal="center", vertical="center")
                borde_fino = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                     top=Side(style='thin'), bottom=Side(style='thin'))

                for col_num, nombre_columna in enumerate(df_filtrado.columns):
                    celda = worksheet.cell(row=1, column=col_num + 1)
                    celda.fill = color_enc
                    celda.font = fuente_enc
                    celda.alignment = alineacion_centro
                    celda.border = borde_fino
                    
                    ancho_col = max(len(str(nombre_columna)), df_filtrado[nombre_columna].astype(str).map(len).max()) + 2
                    letra_col = get_column_letter(col_num + 1)
                    worksheet.column_dimensions[letra_col].width = ancho_col

                for fila in worksheet.iter_rows(min_row=2, max_row=len(df_filtrado) + 1, min_col=1, max_col=len(df_filtrado.columns)):
                    for celda in fila:
                        celda.alignment = alineacion_centro
                        celda.border = borde_fino
                        
            excel_data = buffer.getvalue()
            
            st.download_button(
                label="📥 Descargar Reporte de esta Vista (.xlsx)",
                data=excel_data,
                file_name="reporte_filtrado_finca.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.info("La base de datos está vacía. Utiliza el panel izquierdo para registrar datos.")

except Exception as error:
    st.error(f"Error de conexión: {error} 🔴")

# --- MÓDULO DE CARGA MASIVA EXCEL (AUTOMATIZADO) ---
st.markdown("---")
st.subheader("📁 Cargar datos reales desde Excel")
archivo_subido = st.file_uploader("Sube el archivo Excel de la finca aquí", type=["xlsx", "xls"])

if archivo_subido is not None:
    df_cargado = pd.read_excel(archivo_subido)
    df_cargado.columns = ['id_parcela', 'sector', 'hectareas', 'cultivo', 'humedad_suelo_pct', 'estado']
    
    st.write("Vista previa de los datos a importar:")
    st.dataframe(df_cargado)
    
    if st.button("Guardar y Automatizar Estados en la Nube"):
        try:
            conexion = obtener_conexion()
            cursor = conexion.cursor()
            registros_guardados = 0
            
            for index, fila in df_cargado.iterrows():
                sql = """INSERT IGNORE INTO registro_parcelas 
                         (id_parcela, sector, hectareas, cultivo, humedad_suelo_pct, estado) 
                         VALUES (%s, %s, %s, %s, %s, %s)"""
                
                # ¡LA MAGIA OCURRE AQUÍ! Recalculamos el estado usando el cerebro automático
                estado_inteligente = calcular_estado_automatico(fila['cultivo'], fila['humedad_suelo_pct'])
                
                valores = (
                    str(fila['id_parcela']), 
                    str(fila['sector']), 
                    int(fila['hectareas']), 
                    str(fila['cultivo']), 
                    int(fila['humedad_suelo_pct']), 
                    estado_inteligente # Mandamos a la bóveda el estado calculado, no el del Excel
                )
                
                cursor.execute(sql, valores)
                registros_guardados += cursor.rowcount 
                
            conexion.commit()
            cursor.close()
            conexion.close()
            
            st.success(f"¡Operación exitosa! Se automatizaron y guardaron {registros_guardados} parcelas.")
            st.balloons() 
            
        except Exception as e:
            st.error(f"Hubo un error de conexión al inyectar: {e}")