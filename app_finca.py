import streamlit as st
import pandas as pd
import mysql.connector
import io

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
st.sidebar.header("➕ Registrar Nueva Parcela")

with st.sidebar.form("formulario_parcela", clear_on_submit=True):
    id_parcela = st.text_input("ID de la Parcela (Ej: PAR-006)").strip()
    sector = st.selectbox("Sector", ["Norte", "Sur", "Este", "Oeste", "Central"])
    hectareas = st.number_input("Hectáreas", min_value=1, max_value=300, value=10)
    cultivo = st.text_input("Tipo de Cultivo (Ej: Arroz, Caña)").strip()
    humedad = st.slider("Humedad del Suelo (%)", 0, 100, 50)
    estado = st.selectbox("Estado", ["Óptimo", "Requiere Riego", "Inactivo"])
    
    boton_guardar = st.form_submit_button("Guardar en Base de Datos")

if boton_guardar:
    if id_parcela and cultivo:
        try:
            conn = obtener_conexion()
            cursor = conn.cursor()
            cursor.execute("INSERT INTO registro_parcelas (id_parcela, sector, hectareas, cultivo, humedad_suelo_pct, estado) VALUES (%s, %s, %s, %s, %s, %s)", 
                           (id_parcela, sector, hectareas, cultivo, humedad, estado))
            conn.commit()
            cursor.close()
            conn.close()
            st.sidebar.success(f"¡{id_parcela} guardada correctamente! 🎉")
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
            
            # 1. Creamos la barra de búsqueda
            busqueda = st.text_input("🔎 Escribe el Sector o ID de la parcela para filtrar (Ej: 'Norte' o 'PAR-002'):")
            
            # 2. Lógica del multiplexor: filtramos los datos según lo que escriba el usuario
            if busqueda:
                # Filtramos ignorando mayúsculas y minúsculas
                df_filtrado = df[df['id_parcela'].str.contains(busqueda, case=False, na=False) | 
                                 df['sector'].str.contains(busqueda, case=False, na=False)]
            else:
                df_filtrado = df # Si no hay búsqueda, mostramos todo
            
            # 3. Dibujamos la tabla con los datos filtrados
            st.dataframe(df_filtrado, use_container_width=True)
            
            # 4. El botón de descarga ahora descarga solo lo que está filtrado
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_filtrado.to_excel(writer, index=False, sheet_name='Reporte_Filtro')
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

   
   st.subheader("📁 Cargar datos reales desde Excel")
archivo_subido = st.file_uploader("Sube el archivo Excel de la finca aquí", type=["xlsx", "xls"])

if archivo_subido is not None:
    # 1. Leer el archivo Excel
    df_cargado = pd.read_excel(archivo_subido)
    
    st.write("Vista previa de los datos a importar:")
    st.dataframe(df_cargado)
    
    # 2. Botón de seguridad para confirmar la inyección
    if st.button("Guardar estos datos en la Nube"):
        try:
            conexion = obtener_conexion()
            cursor = conexion.cursor()
            registros_guardados = 0
            
            # 3. Recorrer cada fila del Excel e inyectarla a la base de datos
            for index, fila in df_cargado.iterrows():
                sql = """INSERT IGNORE INTO registro_parcelas 
                         (id_parcela, sector, hectareas, cultivo, humedad_suelo_pct, estado) 
                         VALUES (%s, %s, %s, %s, %s, %s)"""
                
                # Extraemos los datos de la fila actual
                valores = (
                    str(fila['id_parcela']), 
                    str(fila['sector']), 
                    int(fila['hectareas']), 
                    str(fila['cultivo']), 
                    int(fila['humedad_suelo_pct']), 
                    str(fila['estado'])
                )
                
                cursor.execute(sql, valores)
                registros_guardados += cursor.rowcount # Cuenta cuántos se guardaron con éxito
                
            # 4. Sellar la bóveda
            conexion.commit()
            cursor.close()
            conexion.close()
            
            st.success(f"¡Operación exitosa! Se guardaron {registros_guardados} parcelas en el servidor global.")
            st.balloons() # Un efecto visual para celebrar que los datos subieron a la nube
            
        except Exception as e:
            st.error(f"Hubo un error de conexión: {e}")