import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import openpyxl

# Configuración de la página
st.set_page_config(
    page_title="Dashboard - Área Experimental",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Título principal
st.title("🔬 Dashboard Área Experimental - GCF")
st.markdown("---")

# Rutas de los archivos
RUTA_PRODUCCION = "C:\\Users\\caalcivar\\OneDrive - DISTRISODA S.A\\Escritorio\\Area Experimental\\Base de Datos de Producción.xlsx"
RUTA_LABORATORIO = "C:\\Users\\caalcivar\\OneDrive - DISTRISODA S.A\\Escritorio\\Area Experimental\\Resultados EXPERIMENTAL.xlsx"

# Lista de piscinas
PISCINAS = ['LI001', 'LI002', 'LI003', 'LI004', 'LI005', 'LI006']
COLORES_PISCINAS = {
    'LI001': '#1f77b4',
    'LI002': '#ff7f0e', 
    'LI003': '#2ca02c',
    'LI004': '#d62728',
    'LI005': '#9467bd',
    'LI006': '#8c564b'
}

# Función para cargar datos
@st.cache_data(ttl=300)
def cargar_datos_produccion():
    """Carga los datos del archivo de producción"""
    try:
        excel_file = pd.ExcelFile(RUTA_PRODUCCION)
        df = pd.read_excel(RUTA_PRODUCCION, sheet_name=excel_file.sheet_names[0], header=0)
        return df
    except Exception as e:
        st.error(f"Error al cargar datos de producción: {str(e)}")
        return None

@st.cache_data(ttl=300)
def cargar_datos_laboratorio():
    """Carga los datos del archivo de laboratorio"""
    try:
        dfs = {}
        hojas = ['Base_Quimicos', 'Base_Pato', 'Base_ORP', 'Base_Micro', 'Base_Fito']
        
        for hoja in hojas:
            try:
                df_temp = pd.read_excel(RUTA_LABORATORIO, sheet_name=hoja)
                dfs[hoja] = df_temp
            except:
                dfs[hoja] = pd.DataFrame()
                st.warning(f"No se pudo cargar la hoja {hoja}")
        
        return dfs
    except Exception as e:
        st.error(f"Error al cargar datos de laboratorio: {str(e)}")
        return {}

# Función para encontrar columna de fecha
def encontrar_columna_fecha(df):
    """Busca la columna de fecha en un dataframe, priorizando F.muestreo"""
    # Primero buscar F.muestreo
    for col in df.columns:
        if col == 'F.muestreo':
            return col
    
    # Luego buscar otras columnas de fecha
    for col in df.columns:
        col_lower = col.lower()
        if 'muestreo' in col_lower:
            return col
    
    # Finalmente buscar cualquier columna de fecha
    for col in df.columns:
        col_lower = col.lower()
        if 'fecha' in col_lower or 'date' in col_lower:
            return col
    return None

# Cargar datos
df_produccion = cargar_datos_produccion()
dfs_laboratorio = cargar_datos_laboratorio()

# Sidebar para filtros
with st.sidebar:
    # Imagen centrada arriba de los filtros
    try:
        col_img1, col_img2, col_img3 = st.columns([1, 100, 1])
        with col_img2:
            st.image("C:\\Users\\caalcivar\\OneDrive - DISTRISODA S.A\\Escritorio\\Area Experimental\\Imagenes\\GCF.png", width=300)
    except:
        st.warning("No se pudo cargar el logo")
    
    st.markdown("---")
    
    st.header("🎛️ Filtros")
    
    # Selección de piscinas
    st.subheader("Seleccionar Piscinas")
    piscinas_seleccionadas = []
    col1, col2 = st.columns(2)
    
    for i, piscina in enumerate(PISCINAS):
        if i % 2 == 0:
            with col1:
                if st.checkbox(piscina, value=True, key=f"piscina_{piscina}"):
                    piscinas_seleccionadas.append(piscina)
        else:
            with col2:
                if st.checkbox(piscina, value=True, key=f"piscina_{piscina}"):
                    piscinas_seleccionadas.append(piscina)
    
    st.markdown("---")
    
    # Información del dashboard
    st.info("""
    **Dashboard Experimental**
    
    Este dashboard muestra datos de:
    - Producción
    - Parámetros químicos
    - Patógenos
    - ORP
    - Microbiología
    - Fitoplancton
    - VS (Comparador)
    
    Para las piscinas: LI001 - LI006
    """)

# Filtrar datos de producción
if df_produccion is not None:
    df_produccion_filtrado = df_produccion.copy()
else:
    df_produccion_filtrado = None

# Crear tabs para organizar el dashboard
tab_produccion, tab_quimicos, tab_patogenos, tab_orp, tab_micro, tab_fito, tab_vs = st.tabs([
    "📊 Producción",
    "🧪 Químicos",
    "🦠 Patógenos",
    "⚡ ORP",
    "🔬 Microbiología",
    "🌿 Fitoplancton",
    "⚖️ VS"
])

# Tab de Producción
with tab_produccion:
    st.header("📊 Datos de Producción")
    
    if df_produccion_filtrado is not None and not df_produccion_filtrado.empty:
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Total de registros", len(df_produccion_filtrado))
        
        with col2:
            if 'Libras' in df_produccion_filtrado.columns and 'Piscina' in df_produccion_filtrado.columns:
                df_piscinas_seleccionadas = df_produccion_filtrado[df_produccion_filtrado['Piscina'].isin(piscinas_seleccionadas)]
                
                if not df_piscinas_seleccionadas.empty:
                    ultimos_registros = df_piscinas_seleccionadas.groupby('Piscina')['Libras'].last()
                    prod_total = ultimos_registros.sum()
                    st.metric("Producción total (Lbs)", f"{prod_total:,.2f}")
        
        with col3:
            st.metric("Piscinas activas", len(piscinas_seleccionadas))
        
        st.markdown("---")
        
        if 'Piscina' in df_produccion_filtrado.columns:
            df_piscina = df_produccion_filtrado[df_produccion_filtrado['Piscina'].isin(piscinas_seleccionadas)].copy()
            
            metricas = [
                'Peso Actual', 'Incremento semanal', 'Incremento Semanal', 'Delta 1 Sem.', '%Sobrev. Tecnico', '%Sobrev Tabla',
                '%Sobrev Total', 'Densidad', 'Libras', 'Libras/Ha', 'Libras/Ha/Dia',
                'Tipos alimento', 'Proveedor Balanc.', 'Tipo.alim ult. Sem',
                'Tipo de alimentacion', 'FCA', '$.Costo', '$.Costo/Ha', '$.Costo/Lb', '$.Costo/Un'
            ]
            
            metricas_disponibles = [m for m in metricas if m in df_piscina.columns and pd.api.types.is_numeric_dtype(df_piscina[m])]
            
            if metricas_disponibles:
                metrica_seleccionada = st.selectbox(
                    "Seleccionar métrica a visualizar:",
                    metricas_disponibles
                )
                
                if 'F.muestreo' in df_piscina.columns:
                    df_piscina['F.muestreo'] = pd.to_datetime(df_piscina['F.muestreo'], errors='coerce')
                    df_piscina = df_piscina.sort_values('F.muestreo')
                
                fig = go.Figure()
                
                for piscina in piscinas_seleccionadas:
                    df_temp = df_piscina[df_piscina['Piscina'] == piscina].copy()
                    
                    if not df_temp.empty:
                        if 'F.muestreo' in df_temp.columns:
                            df_temp = df_temp.dropna(subset=['F.muestreo', metrica_seleccionada])
                            df_temp = df_temp.sort_values('F.muestreo')
                            x_values = df_temp['F.muestreo']
                        else:
                            df_temp = df_temp.dropna(subset=[metrica_seleccionada])
                            x_values = df_temp.index
                        
                        if len(df_temp) > 0:
                            fig.add_trace(go.Scatter(
                                x=x_values,
                                y=df_temp[metrica_seleccionada],
                                mode='lines+markers',
                                name=piscina,
                                line=dict(color=COLORES_PISCINAS.get(piscina, '#000000'), width=2),
                                marker=dict(size=6),
                                hovertemplate=f'<b>{piscina}</b><br>Fecha: %{{x}}<br>{metrica_seleccionada}: %{{y:,.2f}}<extra></extra>'
                            ))
                
                fig.update_layout(
                    title=f"{metrica_seleccionada} por Piscina",
                    xaxis_title="Fecha de Muestreo" if 'F.muestreo' in df_piscina.columns else "Índice",
                    yaxis_title=metrica_seleccionada,
                    height=500,
                    hovermode='x unified',
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    )
                )
                
                st.plotly_chart(fig, use_container_width=True)
        
        st.subheader("Datos Detallados")
        st.dataframe(df_produccion_filtrado, use_container_width=True)
        
        csv = df_produccion_filtrado.to_csv(index=False)
        st.download_button(
            label="📥 Descargar datos de producción",
            data=csv,
            file_name=f"produccion_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.warning("No hay datos de producción disponibles o no se pudo cargar el archivo.")

# Tab de Químicos
with tab_quimicos:
    st.header("🧪 Parámetros Químicos")
    
    if 'Base_Quimicos' in dfs_laboratorio and not dfs_laboratorio['Base_Quimicos'].empty:
        df_quimicos = dfs_laboratorio['Base_Quimicos']
        df_quimicos_filtrado = df_quimicos.copy()
        
        if 'Piscina' in df_quimicos_filtrado.columns:
            df_quimicos_filtrado = df_quimicos_filtrado[df_quimicos_filtrado['Piscina'].isin(piscinas_seleccionadas)]
        
        cols_numericas = df_quimicos_filtrado.select_dtypes(include=[np.number]).columns.tolist()
        
        columnas_excluir = ['Año', 'HECTÁREAS', 'CICLO', 'Semana']
        cols_numericas = [col for col in cols_numericas if col not in columnas_excluir]
        
        if cols_numericas:
            fecha_col = encontrar_columna_fecha(df_quimicos_filtrado)
            
            parametros_seleccionados = st.multiselect(
                "Seleccionar parámetros a comparar:",
                cols_numericas,
                default=cols_numericas[:2] if len(cols_numericas) >= 2 else cols_numericas
            )
            
            if parametros_seleccionados and 'Piscina' in df_quimicos_filtrado.columns:
                fig = go.Figure()
                
                colores_ejes = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
                
                for i, parametro in enumerate(parametros_seleccionados):
                    for piscina in piscinas_seleccionadas:
                        df_temp = df_quimicos_filtrado[df_quimicos_filtrado['Piscina'] == piscina].copy()
                        
                        if not df_temp.empty:
                            if fecha_col and fecha_col in df_temp.columns:
                                df_temp[fecha_col] = pd.to_datetime(df_temp[fecha_col], errors='coerce')
                                df_temp = df_temp.dropna(subset=[fecha_col, parametro])
                                df_temp = df_temp.sort_values(fecha_col)
                                x_values = df_temp[fecha_col]
                            else:
                                df_temp = df_temp.dropna(subset=[parametro])
                                x_values = df_temp.index
                            
                            if len(df_temp) > 0:
                                y_values = df_temp[parametro].values
                                
                                nombre_traza = f"{parametro} - {piscina}"
                                color_parametro = colores_ejes[i % len(colores_ejes)]
                                
                                fig.add_trace(go.Scatter(
                                    x=x_values,
                                    y=y_values,
                                    mode='lines+markers',
                                    name=nombre_traza,
                                    line=dict(color=color_parametro, width=2, dash='solid' if i == 0 else 'dash'),
                                    marker=dict(size=6),
                                    yaxis='y' if i == 0 else f'y{i+1}',
                                    hovertemplate=f'<b>{piscina}</b><br>Fecha: %{{x}}<br>{parametro}: %{{y:,.2f}}<extra></extra>'
                                ))
                
                fig.update_layout(
                    title="Comparación de Parámetros Químicos por Piscina",
                    xaxis_title="Fecha" if fecha_col else "Índice",
                    height=600,
                    hovermode='x unified',
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    ),
                    xaxis=dict(domain=[0.15, 0.85] if len(parametros_seleccionados) > 1 else [0, 1])
                )
                
                for i, parametro in enumerate(parametros_seleccionados):
                    eje = 'yaxis' if i == 0 else f'yaxis{i+1}'
                    if eje == 'yaxis':
                        fig.update_layout(yaxis=dict(
                            title=dict(text=parametro, font=dict(color=colores_ejes[0])),
                            tickfont=dict(color=colores_ejes[0])
                        ))
                    else:
                        posicion = 1 + (i-1) * 0.15
                        fig.update_layout(**{eje: dict(
                            title=dict(text=parametro, font=dict(color=colores_ejes[i % len(colores_ejes)])),
                            tickfont=dict(color=colores_ejes[i % len(colores_ejes)]),
                            overlaying='y',
                            side='right',
                            position=posicion
                        )})
                
                st.plotly_chart(fig, use_container_width=True)
        
        st.subheader("Datos de Químicos")
        st.dataframe(df_quimicos_filtrado, use_container_width=True)
    else:
        st.warning("No hay datos de químicos disponibles.")

# Tab de Patógenos
with tab_patogenos:
    st.header("🦠 Análisis de Patógenos")
    
    if 'Base_Pato' in dfs_laboratorio and not dfs_laboratorio['Base_Pato'].empty:
        df_patogenos = dfs_laboratorio['Base_Pato']
        df_patogenos_filtrado = df_patogenos.copy()
        
        if 'Piscina' in df_patogenos_filtrado.columns:
            df_patogenos_filtrado = df_patogenos_filtrado[df_patogenos_filtrado['Piscina'].isin(piscinas_seleccionadas)]
        
        if not df_patogenos_filtrado.empty:
            cols_presencia = [col for col in df_patogenos_filtrado.columns if any(x in col.lower() for x in ['presencia', 'positivo', 'negativo', 'deteccion', 'resultado'])]
            
            if cols_presencia:
                for col in cols_presencia:
                    if df_patogenos_filtrado[col].dtype == 'object':
                        valores_unicos = df_patogenos_filtrado[col].value_counts()
                        
                        if len(valores_unicos) <= 10:
                            fig_pato = px.pie(
                                values=valores_unicos.values,
                                names=valores_unicos.index,
                                title=f"Distribución de {col}"
                            )
                            st.plotly_chart(fig_pato, use_container_width=True)
        
        st.subheader("Datos de Patógenos")
        st.dataframe(df_patogenos_filtrado, use_container_width=True)
    else:
        st.warning("No hay datos de patógenos disponibles.")

# Tab de ORP
with tab_orp:
    st.header("⚡ Potencial de Óxido-Reducción (ORP)")
    
    if 'Base_ORP' in dfs_laboratorio and not dfs_laboratorio['Base_ORP'].empty:
        df_orp = dfs_laboratorio['Base_ORP']
        df_orp_filtrado = df_orp.copy()
        
        if 'Piscina' in df_orp_filtrado.columns:
            df_orp_filtrado = df_orp_filtrado[df_orp_filtrado['Piscina'].isin(piscinas_seleccionadas)]
        
        if not df_orp_filtrado.empty:
            cols_orp = [col for col in df_orp_filtrado.columns if 'orp batea' in col.lower()]
            
            if not cols_orp:
                cols_orp = [col for col in df_orp_filtrado.columns if 'orp' in col.lower() or 'oxid' in col.lower() or 'redox' in col.lower()]
            
            if cols_orp and 'Piscina' in df_orp_filtrado.columns:
                fig_orp = go.Figure()
                
                for piscina in piscinas_seleccionadas:
                    df_temp = df_orp_filtrado[df_orp_filtrado['Piscina'] == piscina].copy()
                    
                    if not df_temp.empty:
                        df_temp = df_temp.dropna(subset=[cols_orp[0]])
                        
                        if len(df_temp) > 0:
                            fig_orp.add_trace(go.Scatter(
                                x=df_temp.index,
                                y=df_temp[cols_orp[0]],
                                mode='lines+markers',
                                name=piscina,
                                line=dict(color=COLORES_PISCINAS.get(piscina, '#000000'), width=2),
                                marker=dict(size=6)
                            ))
                
                fig_orp.update_layout(
                    title="Evolución del ORP Batea por Piscina",
                    xaxis_title="Índice",
                    yaxis_title=cols_orp[0],
                    height=400,
                    hovermode='x unified',
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    )
                )
                st.plotly_chart(fig_orp, use_container_width=True)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Estadísticas ORP Batea")
                    stats_orp = df_orp_filtrado.groupby('Piscina')[cols_orp[0]].agg(['mean', 'min', 'max', 'std'])
                    st.dataframe(stats_orp, use_container_width=True)
                
                with col2:
                    st.subheader("Promedio ORP Batea por Piscina")
                    fig_bar_orp = px.bar(
                        df_orp_filtrado.groupby('Piscina')[cols_orp[0]].mean().reset_index(),
                        x='Piscina',
                        y=cols_orp[0],
                        color='Piscina',
                        color_discrete_map=COLORES_PISCINAS
                    )
                    st.plotly_chart(fig_bar_orp, use_container_width=True)
        
        st.subheader("Datos de ORP")
        st.dataframe(df_orp_filtrado, use_container_width=True)
    else:
        st.warning("No hay datos de ORP disponibles.")

# Tab de Microbiología
with tab_micro:
    st.header("🔬 Análisis Microbiológico")
    
    if 'Base_Micro' in dfs_laboratorio and not dfs_laboratorio['Base_Micro'].empty:
        df_micro = dfs_laboratorio['Base_Micro']
        df_micro_filtrado = df_micro.copy()
        
        if 'Piscina' in df_micro_filtrado.columns:
            df_micro_filtrado = df_micro_filtrado[df_micro_filtrado['Piscina'].isin(piscinas_seleccionadas)]
        
        if not df_micro_filtrado.empty:
            cols_numericas_micro = df_micro_filtrado.select_dtypes(include=[np.number]).columns.tolist()
            
            columnas_excluir_micro = [
                'Año', 'EMPRESA', 'ZONA', 'TIPO', 'HECTÁREAS', 'Semana', 'CICLO',
                'Fecha Recepción Larva', 'Análisis', 'Laboratorio', 'Módulo',
                'Ciclo Kraken', 'Piscina', 'Observaciones'
            ]
            cols_numericas_micro = [col for col in cols_numericas_micro if col not in columnas_excluir_micro]
            cols_numericas_micro = [col for col in cols_numericas_micro if 'análisis' not in col.lower()]
            cols_numericas_micro = [col for col in cols_numericas_micro if 'analisis' not in col.lower()]
            cols_numericas_micro = [col for col in cols_numericas_micro if 'modulo' not in col.lower()]
            cols_numericas_micro = [col for col in cols_numericas_micro if 'módulo' not in col.lower()]
            
            if cols_numericas_micro:
                parametros_micro_seleccionados = st.multiselect(
                    "Seleccionar parámetros a comparar:",
                    cols_numericas_micro,
                    default=cols_numericas_micro[:2] if len(cols_numericas_micro) >= 2 else cols_numericas_micro
                )
                
                if parametros_micro_seleccionados and 'Piscina' in df_micro_filtrado.columns:
                    fig = go.Figure()
                    
                    colores_ejes = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
                    
                    for i, parametro in enumerate(parametros_micro_seleccionados):
                        for piscina in piscinas_seleccionadas:
                            df_temp = df_micro_filtrado[df_micro_filtrado['Piscina'] == piscina].copy()
                            
                            if not df_temp.empty:
                                df_temp = df_temp.dropna(subset=[parametro])
                                
                                if len(df_temp) > 0:
                                    nombre_traza = f"{parametro} - {piscina}"
                                    color_parametro = colores_ejes[i % len(colores_ejes)]
                                    
                                    fig.add_trace(go.Scatter(
                                        x=df_temp.index,
                                        y=df_temp[parametro],
                                        mode='lines+markers',
                                        name=nombre_traza,
                                        line=dict(color=color_parametro, width=2, dash='solid' if i == 0 else 'dash'),
                                        marker=dict(size=6),
                                        yaxis='y' if i == 0 else f'y{i+1}',
                                        hovertemplate=f'<b>{piscina}</b><br>Índice: %{{x}}<br>{parametro}: %{{y:,.2f}}<extra></extra>'
                                    ))
                    
                    fig.update_layout(
                        title="Comparación de Parámetros Microbiológicos por Piscina",
                        xaxis_title="Índice",
                        height=600,
                        hovermode='x unified',
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=1.02,
                            xanchor="right",
                            x=1
                        ),
                        xaxis=dict(domain=[0.15, 0.85] if len(parametros_micro_seleccionados) > 1 else [0, 1])
                    )
                    
                    for i, parametro in enumerate(parametros_micro_seleccionados):
                        eje = 'yaxis' if i == 0 else f'yaxis{i+1}'
                        if eje == 'yaxis':
                            fig.update_layout(yaxis=dict(
                                title=dict(text=parametro, font=dict(color=colores_ejes[0])),
                                tickfont=dict(color=colores_ejes[0])
                            ))
                        else:
                            posicion = 1 + (i-1) * 0.15
                            fig.update_layout(**{eje: dict(
                                title=dict(text=parametro, font=dict(color=colores_ejes[i % len(colores_ejes)])),
                                tickfont=dict(color=colores_ejes[i % len(colores_ejes)]),
                                overlaying='y',
                                side='right',
                                position=posicion
                            )})
                    
                    st.plotly_chart(fig, use_container_width=True)
        
        st.subheader("Datos de Microbiología")
        st.dataframe(df_micro_filtrado, use_container_width=True)
    else:
        st.warning("No hay datos de microbiología disponibles.")

# Tab de Fitoplancton
with tab_fito:
    st.header("🌿 Análisis de Fitoplancton")
    
    if 'Base_Fito' in dfs_laboratorio and not dfs_laboratorio['Base_Fito'].empty:
        df_fito = dfs_laboratorio['Base_Fito']
        df_fito_filtrado = df_fito.copy()
        
        if 'Piscina' in df_fito_filtrado.columns:
            df_fito_filtrado = df_fito_filtrado[df_fito_filtrado['Piscina'].isin(piscinas_seleccionadas)]
        
        if not df_fito_filtrado.empty:
            cols_num_fito = df_fito_filtrado.select_dtypes(include=[np.number]).columns.tolist()
            
            columnas_excluir_fito = ['Año', 'HECTÁREAS', 'CICLO', 'Semana']
            cols_num_fito = [col for col in cols_num_fito if col not in columnas_excluir_fito]
            cols_num_fito = [col for col in cols_num_fito if 'observ' not in col.lower()]
            
            if cols_num_fito and 'Piscina' in df_fito_filtrado.columns:
                parametros_fito_seleccionados = st.multiselect(
                    "Seleccionar parámetros a comparar:",
                    cols_num_fito,
                    default=cols_num_fito[:2] if len(cols_num_fito) >= 2 else cols_num_fito
                )
                
                if parametros_fito_seleccionados:
                    fig = go.Figure()
                    
                    colores_ejes = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
                    
                    for i, parametro in enumerate(parametros_fito_seleccionados):
                        for piscina in piscinas_seleccionadas:
                            df_temp = df_fito_filtrado[df_fito_filtrado['Piscina'] == piscina].copy()
                            
                            if not df_temp.empty:
                                df_temp = df_temp.dropna(subset=[parametro])
                                
                                if len(df_temp) > 0:
                                    nombre_traza = f"{parametro} - {piscina}"
                                    color_parametro = colores_ejes[i % len(colores_ejes)]
                                    
                                    fig.add_trace(go.Scatter(
                                        x=df_temp.index,
                                        y=df_temp[parametro],
                                        mode='lines+markers',
                                        name=nombre_traza,
                                        line=dict(color=color_parametro, width=2, dash='solid' if i == 0 else 'dash'),
                                        marker=dict(size=6),
                                        hovertemplate=f'<b>{piscina}</b><br>Índice: %{{x}}<br>{parametro}: %{{y:,.2f}}<extra></extra>'
                                    ))
                    
                    fig.update_layout(
                        title="Comparación de Parámetros de Fitoplancton por Piscina",
                        xaxis_title="Índice",
                        yaxis_title="Conteo de Células",
                        height=600,
                        hovermode='x unified',
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=1.02,
                            xanchor="right",
                            x=1
                        )
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
        
        st.subheader("Datos de Fitoplancton")
        st.dataframe(df_fito_filtrado, use_container_width=True)
    else:
        st.warning("No hay datos de fitoplancton disponibles.")

# Tab de VS (Comparador)
with tab_vs:
    st.header("⚖️ Comparador VS")
    st.markdown("---")
        
    # Crear diccionario con todos los dataframes disponibles
    dfs_disponibles = {}
    
    if df_produccion is not None and not df_produccion.empty:
        dfs_disponibles['Producción'] = df_produccion
    
    if 'Base_Quimicos' in dfs_laboratorio and not dfs_laboratorio['Base_Quimicos'].empty:
        dfs_disponibles['Químicos'] = dfs_laboratorio['Base_Quimicos']
    
    if 'Base_Pato' in dfs_laboratorio and not dfs_laboratorio['Base_Pato'].empty:
        dfs_disponibles['Patógenos'] = dfs_laboratorio['Base_Pato']
    
    if 'Base_ORP' in dfs_laboratorio and not dfs_laboratorio['Base_ORP'].empty:
        dfs_disponibles['ORP'] = dfs_laboratorio['Base_ORP']
    
    if 'Base_Micro' in dfs_laboratorio and not dfs_laboratorio['Base_Micro'].empty:
        dfs_disponibles['Microbiología'] = dfs_laboratorio['Base_Micro']
    
    if 'Base_Fito' in dfs_laboratorio and not dfs_laboratorio['Base_Fito'].empty:
        dfs_disponibles['Fitoplancton'] = dfs_laboratorio['Base_Fito']
    
    if dfs_disponibles:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Filtro 1 (Eje Y Izquierdo)")
            fuente_1 = st.selectbox("Seleccionar fuente de datos:", list(dfs_disponibles.keys()), key="fuente_1")
            
            df_fuente_1 = dfs_disponibles[fuente_1]
            cols_numericas_1 = df_fuente_1.select_dtypes(include=[np.number]).columns.tolist()
            
            if cols_numericas_1:
                columna_1 = st.selectbox("Seleccionar columna:", cols_numericas_1, key="columna_1")
            else:
                st.warning("No hay columnas numéricas en esta fuente")
                columna_1 = None
        
        with col2:
            st.subheader("Filtro 2 (Eje Y Derecho)")
            fuente_2 = st.selectbox("Seleccionar fuente de datos:", list(dfs_disponibles.keys()), key="fuente_2")
            
            df_fuente_2 = dfs_disponibles[fuente_2]
            cols_numericas_2 = df_fuente_2.select_dtypes(include=[np.number]).columns.tolist()
            
            if cols_numericas_2:
                columna_2 = st.selectbox("Seleccionar columna:", cols_numericas_2, key="columna_2")
            else:
                st.warning("No hay columnas numéricas en esta fuente")
                columna_2 = None
        
        if st.button("Generar Comparación", type="primary"):
            if columna_1 and columna_2:
                fig = go.Figure()
                
                # Procesar Filtro 1
                df_1 = dfs_disponibles[fuente_1].copy()
                if 'Piscina' in df_1.columns:
                    df_1 = df_1[df_1['Piscina'].isin(piscinas_seleccionadas)]
                
                fecha_col_1 = encontrar_columna_fecha(df_1)
                if fecha_col_1 is None and 'F.muestreo' in df_1.columns:
                    fecha_col_1 = 'F.muestreo'
                
                for piscina in piscinas_seleccionadas:
                    if 'Piscina' in df_1.columns:
                        df_temp = df_1[df_1['Piscina'] == piscina].copy()
                    else:
                        df_temp = df_1.copy()
                    
                    if not df_temp.empty:
                        if fecha_col_1 and fecha_col_1 in df_temp.columns:
                            df_temp[fecha_col_1] = pd.to_datetime(df_temp[fecha_col_1], errors='coerce')
                            df_temp = df_temp.dropna(subset=[fecha_col_1, columna_1])
                            df_temp = df_temp.sort_values(fecha_col_1)
                            x_values = df_temp[fecha_col_1]
                        else:
                            df_temp = df_temp.dropna(subset=[columna_1])
                            x_values = df_temp.index
                        
                        if len(df_temp) > 0:
                            fig.add_trace(go.Scatter(
                                x=x_values,
                                y=df_temp[columna_1],
                                mode='lines+markers',
                                name=f"{columna_1} - {piscina}",
                                line=dict(color=COLORES_PISCINAS.get(piscina, '#1f77b4'), width=2),
                                marker=dict(size=6),
                                yaxis='y',
                                hovertemplate=f'<b>{piscina}</b><br>Fecha: %{{x}}<br>{columna_1}: %{{y:,.2f}}<extra></extra>'
                            ))
                
                # Procesar Filtro 2
                df_2 = dfs_disponibles[fuente_2].copy()
                if 'Piscina' in df_2.columns:
                    df_2 = df_2[df_2['Piscina'].isin(piscinas_seleccionadas)]
                
                fecha_col_2 = encontrar_columna_fecha(df_2)
                if fecha_col_2 is None and 'F.muestreo' in df_2.columns:
                    fecha_col_2 = 'F.muestreo'
                
                for piscina in piscinas_seleccionadas:
                    if 'Piscina' in df_2.columns:
                        df_temp = df_2[df_2['Piscina'] == piscina].copy()
                    else:
                        df_temp = df_2.copy()
                    
                    if not df_temp.empty:
                        if fecha_col_2 and fecha_col_2 in df_temp.columns:
                            df_temp[fecha_col_2] = pd.to_datetime(df_temp[fecha_col_2], errors='coerce')
                            df_temp = df_temp.dropna(subset=[fecha_col_2, columna_2])
                            df_temp = df_temp.sort_values(fecha_col_2)
                            x_values = df_temp[fecha_col_2]
                        else:
                            df_temp = df_temp.dropna(subset=[columna_2])
                            x_values = df_temp.index
                        
                        if len(df_temp) > 0:
                            fig.add_trace(go.Scatter(
                                x=x_values,
                                y=df_temp[columna_2],
                                mode='lines+markers',
                                name=f"{columna_2} - {piscina}",
                                line=dict(color=COLORES_PISCINAS.get(piscina, '#ff7f0e'), width=2, dash='dash'),
                                marker=dict(size=6, symbol='diamond'),
                                yaxis='y2',
                                hovertemplate=f'<b>{piscina}</b><br>Fecha: %{{x}}<br>{columna_2}: %{{y:,.2f}}<extra></extra>'
                            ))
                
                fig.update_layout(
                    title=f"Comparación: {columna_1} vs {columna_2}",
                    xaxis_title="Fecha",
                    height=600,
                    hovermode='x unified',
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    ),
                    yaxis=dict(
                        title=dict(text=columna_1, font=dict(color='#1f77b4')),
                        tickfont=dict(color='#1f77b4')
                    ),
                    yaxis2=dict(
                        title=dict(text=columna_2, font=dict(color='#ff7f0e')),
                        tickfont=dict(color='#ff7f0e'),
                        overlaying='y',
                        side='right'
                    )
                )
                
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Selecciona columnas válidas para ambos filtros.")
    else:
        st.warning("No hay datos disponibles para comparar.")

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
        <p>Dashboard Área Experimental - Grupo Corporativo Fajardo</p>
        <p>Actualizado: {}</p>
    </div>
    """.format(datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
    unsafe_allow_html=True
)