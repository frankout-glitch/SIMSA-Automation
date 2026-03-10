# 📦 Instalar soporte para Excel
!pip install -q xlrd xlsxwriter

import pandas as pd
from google.colab import files
from datetime import datetime

# 📁 1. Subir archivo original
uploaded = files.upload()
if uploaded:
    archivo = list(uploaded.keys())[0]

    # 🧾 2. Leer con header en fila 3
    xls = pd.ExcelFile(archivo)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=3)
    df.columns = df.columns.str.strip().str.lower()

    # Buscador de columnas por si PeopleSoft corta los nombres
    def buscar_col(lista, texto):
        for c in lista:
            if texto in c: return c
        return None

    c_fab = buscar_col(df.columns, 'fabrican') or 'id fabricante'
    c_par = buscar_col(df.columns, 'par') or 'numero de parte'
    c_um = buscar_col(df.columns, 'unidad medida') or 'unidad medida'

    # 🔍 3. Filtrar aprobadas de SIMSA
    df_f = df[
        (df['unidad negocio'].str.contains('SIMSA', na=False)) & 
        (df['estado solicitud'] == 'Aprobada') & 
        (df['estado actual'] == 'Aprobada')
    ].copy()

    # 🧩 4. Agrupar manteniendo valores originales
    df_agg = df_f.groupby(['id solicitud', 'número línea'], as_index=False).agg({
        'fecha solicitud': 'max',
        'fecha aprobación': 'max',
        'solicitante': 'first',
        'id artículo': 'first',
        'más información': 'first',
        c_fab: 'first', 
        c_par: 'first', 
        c_um: 'first',
        'cantidad solicitud': 'max',
        'cantidad pedido': 'sum'
    })

    # ➕ 5. Calcular Pendientes
    df_agg['cantidad pendiente rq bienes'] = df_agg['cantidad solicitud'] - df_agg['cantidad pedido']
    df_agg = df_agg[df_agg['cantidad pendiente rq bienes'] > 0].copy()

    # 📅 6. CÁLCULO CRÍTICO: Desde APROBACIÓN hasta HOY
    df_agg['fecha_aprob_dt'] = pd.to_datetime(df_agg['fecha aprobación'], errors='coerce')
    hoy = datetime.now()
    df_agg['días_aprobada_sin_atender'] = (hoy - df_agg['fecha_aprob_dt']).dt.days

    # Semáforo limpio (sin números)
    def asignar_semaforo(d):
        if pd.isna(d): return 'Revisar Fecha'
        if d <= 7: return 'Tiempo en proceso ok'
        elif d <= 15: return 'Alerta'
        elif d <= 29: return 'Critico a revisar'
        else: return 'Olvidados'

    df_agg['prioridad tiempo'] = df_agg['días_aprobada_sin_atender'].apply(asignar_semaforo)

    # 📝 7. Formatear fechas para el reporte final
    df_agg['Fecha aprobación RQ'] = pd.to_datetime(df_agg['fecha aprobación']).dt.strftime('%d/%m/%Y')
    df_agg['Fecha solicitud RQ'] = pd.to_datetime(df_agg['fecha solicitud']).dt.strftime('%d/%m/%Y')

    # 📋 8. Reordenar: Tu estructura + Columnas a la DERECHA
    cols_finales = [
        'Fecha aprobación RQ',
        'Fecha solicitud RQ',
        'id solicitud',
        'solicitante',
        'número línea',
        'id artículo',
        'más información',
        c_fab,
        c_par,
        c_um,
        'cantidad solicitud',
        'cantidad pedido',
        'cantidad pendiente rq bienes',
        'prioridad tiempo',          
        'días_aprobada_sin_atender'  
    ]

    df_final = df_agg[cols_finales].rename(columns={
        'cantidad solicitud': 'Cantidad solicitud', 
        'cantidad pedido': 'Cantidad pedido', 
        c_fab: 'ID Fabricante', 
        c_par: 'Número de Parte', 
        c_um: 'Unidad Medida'
    })
    
    # Ordenar por lo más reciente arriba (menor antigüedad arriba)
    df_final = df_final.sort_values(by='días_aprobada_sin_atender', ascending=True)

    # 💾 9. Exportar
    nombre = 'Resultado_RQ_Priorizado_SIMSA.xlsx'
    writer = pd.ExcelWriter(nombre, engine='xlsxwriter')
    df_final.to_excel(writer, sheet_name='Pendientes', index=False)

    wb, ws = writer.book, writer.sheets['Pendientes']
    
    # Formatos
    f_v = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    f_a = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'})
    f_r = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    f_g = wb.add_format({'bg_color': '#F2F2F2', 'font_color': '#7F7F7F'})

    # Aplicar color en columna N (índice 13)
    rango = f'N2:N{len(df_final)+1}'
    ws.conditional_format(rango, {'type': 'cell', 'criteria': 'equal to', 'value': '"Tiempo en proceso ok"', 'format': f_v})
    ws.conditional_format(rango, {'type': 'cell', 'criteria': 'equal to', 'value': '"Alerta"', 'format': f_a})
    ws.conditional_format(rango, {'type': 'cell', 'criteria': 'equal to', 'value': '"Critico a revisar"', 'format': f_r})
    ws.conditional_format(rango, {'type': 'cell', 'criteria': 'equal to', 'value': '"Olvidados"', 'format': f_g})

    # Ajuste de Ancho dinámico
    for i, col in enumerate(df_final.columns):
        if col == 'prioridad tiempo':
            ws.set_column(i, i, 22)
        elif col == 'días_aprobada_sin_atender':
            ws.set_column(i, i, 26)
        else:
            ws.set_column(i, i, 15)

    writer.close()
    files.download(nombre)
