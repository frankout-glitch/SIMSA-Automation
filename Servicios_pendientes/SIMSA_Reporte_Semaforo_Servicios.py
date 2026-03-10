# 📦 Instalar soporte para archivos .xls y formato visual
!pip install -q xlrd xlsxwriter

import pandas as pd
from google.colab import files
from datetime import datetime

# 📁 1. Subir archivo original desde PeopleSoft (LOG_REQ_OC_REC_PU_ALL)
uploaded = files.upload()
if uploaded:
    archivo = list(uploaded.keys())[0]

    # 🧾 2. Leer hoja (desde fila 4) y limpiar columnas
    xls = pd.ExcelFile(archivo)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=3)
    df.columns = df.columns.str.strip().str.lower()

    # 🔍 3. Filtrar por SIMSE y aprobadas (Lógica de Servicios)
    df_filtrado = df[
        (df['unidad negocio'] == 'SIMSE') &
        (df['estado solicitud'] == 'Aprobada') &
        (df['estado actual'] == 'Aprobada')
    ].copy()

    # 🧩 4. Agrupar por ID solicitud + línea
    df_agrupado = df_filtrado.groupby(['id solicitud', 'número línea'], as_index=False).agg({
        'fecha solicitud': 'max',
        'fecha aprobación': 'max',
        'solicitante': 'first',
        'id artículo': 'first',
        'más información': 'first',
        'coment': 'first',
        'unidad medida': 'first',
        'cantidad solicitud': 'max',
        'cantidad pedido': 'sum'
    })

    # ➕ 5. Calcular campo pendiente
    df_agrupado['cantidad pendiente rq servicios'] = df_agrupado['cantidad solicitud'] - df_agrupado['cantidad pedido']
    df_agrupado = df_agrupado[df_agrupado['cantidad pendiente rq servicios'] > 0].copy()

    # 📅 6. CÁLCULO DE ANTIGÜEDAD (Desde Aprobación hasta Hoy)
    df_agrupado['fecha_aprob_dt'] = pd.to_datetime(df_agrupado['fecha aprobación'], errors='coerce')
    hoy = datetime.now()
    df_agrupado['días_aprobada_sin_atender'] = (hoy - df_agrupado['fecha_aprob_dt']).dt.days

    # Semáforo AJUSTADO PARA SERVICIOS (10, 20, 40 días)
    def asignar_semaforo_servicios(d):
        if pd.isna(d): return 'Revisar Fecha'
        if d <= 10: return 'Tiempo en proceso ok'
        elif d <= 20: return 'Alerta'
        elif d <= 40: return 'Critico a revisar'
        else: return 'Olvidados'

    df_agrupado['prioridad tiempo'] = df_agrupado['días_aprobada_sin_atender'].apply(asignar_semaforo_servicios)

    # 📝 7. Formatear fechas y Renombrar columnas
    df_agrupado['Fecha aprobación RQ'] = pd.to_datetime(df_agrupado['fecha aprobación']).dt.strftime('%d/%m/%Y')
    df_agrupado['Fecha solicitud RQ'] = pd.to_datetime(df_agrupado['fecha solicitud']).dt.strftime('%d/%m/%Y')

    # 📋 8. Reordenar: Tu estructura + Semáforos a la DERECHA
    columnas_final = [
        'Fecha aprobación RQ',
        'Fecha solicitud RQ',
        'id solicitud',
        'solicitante',
        'número línea',
        'id artículo',
        'más información',
        'coment',
        'unidad medida',
        'cantidad solicitud',
        'cantidad pedido',
        'cantidad pendiente rq servicios',
        'prioridad tiempo',             
        'días_aprobada_sin_atender'     
    ]

    df_final = df_agrupado[columnas_final].rename(columns={
        'cantidad solicitud': 'Cantidad solicitud',
        'cantidad pedido': 'Cantidad pedido'
    })

    # Ordenar por lo más reciente arriba
    df_final = df_final.sort_values(by='días_aprobada_sin_atender', ascending=True)

    # 💾 9. Exportar con Formato Visual a Excel
    nombre_archivo = 'resultado_rq_servicios_priorizado.xlsx'
    writer = pd.ExcelWriter(nombre_archivo, engine='xlsxwriter')
    df_final.to_excel(writer, sheet_name='Pendientes_Servicios', index=False)

    workbook  = writer.book
    worksheet = writer.sheets['Pendientes_Servicios']

    # Formatos de color
    f_v = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    f_a = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'})
    f_r = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    f_g = workbook.add_format({'bg_color': '#F2F2F2', 'font_color': '#7F7F7F'})

    # Aplicar formato condicional en la columna de 'prioridad tiempo' (índice 12)
    col_idx = 12
    rango = f'M2:M{len(df_final)+1}' 
    
    ws_rules = [
        ('Tiempo en proceso ok', f_v),
        ('Alerta', f_a),
        ('Critico a revisar', f_r),
        ('Olvidados', f_g)
    ]

    for label, fmt in ws_rules:
        worksheet.conditional_format(1, col_idx, len(df_final), col_idx, {
            'type': 'cell', 'criteria': 'equal to', 'value': f'"{label}"', 'format': fmt
        })

    # 📏 Ajuste de Anchos de Columna
    for i, col in enumerate(df_final.columns):
        if col == 'prioridad tiempo':
            worksheet.set_column(i, i, 22)
        elif col == 'días_aprobada_sin_atender':
            worksheet.set_column(i, i, 26)
        else:
            worksheet.set_column(i, i, 15)

    writer.close()
    print(f"Total de servicios SIMSE analizados: {len(df_final)}")
    files.download(nombre_archivo)
