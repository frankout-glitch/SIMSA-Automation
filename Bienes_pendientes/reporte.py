# 📦 Instalar soporte para archivos .xls y formato visual
!pip install -q xlrd xlsxwriter

import pandas as pd
from google.colab import files

# 📁 Subir archivo original desde PeopleSoft
uploaded = files.upload()
if uploaded:
    archivo = list(uploaded.keys())[0]

    # 🧾 Leer hoja (desde fila 4) y limpiar columnas
    xls = pd.ExcelFile(archivo)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=3)
    df.columns = df.columns.str.strip().str.lower()

    # 🔍 Filtrar por SIMSA y aprobadas
    df_filtrado = df[
        (df['unidad negocio'] == 'SIMSA') &
        (df['estado solicitud'] == 'Aprobada') &
        (df['estado actual'] == 'Aprobada')
    ].copy()

    # 🧩 Agrupar por ID solicitud + línea
    columnas_grupo = ['id solicitud', 'número línea']
    df_agrupado = df_filtrado.groupby(columnas_grupo, as_index=False).agg({
        'fecha solicitud': 'max',
        'fecha aprobación': 'max',
        'solicitante': 'first',
        'id artículo': 'first',
        'más información': 'first',
        'id fabricante': 'first',
        'numero de parte': 'first',
        'unidad medida': 'first',
        'cantidad solicitud': 'max',
        'cantidad pedido': 'sum'
    })

    # ➕ Calcular campo pendiente
    df_agrupado['cantidad pendiente rq bienes'] = df_agrupado['cantidad solicitud'] - df_agrupado['cantidad pedido']
    df_agrupado = df_agrupado[df_agrupado['cantidad pendiente rq bienes'] > 0].copy()

    # 📝 Renombrar columnas
    df_agrupado.rename(columns={
        'fecha aprobación': 'Fecha aprobación RQ',
        'fecha solicitud': 'Fecha solicitud RQ',
        'cantidad solicitud': 'Cantidad solicitud',
        'cantidad pedido': 'Cantidad pedido'
    }, inplace=True)

    # 📅 CONVERSIÓN A FECHA PURA (Sin hora y sin convertir a texto)
    df_agrupado['Fecha aprobación RQ'] = pd.to_datetime(df_agrupado['Fecha aprobación RQ'], errors='coerce').dt.normalize()
    df_agrupado['Fecha solicitud RQ'] = pd.to_datetime(df_agrupado['Fecha solicitud RQ'], errors='coerce').dt.normalize()

    # 📋 Reordenar y ORDENAR (De la más reciente a la más antigua)
    columnas_final = [
        'Fecha aprobación RQ',
        'Fecha solicitud RQ',
        'id solicitud',
        'solicitante',
        'número línea',
        'id artículo',
        'más información',
        'id fabricante',
        'numero de parte',
        'unidad medida',
        'Cantidad solicitud',
        'Cantidad pedido',
        'cantidad pendiente rq bienes'
    ]

    # Ordenamos por objeto fecha real
    df_final = df_agrupado[columnas_final].sort_values(by='Fecha aprobación RQ', ascending=False)

    # 💾 Exportar forzando formato de fecha corto y ancho de columna
    nombre_archivo = 'resultado_rq_bienes_pendientes.xlsx'
    
    # datetime_format asegura que Excel muestre dd/mm/yyyy automáticamente
    writer = pd.ExcelWriter(nombre_archivo, engine='xlsxwriter', datetime_format='dd/mm/yyyy')
    
    df_final.to_excel(writer, index=False, sheet_name='Pendientes')

    workbook  = writer.book
    worksheet = writer.sheets['Pendientes']

    # Aplicar ancho 15 a todas las columnas
    for i in range(len(columnas_final)):
        worksheet.set_column(i, i, 15)

    writer.close()
    print("Total de filas exportadas:", len(df_final))
    files.download(nombre_archivo)
