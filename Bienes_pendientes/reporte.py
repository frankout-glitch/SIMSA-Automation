# 📦 Instalar soporte para archivos .xls
!pip install -q xlrd

import pandas as pd
from google.colab import files

# 📁 Subir archivo original desde PeopleSoft
uploaded = files.upload()
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

# ➕ Calcular campo pendiente con nuevo nombre
df_agrupado['cantidad pendiente rq bienes'] = df_agrupado['cantidad solicitud'] - df_agrupado['cantidad pedido']

# 🧼 Filtrar solo pendientes
df_agrupado = df_agrupado[df_agrupado['cantidad pendiente rq bienes'] > 0]

# 📝 Renombrar columnas
df_agrupado.rename(columns={
    'fecha aprobación': 'Fecha aprobación RQ',
    'fecha solicitud': 'Fecha solicitud RQ',
    'cantidad solicitud': 'Cantidad solicitud',
    'cantidad pedido': 'Cantidad pedido'
}, inplace=True)

# 📅 Formatear fechas como dd/mm/yyyy
df_agrupado['Fecha aprobación RQ'] = pd.to_datetime(df_agrupado['Fecha aprobación RQ']).dt.strftime('%d/%m/%Y')
df_agrupado['Fecha solicitud RQ'] = pd.to_datetime(df_agrupado['Fecha solicitud RQ']).dt.strftime('%d/%m/%Y')

# 📋 Reordenar columnas
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

df_final = df_agrupado[columnas_final].sort_values(by='Fecha aprobación RQ', ascending=False)

# 📊 Mostrar cuántas filas hay
print("Total de filas exportadas:", len(df_final))

# 💾 Exportar archivo final
nombre_archivo = 'resultado_rq_bienes_pendientes.xlsx'
df_final.to_excel(nombre_archivo, index=False)
files.download(nombre_archivo)
