# 📦 Instalar soporte para Excel
!pip install -q xlrd xlsxwriter

import pandas as pd
from google.colab import files
from datetime import datetime

# 💰 Ingreso de Tipo de Cambio
try:
    tc = float(input("Ingrese el tipo de cambio (ejemplo 3.75): "))
except:
    tc = 3.75
    print("Usando TC por defecto: 3.75")

# 📁 1. Subir archivo
uploaded = files.upload()
if uploaded:
    archivo = list(uploaded.keys())[0]

    # 🧾 2. Leer desde fila 4
    xls = pd.ExcelFile(archivo)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=3)
    df.columns = df.columns.str.strip().str.lower()

    # 🔍 3. Filtro Inicial: Solo SIMSA y Estado D
    df = df[(df['un'] == 'SIMSA') & (df['estado'].str.strip().str.upper() == 'D')].copy()

    # 📅 4. Normalizar Fechas
    df['f pedido'] = pd.to_datetime(df['f pedido'], errors='coerce').dt.normalize()

    # 🧩 5. AGRUPACIÓN PARA EVITAR DUPLICADOS (Consolidar recepciones parciales)
    columnas_id = ['un', 'f pedido', 'nº pedido', 'estado', 'nom 1', 'artículo', 'más info', 'um', 'moneda', 'precio']
    
    df_agrupado = df.groupby(columnas_id, as_index=False).agg({
        'cant ped': 'max',
        'neto recep': 'sum'
    })

    # ➕ 6. Cálculos de Saldo y Aging
    df_agrupado['cant_pendiente'] = df_agrupado['cant ped'] - df_agrupado['neto recep']
    df_p = df_agrupado[df_agrupado['cant_pendiente'] > 0.01].copy()

    hoy = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    df_p['días_emisión'] = (hoy - df_p['f pedido']).dt.days

    # Nueva lógica de estados
    def definir_prioridad(d):
        if pd.isna(d): return 'REVISAR'
        if d <= 15: return 'OK'
        elif d <= 30: return 'ALERTA'
        elif d <= 60: return 'VENCIDO'
        else: return 'OLVIDADO'

    df_p['prioridad'] = df_p['días_emisión'].apply(definir_prioridad)

    # 💵 7. Conversión a USD
    df_p['pu usd'] = df_p.apply(lambda r: r['precio'] if 'USD' in str(r['moneda']).upper() else r['precio'] / tc, axis=1)
    df_p['TOTAL_pendiente_USD'] = df_p['pu usd'] * df_p['cant_pendiente']

    # 📋 8. Reordenar Columnas
    cols_finales = [
        'un', 'f pedido', 'nº pedido', 'estado', 'nom 1', 'artículo', 'más info',
        'um', 'cant ped', 'neto recep', 'cant_pendiente', 
        'prioridad', 'días_emisión', 
        'moneda', 'precio', 'pu usd', 'TOTAL_pendiente_USD'
    ]
    
    df_final = df_p[cols_finales].rename(columns={
        'un': 'Sede', 'f pedido': 'Fecha Emision', 'nº pedido': 'OC #',
        'estado': 'Estado Orig', 'nom 1': 'Proveedor', 'artículo': 'Codigo',
        'más info': 'Descripcion', 'um': 'UM', 'precio': 'Precio Unit Orig'
    })

    df_final = df_final.sort_values(by='Fecha Emision', ascending=False)

    # 💾 9. Exportar con Formatos Personalizados
    nombre_excel = 'Gestion_OC_SIMSA_Consolidado.xlsx'
    writer = pd.ExcelWriter(nombre_excel, engine='xlsxwriter', datetime_format='dd/mm/yyyy')
    df_final.to_excel(writer, sheet_name='Pendientes_D', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Pendientes_D']
    
    # FORMATOS DE COLOR
    f_gris = workbook.add_format({'bg_color': '#D3D3D3', 'font_color': '#545454'}) # Olvidado
    f_rojo = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Vencido
    f_amar = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'}) # Alerta
    f_verd = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'}) # OK
    f_dolar = workbook.add_format({'num_format': '$#,##0.00'})

    # --- APLICAR COLOR A COLUMNAS L (11) y M (12) ---
    row_count = len(df_final)
    for col in [11, 12]:
        # Formato para OLVIDADO (Gris)
        worksheet.conditional_format(1, col, row_count, col,
                                     {'type': 'formula', 'criteria': '=$L2="OLVIDADO"', 'format': f_gris})
        # Formato para VENCIDO (Rojo)
        worksheet.conditional_format(1, col, row_count, col,
                                     {'type': 'formula', 'criteria': '=$L2="VENCIDO"', 'format': f_rojo})
        # Formato para ALERTA (Amarillo)
        worksheet.conditional_format(1, col, row_count, col,
                                     {'type': 'formula', 'criteria': '=$L2="ALERTA"', 'format': f_amar})
        # Formato para OK (Verde)
        worksheet.conditional_format(1, col, row_count, col,
                                     {'type': 'formula', 'criteria': '=$L2="OK"', 'format': f_verd})

    # Ajustes de ancho y formato moneda
    worksheet.set_column('P:P', 13, f_dolar) # pu usd con ancho 13
    worksheet.set_column('Q:Q', 18, f_dolar) # TOTAL_pendiente_USD
    worksheet.set_column('A:K', 14)
    worksheet.set_column('L:M', 15)

    writer.close()
    print(f"Éxito: Reporte consolidado generado con {len(df_final)} líneas reales.")
    files.download(nombre_excel)
