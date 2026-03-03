import pandas as pd
from google.colab import files

# 📁 Subir archivo
uploaded = files.upload()
archivo = list(uploaded.keys())[0]

# 🧾 Leer desde fila 4 (header=3) y limpiar nombres de columnas
df = pd.read_excel(archivo, sheet_name=0, header=3)
df.columns = df.columns.str.strip()

# 🔍 Filtrar solo SIMSA con recepción
df = df[(df["UN"] == "SIMSA") & (df["Neto Recep"] > 0)].copy()

# 💱 Tasa de cambio variable input PEN a USD
try:
    tasa_cambio = float(input("💲 Ingresa la tasa de cambio PEN a USD (ej. 3.5): "))
except:
    print("⚠️ Entrada inválida. Se usará tasa por defecto de 3.5")
    tasa_cambio = 3.5

print(f"✅ Tasa de cambio usada: {tasa_cambio}")

# 💲 Convertir precio a USD si la moneda es PEN
df["Precio USD"] = df.apply(
    lambda x: x["Precio"] / tasa_cambio if x["Moneda"] == "PEN" else x["Precio"], axis=1
)

# 🧮 Calcular importe en USD
df["Importe USD"] = df["Neto Recep"] * df["Precio USD"]

# 📝 Renombrar columnas para el reporte
df.rename(columns={
    "Más Info": "Descripción",
    "Nom 1": "Nombre Proveedor",
    "Nº ID": "RUC",
    "Descr": "Nombre Familia"
}, inplace=True)

# 📋 Seleccionar y ordenar columnas finales
columnas_finales = [
    "UN",
    "UN IN",
    "F/H Recep",
    "F Pedido",
    "Nº Pedido",
    "Línea",
    "Artículo",
    "Descripción",
    "Neto Recep",
    "UM",
    "Precio USD",
    "Importe USD",
    "No. de Parte",
    "Nombre Proveedor",
    "RUC",
    "Fam",
    "Nombre Familia"
]

# ✅ Generar DataFrame final
df_final = df[columnas_finales].copy()

# 💾 Exportar a Excel
nombre_archivo = "resultado_oc_recepcion_validadas.xlsx"
df_final.to_excel(nombre_archivo, index=False)
files.download(nombre_archivo)

