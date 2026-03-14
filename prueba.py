import pandas as pd

archivo_entrada = "INVF COMFY.TXT"
archivo_salida = "datos_ordenado.txt"

#El .txt viene con las columnas en el siguiente orden
columnas_raw = [
    "codigo",
    "departamento",
    "descripcion",
    "color",
    "precio",
    "costo",
    "cantidad",
    "talla",
    "referencia"
]

#leer el archivo sin header
df = pd.read_csv(
    archivo_entrada,
    sep="\t",
    header=None,
    names=columnas_raw
)

#orden de columnas deseado
orden_columnas = [
    "codigo",
    "referencia",
    "departamento",
    "descripcion",
    "precio",
    "costo",
    "color",
    "talla",
    "cantidad"
]

df = df[orden_columnas]

#guardar nuevamente

df.to_csv(
    archivo_salida,
    sep="\t",
    index=False,
    header=False
)

print("archivo Procesado")