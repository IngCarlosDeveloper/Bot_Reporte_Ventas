import os
import subprocess
import time
import pandas as pd
import re
from openpyxl import load_workbook
import traceback
import shutil
import fnmatch
from datetime import datetime

#Ruta donde estan los .txt

ruta_base = os.path.dirname(os.path.abspath(__file__))

carpeta_txt = os.path.join(ruta_base, "Data")

#Enlista cada .txt que hay en la ruta especificada
archivos_txt = [f for f in os.listdir(carpeta_txt) if f.endswith(".TXT")]

#---------------------------------------------------BUCLE PARA LEER LAS LINEAS DEL .TXT-------------------------

for archivo in archivos_txt:
    try:

        ruta_archivo = os.path.join(carpeta_txt, archivo)

        with open(ruta_archivo, "r", encoding="latin-1") as f:

            linea = [line.strip() for line in f if line.strip()]

            resultado = []

            for l in linea:
                
                #eliminar los espacios sobrantes
                space = re.sub(r" {3,}", "", l)

                #Cambiar puntos por comas (para estandarizar)
                space = space.replace(".", "¦")
                space = space.replace(",", ".")
                space = space.replace("¦", ",")

                #divide por tabs
                tabs = space.split("\t")

                #inserta el resultado en un array
                resultado.append(tabs)

            #-------------------------------------------------PREPARAR DATOS----------------------------------------------

            if fnmatch.fnmatch(archivo, "*RVM*"):

                producto = []
                venta = []
                concatenacion = []
                
                #Bucle para leer cada fila del .txt
                for dato in resultado:

                    # >5 valores = descripcion de producto
                    if len(dato) >= 5:
                        producto = []
                        venta = []
                        producto = dato
                    
                    # == 4 es una venta
                    elif len(dato) == 4:
                        venta = []
                        venta = dato

                        venta.insert(2, producto[0]) #agregar codigo

                        if len(producto) == 8:
                            venta.insert(3, producto[7]) #agregar referencia
                        elif len(producto) == 7:
                            venta.insert(3, "") #agregar referencia vacia

                        venta.insert(4, producto[2]) #agregar descripcion
                        venta.insert(5, producto[3]) #agregar color
                        venta.insert(6, producto[4]) #agregar talla
                        venta.insert(7, producto[1]) #agregar departamento
                        venta.insert(8, producto[5]) #agregar pvp
                        venta.insert(10, producto[6]) #agregar costo

                        #------------------------------------- corregido: limpiar antes de convertir
                        try:
                            cantidad = float(str(venta[9]).replace(",", ".").strip())
                            precio = float(str(venta[8]).replace(",", ".").strip())
                            venta_bruta = precio * cantidad
                        except ValueError:
                            venta_bruta = 0.0

                        venta.insert(9, venta_bruta)

                    if len(venta) > 0:
                        concatenacion.append(venta) #insertamos la venta dentro del array completo [concatenacion]

                #---------------------------------------------------CREAR EL EXCEL---------------------------------------
                columnas_excel = ["NUMERO_TRANSACCION", "FECHA_TRANSACCION", "CODIGO_BARRA", "REFERENCIA", "DESCRIPCION", "COLOR", "TALLA", "DEPARTAMENTO", "PRECIO_UNITARIO", "VENTAS_BRUTAS_USD", "UNIDADES_VENDIDAS", "COSTO_VENTA_USD", "PORCENTAJE_DESCUENTO"]

                df_venta = pd.DataFrame(concatenacion, columns=columnas_excel)

                df_venta["CODIGO_BARRA"] = df_venta["CODIGO_BARRA"].astype(str)
                
                titulo = ruta_archivo.replace(".TXT", ".xlsx")

                with pd.ExcelWriter(titulo, engine="openpyxl") as writer:
                    df_venta.to_excel(writer, sheet_name="Hoja1", index=False)
                    print("Excel generado exitosamente")

                #---------------------------------------------------PASAR LA COLUMNA A TIPO TEXTO---------------------------------
                wb = load_workbook(titulo) #abre el libro
                ws = wb["Hoja1"] #abre la hoja

                columnas_texto = ["NUMERO_TRANSACCION", "CODIGO_BARRA"]

                for columna in columnas_texto:

                    col_idx = None

                    for idx, cell in enumerate(ws[1], start=1):
                        if cell.value == columna:
                            col_idx = idx
                            break

                    if col_idx:
                        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                            for cell in row:
                                cell.number_format = "@"

                wb.save(titulo)
                print("Formato de texto aplicado a CODIGO_BARRA")

            elif fnmatch.fnmatch(archivo, "*INVF*"):

                columnas_inventario = ["CODIGO_BARRA", "REFERENCIA", "DEPARTAMENTO", "DESCRIPCION", "PRECIO_NETO", "COSTO_UNITARIO", "COLOR", "TALLA", "UNIDADES_INVENTARIO"]
                hoy = datetime.now().date()

                #-------------------------------IDENTIFICAR TIENDA Y MARCA------------------------------                
                lista_tienda = [
                    "DKNY TOLON",
                    "DKNY CV",
                    "PENGUIN CV",
                    "PENGUIN TOLON",
                    "PENGUIN SAMBIL",
                    "PENGUIN LIDER",
                    "PURO EGO TOLON",
                    "PENGUIN BARQUISIMETO",
                    "PENGUIN AEROPUERTO",
                    "PENGUIN MARGARITA",
                    "PURO EGO MARGARITA"
                ]

                lista_marcas = ["DKNY", "PENGUIN", "PURO EGO"]

                for tienda in lista_tienda:
                    if fnmatch.fnmatch(archivo, f"*{tienda}*"):
                        tienda_actual = tienda
                        print(f"Tienda identificada: {tienda_actual}")
                        break

                for marca in lista_marcas:
                    if fnmatch.fnmatch(tienda_actual, f"*{marca}*"):
                        marca_actual = marca
                        print(f"Marca identificada: {marca_actual}")
                        break

                df_inv = pd.DataFrame(resultado, columns=columnas_inventario)

                #------------------- corregido: limpiar y convertir
                df_inv['COSTO_UNITARIO'] = pd.to_numeric(
                    df_inv['COSTO_UNITARIO'].astype(str).str.replace(",", ".").str.strip(),
                    errors="coerce"
                )
                df_inv['UNIDADES_INVENTARIO'] = pd.to_numeric(
                    df_inv['UNIDADES_INVENTARIO'].astype(str).str.replace(",", ".").str.strip(),
                    errors="coerce"
                )

                df_inv["CODIGO_BARRA"] = df_inv["CODIGO_BARRA"].astype(str)
                df_inv["COLOR"] = df_inv["COLOR"].astype(str)

                df_inv["PAIS"] = "VENEZUELA"
                df_inv["FECHA"] = hoy
                df_inv["NOMBRE_TIENDA/BODEGA"] = tienda_actual
                df_inv["MARCA"] = marca_actual
                df_inv["COSTO_TOTAL"] = df_inv["COSTO_UNITARIO"] * df_inv["UNIDADES_INVENTARIO"]

                orden_columnas = ["PAIS", "FECHA", "NOMBRE_TIENDA/BODEGA", "MARCA", "CODIGO_BARRA", "REFERENCIA", "DESCRIPCION", "COLOR", "TALLA", "DEPARTAMENTO", "PRECIO_NETO", "UNIDADES_INVENTARIO", "COSTO_UNITARIO", "COSTO_TOTAL"]

                df_inv = df_inv[orden_columnas]

                titulo = ruta_archivo.replace(".TXT", ".xlsx")

                with pd.ExcelWriter(titulo, engine="openpyxl") as writer:
                    df_inv.to_excel(writer, sheet_name="Hoja1", index=False)
                    print("Excel generado exitosamente")

                wb = load_workbook(titulo) #abre el libro
                ws = wb["Hoja1"] #abre la hoja

                columnas_texto = ["CODIGO_BARRA", "COLOR"]

                for columna in columnas_texto:

                    col_idx = None

                    for idx, cell in enumerate(ws[1], start=1):
                        if cell.value == columna:
                            col_idx = idx
                            break

                    if col_idx:
                        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                            for cell in row:
                                cell.number_format = "@"

                wb.save(titulo)
                print("Formato de texto aplicado a CODIGO_BARRA")

        #Terminado el proceso, cambiamos los TXT.
        ruta_destino = os.path.join(ruta_base, "Data", "Completados_txt")
        shutil.move(ruta_archivo, ruta_destino)
        print(f"Documento: {archivo} enviado a 'Completados_txt' exitosamente")

    except Exception as e:
        print(f"Fallo en archivo: {archivo}")
        print(f"Tipo de error: {type(e).__name__}")
        print(f"Descripción: {e}")
        traceback.print_exc()  # Muestra la traza completa del error

next = os.path.join(ruta_base, "3-preparador.py")

print("2-generador.py ejecutador exitosamente\n\n")
subprocess.run(["python", next])
