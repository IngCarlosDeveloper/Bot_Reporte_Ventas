import os
import sys
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

carpeta = os.path.join(ruta_base, "Data")

#Enlista cada .xlsx que hay en la ruta especificada
archivos_xlsx = [f for f in os.listdir(carpeta) if f.endswith(".xlsx") and not f.startswith("~$")]

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

for archivo in archivos_xlsx: #Bucle por cada excel que hay
    try:

        rvm_actual = 0
        tienda_actual = 0
        inventario_actual = 0

        if fnmatch.fnmatch(archivo, "*RVM*"): #Partimos siempre por un Reporte de ventas

            rvm_actual = archivo
            print(f"RVM encontrado: {rvm_actual}")

            for tienda in lista_tienda:
                if fnmatch.fnmatch(rvm_actual, f"*{tienda}*"): #Identificamos la tienda del RVM
                    tienda_actual = tienda
                    print(f"Tienda Identificada: {tienda_actual}")
                    break

            for sub_archivo in archivos_xlsx: #Repetimos el bucle, queremos buscar el inventario del RVM actual

                #Debe ser tipo Inventario  y debe ser de la tienda actual
                if fnmatch.fnmatch(sub_archivo, "*INVF*") and fnmatch.fnmatch(sub_archivo, f"*{tienda_actual}*"): 

                    inventario_actual = sub_archivo
                    print(f"INVF encontrado: {inventario_actual}")

            if rvm_actual == 0 or tienda_actual == 0 or inventario_actual == 0: #Dato Vacio = error
                print(f"Falta algun dato... Archivo actual {archivo}")
                print(f"RVM: {rvm_actual}\nTienda: {tienda_actual}\nInventario: {inventario_actual}")
                sys.exit()

            else: #Continuamos, todo bien

                #Unimos la ruta relativa con el nombre de los archivos
                rvm = os.path.join(carpeta, rvm_actual)
                invf = os.path.join(carpeta, inventario_actual)

                #Inicializamos los DataFrame
                df1 = pd.read_excel(rvm, dtype={"NUMERO_TRANSACCION": str, "CODIGO_BARRA": str})
                df2 = pd.read_excel(invf, dtype={"CODIGO_BARRA": str})

                #Depuramos y conservamos la integridad de los datos

                columnas_texto_1 = ["NUMERO_TRANSACCION", "CODIGO_BARRA", "COLOR"]
                columnas_texto_2 = ["FECHA", "CODIGO_BARRA", "COLOR"]

                columnas_comas_1= ["UNIDADES_VENDIDAS", "PRECIO_UNITARIO", "VENTAS_BRUTAS_USD", "PORCENTAJE_DESCUENTO", "VENTAS_NETAS_USD", "UTILIDAD_NETA_USD", "COSTO_VENTA_USD"]
                columnas_comas_2= ["PRECIO_NETO", "COSTO_UNITARIO", "COSTO_TOTAL"]

                def convertir_texto(columnas, df):

                    for columna in columnas:
                        df[columna] = df[columna].fillna("").astype(str)

                    print("Columnas convertidas a texto exitosamente")
                    return df
                
                def convertir_comas(columnas, df):

                    for columna in columnas:
                        df[columna] = df[columna].astype(str).replace(",",".", regex=True)
                        df[columna] = df[columna].astype(float)
                    
                    print("Comas transformadas a puntos exitosamente")
                    return df
                
                convertir_texto(columnas_texto_1, df1)
                convertir_texto(columnas_texto_2, df2)

                convertir_comas(columnas_comas_1, df1)
                convertir_comas(columnas_comas_2, df2)






                hoy = str(datetime.now().date())
                titulo = "REPORTE VENTAS MENSUALES " + tienda_actual + " " + hoy + ".xlsx"

                ruta_titulo = os.path.join(carpeta, titulo)

                with pd.ExcelWriter(ruta_titulo, engine="openpyxl") as writer:
                    df1.to_excel(writer, sheet_name="Hoja1", index=False)
                    df2.to_excel(writer, sheet_name="Hoja2", index=False)

                print(f"Archivos unidos exitosamente en {titulo} con 2 Hojas")

#------------------------------------------------CONVERTIR COLUMNAS A TEXTO--------------------------------------

                wb = load_workbook(ruta_titulo) #abre el libro
                ws_varias = [ wb["Hoja1"], wb["Hoja2"] ] #abre la hoja
            
                contador = 0
                columnas_texto = 0
                
                for ws in ws_varias:

                    if contador == 0:
                        columnas_texto = ["NUMERO_TRANSACCION", "CODIGO_BARRA", "COLOR"]
                    elif contador == 1:
                        columnas_texto = 0
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

                    contador += 1

                    print("Columnas en el excel convertidas a Texto exitosamente")



                wb.save(ruta_titulo)

                completados_xlsx = os.path.join(carpeta, "Completados_xlsx")
                reportes_finales = os.path.join(carpeta, "Reportes_Finales")

                shutil.move(rvm, completados_xlsx)
                shutil.move(invf, completados_xlsx)
                shutil.move(ruta_titulo, reportes_finales)




    except Exception as e:
        print(f"Fallo en archivo {archivo}")
        print(f"Tipo de error: {type(e).__name__}")
        print(f"Descripción: {e}")
        traceback.print_exc()  # Muestra la traza completa del error


print("4-union.py ejecutador exitosamente\n\n:D")


                    
