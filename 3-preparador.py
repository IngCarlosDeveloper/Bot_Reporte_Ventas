import os
import subprocess
import time
import pandas as pd
import re
from openpyxl import load_workbook
import traceback
import shutil
import fnmatch

#Ruta donde estan los .txt

ruta_base = os.path.dirname(os.path.abspath(__file__))

carpeta = os.path.join(ruta_base, "Data")

#Enlista cada .xlsx que hay en la ruta especificada
archivos_xlsx = [f for f in os.listdir(carpeta) if f.endswith(".xlsx")]

for archivo in archivos_xlsx:
    try:

        #Filtramos para RVM

        if fnmatch.fnmatch(archivo, "*RVM*"):

            #Identificar marca y tienda

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
        

            excel = os.path.join(carpeta, archivo)

            # 🔹 Forzar a que NUMERO_TRANSACCION y CODIGO_BARRA sean texto desde la lectura
            excel_cargar = pd.read_excel(
                excel,
                dtype={"NUMERO_TRANSACCION": str, "CODIGO_BARRA": str}
            ) #Inicializamos un data frame del excel

            last_index = excel_cargar["NUMERO_TRANSACCION"].last_valid_index()
            if last_index is not None:
                excel_cargar = excel_cargar.loc[:last_index]
            
            # 🔹 Conversión de columnas numéricas
            excel_cargar['VENTAS_BRUTAS_USD'] = (
                excel_cargar['VENTAS_BRUTAS_USD']
                .astype(str)
                .str.replace(",", ".", regex=False)
                .astype(float)
            )
            excel_cargar['PORCENTAJE_DESCUENTO'] = (
                excel_cargar['PORCENTAJE_DESCUENTO']
                .astype(str)
                .str.replace(",", ".", regex=False)
                .astype(float)
            ) 
            excel_cargar['COSTO_VENTA_USD'] = (
                excel_cargar['COSTO_VENTA_USD']
                .astype(str)
                .str.replace(",", ".", regex=False)
                .astype(float)
            ) 
            excel_cargar['UNIDADES_VENDIDAS'] = (
                excel_cargar['UNIDADES_VENDIDAS']
                .astype(str)
                .str.replace(",", ".", regex=False)
                .astype(float)
            )  # 🔹 Conversión agregada para evitar error en multiplicación

            #Creamos las columnas necesarias y les asignamos valores
            excel_cargar['PAIS'] = "VENEZUELA"
            excel_cargar['TIPO_TRANSACCION'] = "VENTAS"
            excel_cargar['NOMBRE_TIENDA/BODEGA'] = tienda_actual
            excel_cargar['MARCA'] = marca_actual
            excel_cargar['VENTAS_NETAS_USD'] = excel_cargar['VENTAS_BRUTAS_USD'] - (excel_cargar['VENTAS_BRUTAS_USD'] * (excel_cargar['PORCENTAJE_DESCUENTO'] / 100))
            excel_cargar['UTILIDAD_NETA_USD'] = excel_cargar['VENTAS_NETAS_USD'] - (excel_cargar['UNIDADES_VENDIDAS'] * excel_cargar['COSTO_VENTA_USD'])

            orden_columnas = [
                'PAIS',
                'FECHA_TRANSACCION',
                'TIPO_TRANSACCION',
                'NUMERO_TRANSACCION',
                'NOMBRE_TIENDA/BODEGA',
                'MARCA',
                'CODIGO_BARRA',
                'REFERENCIA',
                'DESCRIPCION',
                'COLOR',
                'TALLA',
                'DEPARTAMENTO',
                'UNIDADES_VENDIDAS',
                'PRECIO_UNITARIO',
                'VENTAS_BRUTAS_USD',
                'PORCENTAJE_DESCUENTO',
                'VENTAS_NETAS_USD',
                'UTILIDAD_NETA_USD',
                'COSTO_VENTA_USD'
            ]

            excel_cargar = excel_cargar[orden_columnas]

            #Guardar el DataGrame completo

            excel_cargar.to_excel(excel, index=False)
            print(f"Archivo {archivo} generado exitosamente")

            wb = load_workbook(excel) #abre el libro
            ws = wb["Sheet1"] #abre la hoja

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
                            #Forzar formato texto
                            if cell.value is not None:
                                cell.value = str(cell.value)
                            cell.number_format = "@"

            wb.save(excel)
            print("Formato de texto aplicado a NUMERO_TRANSACCION y CODIGO_BARRA")

    except Exception as e:
        print(f"Fall en archivo {archivo}")
        print(f"Tipo de error: {type(e).__name__}")
        print(f"Descripción: {e}")
        traceback.print_exc()  # Muestra la traza completa del error

next = os.path.join(ruta_base, "4-union.py")

print("3-preparador.py ejecutador exitosamente\n\n")
subprocess.run(["python", next])