import os
import pandas as pd

def obtener_nombre_mes_y_anio(nombre_archivo):
    """
    Convierte el nombre del archivo en un formato de mes y año legible.
    Ejemplo: SF-ab2023.xls -> abril 2023

    :param nombre_archivo: Nombre del archivo.
    :return: Nombre del mes y año en formato legible.
    """
    meses = {
        "en": "enero", "fe": "febrero", "ma": "marzo", "ab": "abril",
        "my": "mayo", "jn": "junio", "jl": "julio", "ag": "agosto",
        "se": "septiembre", "oc": "octubre", "no": "noviembre", "di": "diciembre"
    }
    partes = nombre_archivo.split("-")
    mes_codigo = partes[1][:2]
    anio = partes[1][2:6]
    return f"{meses.get(mes_codigo, 'mes desconocido')} {anio}"

def consolidar_depositos(carpeta_fuente, archivo_salida):
    """
    Itera sobre los archivos Excel en la carpeta fuente, extrae la información de depósitos
    de la hoja "Ctas BM" y consolida los datos en un único archivo de salida en formato horizontal.

    :param carpeta_fuente: Carpeta donde están los archivos Excel.
    :param archivo_salida: Nombre del archivo consolidado de salida.
    """
    # Diccionario para almacenar los datos consolidados
    datos_consolidados = {}

    # Iterar sobre los archivos en la carpeta fuente
    for archivo in os.listdir(carpeta_fuente):
        if archivo.endswith(".xls"):
            ruta_archivo = os.path.join(carpeta_fuente, archivo)

            try:
                # Leer la hoja "Ctas BM"
                print(f"Procesando archivo: {ruta_archivo}")
                hoja_ctas_bm = pd.read_excel(ruta_archivo, sheet_name="Ctas BM", header=0, engine="xlrd")

                # Buscar la sección de "Depósitos totales"
                if hoja_ctas_bm.iloc[:, 0].str.contains("Depósitos totales", na=False).any():
                    indice_depositos = hoja_ctas_bm[hoja_ctas_bm.iloc[:, 0].str.contains("Depósitos totales", na=False)].index[0]

                    # Extraer la fila relevante (columna 5 que contiene los datos requeridos)
                    datos_deposito = hoja_ctas_bm.iloc[indice_depositos, 5]

                    # Usar el mes y año como encabezado
                    mes_y_anio = obtener_nombre_mes_y_anio(archivo)
                    datos_consolidados[mes_y_anio] = [datos_deposito]
                else:
                    print(f"La sección 'Depósitos totales' no se encontró en el archivo: {archivo}")

            except Exception as e:
                print(f"Error procesando el archivo {archivo}: {e}")

    # Crear un DataFrame consolidado con los datos en formato horizontal
    if datos_consolidados:
        df_consolidado = pd.DataFrame(datos_consolidados)

        # Guardar el DataFrame consolidado en un archivo Excel
        df_consolidado.to_excel(archivo_salida, index=False)
        print(f"Archivo consolidado guardado en: {archivo_salida}")
    else:
        print("No se encontraron datos para consolidar.")

# Configuración
carpeta_fuente = r"C:\Users\José Estrada\OneDrive - ABC Capital\Escritorio\Sistema Financiero\SBS\Data"
archivo_salida = r"C:\Users\José Estrada\OneDrive - ABC Capital\Escritorio\Sistema Financiero\SBS\Consolidado_Depositos.xlsx"

# Ejecutar la función
consolidar_depositos(carpeta_fuente, archivo_salida)