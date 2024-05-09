import openpyxl
import pandas as pd
import os

def procesar_archivo(ruta_archivo):
    # Cargar el archivo Excel en un DataFrame
    df = pd.read_excel(ruta_archivo, header=None)  # Sin cabeceras

    # Encontrar la última fila con datos en la segunda columna
    ultima_fila = df[1].last_valid_index()

    # Si no hay datos en la segunda columna, no hacemos nada con este archivo
    if ultima_fila is None:
        return None

    # Obtener los datos desde la celda A2 hasta la última fila con datos en la segunda columna,
    # y seleccionar solo las columnas 0 A 11
    datos = df.loc[4:ultima_fila, 0:11]  # Columnas de 0 a 11
    # datos.columns = ["Proveedor", "Fecha", "Fecha Comprobante", "Comprobante", "Condicion", "B/S", "Codigo", "Cuenta", "Obra", "No gravados",
    #                  "No gravado C Comp", "Imp. Interno", "Neto Gravado", "Tasa", "IVA", "Perc. IIBB", "Perc. IVA", "Perc. SeH", "Total"]
    datos.columns = ['Fecha', 'Comprobante', 'Obra', 'Cliente', 'Condicion', 'CUIT', 'Gravado', 'No gravado', 'Tasa', 'IVA', 'Total']

    # Crear la nueva columna 'Tipo de comprobante'
    datos.insert(2, 'Tipo de comprobante', datos['Comprobante'].str[:4])

    # Convertir la columna de fecha al formato deseado
    datos['Fecha'] = pd.to_datetime(datos['Fecha']).dt.strftime('%d/%m/%Y')
    # datos['Fecha Comprobante'] = pd.to_datetime(datos['Fecha Comprobante']).dt.strftime('%d/%m/%Y')

    # Reemplazar los valores vacíos con ceros
    datos = datos.infer_objects(copy=False)
    datos = datos.fillna(0)

    return datos


def main(carpeta):
    # Crear un DataFrame vacío para almacenar los datos acumulativos
    datos_acumulativos = pd.DataFrame()

    # Obtener la lista de archivos en la carpeta especificada
    archivos = os.listdir(carpeta)

    # Iterar sobre cada archivo en la carpeta
    for archivo in archivos:
        # Comprobar si el archivo es un archivo Excel
        if archivo.endswith('.xlsx'):
            # Construir la ruta completa del archivo
            ruta_archivo = os.path.join(carpeta, archivo)

            # Procesar el archivo
            datos_nuevos = procesar_archivo(ruta_archivo)

            # Si se obtuvieron datos del archivo, agregarlos al DataFrame acumulativo
            if datos_nuevos is not None:
                datos_acumulativos = pd.concat([datos_acumulativos, datos_nuevos], ignore_index=True)

    # Guardar los datos acumulativos en un nuevo archivo Excel
    datos_acumulativos.to_excel('datos_acumulativos_vtas.xlsx', index=False)
    print("Datos acumulativos guardados correctamente.")


if __name__ == "__main__":
    carpeta = r"" #INGRESAR RUTA DEL ARCHIVO
    main(carpeta)