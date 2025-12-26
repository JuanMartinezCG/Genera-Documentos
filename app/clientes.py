import pandas as pd
import os

def cargar_datos_clientes() -> pd.DataFrame:
    """
    Carga los datos de clientes desde un archivo CSV.

    Parámetros:
    ruta_archivo (str): La ruta al archivo CSV que contiene los datos de clientes.

    Retorna:
    pd.DataFrame: Un DataFrame de pandas con los datos de clientes.
    """
    try: # Intentar cargar el archivo clientes.xlsx

        # Ruta absoluta del archivo clientes.py
        base_dir = os.path.dirname(os.path.abspath(__file__))

        # clientes.xlsx debe estar en la misma carpeta
        ruta_archivo = os.path.join(base_dir, "clientes.xlsx")
        
        df = pd.read_excel(ruta_archivo) # Cargar el archivo Excel en un DataFrame de pandas

        clientes = [] 

        for _, fila in df.iterrows(): # Iterar sobre cada fila del DataFrame
            cliente = fila.to_dict() # Convertir la fila a un diccionario
            clientes.append(cliente) # Agregar el diccionario a la lista de clientes

        return clientes

    # Manejo de errores comunes
    except FileNotFoundError:
        print("Error: No se encontró el archivo clientes.xlsx en la carpeta.")
        return []

    except pd.errors.EmptyDataError:
        print("Error: El archivo clientes.xlsx está vacío.")
        return []

    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")
        return []
    