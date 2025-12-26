import pandas as pd
import os

def cargar_clientes_excel(nombre_archivo="clientes.xlsx") -> pd.DataFrame:
    """
    Carga los datos de clientes desde un archivo CSV.

    Parámetros:
    ruta_archivo (str): La ruta al archivo CSV que contiene los datos de clientes.

    Retorna:
    pd.DataFrame: Un DataFrame de pandas con los datos de clientes.
    """
    try: # Intentar cargar el archivo clientes.xlsx

        """
        Carga clientes desde un archivo Excel ubicado en la carpeta raíz del proyecto.
        Retorna una lista de diccionarios.
        """
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) # Directorio base del proyecto
        ruta = os.path.join(base_dir, nombre_archivo) # Ruta completa al archivo clientes.xlsx

        if not os.path.exists(ruta):
            raise FileNotFoundError(f"No se encontró el archivo: {ruta}")

        df = pd.read_excel(ruta) # Cargar el archivo Excel

        # Normalizar nombres de columnas 
        df.columns = df.columns.str.strip().str.upper() # Eliminar espacios y convertir a mayúsculas

        clientes = []
        for _, fila in df.iterrows(): # Iterar sobre cada fila del DataFrame
            clientes.append(fila.to_dict()) # Convertir la fila a diccionario y agregar a la lista

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
    