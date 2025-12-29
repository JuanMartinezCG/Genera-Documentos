# Logic/clientes.py
import pandas as pd
import os

def cargar_clientes_excel(nombre_archivo="clientes.xlsx"):
    """
    Carga clientes desde un archivo Excel ubicado en la carpeta raíz del proyecto.
    Retorna una lista de diccionarios.
    """
    try:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) 
        ruta = os.path.join(base_dir, nombre_archivo)
        
        print(f"DEBUG: Buscando archivo en: {ruta}")
        print(f"DEBUG: ¿Existe el archivo? {os.path.exists(ruta)}")
        
        if not os.path.exists(ruta):
            print(f"DEBUG: Archivo no encontrado: {ruta}")
            return []
        
        # Leer el Excel
        df = pd.read_excel(ruta)
        print(f"DEBUG: DataFrame cargado. Forma: {df.shape}")
        print(f"DEBUG: Columnas originales: {list(df.columns)}")
        
        # Normalizar nombres de columnas
        df.columns = df.columns.str.strip().str.upper()
        print(f"DEBUG: Columnas normalizadas: {list(df.columns)}")
        
        # Convertir a lista de diccionarios
        clientes = []
        for i, fila in df.iterrows():
            cliente = fila.to_dict()
            print(f"DEBUG: Fila {i} -> {cliente.get('NOMBRE_EMPRESA', 'NO ENCONTRADO')}")
            clientes.append(cliente)
        
        print(f"DEBUG: Total clientes cargados: {len(clientes)}")
        return clientes

    except FileNotFoundError as e:
        print(f"Error: No se encontró el archivo clientes.xlsx: {e}")
        return []
    except pd.errors.EmptyDataError:
        print("Error: El archivo clientes.xlsx está vacío.")
        return []
    except Exception as e:
        print(f"Error inesperado al cargar clientes: {e}")
        return []