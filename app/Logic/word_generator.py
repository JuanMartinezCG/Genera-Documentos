from docx import Document
from datetime import datetime
import os

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
    }

def generar_documento_word(datos_cliente: dict, nombre_salida: str):
    """Genera un documento de Word a partir de una plantilla y datos del cliente."""

    # Mapeo de meses en español
    fecha_actual = datetime.now()

    datos_cliente = datos_cliente.copy()
    datos_cliente["DIA"] = str(fecha_actual.day)
    datos_cliente["MES"] = MESES_ES[fecha_actual.month]
    datos_cliente["ANIO"] = str(fecha_actual.year)

    # Ruta de la plantilla de Word
    base_dir = os.path.dirname(os.path.abspath(__file__)) # Directorio base del script
    ruta_plantilla = os.path.join(base_dir, "plantilla.docx") # Ruta de la plantilla de Word

    doc = Document(ruta_plantilla) # Cargar la plantilla de Word

    def reemplazar_texto_en_parrafo(parrafo, datos):
        texto_completo = "".join(run.text for run in parrafo.runs) # Texto completo del párrafo

        for clave, valor in datos.items(): # Iterar sobre cada clave y valor en los datos del cliente
            marcador = f"{{{{{clave}}}}}" # Formatear el marcador a buscar
            if marcador in texto_completo: # Si el marcador está en el texto del párrafo
                texto_completo = texto_completo.replace(marcador, str(valor)) # Reemplazar el marcador con el valor correspondiente

        if texto_completo != "".join(run.text for run in parrafo.runs): # Si hubo un cambio en el texto
            for run in parrafo.runs: # Limpiar el texto de todos los runs
                run.text = "" # Vaciar el texto del run
            parrafo.runs[0].text = texto_completo # Asignar el texto completo al primer run


    # Reemplazo de marcadores en el documento
    for parrafo in doc.paragraphs: # Iterar sobre cada párrafo en el documento
        reemplazar_texto_en_parrafo(parrafo, datos_cliente)

    # Tablas
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    reemplazar_texto_en_parrafo(parrafo, datos_cliente)

    ruta_salida = os.path.join(base_dir, nombre_salida) # Ruta de salida del documento generado
    doc.save(ruta_salida) # Guardar el documento generado 