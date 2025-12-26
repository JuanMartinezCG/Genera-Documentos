from clientes import cargar_datos_clientes
from selector_clientes import preparar_opciones_clientes
from word_generator import generar_documento_word

clientes = cargar_datos_clientes()

if not clientes:
    print("No se cargaron clientes.")
else:
    print("==============================")
    print(clientes[0]["NOMBRE_EMPRESA"])
    print(clientes[0]["NIT"])
    print("==============================")


opciones = preparar_opciones_clientes(clientes)

if not opciones:
    print("No hay clientes para mostrar.")
else:
    for opcion in opciones:
        print("==============================")
        print(opcion)
        print("==============================")


# Tomamos el primer cliente como prueba
cliente = list(opciones.values())[0]

generar_documento_word(
    datos_cliente=cliente,
    nombre_salida="documento_generado.docx"
)

print("Documento generado correctamente.")


