from clientes import cargar_datos_clientes
from selector_clientes import preparar_opciones_clientes

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


