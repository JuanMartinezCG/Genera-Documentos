from clientes import cargar_datos_clientes

clientes = cargar_datos_clientes()

if not clientes:
    print("No se cargaron clientes.")
else:
    print(clientes[0]["NOMBRE_EMPRESA"])
    print(clientes[0]["NIT"])
