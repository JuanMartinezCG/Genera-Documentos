def preparar_opciones_clientes(clientes: list) -> dict: # Prepara un diccionario de opciones para un selector de clientes.
    opciones = {}

    for cliente in clientes: # Iterar sobre cada cliente
        nombre = str(cliente.get("NOMBRE_EMPRESA", "")).strip() # Obtener y limpiar el nombre de la empresa
        nit = nit = str(cliente.get("NIT", "")).strip() # Obtener y limpiar el NIT

        texto_opcion = f"{nombre} – NIT {nit}" # Formatear el texto de la opción
        opciones[texto_opcion] = cliente # Agregar la opción al diccionario

    return opciones
