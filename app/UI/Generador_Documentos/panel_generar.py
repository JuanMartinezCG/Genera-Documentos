import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from Logic.clientes import cargar_clientes_excel
from Logic.word_generator import generar_documento_word
import os


class PanelGenerarDocumento(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        if not hasattr(self.parent, "carpeta_salida"):
            self.parent.carpeta_salida = None

        self.clientes = []
        self.documentos = []

        # NUEVO: selección
        self.cliente_seleccionado = None
        self.documentos_seleccionados = set()

        self._crear_interfaz()

    # --------------------------------------------------
    # CREAR INTERFAZ
    # --------------------------------------------------
    def _crear_interfaz(self):
        ttk.Label(self, text="Generar Documento", font=("Segoe UI", 16, "bold")).pack(pady=15)

        contenedor = ttk.Frame(self)
        contenedor.pack(fill="both", expand=True, padx=20, pady=10)

        contenedor.columnconfigure(0, weight=0)
        contenedor.columnconfigure(1, weight=1)
        contenedor.columnconfigure(2, weight=0)
        contenedor.rowconfigure(0, weight=1)

        col_clientes = ttk.Frame(contenedor, width=220)
        col_documentos = ttk.Frame(contenedor)
        col_acciones = ttk.Frame(contenedor, width=200)

        col_clientes.grid(row=0, column=0, sticky="ns", padx=(0, 15))
        col_documentos.grid(row=0, column=1, sticky="nsew", padx=(0, 15))
        col_acciones.grid(row=0, column=2, sticky="ns")

        col_clientes.grid_propagate(False)
        col_clientes.config(width=220)

        col_acciones.grid_propagate(False)
        col_acciones.config(width=200)

        # CLIENTES
        ttk.Label(col_clientes, text="Clientes", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        self.frame_clientes_scroll = ttk.Frame(col_clientes)
        self.frame_clientes_scroll.pack(fill="both", expand=True, pady=5)

        # DOCUMENTOS
        ttk.Label(col_documentos, text="Documentos", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        self.frame_documentos_scroll = ttk.Frame(col_documentos)
        self.frame_documentos_scroll.pack(fill="both", expand=True, pady=5)

        # ACCIONES
        ttk.Label(col_acciones, text="Acciones", font=("Segoe UI", 12, "bold")).pack(pady=5)
        frame_acciones = ttk.Frame(col_acciones)
        frame_acciones.pack(fill="both", pady=5)

        ttk.Button(frame_acciones, text="Generar Documento", command=self._generar_documento).pack(fill="x", pady=5)
        ttk.Button(frame_acciones, text="Refrescar", command=self._cargar_datos).pack(fill="x", pady=5)
        ttk.Button(frame_acciones, text="Elegir carpeta de salida", command=self._elegir_carpeta).pack(fill="x", pady=5)

        self.lbl_carpeta = ttk.Label(
            frame_acciones,
            text="Carpeta no seleccionada",
            wraplength=160,
            foreground="gray"
        )
        self.lbl_carpeta.pack(pady=5)

        if self.parent.carpeta_salida:
            self.lbl_carpeta.config(text=self.parent.carpeta_salida, foreground="black")

        self._cargar_datos()

    # --------------------------------------------------
    # CARGAR DATOS
    # --------------------------------------------------
    def _cargar_datos(self):
        for w in self.frame_clientes_scroll.winfo_children():
            w.destroy()
        for w in self.frame_documentos_scroll.winfo_children():
            w.destroy()

        self.clientes = cargar_clientes_excel()
        self._crear_lista_modular(self.frame_clientes_scroll, self.clientes, "cliente")

        carpeta_docs = os.path.join(os.getcwd(), "Documentos")
        os.makedirs(carpeta_docs, exist_ok=True)

        self.documentos = [f for f in os.listdir(carpeta_docs) if f.lower().endswith(".docx")]
        self._crear_lista_modular(self.frame_documentos_scroll, self.documentos, "documento")

    # --------------------------------------------------
    # LISTA MODULAR CON SELECCIÓN
    # --------------------------------------------------
    def _crear_lista_modular(self, parent_frame, items, tipo):
        canvas = tk.Canvas(parent_frame)
        scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        COLOR_1 = "#F0F0F0"
        COLOR_2 = "#FFFFFF"
        COLOR_SEL = "#CDE8FF"

        for idx, item in enumerate(items):
            if tipo == "cliente":
                seleccionado = self.cliente_seleccionado == idx
                nombre = item.get("NOMBRE_EMPRESA", "SIN NOMBRE")
                nit = item.get("NIT", "")
                texto = f"{nombre}\nNIT: {nit}"
                icono = "👤"
            else:
                seleccionado = idx in self.documentos_seleccionados
                texto = item
                icono = "📄"

            base_color = COLOR_SEL if seleccionado else (COLOR_1 if idx % 2 == 0 else COLOR_2)

            card = tk.Frame(scroll_frame, bg=base_color, relief="solid", bd=2 if seleccionado else 1)
            card.pack(fill="x", padx=5, pady=4)

            label = tk.Label(
                card,
                text=f"{icono} {texto}",
                bg=base_color,
                anchor="w",
                justify="left",
                padx=6,
                pady=6,
                wraplength=360 if tipo == "documento" else 200
            )
            label.pack(fill="both")

            # CLICK = SELECCIÓN
            if tipo == "cliente":
                card.bind("<Button-1>", lambda e, i=idx: self._seleccionar_cliente(i))
                label.bind("<Button-1>", lambda e, i=idx: self._seleccionar_cliente(i))

                card.bind("<Double-Button-1>", lambda e, i=idx: self._previsualizar_cliente_tarjeta(i))
                label.bind("<Double-Button-1>", lambda e, i=idx: self._previsualizar_cliente_tarjeta(i))
            else:
                card.bind("<Button-1>", lambda e, i=idx: self._toggle_documento(i))
                label.bind("<Button-1>", lambda e, i=idx: self._toggle_documento(i))

                card.bind("<Double-Button-1>", lambda e, i=idx: self._abrir_documento_tarjeta(i))
                label.bind("<Double-Button-1>", lambda e, i=idx: self._abrir_documento_tarjeta(i))

    # --------------------------------------------------
    # SELECCIÓN
    # --------------------------------------------------
    def _seleccionar_cliente(self, idx):
        self.cliente_seleccionado = idx
        self._cargar_datos()

    def _toggle_documento(self, idx):
        if idx in self.documentos_seleccionados:
            self.documentos_seleccionados.remove(idx)
        else:
            self.documentos_seleccionados.add(idx)
        self._cargar_datos()

    # --------------------------------------------------
    # PREVISUALIZAR CLIENTE
    # --------------------------------------------------
    def _previsualizar_cliente_tarjeta(self, idx):
        cliente = self.clientes[idx]
        ventana = tk.Toplevel(self)
        ventana.title("Vista previa del cliente")
        ventana.geometry("420x420")

        tree = ttk.Treeview(ventana, columns=("Campo", "Valor"), show="headings")
        tree.heading("Campo", text="Campo")
        tree.heading("Valor", text="Valor")
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        for k, v in cliente.items():
            tree.insert("", tk.END, values=(k, v))

    # --------------------------------------------------
    # ABRIR DOCUMENTO
    # --------------------------------------------------
    def _abrir_documento_tarjeta(self, idx):
        ruta = os.path.join(os.getcwd(), "Documentos", self.documentos[idx])
        if os.path.exists(ruta):
            os.startfile(ruta)

    # --------------------------------------------------
    # GENERAR DOCUMENTOS
    # --------------------------------------------------
    def _generar_documento(self):
        if self.cliente_seleccionado is None:
            messagebox.showwarning("Atención", "Seleccione un cliente.")
            return

        if not self.documentos_seleccionados:
            messagebox.showwarning("Atención", "Seleccione al menos un documento.")
            return

        if not self.parent.carpeta_salida or not os.path.exists(self.parent.carpeta_salida):
            messagebox.showwarning("Atención", "Seleccione una carpeta de salida válida.")
            return

        cliente = self.clientes[self.cliente_seleccionado]
        empresa = cliente.get("NOMBRE_EMPRESA", "cliente").replace(" ", "_")

        for idx in self.documentos_seleccionados:
            nombre_doc = self.documentos[idx]

            ruta_base = os.path.join(os.getcwd(), "Documentos", nombre_doc)

            empresa = cliente.get("NOMBRE_EMPRESA", "cliente").replace(" ", "_") # Parte del nombre de archivo
            salida = f"{empresa}_{nombre_doc}" # Nombre de archivo de salida

            ruta_salida = os.path.join(self.parent.carpeta_salida, salida) # Ruta completa de salida

            # confirmación de sobrescritura
            if os.path.exists(ruta_salida):
                respuesta = messagebox.askyesno(
                    "Archivo existente",
                    f"El archivo:\n\n{salida}\n\nya existe.\n¿Desea sobrescribirlo?"
                )
                if not respuesta:
                    continue  # Saltar este documento
                
            generar_documento_word(
                ruta_documento_base=ruta_base,
                ruta_salida=ruta_salida,
                datos_cliente=cliente
            )


        messagebox.showinfo(
            "Proceso completado",
            f"Se generaron {len(self.documentos_seleccionados)} documentos correctamente."
        )

    # --------------------------------------------------
    # CARPETA SALIDA
    # --------------------------------------------------
    def _elegir_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if carpeta:
            self.parent.carpeta_salida = carpeta
            self.lbl_carpeta.config(text=carpeta, foreground="black")
