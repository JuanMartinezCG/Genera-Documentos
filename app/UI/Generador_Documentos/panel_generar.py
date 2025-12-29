# UI/Generator_Documents/GenerarDoc/panel_generar.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from Logic.clientes import cargar_clientes_excel
from Logic.word_generator import generar_documento_word
import os
import json
from datetime import datetime
from pathlib import Path


class PanelGenerarDocumento(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        # Configuración persistente
        self.config_file = Path.home() / ".document_generator_config.json"
        self._cargar_configuracion()
        
        # Estilos básicos (simplificado - todo en un archivo)
        self._configurar_estilos_basicos()
        
        self.clientes = []
        self.documentos = []
        self.cliente_seleccionado = None
        self.documentos_seleccionados = set()
        
        self._crear_interfaz()
        self._configurar_bindings()
    
    def _configurar_estilos_basicos(self):
        """Configura estilos básicos sin archivo externo"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Colores base
        self.COLOR_PRIMARIO = "#2C3E50"
        self.COLOR_SECUNDARIO = "#3498DB"
        self.COLOR_EXITO = "#27AE60"
        self.COLOR_PELIGRO = "#E74C3C"
        self.COLOR_FONDO = "#F8F9FA"
        self.COLOR_TARJETA = "#FFFFFF"
        self.COLOR_TARJETA_SEL = "#E3F2FD"
        self.COLOR_BORDE = "#E0E0E0"
        self.COLOR_TEXTO = "#2C3E50"
        self.COLOR_HOVER = "#F5F5F5"
        
        # Configurar estilos básicos
        style.configure("TFrame", background=self.COLOR_FONDO)
        style.configure("TLabel", background=self.COLOR_FONDO)
        
        style.configure("Titulo.TLabel", 
                       font=("Segoe UI", 18, "bold"),
                       foreground=self.COLOR_PRIMARIO)
        
        style.configure("Subtitulo.TLabel",
                       font=("Segoe UI", 12, "bold"),
                       foreground=self.COLOR_SECUNDARIO)
        
        style.configure("Accion.TButton",
                       font=("Segoe UI", 10),
                       padding=8)
        
        style.configure("Primary.TButton",
                       font=("Segoe UI", 10, "bold"),
                       foreground="white",
                       background=self.COLOR_SECUNDARIO)
        
        style.configure("Card.TFrame",
                       background=self.COLOR_TARJETA,
                       relief="solid",
                       borderwidth=1)
    
    def _cargar_configuracion(self):
        """Carga configuración sin archivo externo"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    if hasattr(self.parent, 'carpeta_salida'):
                        self.parent.carpeta_salida = config.get('carpeta_salida')
                    else:
                        self.parent.carpeta_salida = None
        except:
            if hasattr(self.parent, 'carpeta_salida'):
                self.parent.carpeta_salida = None
            else:
                self.parent.carpeta_salida = None
    
    def _guardar_configuracion(self):
        """Guarda configuración sin archivo externo"""
        try:
            config = {'carpeta_salida': self.parent.carpeta_salida}
            with open(self.config_file, 'w') as f:
                json.dump(config, f)
        except:
            pass
    
    def _crear_interfaz(self):
        # Configurar fondo del frame principal
        self.config(style="TFrame")
        
        # Título principal
        frame_titulo = ttk.Frame(self)
        frame_titulo.pack(fill="x", padx=20, pady=(15, 10))
        
        ttk.Label(frame_titulo, 
                 text="📄 Generador de Documentos",
                 style="Titulo.TLabel").pack(side="left")
        
        # Info de estado
        self.lbl_estado = ttk.Label(frame_titulo, 
                                   text="Listo",
                                   foreground="gray",
                                   font=("Segoe UI", 9))
        self.lbl_estado.pack(side="right")
        
        # Contenedor principal
        contenedor = ttk.Frame(self)
        contenedor.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Configurar grid responsivo
        contenedor.columnconfigure(0, weight=1, minsize=300)
        contenedor.columnconfigure(1, weight=2)
        contenedor.columnconfigure(2, weight=0, minsize=250)
        contenedor.rowconfigure(0, weight=1)
        
        # PANEL CLIENTES
        frame_clientes = self._crear_panel_con_scroll(
            contenedor, "👥 Clientes", 0, 0, "clientes"
        )
        
        # PANEL DOCUMENTOS
        frame_documentos = self._crear_panel_con_scroll(
            contenedor, "📄 Documentos Disponibles", 0, 1, "documentos"
        )
        
        # PANEL ACCIONES
        frame_acciones = ttk.LabelFrame(contenedor, text="⚡ Acciones", padding=15)
        frame_acciones.grid(row=0, column=2, sticky="nsew", padx=(10, 0))
        
        self._crear_panel_acciones(frame_acciones)
        
        # Cargar datos
        self._cargar_datos()
        
        # Actualizar estado
        self._actualizar_estado()
    
    def _crear_panel_con_scroll(self, parent, titulo, row, column, tipo):
        """Crea un panel con scroll de forma consistente"""
        frame = ttk.LabelFrame(parent, text=titulo, padding=10)
        frame.grid(row=row, column=column, sticky="nsew", padx=(0, 10))
        
        # Configurar grid interno
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        
        # Contador
        contador = ttk.Label(frame, text="0 items")
        contador.grid(row=0, column=0, sticky="w", pady=(0, 10))
        
        # Frame para canvas y scrollbar
        frame_contenido = ttk.Frame(frame)
        frame_contenido.grid(row=1, column=0, sticky="nsew")
        
        # Canvas y scrollbar
        canvas = tk.Canvas(frame_contenido, 
                          bg=self.COLOR_FONDO, 
                          highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame_contenido, 
                                 orient="vertical", 
                                 command=canvas.yview)
        
        # Frame interno para los items
        inner_frame = ttk.Frame(canvas)
        inner_frame.bind(
            "<Configure>",
            lambda e, c=canvas: c.configure(scrollregion=c.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=inner_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Guardar referencias según el tipo
        if tipo == "clientes":
            self.frame_clientes_inner = inner_frame
            self.lbl_contador_clientes = contador
        else:
            self.frame_documentos_inner = inner_frame
            self.lbl_contador_docs = contador
        
        return frame
    
    def _crear_panel_acciones(self, parent):
        """Crea el panel de acciones"""
        # Botón principal de generación
        btn_generar = ttk.Button(parent,
                                text="🚀 Generar Documentos",
                                command=self._generar_documento,
                                style="Primary.TButton")
        btn_generar.pack(fill="x", pady=5)
        
        # Botones secundarios
        ttk.Button(parent,
                  text="🔄 Refrescar Datos",
                  command=self._cargar_datos).pack(fill="x", pady=5)
        
        ttk.Button(parent,
                  text="📁 Elegir Carpeta de Salida",
                  command=self._elegir_carpeta).pack(fill="x", pady=5)
        
        # Botón para abrir carpeta
        self.btn_abrir_carpeta = ttk.Button(parent,
                                           text="🔍 Abrir Carpeta",
                                           command=self._abrir_carpeta_salida,
                                           state="disabled")
        self.btn_abrir_carpeta.pack(fill="x", pady=5)
        
        # Separador
        ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=15)
        
        # Info de carpeta
        ttk.Label(parent,
                 text="Carpeta de salida:",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        
        self.lbl_carpeta = ttk.Label(parent,
                                    text="No seleccionada",
                                    wraplength=200,
                                    foreground="red",
                                    font=("Segoe UI", 9))
        self.lbl_carpeta.pack(anchor="w", pady=(5, 10))
        
        if self.parent.carpeta_salida:
            self.lbl_carpeta.config(text=self.parent.carpeta_salida, 
                                   foreground="green")
            self.btn_abrir_carpeta.config(state="normal")
        
        # Contador de selección
        self.lbl_seleccion = ttk.Label(parent,
                                      text="Cliente: Ninguno\nDocumentos: 0 seleccionados",
                                      font=("Segoe UI", 9))
        self.lbl_seleccion.pack(anchor="w", pady=10)
        
        # Botón para limpiar
        ttk.Button(parent,
                  text="🗑️ Limpiar Selección",
                  command=self._limpiar_seleccion).pack(fill="x", pady=5)
    
    def _configurar_bindings(self):
        """Configura los bindings globales"""
        # Atajos de teclado
        self.bind_all("<Control-r>", lambda e: self._cargar_datos())
        self.bind_all("<Control-g>", lambda e: self._generar_documento())
        self.bind_all("<Control-l>", lambda e: self._limpiar_seleccion())
        self.bind_all("<Control-o>", lambda e: self._elegir_carpeta())
    
    # --------------------------------------------------
    # CARGAR DATOS
    # --------------------------------------------------
    def _cargar_datos(self):
        try:
            self.lbl_estado.config(text="Cargando datos...", foreground="blue")
            self.update()
            
            # Limpiar contenido anterior
            for widget in self.frame_clientes_inner.winfo_children():
                widget.destroy()
            for widget in self.frame_documentos_inner.winfo_children():
                widget.destroy()
            
            # Cargar clientes
            self.clientes = cargar_clientes_excel()
            self._crear_items_clientes()
            
            # Cargar documentos
            carpeta_docs = os.path.join(os.getcwd(), "Documentos")
            os.makedirs(carpeta_docs, exist_ok=True)
            self.documentos = [f for f in os.listdir(carpeta_docs) 
                             if f.lower().endswith(".docx")]
            self._crear_items_documentos()
            
            # Actualizar contadores
            self.lbl_contador_clientes.config(
                text=f"{len(self.clientes)} cliente(s) cargados"
            )
            self.lbl_contador_docs.config(
                text=f"{len(self.documentos)} documento(s) disponibles"
            )
            
            self.lbl_estado.config(text="Datos cargados", 
                                  foreground=self.COLOR_EXITO)
            self._actualizar_estado()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar datos: {str(e)}")
            self.lbl_estado.config(text="Error al cargar datos", 
                                  foreground=self.COLOR_PELIGRO)
    
    # --------------------------------------------------
    # CREAR ITEMS CON CLICK DERECHO
    # --------------------------------------------------
    def _crear_items_clientes(self):
        """Crea los items de clientes"""
        for idx, cliente in enumerate(self.clientes):
            seleccionado = (self.cliente_seleccionado == idx)
            self._crear_item(cliente, idx, seleccionado, "cliente")
    
    def _crear_items_documentos(self):
        """Crea los items de documentos"""
        for idx, doc in enumerate(self.documentos):
            seleccionado = (idx in self.documentos_seleccionados)
            ruta_doc = os.path.join(os.getcwd(), "Documentos", doc)
            tamaño = self._obtener_tamaño_archivo(ruta_doc)
            
            datos_doc = {
                "nombre": doc,
                "tamaño": tamaño,
                "ruta": ruta_doc
            }
            self._crear_item(datos_doc, idx, seleccionado, "documento")
    
    def _crear_item(self, datos, idx, seleccionado, tipo):
        """Crea un item seleccionable"""
        # Determinar frame padre
        frame_padre = self.frame_clientes_inner if tipo == "cliente" else self.frame_documentos_inner
        
        # Crear tarjeta
        bg_color = self.COLOR_TARJETA_SEL if seleccionado else self.COLOR_TARJETA
        border_color = self.COLOR_SECUNDARIO if seleccionado else self.COLOR_BORDE
        
        frame = tk.Frame(frame_padre,
                        bg=bg_color,
                        relief="solid",
                        borderwidth=2 if seleccionado else 1,
                        highlightbackground=border_color,
                        highlightthickness=1)
        frame.pack(fill="x", padx=5, pady=3)
        
        # Configurar cursor
        frame.config(cursor="hand2")
        
        # Crear contenido
        if tipo == "cliente":
            icono = "✅" if seleccionado else "🏢"
            texto = (f"{icono} {datos.get('NOMBRE_EMPRESA', 'Sin nombre')}\n"
                    f"📋 NIT: {datos.get('NIT', 'Sin NIT')}\n"
                    f"👤 {datos.get('CONTACTO', 'Sin contacto') if datos.get('CONTACTO') else 'Sin contacto'}")
        else:
            icono = "✅" if seleccionado else "📄"
            texto = (f"{icono} {datos['nombre']}\n"
                    f"📊 Tamaño: {datos['tamaño']}")
        
        label = tk.Label(frame,
                        text=texto,
                        bg=bg_color,
                        anchor="w",
                        justify="left",
                        font=("Segoe UI", 9),
                        wraplength=280 if tipo == "cliente" else 450,
                        cursor="hand2")
        label.pack(fill="both", padx=10, pady=8)
        
        # Guardar datos para acceso rápido
        frame.datos = datos
        frame.idx = idx
        frame.tipo = tipo
        
        # BINDINGS - Click izquierdo = seleccionar
        frame.bind("<Button-1>", self._on_item_click)
        label.bind("<Button-1>", self._on_item_click)
        
        # BINDINGS - Click derecho = menú contextual
        frame.bind("<Button-3>", self._on_right_click)
        label.bind("<Button-3>", self._on_right_click)
        
        # Efecto hover
        frame.bind("<Enter>", lambda e, f=frame: self._aplicar_hover(f, True))
        frame.bind("<Leave>", lambda e, f=frame: self._aplicar_hover(f, False))
        label.bind("<Enter>", lambda e, f=frame: self._aplicar_hover(f, True))
        label.bind("<Leave>", lambda e, f=frame: self._aplicar_hover(f, False))
        
        return frame
    
    def _on_item_click(self, event):
        """Maneja el click izquierdo en un item"""
        widget = event.widget
        
        # Encontrar el frame padre si se hizo click en el label
        if isinstance(widget, tk.Label):
            widget = widget.master
        
        idx = widget.idx
        tipo = widget.tipo
        
        if tipo == "cliente":
            self._seleccionar_cliente(idx)
        else:
            self._toggle_documento(idx)
    
    def _on_right_click(self, event):
        """Muestra menú contextual con click derecho"""
        widget = event.widget
        
        # Encontrar el frame padre si se hizo click en el label
        if isinstance(widget, tk.Label):
            widget = widget.master
        
        # Crear menú contextual
        menu = tk.Menu(self, tearoff=0)
        
        if widget.tipo == "cliente":
            menu.add_command(label="👁️ Ver Detalles", 
                        command=lambda: self._mostrar_detalles_cliente(widget.idx))
            menu.add_command(label="📋 Copiar NIT", 
                        command=lambda: self._copiar_al_portapapeles(widget.datos.get('NIT', '')))
            menu.add_separator()
            menu.add_command(label="✅ Seleccionar", 
                        command=lambda: self._seleccionar_cliente(widget.idx))
        else:
            menu.add_command(label="📄 Abrir Documento", 
                        command=lambda: self._abrir_documento(widget.idx))
            menu.add_command(label="📋 Copiar Nombre", 
                        command=lambda: self._copiar_al_portapapeles(widget.datos['nombre']))
            menu.add_separator()
            menu.add_command(label="✅ Seleccionar/Deseleccionar", 
                        command=lambda: self._toggle_documento(widget.idx))
        
        # Mostrar menú en la posición del click
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
    
    def _aplicar_hover(self, frame, entrar):
        """Aplica efecto hover a la tarjeta"""
        if entrar:
            frame.config(bg=self.COLOR_HOVER)
            for widget in frame.winfo_children():
                widget.config(bg=self.COLOR_HOVER)
        else:
            # Restaurar color original
            current_bg = frame.cget("bg")
            if current_bg == self.COLOR_TARJETA_SEL:
                new_bg = self.COLOR_TARJETA_SEL
            else:
                new_bg = self.COLOR_TARJETA
            
            frame.config(bg=new_bg)
            for widget in frame.winfo_children():
                widget.config(bg=new_bg)
    
    # --------------------------------------------------
    # FUNCIONALIDADES PRINCIPALES
    # --------------------------------------------------
    def _seleccionar_cliente(self, idx):
        """Selecciona un cliente"""
        self.cliente_seleccionado = idx
        self._cargar_datos()
        self._actualizar_estado()
    
    def _toggle_documento(self, idx):
        """Alterna selección de documento"""
        if idx in self.documentos_seleccionados:
            self.documentos_seleccionados.remove(idx)
        else:
            self.documentos_seleccionados.add(idx)
        self._cargar_datos()
        self._actualizar_estado()
    
    # --------------------------------------------------
    # MOSTRAR DETALLES CLIENTES
    # --------------------------------------------------
        # --------------------------------------------------
    # MOSTRAR DETALLES CLIENTES
    # --------------------------------------------------
    def _mostrar_detalles_cliente(self, idx):
        """Muestra detalles del cliente en ventana emergente"""
        cliente = self.clientes[idx]

        ventana = tk.Toplevel(self)
        ventana.title(f"Detalles - {cliente.get('NOMBRE_EMPRESA', 'Cliente')}")
        ventana.geometry("550x600")
        ventana.transient(self)
        ventana.grab_set()

        # Frame principal
        main_frame = ttk.Frame(ventana, padding=20)
        main_frame.pack(fill="both", expand=True)

        # Frame para título y botón cerrar
        frame_titulo = ttk.Frame(main_frame)
        frame_titulo.pack(fill="x", pady=(0, 20))

        # Título
        ttk.Label(frame_titulo,
                text=f"🏢 {cliente.get('NOMBRE_EMPRESA', 'Sin nombre')}",
                font=("Segoe UI", 14, "bold")).pack(side="left")

        # Botón cerrar al lado del título
        ttk.Button(frame_titulo,
                text="✕ Cerrar",
                command=ventana.destroy,
                style="Accion.TButton").pack(side="right", padx=(10, 0))

        # Frame para controles adicionales
        frame_controles = ttk.Frame(main_frame)
        frame_controles.pack(fill="x", pady=(0, 10))

        # Frame para treeview con scrollbar
        frame_tree = ttk.Frame(main_frame)
        frame_tree.pack(fill="both", expand=True)

        # Treeview para mostrar datos
        tree = ttk.Treeview(frame_tree, columns=("Campo", "Valor"), 
                        show="headings", height=20)
        tree.heading("Campo", text="Campo")
        tree.heading("Valor", text="Valor")
        tree.column("Campo", width=200)
        tree.column("Valor", width=250)

        # Scrollbars vertical y horizontal
        v_scrollbar = ttk.Scrollbar(frame_tree, orient="vertical", command=tree.yview)
        h_scrollbar = ttk.Scrollbar(frame_tree, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout para treeview y scrollbars
        tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        frame_tree.grid_rowconfigure(0, weight=1)
        frame_tree.grid_columnconfigure(0, weight=1)

        # Insertar datos
        for key, value in cliente.items():
            tree.insert("", "end", values=(key, str(value)))

        # Menú contextual para el treeview
        menu_contextual = tk.Menu(tree, tearoff=0)
        menu_contextual.add_command(label="📋 Copiar campo", 
                                command=lambda: self._copiar_item_treeview(tree, "campo"))
        menu_contextual.add_command(label="📋 Copiar valor", 
                                command=lambda: self._copiar_item_treeview(tree, "valor"))
        menu_contextual.add_command(label="📋 Copiar ambos", 
                                command=lambda: self._copiar_item_treeview(tree, "ambos"))
    
        def mostrar_menu_contextual(event):
            """Muestra el menú contextual en la posición del click"""
            item = tree.identify_row(event.y)
            if item:
                tree.selection_set(item)
                try:
                    menu_contextual.tk_popup(event.x_root, event.y_root)
                finally:
                    menu_contextual.grab_release()

        # Bindings para el treeview
        tree.bind("<Button-3>", mostrar_menu_contextual)  # Click derecho
        tree.bind("<Double-Button-1>", lambda e: self._copiar_item_treeview(tree, "valor"))  # Doble click
        tree.bind("<Control-c>", lambda e: self._copiar_item_treeview(tree, "ambos"))  # Ctrl+C

        # Etiqueta de ayuda
        ttk.Label(main_frame,
                text="💡 Consejo: Haz clic derecho para opciones de copia | Doble clic para copiar valor",
                font=("Segoe UI", 8),
                foreground="gray").pack(pady=(10, 0))

    def _copiar_item_treeview(self, treeview, tipo="valor"):
        """Copia el item seleccionado del treeview"""
        seleccion = treeview.selection()
        if not seleccion:
            return
        
        item = seleccion[0]
        valores = treeview.item(item, "values")
        
        if not valores or len(valores) < 2:
            return
        
        campo, valor = valores[0], valores[1]
        
        if tipo == "campo":
            texto = campo
        elif tipo == "valor":
            texto = valor
        elif tipo == "ambos":
            texto = f"{campo}: {valor}"
        else:
            texto = valor
        
        self.clipboard_clear()
        self.clipboard_append(str(texto))
        
        # Feedback visual
        for i in range(2):  # Parpadeo rápido
            treeview.item(item, tags=("copiado",))
            treeview.update()
            self.after(100)
            treeview.item(item, tags=())
            treeview.update()
            if i == 0:
                self.after(100)
    
    # --------------------------------------------------
    # ACCIONES VARIAS
    # --------------------------------------------------
    def _abrir_documento(self, idx):
        """Abre el documento con el programa predeterminado"""
        ruta = os.path.join(os.getcwd(), "Documentos", self.documentos[idx])
        if os.path.exists(ruta):
            try:
                os.startfile(ruta)
                self.lbl_estado.config(text=f"Abriendo documento...", 
                                    foreground="blue")
                self.after(2000, lambda: self.lbl_estado.config(text="Listo", foreground="gray"))
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir: {str(e)}")
        else:
            messagebox.showerror("Error", "Documento no encontrado")
    
    def _copiar_al_portapapeles(self, texto):
        """Copia texto al portapapeles"""
        self.clipboard_clear()
        self.clipboard_append(str(texto))
        self.lbl_estado.config(text="Texto copiado", foreground="blue")
        self.after(2000, lambda: self.lbl_estado.config(text="Listo", foreground="gray"))
    
    def _generar_documento(self):
        """Genera documentos"""
        if self.cliente_seleccionado is None:
            messagebox.showwarning("Atención", "Seleccione un cliente.")
            return
        
        if not self.documentos_seleccionados:
            messagebox.showwarning("Atención", "Seleccione al menos un documento.")
            return
        
        if not self.parent.carpeta_salida:
            messagebox.showwarning("Atención", "Seleccione una carpeta de salida.")
            return
        
        cliente = self.clientes[self.cliente_seleccionado]
        empresa = cliente.get("NOMBRE_EMPRESA", "cliente").replace(" ", "_")
        
        exitosos = 0
        errores = 0
        
        for idx in self.documentos_seleccionados:
            doc = self.documentos[idx]
            ruta_base = os.path.join(os.getcwd(), "Documentos", doc)
            
            # Nombre único con timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_salida = f"{empresa}_{doc.replace('.docx', '')}_{timestamp}.docx"
            ruta_salida = os.path.join(self.parent.carpeta_salida, nombre_salida)
            
            try:
                generar_documento_word(ruta_base, ruta_salida, cliente)
                exitosos += 1
            except Exception as e:
                errores += 1
                messagebox.showerror("Error", f"Error con {doc}: {str(e)}")
        
        # Mostrar resumen
        mensaje = f"""
        Proceso completado:
        
        ✅ Documentos exitosos: {exitosos}
        ❌ Documentos con error: {errores}
        📁 Carpeta: {self.parent.carpeta_salida}
        """
        
        if exitosos > 0:
            respuesta = messagebox.askyesno("Éxito", f"{mensaje}\n\n¿Abrir carpeta?")
            if respuesta:
                self._abrir_carpeta_salida()
        
        self.lbl_estado.config(text="Generación completada", 
                              foreground=self.COLOR_EXITO)
    
    def _elegir_carpeta(self):
        """Elige carpeta de salida (persistente)"""
        carpeta = filedialog.askdirectory(
            title="Seleccionar carpeta de salida",
            initialdir=self.parent.carpeta_salida
        )
        
        if carpeta:
            self.parent.carpeta_salida = carpeta
            self._guardar_configuracion()
            self.lbl_carpeta.config(text=carpeta, foreground="green")
            self.btn_abrir_carpeta.config(state="normal")
            self._actualizar_estado()
    
    def _abrir_carpeta_salida(self):
        """Abre la carpeta de salida"""
        if self.parent.carpeta_salida and os.path.exists(self.parent.carpeta_salida):
            os.startfile(self.parent.carpeta_salida)
    
    def _limpiar_seleccion(self):
        """Limpia todas las selecciones"""
        self.cliente_seleccionado = None
        self.documentos_seleccionados.clear()
        self._cargar_datos()
        self._actualizar_estado()
    
    def _actualizar_estado(self):
        """Actualiza el estado de selección"""
        cliente_text = "Ninguno"
        if self.cliente_seleccionado is not None and self.clientes:
            cliente = self.clientes[self.cliente_seleccionado]
            cliente_text = cliente.get("NOMBRE_EMPRESA", "Cliente")[:30]
            if len(cliente_text) == 30:
                cliente_text += "..."
        
        self.lbl_seleccion.config(
            text=f"Cliente: {cliente_text}\n"
                f"Documentos: {len(self.documentos_seleccionados)} seleccionados"
        )
    
    def _obtener_tamaño_archivo(self, ruta):
        """Obtiene el tamaño legible de un archivo"""
        if os.path.exists(ruta):
            size = os.path.getsize(ruta)
            if size < 1024:
                return f"{size} B"
            elif size < 1024 * 1024:
                return f"{size/1024:.1f} KB"
            else:
                return f"{size/(1024*1024):.1f} MB"
        return "Desconocido"