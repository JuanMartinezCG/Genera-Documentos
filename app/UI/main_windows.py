import tkinter as tk
from tkinter import ttk
from UI.Generador_Documentos.panel_generar import PanelGenerarDocumento
from UI.Clients import panel_clientes


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        self.carpeta_salida = None
        
        # Configuración básica
        self.title("📄 Sistema de Generación de Documentos")
        self.geometry("1200x700")
        self.minsize(1100, 650)
        
        # Centrar ventana
        self._centrar_ventana()
        
        # Configurar estilo
        self._configurar_estilos()
        
        # Variables de estado
        self.boton_activo = None
        self.panel_activo = None
        
        # Layout principal
        self._crear_layout()
        
        # Mostrar inicio por defecto
        self.after(100, self._mostrar_inicio)
    
    def _centrar_ventana(self):
        """Centra la ventana en la pantalla"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
    
    def _configurar_estilos(self):
        """Configura estilos personalizados"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Colores base
        self.COLOR_PRIMARIO = "#2C3E50"
        self.COLOR_SECUNDARIO = "#3498DB"
        self.COLOR_FONDO = "#F8F9FA"
        self.COLOR_SIDEBAR = "#2C3E50"
        self.COLOR_BOTON_ACTIVO = "#34495E"
        self.COLOR_BOTON_HOVER = "#3D566E"
        self.COLOR_TEXTO_BOTON = "#FFFFFF"
        
        # Estilo de la ventana
        self.configure(bg=self.COLOR_FONDO)
        
        # Estilos para widgets
        style.configure("TFrame", background=self.COLOR_FONDO)
        style.configure("Sidebar.TFrame", background=self.COLOR_SIDEBAR)
        style.configure("Content.TFrame", background=self.COLOR_FONDO)
        
        # Estilo para botones del sidebar
        style.configure("Sidebar.TButton",
                    background=self.COLOR_SIDEBAR,
                    foreground=self.COLOR_TEXTO_BOTON,
                    borderwidth=0,
                    focuscolor="none",
                    font=("Segoe UI", 10),
                    padding=(20, 12))
        
        style.map("Sidebar.TButton",
                background=[("active", self.COLOR_BOTON_HOVER),
                        ("selected", self.COLOR_BOTON_ACTIVO)])
        
        style.configure("Active.TButton",
                    background=self.COLOR_BOTON_ACTIVO,
                    foreground=self.COLOR_TEXTO_BOTON)
    
    def _crear_layout(self):
        # Configurar grid principal
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Sidebar (menú lateral)
        self.sidebar = ttk.Frame(self, width=260, style="Sidebar.TFrame")
        self.sidebar.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        self.sidebar.grid_propagate(False)  # Mantener ancho fijo
        
        # Área de contenido principal
        self.content = ttk.Frame(self, style="Content.TFrame")
        self.content.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        
        # Configurar expansión
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)
        
        self._crear_sidebar()
        self._crear_header()
    
    def _crear_header(self):
        """Crea el encabezado del contenido"""
        header_frame = ttk.Frame(self.content, style="Content.TFrame")
        header_frame.pack(fill="x", padx=30, pady=(30, 20))
        
        # Título dinámico
        self.lbl_titulo = ttk.Label(
            header_frame,
            text="Bienvenido",
            font=("Segoe UI", 24, "bold"),
            foreground=self.COLOR_PRIMARIO,
            background=self.COLOR_FONDO
        )
        self.lbl_titulo.pack(side="left")
        
        # Barra de estado (derecha)
        status_frame = ttk.Frame(header_frame, style="Content.TFrame")
        status_frame.pack(side="right")
        
        # Indicador de carpeta
        self.lbl_carpeta_estado = ttk.Label(
            status_frame,
            text="📁 Sin carpeta seleccionada",
            font=("Segoe UI", 9),
            foreground="#95A5A6",
            background=self.COLOR_FONDO
        )
        self.lbl_carpeta_estado.pack(side="right", padx=(10, 0))
        
        # Separador
        ttk.Separator(self.content, orient="horizontal").pack(fill="x", padx=30)
    
    def _crear_sidebar(self):
        """Crea el menú lateral mejorado"""
        
        # Logo/Header del sidebar
        header_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        header_frame.pack(fill="x", pady=(30, 20))
        
        ttk.Label(
            header_frame,
            text="📄",
            font=("Segoe UI", 28),
            background=self.COLOR_SIDEBAR,
            foreground="white"
        ).pack()
        
        ttk.Label(
            header_frame,
            text="DOCUMENTOS",
            font=("Segoe UI", 12, "bold"),
            background=self.COLOR_SIDEBAR,
            foreground="white"
        ).pack(pady=(5, 0))
        
        ttk.Label(
            header_frame,
            text="Generador Inteligente",
            font=("Segoe UI", 9),
            background=self.COLOR_SIDEBAR,
            foreground="#BDC3C7"
        ).pack()
        
        # Separador
        ttk.Separator(self.sidebar, orient="horizontal").pack(fill="x", padx=20, pady=20)
        
        # Botones del menú
        menu_items = [
            ("🏠", "Inicio", self._mostrar_inicio),
            ("⚡", "Generar Documentos", self._mostrar_generar),
            ("👥", "Clientes", self._mostrar_clientes),
            ("📋", "Plantillas", self._mostrar_documentos),
        ]
        
        for icono, texto, comando in menu_items:
            self._crear_boton_menu(icono, texto, comando)
        
        # Separador antes del botón salir
        ttk.Separator(self.sidebar, orient="horizontal").pack(fill="x", padx=20, pady=20)
        
        # Botón salir
        self._crear_boton_menu("🚪", "Salir", self.quit)
        
        # Espaciador para empujar todo hacia arriba
        ttk.Frame(self.sidebar, height=20, style="Sidebar.TFrame").pack(side="bottom")
        
        # Información de versión/estado
        footer_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        footer_frame.pack(side="bottom", fill="x", pady=20)
        
        ttk.Label(
            footer_frame,
            text="v1.0.0",
            font=("Segoe UI", 8),
            background=self.COLOR_SIDEBAR,
            foreground="#7F8C8D"
        ).pack()
    
    def _crear_boton_menu(self, icono, texto, comando):
        """Crea un botón del menú con estilo"""
        btn_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        btn_frame.pack(fill="x", padx=15, pady=2)
        
        # Botón personalizado
        btn = tk.Button(
            btn_frame,
            text=f"   {icono}  {texto}",
            command=lambda: self._seleccionar_menu(btn, comando),
            font=("Segoe UI", 10),
            bg=self.COLOR_SIDEBAR,
            fg=self.COLOR_TEXTO_BOTON,
            activebackground=self.COLOR_BOTON_HOVER,
            activeforeground="white",
            borderwidth=0,
            highlightthickness=0,
            anchor="w",
            cursor="hand2",
            padx=20,
            pady=12
        )
        btn.pack(fill="x")
        
        # Efecto hover
        btn.bind("<Enter>", lambda e, b=btn: self._aplicar_hover_boton(b, True))
        btn.bind("<Leave>", lambda e, b=btn: self._aplicar_hover_boton(b, False))
    
    def _seleccionar_menu(self, boton, comando):
        """Maneja la selección de un botón del menú"""
        # Restaurar estilo del botón anterior
        if self.boton_activo:
            self.boton_activo.config(bg=self.COLOR_SIDEBAR, fg=self.COLOR_TEXTO_BOTON)
        
        # Resaltar botón actual
        boton.config(bg=self.COLOR_BOTON_ACTIVO, fg="white")
        self.boton_activo = boton
        
        # Ejecutar comando
        comando()
    
    def _aplicar_hover_boton(self, boton, entrar):
        """Aplica efecto hover al botón"""
        if boton != self.boton_activo:
            if entrar:
                boton.config(bg=self.COLOR_BOTON_HOVER)
            else:
                boton.config(bg=self.COLOR_SIDEBAR)
    
    def _limpiar_contenido(self):
        """Limpia el contenido del panel principal"""
        if self.panel_activo:
            self.panel_activo.destroy()
            self.panel_activo = None
        
        # También limpiar cualquier otro widget en content
        for widget in self.content.winfo_children()[2:]:  # Saltar header y separador
            widget.destroy()
    
    def _mostrar_inicio(self):
        """Muestra la pantalla de inicio"""
        self._limpiar_contenido()
        self.lbl_titulo.config(text="Inicio")
        
        # Frame principal para contenido de inicio
        inicio_frame = ttk.Frame(self.content, style="Content.TFrame")
        inicio_frame.pack(fill="both", expand=True, padx=50, pady=30)
        
        # Tarjeta de bienvenida
        welcome_card = tk.Frame(
            inicio_frame,
            bg="white",
            relief="flat",
            highlightbackground="#E0E0E0",
            highlightthickness=1
        )
        welcome_card.pack(fill="x", pady=(0, 30))
        
        tk.Label(
            welcome_card,
            text="🚀 Bienvenido al Sistema",
            font=("Segoe UI", 22, "bold"),
            bg="white",
            fg=self.COLOR_PRIMARIO,
            padx=30,
            pady=30
        ).pack(anchor="w")
        
        tk.Label(
            welcome_card,
            text="Genera documentos profesionales de forma rápida y eficiente\nutilizando nuestras plantillas personalizables.",
            font=("Segoe UI", 11),
            bg="white",
            fg="#666666",
            justify="left",
            padx=30,
            pady=(0, 30)
        ).pack(anchor="w")
        
        # Grid de acciones rápidas
        acciones_frame = ttk.Frame(inicio_frame, style="Content.TFrame")
        acciones_frame.pack(fill="x", pady=(0, 30))
        
        acciones = [
            ("📄", "Generar Documentos", "Crea nuevos documentos personalizados", self._mostrar_generar),
            ("👥", "Gestionar Clientes", "Administra tu lista de clientes", self._mostrar_clientes),
            ("📋", "Ver Plantillas", "Explora plantillas disponibles", self._mostrar_documentos),
        ]
        
        for i, (icono, titulo, desc, comando) in enumerate(acciones):
            frame_accion = tk.Frame(
                acciones_frame,
                bg="white",
                relief="flat",
                highlightbackground="#E0E0E0",
                highlightthickness=1,
                cursor="hand2"
            )
            frame_accion.grid(row=0, column=i, padx=(0, 15) if i < 2 else 0, sticky="nsew")
            acciones_frame.columnconfigure(i, weight=1)
            
            # Contenido de la tarjeta
            tk.Label(
                frame_accion,
                text=icono,
                font=("Segoe UI", 24),
                bg="white",
                padx=20,
                pady=(20, 10)
            ).pack(anchor="w")
            
            tk.Label(
                frame_accion,
                text=titulo,
                font=("Segoe UI", 12, "bold"),
                bg="white",
                fg=self.COLOR_PRIMARIO,
                padx=20,
                pady=(0, 5)
            ).pack(anchor="w")
            
            tk.Label(
                frame_accion,
                text=desc,
                font=("Segoe UI", 9),
                bg="white",
                fg="#666666",
                wraplength=200,
                justify="left",
                padx=20
            ).pack(anchor="w", pady=(0, 20))
            
            # Bind para hacer click en toda la tarjeta
            frame_accion.bind("<Button-1>", lambda e, c=comando: c())
            for child in frame_accion.winfo_children():
                child.bind("<Button-1>", lambda e, c=comando: c())
        
        # Estadísticas/información
        info_frame = ttk.Frame(inicio_frame, style="Content.TFrame")
        info_frame.pack(fill="x")
        
        info_card = tk.Frame(
            info_frame,
            bg="#F8F9FA",
            relief="flat",
            highlightbackground="#D6DBDF",
            highlightthickness=1
        )
        info_card.pack(fill="x", pady=20)
        
        tk.Label(
            info_card,
            text="💡 Consejos Rápidos",
            font=("Segoe UI", 12, "bold"),
            bg="#F8F9FA",
            fg=self.COLOR_PRIMARIO,
            padx=20,
            pady=20
        ).pack(anchor="w")
        
        consejos = [
            "• Selecciona un cliente antes de generar documentos",
            "• Verifica que tienes la carpeta de salida configurada",
            "• Revisa las plantillas disponibles antes de comenzar",
            "• Usa Ctrl+R para refrescar la lista de clientes"
        ]
        
        for consejo in consejos:
            tk.Label(
                info_card,
                text=consejo,
                font=("Segoe UI", 9),
                bg="#F8F9FA",
                fg="#666666",
                justify="left",
                padx=20
            ).pack(anchor="w", pady=(0, 8)) 
    
    def _mostrar_generar(self):
        """Muestra el panel de generación de documentos"""
        self._limpiar_contenido()
        self.lbl_titulo.config(text="Generar Documentos")
        
        panel = PanelGenerarDocumento(self.content)
        panel.pack(fill="both", expand=True, padx=30, pady=(0, 30))
        self.panel_activo = panel
    
    # En tu MainWindow, en el método _mostrar_clientes:
    def _mostrar_clientes(self):
        """Muestra el módulo de gestión de clientes"""
        self._limpiar_contenido()
        self.lbl_titulo.config(text="Gestión de Clientes")

        panel = panel_clientes.PanelClientes(self.content)
        panel.pack(fill="both", expand=True, padx=30, pady=(0, 30))
        self.panel_activo = panel
    
    def _mostrar_documentos(self):
        """Muestra el módulo de plantillas (placeholder)"""
        self._limpiar_contenido()
        self.lbl_titulo.config(text="Gestión de Plantillas")
        
        contenido_frame = ttk.Frame(self.content, style="Content.TFrame")
        contenido_frame.pack(fill="both", expand=True, padx=50, pady=50)
        
        # Mensaje temporal
        tk.Label(
            contenido_frame,
            text="📋 Módulo en Desarrollo",
            font=("Segoe UI", 24, "bold"),
            bg=self.COLOR_FONDO,
            fg=self.COLOR_PRIMARIO
        ).pack(pady=20)
        
        tk.Label(
            contenido_frame,
            text="Esta funcionalidad estará disponible próximamente",
            font=("Segoe UI", 12),
            bg=self.COLOR_FONDO,
            fg="#666666"
        ).pack()
        
        # Botón para volver a inicio
        ttk.Button(
            contenido_frame,
            text="← Volver al Inicio",
            command=self._mostrar_inicio,
            style="Primary.TButton"
        ).pack(pady=30)
