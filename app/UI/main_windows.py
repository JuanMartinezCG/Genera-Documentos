import tkinter as tk
from tkinter import ttk
from UI.Generador_Documentos.panel_generar import PanelGenerarDocumento


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        self.carpeta_salida = None # Carpeta donde se guardarán los documentos generados
        
        # Configuración básica
        self.title("Sistema de Generación de Documentos")
        self.geometry("1100x650") # Tamaño inicial de la ventana
        self.minsize(1000, 600) # Tamaño mínimo de la ventana

        # Layout principal
        self._crear_layout()

    def _crear_layout(self):
        # Contenedor principal
        self.container = ttk.Frame(self) # Contenedor principal
        self.container.pack(fill="both", expand=True) # Expandir para llenar toda la ventana

        # Panel lateral
        self.sidebar = ttk.Frame(self.container, width=220) # Ancho fijo para la barra lateral
        self.sidebar.pack(side="left", fill="y") # Llenar verticalmente

        # Panel de contenido
        self.content = ttk.Frame(self.container) # Panel de contenido principal
        self.content.pack(side="right", fill="both", expand=True) # Expandir para llenar el espacio restante

        self._crear_sidebar()
        self._mostrar_inicio()

    def _crear_sidebar(self): # Crear los botones en la barra lateral
        ttk.Label(
            self.sidebar,
            text="📄 MENÚ",
            font=("Segoe UI", 14, "bold")
        ).pack(pady=20)

        ttk.Separator(self.sidebar).pack(fill="x", pady=20)
        
        ttk.Button(
            self.sidebar,
            text="Inicio",
            command=self._mostrar_inicio
        ).pack(fill="x", padx=20, pady=5)

        ttk.Separator(self.sidebar).pack(fill="x", pady=20)

        ttk.Button(
            self.sidebar,
            text="Generar Documento",
            command=self._mostrar_generar
        ).pack(fill="x", padx=20, pady=5)

        ttk.Button(
            self.sidebar,
            text="Clientes",
            command=self._mostrar_clientes
        ).pack(fill="x", padx=20, pady=5)

        ttk.Button(
            self.sidebar,
            text="Plantillas",
            command=self._mostrar_documentos
        ).pack(fill="x", padx=20, pady=5)

        ttk.Separator(self.sidebar).pack(fill="x", pady=20)

        ttk.Button(
            self.sidebar,
            text="Salir",
            command=self.quit
        ).pack(fill="x", padx=20, pady=5)

    def _limpiar_contenido(self): # Limpiar el contenido del panel principal
        for widget in self.content.winfo_children():
            widget.destroy()

    def _mostrar_inicio(self): # Mostrar la pantalla de inicio
        self._limpiar_contenido()
        ttk.Label(
            self.content,
            text="Bienvenido al sistema de generación de documentos",
            font=("Segoe UI", 18)
        ).pack(pady=50)

    def _mostrar_generar(self):
        self._limpiar_contenido()
        panel = PanelGenerarDocumento(self.content)
        panel.pack(fill="both", expand=True)


    def _mostrar_clientes(self):
        self._limpiar_contenido()
        ttk.Label(
            self.content,
            text="Módulo: Gestión de Clientes",
            font=("Segoe UI", 16)
        ).pack(pady=20)

    def _mostrar_documentos(self):
        self._limpiar_contenido()
        ttk.Label(
            self.content,
            text="Módulo: Gestión de Plantillas",
            font=("Segoe UI", 16)
        ).pack(pady=20)
