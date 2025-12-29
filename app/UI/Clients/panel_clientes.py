# UI/Clients/panel_clientes.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
from datetime import datetime
from pathlib import Path
from Logic.clientes import cargar_clientes_excel


class PanelClientes(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        # Configuración
        self.clientes = []
        self.cliente_editando = None
        self.archivo_clientes = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))) / "clientes.xlsx"
        
        # Estilos
        self._configurar_estilos()
        
        self._crear_interfaz()
        self._cargar_clientes()
    
    def _configurar_estilos(self):
        """Configura estilos para el panel de clientes"""
        self.COLOR_PRIMARIO = "#2C3E50"
        self.COLOR_SECUNDARIO = "#3498DB"
        self.COLOR_EXITO = "#27AE60"
        self.COLOR_PELIGRO = "#E74C3C"
        self.COLOR_ADVERTENCIA = "#F39C12"
        self.COLOR_FONDO = "#F8F9FA"
        self.COLOR_TARJETA = "#FFFFFF"
        self.COLOR_BORDE = "#E0E0E0"
        self.COLOR_HOVER = "#F5F5F5"
        
    def _crear_interfaz(self):
        """Crea la interfaz del panel de clientes"""
        # Configurar fondo
        self.configure(style="TFrame")
        
        # Header con título y botones
        frame_header = ttk.Frame(self)
        frame_header.pack(fill="x", padx=20, pady=(15, 10))
        
        ttk.Label(frame_header,
                 text="👥 Gestión de Clientes",
                 font=("Segoe UI", 18, "bold"),
                 foreground=self.COLOR_PRIMARIO).pack(side="left")
        
        # Botones de acción en el header
        frame_botones = ttk.Frame(frame_header)
        frame_botones.pack(side="right")
        
        ttk.Button(frame_botones,
                  text="➕ Nuevo Cliente",
                  command=self._mostrar_formulario_nuevo_cliente,  # Cambiado el nombre
                  style="Primary.TButton").pack(side="left", padx=5)
        
        ttk.Button(frame_botones,
                  text="📥 Importar Excel",
                  command=self._importar_excel).pack(side="left", padx=5)
        
        ttk.Button(frame_botones,
                  text="📤 Exportar Excel",
                  command=self._exportar_excel).pack(side="left", padx=5)
        
        # Barra de búsqueda
        frame_busqueda = ttk.Frame(self)
        frame_busqueda.pack(fill="x", padx=20, pady=(0, 15))
        
        ttk.Label(frame_busqueda,
                 text="🔍 Buscar:",
                 font=("Segoe UI", 10)).pack(side="left", padx=(0, 10))
        
        self.entry_busqueda = ttk.Entry(frame_busqueda, width=40, font=("Segoe UI", 10))
        self.entry_busqueda.pack(side="left")
        self.entry_busqueda.bind("<KeyRelease>", self._filtrar_clientes)
        
        ttk.Button(frame_busqueda,
                  text="Limpiar",
                  command=self._limpiar_busqueda).pack(side="left", padx=10)
        
        # Contenedor principal para el scroll
        contenedor = ttk.Frame(self)
        contenedor.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # Canvas y scrollbar
        self.canvas = tk.Canvas(contenedor, bg=self.COLOR_FONDO, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(contenedor, orient="vertical", command=self.canvas.yview)
        
        # Frame interno para las tarjetas
        self.frame_clientes = tk.Frame(self.canvas, bg=self.COLOR_FONDO)
        
        # Configurar scroll
        self.frame_clientes.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.frame_clientes, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Layout para canvas y scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Contador de clientes
        self.lbl_contador = ttk.Label(self, text="0 clientes cargados", font=("Segoe UI", 9))
        self.lbl_contador.pack(side="bottom", pady=(0, 10))
        
        # Bind para scroll con rueda del mouse
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
    
    def _on_mousewheel(self, event):
        """Maneja el scroll con la rueda del mouse"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def _cargar_clientes(self):
        """Carga los clientes usando tu función existente"""
        try:
            # USAR TU FUNCIÓN EXISTENTE
            self.clientes = cargar_clientes_excel()
            
            if not self.clientes:
                self._mostrar_mensaje_vacio()
                return
            
            self._actualizar_lista()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar clientes: {str(e)}")
            self.clientes = []
            self._mostrar_mensaje_vacio()
    
    def _mostrar_mensaje_vacio(self):
        """Muestra mensaje cuando no hay clientes"""
        for widget in self.frame_clientes.winfo_children():
            widget.destroy()
        
        ttk.Label(self.frame_clientes,
                 text="No hay clientes registrados.\nHaz clic en 'Nuevo Cliente' para agregar uno.",
                 font=("Segoe UI", 12),
                 foreground="gray",
                 background=self.COLOR_FONDO).pack(pady=50)
        self.lbl_contador.config(text="0 clientes cargados")
    
    def _actualizar_lista(self):
        """Actualiza la lista visual de clientes"""
        # Limpiar frame
        for widget in self.frame_clientes.winfo_children():
            widget.destroy()
        
        if not self.clientes:
            self._mostrar_mensaje_vacio()
            return
        
        # Organizar clientes en grid (2 columnas para más espacio)
        for i, cliente in enumerate(self.clientes):
            row = i // 2
            col = i % 2
            
            # Crear frame para la tarjeta
            frame_tarjeta = tk.Frame(
                self.frame_clientes,
                bg=self.COLOR_TARJETA,
                relief="solid",
                borderwidth=1,
                highlightbackground=self.COLOR_BORDE,
                highlightthickness=1,
                width=350,
                height=160
            )
            frame_tarjeta.grid(row=row, column=col, padx=15, pady=15, sticky="nsew")
            frame_tarjeta.grid_propagate(False)
            
            # Crear contenido de la tarjeta
            self._crear_tarjeta_cliente(frame_tarjeta, cliente, i)
        
        # Configurar columnas del grid
        for i in range(2):
            self.frame_clientes.grid_columnconfigure(i, weight=1, minsize=380)
        
        self.lbl_contador.config(text=f"{len(self.clientes)} cliente(s) cargados")
    
    def _crear_tarjeta_cliente(self, parent, cliente, idx):
        """Crea una tarjeta visual para un cliente"""
        try:
            parent.config(cursor="hand2")
            
            # Encabezado con nombre de empresa
            nombre = cliente.get('NOMBRE_EMPRESA', 'Sin nombre')
            nombre_display = nombre[:30] + "..." if len(nombre) > 30 else nombre
            
            frame_header = tk.Frame(parent, bg=self.COLOR_PRIMARIO)
            frame_header.pack(fill="x", pady=(0, 10))
            
            lbl_nombre = tk.Label(
                frame_header,
                text=f"🏢 {nombre_display}",
                font=("Segoe UI", 11, "bold"),
                bg=self.COLOR_PRIMARIO,
                fg="white",
                padx=12,
                pady=8
            )
            lbl_nombre.pack(side="left")
            
            # Información básica
            contenido_frame = tk.Frame(parent, bg=self.COLOR_TARJETA)
            contenido_frame.pack(fill="both", expand=True, padx=12, pady=(0, 10))
            
            # NIT
            nit = cliente.get('NIT', 'Sin NIT')
            nit_display = nit[:25] + "..." if len(nit) > 25 else nit
            
            frame_nit = tk.Frame(contenido_frame, bg=self.COLOR_TARJETA)
            frame_nit.pack(fill="x", pady=3)
            
            lbl_nit_icon = tk.Label(
                frame_nit,
                text="📋 NIT:",
                font=("Segoe UI", 9, "bold"),
                bg=self.COLOR_TARJETA,
                fg="#333333"
            )
            lbl_nit_icon.pack(side="left", padx=(0, 5))
            
            lbl_nit_valor = tk.Label(
                frame_nit,
                text=nit_display,
                font=("Segoe UI", 9),
                bg=self.COLOR_TARJETA,
                fg="#666666"
            )
            lbl_nit_valor.pack(side="left")
            
            # Contacto (si existe REPRESENTANTE_LEGAL)
            contacto = cliente.get('REPRESENTANTE_LEGAL', cliente.get('CONTACTO', 'Sin contacto'))
            contacto_display = contacto[:30] + "..." if len(contacto) > 30 else contacto
            
            frame_contacto = tk.Frame(contenido_frame, bg=self.COLOR_TARJETA)
            frame_contacto.pack(fill="x", pady=3)
            
            lbl_contacto_icon = tk.Label(
                frame_contacto,
                text="👤 Contacto:",
                font=("Segoe UI", 9, "bold"),
                bg=self.COLOR_TARJETA,
                fg="#333333"
            )
            lbl_contacto_icon.pack(side="left", padx=(0, 5))
            
            lbl_contacto_valor = tk.Label(
                frame_contacto,
                text=contacto_display,
                font=("Segoe UI", 9),
                bg=self.COLOR_TARJETA,
                fg="#666666"
            )
            lbl_contacto_valor.pack(side="left")
            
            # Ciudad
            ciudad = cliente.get('CIUDAD', '')
            if ciudad and not pd.isna(ciudad):
                frame_ciudad = tk.Frame(contenido_frame, bg=self.COLOR_TARJETA)
                frame_ciudad.pack(fill="x", pady=3)
                
                lbl_ciudad_icon = tk.Label(
                    frame_ciudad,
                    text="📍 Ciudad:",
                    font=("Segoe UI", 9, "bold"),
                    bg=self.COLOR_TARJETA,
                    fg="#333333"
                )
                lbl_ciudad_icon.pack(side="left", padx=(0, 5))
                
                lbl_ciudad_valor = tk.Label(
                    frame_ciudad,
                    text=str(ciudad),
                    font=("Segoe UI", 9),
                    bg=self.COLOR_TARJETA,
                    fg="#666666"
                )
                lbl_ciudad_valor.pack(side="left")
            
            # Botones de acción
            frame_botones = tk.Frame(parent, bg=self.COLOR_TARJETA)
            frame_botones.pack(fill="x", padx=12, pady=(0, 10))
            
            # Botón editar
            btn_editar = tk.Button(
                frame_botones,
                text="✏️ Editar",
                font=("Segoe UI", 9),
                bg=self.COLOR_SECUNDARIO,
                fg="white",
                borderwidth=0,
                padx=12,
                pady=4,
                cursor="hand2"
            )
            btn_editar.pack(side="left", padx=(0, 8))
            btn_editar.bind("<Button-1>", lambda e, i=idx: self._editar_cliente(i))
            
            # Botón eliminar
            btn_eliminar = tk.Button(
                frame_botones,
                text="🗑️ Eliminar",
                font=("Segoe UI", 9),
                bg=self.COLOR_PELIGRO,
                fg="white",
                borderwidth=0,
                padx=12,
                pady=4,
                cursor="hand2"
            )
            btn_eliminar.pack(side="left")
            btn_eliminar.bind("<Button-1>", lambda e, i=idx: self._confirmar_eliminar(i))
            
            # Función para mostrar detalles
            def mostrar_detalles(e=None):
                self._mostrar_detalles(idx)
            
            # BINDINGS SIMPLIFICADOS
            parent.bind("<Button-1>", mostrar_detalles)
            frame_header.bind("<Button-1>", mostrar_detalles)
            lbl_nombre.bind("<Button-1>", mostrar_detalles)
            contenido_frame.bind("<Button-1>", mostrar_detalles)
            frame_nit.bind("<Button-1>", mostrar_detalles)
            lbl_nit_icon.bind("<Button-1>", mostrar_detalles)
            lbl_nit_valor.bind("<Button-1>", mostrar_detalles)
            frame_contacto.bind("<Button-1>", mostrar_detalles)
            lbl_contacto_icon.bind("<Button-1>", mostrar_detalles)
            lbl_contacto_valor.bind("<Button-1>", mostrar_detalles)
            
            if 'frame_ciudad' in locals():
                frame_ciudad.bind("<Button-1>", mostrar_detalles)
                lbl_ciudad_icon.bind("<Button-1>", mostrar_detalles)
                lbl_ciudad_valor.bind("<Button-1>", mostrar_detalles)
            
            # Evitar que los botones activen mostrar_detalles
            btn_editar.bind("<Button-1>", lambda e: "break")  # Detener propagación
            btn_eliminar.bind("<Button-1>", lambda e: "break")  # Detener propagación
            frame_botones.bind("<Button-1>", lambda e: "break")  # Detener propagación
            
        except Exception as e:
            print(f"Error al crear tarjeta para cliente {idx}: {e}")
    
    def _mostrar_detalles(self, idx):
        """Muestra detalles completos del cliente"""
        try:
            if idx >= len(self.clientes):
                return
                
            cliente = self.clientes[idx]
            
            ventana = tk.Toplevel(self)
            ventana.title(f"Detalles - {cliente.get('NOMBRE_EMPRESA', 'Cliente')}")
            ventana.geometry("600x600")
            ventana.transient(self)
            ventana.grab_set()
            ventana.resizable(False, False)
            
            # Frame principal
            main_frame = ttk.Frame(ventana, padding=20)
            main_frame.pack(fill="both", expand=True)
            
            # Header
            frame_header = ttk.Frame(main_frame)
            frame_header.pack(fill="x", pady=(0, 20))
            
            ttk.Label(frame_header,
                     text=f"🏢 {cliente.get('NOMBRE_EMPRESA', 'Sin nombre')}",
                     font=("Segoe UI", 16, "bold"),
                     foreground=self.COLOR_PRIMARIO).pack(side="left")
            
            ttk.Button(frame_header,
                      text="✕ Cerrar",
                      command=ventana.destroy).pack(side="right")
            
            # Frame para botones de acción
            frame_acciones = ttk.Frame(main_frame)
            frame_acciones.pack(fill="x", pady=(0, 20))
            
            ttk.Button(frame_acciones,
                      text="✏️ Editar Cliente",
                      command=lambda: [ventana.destroy(), self._editar_cliente(idx)],
                      style="Primary.TButton").pack(side="left", padx=(0, 10))
            
            ttk.Button(frame_acciones,
                      text="🗑️ Eliminar Cliente",
                      command=lambda: [ventana.destroy(), self._confirmar_eliminar(idx)]).pack(side="left")
            
            # Frame scrollable para los campos
            frame_contenido = ttk.Frame(main_frame)
            frame_contenido.pack(fill="both", expand=True)
            
            # Canvas y scrollbar
            canvas = tk.Canvas(frame_contenido, bg="white", highlightthickness=0)
            scrollbar = ttk.Scrollbar(frame_contenido, orient="vertical", command=canvas.yview)
            
            frame_campos = ttk.Frame(canvas)
            frame_campos.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=frame_campos, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            # Mostrar campos importantes
            campos_importantes = [
                ('NIT', 'NIT'),
                ('CC', 'Cédula'),
                ('RUT', 'RUT'),
                ('DIRECCION', 'Dirección'),
                ('CIUDAD', 'Ciudad'),
                ('TELEFONO', 'Teléfono'),
                ('EMAIL', 'Email'),
                ('REPRESENTANTE_LEGAL', 'Representante Legal'),
                ('REP_LEGAL_DIRECCION', 'Dirección Representante'),
                ('ACTIVIDAD_ECONOMICA', 'Actividad Económica'),
                ('CAPITAL_REGISTRADO', 'Capital Registrado'),
                ('NOMBRE_BANCO', 'Banco'),
                ('NUMERO_CUENTA', 'Número de Cuenta'),
                ('TELEFONO_BANCO', 'Teléfono Banco'),
                ('NOMBRE_FIRMANTE', 'Firmante'),
                ('CEDULA_FIRMANTE', 'Cédula Firmante'),
            ]
            
            for campo_key, campo_label in campos_importantes:
                valor = cliente.get(campo_key, '')
                # Verificar si el valor existe y no es NaN
                if valor is not None and not pd.isna(valor) and str(valor).strip():
                    frame_campo = ttk.Frame(frame_campos)
                    frame_campo.pack(fill="x", pady=8, padx=5)
                    
                    ttk.Label(frame_campo,
                             text=f"{campo_label}:",
                             font=("Segoe UI", 10, "bold"),
                             foreground=self.COLOR_PRIMARIO).pack(anchor="w")
                    
                    valor_str = str(valor).strip()
                    ttk.Label(frame_campo,
                             text=valor_str,
                             font=("Segoe UI", 10),
                             wraplength=500,
                             foreground="#666666",
                             justify="left").pack(anchor="w", pady=(2, 0))
            
            # Si no hay campos para mostrar
            if not frame_campos.winfo_children():
                ttk.Label(frame_campos,
                         text="No hay información adicional disponible",
                         font=("Segoe UI", 11),
                         foreground="gray",
                         justify="center").pack(pady=50)
            
            # Layout para canvas y scrollbar
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
        except Exception as e:
            print(f"Error al mostrar detalles del cliente {idx}: {e}")
            messagebox.showerror("Error", f"No se pudieron mostrar los detalles: {str(e)}")
    
    def _mostrar_formulario_nuevo_cliente(self):
        """Muestra formulario para crear nuevo cliente (placeholder)"""
        messagebox.showinfo("Información", "Funcionalidad en desarrollo. Próximamente podrás agregar nuevos clientes.")
    
    def _editar_cliente(self, idx):
        """Abre formulario para editar cliente (placeholder)"""
        if idx < len(self.clientes):
            cliente = self.clientes[idx]
            messagebox.showinfo("Información", 
                              f"Funcionalidad en desarrollo.\nPróximamente podrás editar a:\n\n"
                              f"🏢 {cliente.get('NOMBRE_EMPRESA')}\n"
                              f"📋 NIT: {cliente.get('NIT')}")
    
    def _confirmar_eliminar(self, idx):
        """Confirma eliminación de cliente (placeholder)"""
        if idx < len(self.clientes):
            cliente = self.clientes[idx]
            messagebox.showinfo("Información", 
                              f"Funcionalidad en desarrollo.\nPróximamente podrás eliminar a:\n\n"
                              f"🏢 {cliente.get('NOMBRE_EMPRESA')}\n"
                              f"📋 NIT: {cliente.get('NIT')}")
    
    def _filtrar_clientes(self, event=None):
        """Filtra clientes según búsqueda"""
        texto = self.entry_busqueda.get().lower()
        
        if not texto:
            self._actualizar_lista()
            return
        
        # Filtrar clientes
        clientes_filtrados = []
        for cliente in self.clientes:
            if (texto in cliente.get('NOMBRE_EMPRESA', '').lower() or
                texto in cliente.get('NIT', '').lower() or
                texto in cliente.get('REPRESENTANTE_LEGAL', '').lower() or
                texto in cliente.get('EMAIL', '').lower()):
                clientes_filtrados.append(cliente)
        
        # Mostrar resultados filtrados
        self._mostrar_clientes_filtrados(clientes_filtrados)
    
    def _mostrar_clientes_filtrados(self, clientes_filtrados):
        """Muestra clientes filtrados"""
        # Limpiar frame
        for widget in self.frame_clientes.winfo_children():
            widget.destroy()
        
        if not clientes_filtrados:
            ttk.Label(self.frame_clientes,
                     text="No se encontraron clientes que coincidan con la búsqueda.",
                     font=("Segoe UI", 12),
                     foreground="gray",
                     background=self.COLOR_FONDO).pack(pady=50)
            self.lbl_contador.config(text="0 clientes encontrados")
            return
        
        # Mostrar clientes filtrados
        for i, cliente in enumerate(clientes_filtrados):
            row = i // 2
            col = i % 2
            
            frame_tarjeta = tk.Frame(
                self.frame_clientes,
                bg=self.COLOR_TARJETA,
                relief="solid",
                borderwidth=1,
                highlightbackground=self.COLOR_BORDE,
                highlightthickness=1,
                width=350,
                height=160
            )
            frame_tarjeta.grid(row=row, column=col, padx=15, pady=15, sticky="nsew")
            frame_tarjeta.grid_propagate(False)
            
            # Encontrar índice real del cliente
            idx = next((i for i, c in enumerate(self.clientes) 
                       if c.get('NIT') == cliente.get('NIT')), -1)
            
            self._crear_tarjeta_cliente(frame_tarjeta, cliente, idx)
        
        # Configurar columnas
        for i in range(2):
            self.frame_clientes.grid_columnconfigure(i, weight=1, minsize=380)
        
        self.lbl_contador.config(text=f"{len(clientes_filtrados)} cliente(s) encontrados")
    
    def _limpiar_busqueda(self):
        """Limpia la búsqueda y muestra todos los clientes"""
        self.entry_busqueda.delete(0, tk.END)
        self._actualizar_lista()
    
    def _importar_excel(self):
        """Importa clientes desde un archivo Excel"""
        messagebox.showinfo("Información", "Funcionalidad en desarrollo.")
    
    def _exportar_excel(self):
        """Exporta clientes a un archivo Excel"""
        messagebox.showinfo("Información", "Funcionalidad en desarrollo.")