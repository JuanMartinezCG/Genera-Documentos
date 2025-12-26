import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from Logic.clientes import cargar_clientes_excel
from Logic.word_generator import generar_documento_word
from docx import Document
import os

class PanelGenerarDocumento(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        # Inicializar carpeta_salida si no existe
        if not hasattr(self.parent, "carpeta_salida"):
            self.parent.carpeta_salida = None

        self._crear_interfaz()

    def _crear_interfaz(self):
        # Título
        ttk.Label(self, text="Generar Documento", font=("Segoe UI", 16, "bold")).pack(pady=15)

        contenedor = ttk.Frame(self)
        contenedor.pack(fill="both", expand=True, padx=20, pady=10)

        # --- Columnas ---
        col_clientes = ttk.Frame(contenedor)
        col_documentos = ttk.Frame(contenedor)
        col_acciones = ttk.Frame(contenedor)

        col_clientes.pack(side="left", fill="both", expand=True, padx=10)
        col_documentos.pack(side="left", fill="both", expand=True, padx=10)
        col_acciones.pack(side="left", fill="y", padx=10)

        # ================= CLIENTES =================
        ttk.Label(col_clientes, text="Clientes", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        self.lista_clientes = tk.Listbox(col_clientes, height=20)
        self.lista_clientes.pack(fill="both", expand=True, pady=5)

        # ================= DOCUMENTOS =================
        ttk.Label(col_documentos, text="Documentos", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        self.lista_documentos = tk.Listbox(col_documentos, height=20)
        self.lista_documentos.pack(fill="both", expand=True, pady=5)

        # ================= ACCIONES =================
        ttk.Label(col_acciones, text="Acciones", font=("Segoe UI", 12, "bold")).pack(pady=5)
        ttk.Button(col_acciones, text="Generar Documento", command=self._generar_documento).pack(fill="x", pady=10)
        ttk.Button(col_acciones, text="Refrescar", command=self._cargar_datos).pack(fill="x")
        ttk.Button(col_acciones, text="Elegir carpeta de salida", command=self._elegir_carpeta).pack(fill="x", pady=5)
        self.lista_documentos.bind("<Double-Button-1>", self._previsualizar_documento)
        self.lista_clientes.bind("<Double-Button-1>", self._previsualizar_cliente)

        self.lbl_carpeta = ttk.Label(col_acciones, text="Carpeta no seleccionada", wraplength=180, foreground="gray")
        self.lbl_carpeta.pack(pady=5)

        # Mostrar carpeta si ya está seleccionada
        if self.parent.carpeta_salida:
            self.lbl_carpeta.config(text=self.parent.carpeta_salida, foreground="black")

        # Cargar datos iniciales
        self._cargar_datos()

    # --------------------------------------------------
    # CARGA DE DATOS
    # --------------------------------------------------
    def _cargar_datos(self):
        self.lista_clientes.delete(0, tk.END)
        self.lista_documentos.delete(0, tk.END)

        self.clientes = cargar_clientes_excel()

        for cliente in self.clientes:
            nombre = cliente.get("NOMBRE_EMPRESA", "SIN NOMBRE")
            nit = cliente.get("NIT", "")
            texto = f"{nombre} - NIT {nit}"
            self.lista_clientes.insert(tk.END, texto)

        carpeta_docs = os.path.join(os.getcwd(), "Documentos")
        if not os.path.exists(carpeta_docs):
            os.makedirs(carpeta_docs)

        for archivo in os.listdir(carpeta_docs):
            if archivo.lower().endswith(".docx"):
                self.lista_documentos.insert(tk.END, archivo)

    # --------------------------------------------------
    # GENERAR DOCUMENTO
    # --------------------------------------------------
    def _generar_documento(self):
        if not self.parent.carpeta_salida:
            messagebox.showwarning("Atención", "Seleccione una carpeta de salida.")
            return

        if not os.path.exists(self.parent.carpeta_salida):
            messagebox.showerror("Error", "La carpeta de salida ya no existe. Seleccione otra.")
            self.parent.carpeta_salida = None
            self.lbl_carpeta.config(text="Carpeta no seleccionada", foreground="gray")
            return

        if not self.lista_clientes.curselection():
            messagebox.showwarning("Atención", "Seleccione un cliente.")
            return

        if not self.lista_documentos.curselection():
            messagebox.showwarning("Atención", "Seleccione un documento.")
            return

        idx_cliente = self.lista_clientes.curselection()[0]
        cliente = self.clientes[idx_cliente]

        nombre_doc = self.lista_documentos.get(self.lista_documentos.curselection())
        empresa = cliente.get("NOMBRE_EMPRESA", "cliente").replace(" ", "_")
        salida = f"{empresa}_{nombre_doc}"

        ruta_salida = os.path.join(self.parent.carpeta_salida, salida)

        try:
            generar_documento_word(ruta_salida=ruta_salida, datos_cliente=cliente, nombre_salida=salida)
            messagebox.showinfo("Documento generado", f"Documento creado correctamente:\n\n{ruta_salida}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el documento:\n{str(e)}")

    # --------------------------------------------------
    # ELEGIR CARPETA DE SALIDA
    # --------------------------------------------------
    def _elegir_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if carpeta:
            self.parent.carpeta_salida = carpeta
            self.lbl_carpeta.config(text=carpeta, foreground="black")

    # --------------------------------------------------
    # PREVISUALIZAR DOCUMENTO
    # --------------------------------------------------
    def _previsualizar_documento(self, event):
        if not self.lista_documentos.curselection():
            return
        nombre_doc = self.lista_documentos.get(self.lista_documentos.curselection())
        ruta_doc = os.path.join(os.getcwd(), "Documentos", nombre_doc)

        if os.path.exists(ruta_doc):
            try:
                os.startfile(ruta_doc)  # Abre el documento con la app predeterminada
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el documento:\n{str(e)}")
        else:
            messagebox.showerror("Error", "El documento no existe.")

    # --------------------------------------------------
    # PREVISUALIZAR CLIENTE
    # --------------------------------------------------
    def _previsualizar_cliente(self, event):
        if not self.lista_clientes.curselection():
            return
        idx = self.lista_clientes.curselection()[0]
        cliente = self.clientes[idx]
    
        ventana = tk.Toplevel(self)
        ventana.title("Vista previa del cliente")
        ventana.geometry("600x400")
    
        tree = ttk.Treeview(ventana, columns=("Campo", "Valor"), show="headings")
        tree.heading("Campo", text="Campo")
        tree.heading("Valor", text="Valor")
        tree.pack(fill="both", expand=True, padx=10, pady=10)
    
        for k, v in cliente.items():
            tree.insert("", tk.END, values=(k, v))



