import tkinter as tk

class DocumentosView(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        tk.Label(self, text="Vista de Documentos", font=("Arial", 18)).pack(pady=20)

        tk.Button(
            self,
            text="Ir a Clientes",
            command=lambda: controller.mostrar_vista("ClientesView")
        ).pack(pady=10)

        tk.Button(
            self,
            text="Ir a Rellenar Documento",
            command=lambda: controller.mostrar_vista("RellenarView")
        ).pack(pady=10)
