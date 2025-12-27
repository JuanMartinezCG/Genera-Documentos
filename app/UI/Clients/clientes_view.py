import tkinter as tk

class ClientesView(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        tk.Label(self, text="Vista de Clientes", font=("Arial", 18)).pack(pady=20)

        tk.Button(
            self,
            text="Volver a Documentos",
            command=lambda: controller.mostrar_vista("DocumentosView")
        ).pack(pady=10)
