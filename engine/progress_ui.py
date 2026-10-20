import tkinter as tk
from tkinter import ttk


class ProgressWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("TBG CALIDAD - Progreso")
        self.root.geometry("430x205")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)

        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text="TBG CALIDAD",
            font=("Segoe UI", 15, "bold"),
        ).pack(anchor="w")

        self.mensaje = tk.StringVar(value="Iniciando proceso...")
        ttk.Label(
            frame,
            textvariable=self.mensaje,
            font=("Segoe UI", 10),
            wraplength=390,
        ).pack(anchor="w", pady=(12, 8))

        self.barra = ttk.Progressbar(
            frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            length=390,
        )
        self.barra.pack(fill="x", pady=(2, 5))

        self.porcentaje = tk.StringVar(value="2 %")
        ttk.Label(
            frame,
            textvariable=self.porcentaje,
            font=("Segoe UI", 9),
        ).pack(anchor="e")

        self.actualizar("Iniciando proceso...", 2)

    def actualizar(self, mensaje, porcentaje):
        porcentaje = max(0, min(100, int(porcentaje)))
        self.mensaje.set(mensaje)
        self.barra["value"] = porcentaje
        self.porcentaje.set(f"{porcentaje} %")
        self.root.update_idletasks()
        self.root.update()

    def indicador(self, actual, total):
        if total <= 0:
            self.actualizar("No hay indicadores para calcular.", 75)
            return

        porcentaje = 20 + round(55 * actual / total)
        self.actualizar(
            f"Calculando indicador {actual} de {total}...",
            porcentaje,
        )

    def cerrar(self):
        try:
            self.root.destroy()
        except tk.TclError:
            pass
