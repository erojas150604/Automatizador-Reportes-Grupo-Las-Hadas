import tkinter as tk
from tkinter import messagebox
import threading
from reportes.reporte2 import Reporte2

class Reporte2UI(tk.Toplevel):
    def __init__(self, master, funciones, icono_path=None):
        super().__init__(master)
        self.funciones = funciones
        self.reporte2 = Reporte2()
        self.icono_path = icono_path
        if icono_path:
            self.iconbitmap(icono_path)
        self.title("Reporte: Cuentas vencidas")
        self.geometry("450x350")
        self.configure(bg="white")

        self.tipos_reporte = {
            "terrenos": tk.BooleanVar(value=True),
            "construcciones": tk.BooleanVar(value=True)
        }
        self.tablas_pdf = {
            "menos_de_1_mes": tk.BooleanVar(value=True),
            "1_mes": tk.BooleanVar(value=True),
            "2_meses": tk.BooleanVar(value=True),
            "3_meses": tk.BooleanVar(value=True),
            "4_a_6_meses": tk.BooleanVar(value=True),
            "más_de_6_meses": tk.BooleanVar(value=True)
        }

        self.crear_widgets()

    def crear_widgets(self):
        # ==== SELECCIÓN DE TIPO DE REPORTE ====
        frame_tipo = tk.LabelFrame(self, text="Selecciona el tipo de reporte", bg="white")
        frame_tipo.pack(padx=10, pady=10, fill="x")

        tk.Checkbutton(frame_tipo, text="Terrenos", variable=self.tipos_reporte["terrenos"], bg="white").pack(anchor="w", padx=10)
        tk.Checkbutton(frame_tipo, text="Construcciones", variable=self.tipos_reporte["construcciones"], bg="white").pack(anchor="w", padx=10)


        # ==== SELECCIÓN DE TABLAS A INCLUIR EN PDF ====
        frame_tablas = tk.LabelFrame(self, text="Tablas a incluir en el PDF", bg="white")
        frame_tablas.pack(padx=10, pady=10, fill="x")

        for nombre, var in self.tablas_pdf.items():
            tk.Checkbutton(frame_tablas, text=nombre.replace("_", " ").capitalize(), variable=var, bg="white").pack(anchor="w", padx=10)

        # ==== BOTONES DE ACCIÓN ====
        frame_botones = tk.Frame(self, bg="white")
        frame_botones.pack(pady=20)

        btn_generar = tk.Button(frame_botones, text="Generar Reporte", width=20, command=self.generar_reporte)
        btn_generar.grid(row=0, column=0, padx=10)

        btn_reset = tk.Button(frame_botones, text="Resetear interfaz", width=20, command=self.resetear_interfaz)
        btn_reset.grid(row=0, column=1, padx=10)


    def generar_reporte(self):
        # 1. Mostrar ventana de carga
        ventana_carga = self.funciones.mostrar_ventana_cargando(self.master, title="Generando reporte", label_text="Tu reporte ya casi está listo...", icono_path=self.icono_path)
        self.master.update()

        def tarea():
            try:
                # 2. Configurar los valores necesarios
                # Calcular tipo de reporte según checkboxes activos
                tipos_seleccionados = [k for k, v in self.tipos_reporte.items() if v.get()]
                if len(tipos_seleccionados) == 2:
                    self.reporte2.tipo_reporte = "ambos"
                elif len(tipos_seleccionados) == 1:
                    self.reporte2.tipo_reporte = tipos_seleccionados[0]
                else:
                    messagebox.showwarning("Atención", "Selecciona al menos un tipo de reporte.")
                    ventana_carga.destroy()
                    return

                self.reporte2.tablas_seleccionadas_pdf = [
                    k for k, v in self.tablas_pdf.items() if v.get()
                ]

                # 3. Generar tablas
                exito = self.reporte2.generar_tablas_cuentas_vencidas()
                if not exito:
                    return

                # 4. Generar PDF
                generado = self.reporte2.generar_pdf_reporte2()

                if generado:
                    self.master.after(0, lambda: messagebox.showinfo("Éxito", "✅ El reporte PDF fue generado correctamente."))
                else:
                    print("ℹ️ El usuario canceló el guardado o no se generó el PDF.")

            except Exception as e:
                print(f"❌ Error al generar el reporte: {e}")
                self.master.after(0, lambda: messagebox.showerror("Error", f"Ocurrió un error al generar el reporte:\n{e}"))

            finally:
                # 5. Cerrar ventana de carga al terminar
                self.master.after(0, ventana_carga.destroy)

        # 6. Ejecutar en hilo separado
        threading.Thread(target=tarea).start()
        
    def resetear_interfaz(self):
        self.funciones.resetear_widgets([
            *self.tipos_reporte.values(),
            *self.tablas_pdf.values()
        ])

        tablas_permanentes = ["admDocumentos", "admMovimientos", "admProductos"]
        self.funciones.eliminar_tablas_temporales(tablas_permanentes)


