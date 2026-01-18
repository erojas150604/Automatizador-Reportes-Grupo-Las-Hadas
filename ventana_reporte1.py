import tkinter as tk
from tkinter import ttk, Toplevel, Canvas, Scrollbar
from tkinter import messagebox
import threading
# from ventanas.ventana_exportar_excel_reporte1 import VentanaExportarExcelReporte1
from reportes.reporte1 import Reporte1
from reportes.config import empresas, cargar_json

class Reporte1UI:
    def __init__(self, master, funciones, icono_path=None):
        self.master = master
        self.funciones = funciones
        # Excluir "Otros" antes de usarlo
        proyectos_filtrados = {
            privada: proyectos
            for privada, proyectos in cargar_json("proyectos_por_privada.json").items()
            if privada != "Otros"
        }
        self.proyectos_por_privada = proyectos_filtrados
        self.reporte1 = Reporte1()
        self.ventana = Toplevel(master)
        self.icono_path = icono_path
        if icono_path:
            self.ventana.iconbitmap(icono_path)
        self.ventana.title("Reporte: Gastos por proyecto")
        self.ventana.geometry("450x600")

        # tk.Label(self.ventana, text="Generar Reporte 1", font=("Arial", 14)).pack(pady=10)

        # Obtener datos dinámicos
        años, meses, _ = self.funciones.obtener_años_meses_proyectos(empresas)

        # Generar lista de privadas desde el diccionario
        privadas = cargar_json("privadas.json")

        # Obtener todos los proyectos de todas las privadas en formato "codigo - nombre"
        lista_proyectos = sorted(
            [f"{codigo} - {nombre}"
             for proyectos in self.proyectos_por_privada.values()
             for codigo, nombre in proyectos.items()],
            key=lambda x: int(x.split(" - ")[0])
        )

        # Scroll general
        contenedor = tk.Frame(self.ventana)
        contenedor.pack(fill="both", expand=True)

        canvas = Canvas(contenedor)
        scrollbar = Scrollbar(contenedor, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Parámetros en vertical
        self.vars_años, self.var_todos_años = self.crear_scroll_parametro("Año(s)", años)
        self.vars_meses, self.var_todos_meses = self.crear_scroll_parametro("Mes(es)", meses)
        self.vars_empresas, self.var_todos_empresas = self.crear_scroll_parametro("Empresa(s)", empresas)
        self.vars_privadas, self.var_todos_privadas = self.crear_scroll_parametro("Privada(s)", privadas)
        self.vars_proyectos, self.var_todos_proyectos = self.crear_scroll_parametro("Proyecto(s)", lista_proyectos)

        tk.Button(self.scrollable_frame, text="Generar Reporte", command=self.generar_reporte).pack(pady=10)
        
        self.boton_reset = tk.Button(self.ventana, text="Resetear interfaz", command=self.resetear_interfaz)
        self.boton_reset.pack(pady=5, padx=10, anchor="ne")  # Ajusta el anchor según lo quieras (e.g. "center", "e", etc.)

        self.ventana.protocol("WM_DELETE_WINDOW", self.on_close)

    def crear_scroll_parametro(self, titulo, opciones):
        frame = tk.LabelFrame(self.scrollable_frame, text=titulo)
        frame.pack(padx=10, pady=5, fill="x")

        canvas = Canvas(frame)
        scrollbar = Scrollbar(frame, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set, height=150)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        vars_opciones = {}

        def toggle_todos():
            estado = var_todos.get()
            for var in vars_opciones.values():
                var.set(estado)

        var_todos = tk.BooleanVar()
        tk.Checkbutton(scroll_frame, text="Todos", variable=var_todos, command=toggle_todos).pack(anchor="w")

        for opcion in opciones:
            var = tk.BooleanVar()
            def actualizar_todos(var=var):  # var=var es clave para capturar correctamente
                seleccionadas = sum(v.get() for v in vars_opciones.values())
                if seleccionadas == len(vars_opciones):
                    var_todos.set(True)
                else:
                    var_todos.set(False)
            chk = tk.Checkbutton(scroll_frame, text=opcion, variable=var, command=actualizar_todos)
            chk.pack(anchor="w")
            vars_opciones[opcion] = var



        return vars_opciones, var_todos

    def get_parametros_seleccionados(self):
        años = list(self.vars_años.keys()) if self.var_todos_años.get() else [a for a, v in self.vars_años.items() if v.get()]
        meses = list(self.vars_meses.keys()) if self.var_todos_meses.get() else [m for m, v in self.vars_meses.items() if v.get()]
        empresas = list(self.vars_empresas.keys()) if self.var_todos_empresas.get() else [e for e, v in self.vars_empresas.items() if v.get()]
        privadas = list(self.vars_privadas.keys()) if self.var_todos_privadas.get() else [p for p, v in self.vars_privadas.items() if v.get()]
        if self.var_todos_proyectos.get():
            proyectos = [int(p.split(" - ")[0]) for p in self.vars_proyectos.keys()]
        else:
            proyectos = [int(p.split(" - ")[0]) for p, v in self.vars_proyectos.items() if v.get()]

        return {
            "años": años,
            "meses": meses,
            "empresas": empresas,
            "privadas": privadas,
            "proyectos": proyectos
        }

    

    def generar_reporte(self):
        # 1. Mostrar ventana de carga
        ventana_carga = self.funciones.mostrar_ventana_cargando(self.master, title="Generando reporte", label_text="Tu reporte ya casi está listo...", icono_path=self.icono_path)
        self.master.update()

        def tarea_generacion():
            try:
                params = self.get_parametros_seleccionados()

                df1 = self.reporte1.generar_tabla_gastos_por_proyecto(**params)
                if df1.empty:
                    return

                exito = self.reporte1.generar_tablas_gastos_por_proyecto_por_privada(**params)
                if not exito:
                    return

                generado = self.reporte1.generar_pdf_reporte1(**params)
                if generado:
                    self.master.after(0, lambda: messagebox.showinfo("Éxito", "✅ El reporte PDF fue generado correctamente."))
                else:
                    print("ℹ️ El usuario canceló el guardado o no se generó el PDF.")

            except Exception as e:
                print(f"❌ Error al generar el PDF: {e}")
                self.master.after(0, lambda: messagebox.showerror("Error", f"Ocurrió un error al generar el reporte:\n{e}"))

            finally:
                self.master.after(0, ventana_carga.destroy)

        # 2. Ejecutar la carga en segundo plano
        threading.Thread(target=tarea_generacion).start()




    # def abrir_ventana_exportar(self):
    #     VentanaExportarExcelReporte1(self.ventana, self.reporte1)

    def on_close(self):
        self.master.ventana_reporte_abierta = False
        self.ventana.destroy()
        
    
    def resetear_interfaz(self):
        self.funciones.resetear_widgets([
        self.var_todos_años,
        self.var_todos_meses,
        self.var_todos_empresas,
        self.var_todos_privadas,
        self.var_todos_proyectos,
        *self.vars_años.values(),
        *self.vars_meses.values(),
        *self.vars_empresas.values(),
        *self.vars_privadas.values(),
        *self.vars_proyectos.values()
        ])

        tablas_permanentes = ["admDocumentos", "admMovimientos", "admProductos"]

        self.funciones.eliminar_tablas_temporales(tablas_permanentes)
