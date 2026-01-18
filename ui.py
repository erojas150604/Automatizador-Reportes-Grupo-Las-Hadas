import tkinter as tk
from tkinter import ttk, messagebox
import threading
import pandas as pd
from funciones import AppFunciones, name_db
from base_de_datos import BaseDeDatos
from ventanas.ventana_reporte1 import Reporte1UI  
from ventanas.ventana_reporte2 import Reporte2UI 
from ventanas.ventana_reporte3 import Reporte3UI 
from ventanas.ventana_reporte4 import Reporte4UI
from ventanas.ventana_reporte5 import Reporte5UI  
from ventanas.ventana_reporte6 import Reporte6UI  
from ventanas.ventana_reporte7 import Reporte7UI
from ventanas.ventana_reporte8 import Reporte8UI
from ventanas.ventana_reporte9 import Reporte9UI
from ventanas.ventana_config import VentanaConfiguracion
from version import __version__, __app_name__, __client__



# En tu archivo principal (como main.py o app.py)
import os
import sys

def obtener_ruta_recurso(rel_path):
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, rel_path)

ruta_icono = obtener_ruta_recurso("Logo CUBE.ico")


class AppUI:
    def __init__(self, master):
        self.master = master
        self.funciones = AppFunciones()
        self.db = BaseDeDatos(name_db)

        self.master.title(f"{__app_name__} – {__client__} v{__version__}")
        master.iconbitmap(ruta_icono)
        master.geometry("650x400")
        self.master.protocol("WM_DELETE_WINDOW", self.cerrar_aplicacion)

        self.archivos_cargados = False
        self.ventana_config = None  # <- evita duplicados de configuración
        
        self.tablas = {}
        self.ventanas_abiertas = {}  # clave = nombre del reporte, valor = instancia de la ventana


        # Frame principal dividido en 2 columnas
        frame_principal = tk.Frame(master)
        frame_principal.pack(fill="both", expand=True, padx=20, pady=20)

        # Frame izquierdo (botones y controles)
        frame_izquierdo = tk.Frame(frame_principal)
        frame_izquierdo.grid(row=0, column=0, sticky="n")

        # Frame derecho (lista de tablas)
        frame_derecho = tk.Frame(frame_principal)
        frame_derecho.grid(row=0, column=1, padx=40, sticky="n")

        # --- Panel izquierdo ---

        # Combobox de selección de reporte
        self.area_var = tk.StringVar(value="Seleccionar reporte...")
        self.boton_desplegable = ttk.Combobox(
            frame_izquierdo,
            textvariable=self.area_var,
            values=sorted([
                "Gastos por proyecto",
                "Cuentas vencidas",
                "Estado de cuenta",
                "Gastos por encargado de obra",
                "Gastos de materiales y servicios",
                "Costos por proyecto",
                "Crédito y cobranza",
                "Costo de IMSS, SAR e INFONAVIT",
                "Costo por proyecto por fases"
            ]),
            state="readonly",
            width=30
        )
        self.boton_desplegable.pack(pady=(0, 15))
        self.boton_desplegable.bind("<<ComboboxSelected>>", self.manejar_cambio_combobox)

        # Botones de acción
        self.boton_cargar_archivo = tk.Button(frame_izquierdo, text="Cargar archivos Excel", command=self.cargar_archivos)
        self.boton_cargar_archivo.pack(fill="x", pady=5)

        self.boton_reset = tk.Button(frame_izquierdo, text="Resetear interfaz", command=self.resetear_interfaz)
        self.boton_reset.pack(fill="x", pady=5)

        btn_editar_config = tk.Button(frame_izquierdo, text="Editar configuración", command=self.abrir_editor_config)
        btn_editar_config.pack(fill="x", pady=5)

        # --- Panel derecho ---

        self.label_tablas = tk.Label(frame_derecho, text="Tablas cargadas:", anchor="w")
        self.label_tablas.pack(anchor="w")

        # Scroll para la lista
        scroll = tk.Scrollbar(frame_derecho)
        scroll.pack(side="right", fill="y")

        self.lista_tablas = tk.Listbox(frame_derecho, height=20, width=50, yscrollcommand=scroll.set)
        self.lista_tablas.pack()
        scroll.config(command=self.lista_tablas.yview)



    def cargar_archivos(self):
        file_paths = self.funciones.seleccionar_archivos()
        if not file_paths:
            messagebox.showwarning("Aviso", "No se seleccionaron archivos.")
            return

        # Mostrar ventana de carga y mantenerla accesible
        ventana_carga = self.funciones.mostrar_ventana_cargando(self.master, title="Cargando archivos", label_text="Dame unos segundos más...", icono_path=ruta_icono)
        self.master.update()

        def tarea_de_carga():
            try:
                tablas_cargadas = self.funciones.cargar_archivos_excel(file_paths)

                # Regresar al hilo principal para actualizar UI
                self.master.after(0, lambda: self.procesar_tablas(tablas_cargadas))

            except Exception as e:
                self.master.after(0, lambda: messagebox.showerror("Error", f"Error al cargar archivos:\n{e}"))

            finally:
                self.master.after(0, ventana_carga.destroy)

        # Ejecutar la carga en segundo plano
        threading.Thread(target=tarea_de_carga).start()
        
    def procesar_tablas(self, tablas_cargadas):
        tablas_nuevas = []

        for nombre_tabla, tabla in tablas_cargadas.items():
            if not self.db.tabla_existe(nombre_tabla):
                try:
                    self.db.agregar_tabla(nombre_tabla, tabla)
                    self.lista_tablas.insert(tk.END, nombre_tabla)
                    tablas_nuevas.append(nombre_tabla)
                except Exception as e:
                    print(f"❌ Error al insertar tabla '{nombre_tabla}': {e}")
            else:
                print(f"⚠️ Tabla '{nombre_tabla}' ya existe. Se omite.")

        if tablas_nuevas:
            messagebox.showinfo("Éxito", f"Se agregaron {len(tablas_nuevas)} nuevas tabla(s).")
        else:
            messagebox.showinfo("Aviso", "No se agregó ninguna tabla nueva.")

        self.archivos_cargados = True


        
    def manejar_cambio_combobox(self, event=None):
        seleccion = self.area_var.get()
        if self.archivos_cargados:
            self.abrir_ventana_reporte(seleccion)
        else:
            messagebox.showwarning("Advertencia", "Primero debes cargar archivos.")

        
    def intentar_abrir_ventana_reporte(self):
        if (
            self.archivos_cargados
            and self.reporte_seleccionado_pendiente == "Gastos por proyecto"
            and not self.ventana_reporte_abierta
        ):
            Reporte1UI(self.master, self.funciones, icono_path=ruta_icono)
            self.ventana_reporte_abierta = True
        
        elif (
            self.archivos_cargados
            and self.reporte_seleccionado_pendiente == "Cuentas vencidas"
            and not self.ventana_reporte_abierta
        ):
            Reporte2UI(self.master, self.funciones, icono_path=ruta_icono)
            self.ventana_reporte_abierta = True
            
        elif (
            self.archivos_cargados
            and self.reporte_seleccionado_pendiente == "Estado de cuenta"
            and not self.ventana_reporte_abierta
        ):
            Reporte3UI(self.master, icono_path=ruta_icono)
            self.ventana_reporte_abierta = True
        
        elif (
            self.archivos_cargados
            and self.reporte_seleccionado_pendiente == "Gastos por encargado de obra"
            and not self.ventana_reporte_abierta
        ):
            Reporte4UI(self.master, self.funciones, icono_path=ruta_icono)
            self.ventana_reporte_abierta = True

        elif (
            self.archivos_cargados
            and self.reporte_seleccionado_pendiente == "Gastos de materiales y servicios"
            and not self.ventana_reporte_abierta
        ):
            Reporte5UI(self.master, self.funciones, icono_path=ruta_icono)
            self.ventana_reporte_abierta = True
            
        elif (
            self.archivos_cargados
            and self.reporte_seleccionado_pendiente == "Costos por proyecto"
            and not self.ventana_reporte_abierta
        ):
            Reporte6UI(self.master, self.funciones, icono_path=ruta_icono)
            self.ventana_reporte_abierta = True

    def abrir_ventana_reporte(self, seleccion):
        if seleccion in self.ventanas_abiertas:
            ventana = self.ventanas_abiertas[seleccion]
            if ventana.winfo_exists():
                messagebox.showinfo("Ventana abierta", f"Ya tienes abierta la ventana de '{seleccion}'.")
                return
            else:
                self.ventanas_abiertas.pop(seleccion)  # limpiar si se cerró manualmente

        # Crear nueva instancia
        instancia = None

        if seleccion == "Gastos por proyecto":
            instancia = Reporte1UI(self.master, self.funciones, icono_path=ruta_icono)

        elif seleccion == "Cuentas vencidas":
            instancia = Reporte2UI(self.master, self.funciones, icono_path=ruta_icono)

        elif seleccion == "Estado de cuenta":
            instancia = Reporte3UI(self.master, icono_path=ruta_icono)

        elif seleccion == "Gastos por encargado de obra":
            instancia = Reporte4UI(self.master, self.funciones, icono_path=ruta_icono)

        elif seleccion == "Gastos de materiales y servicios":
            instancia = Reporte5UI(self.master, self.funciones, icono_path=ruta_icono)

        elif seleccion == "Costos por proyecto":
            instancia = Reporte6UI(self.master, self.funciones, icono_path=ruta_icono)
        
        elif seleccion == "Crédito y cobranza":
            instancia = Reporte7UI(self.master, self.funciones, icono_path=ruta_icono)
            
        elif seleccion == "Costo de IMSS, SAR e INFONAVIT":
            instancia = Reporte8UI(self.master, self.funciones, icono_path=ruta_icono)

        elif seleccion == "Costo por proyecto por fases":
            instancia = Reporte9UI(self.master, self.funciones, icono_path=ruta_icono)

        else:
            return

        # Detectar si la instancia ES una ventana (Toplevel) o la TIENE como atributo
        if isinstance(instancia, tk.Toplevel):
            ventana = instancia
        elif hasattr(instancia, "ventana"):
            ventana = instancia.ventana
        else:
            print(f"❌ No se pudo identificar ventana para '{seleccion}'")
            return

        self.ventanas_abiertas[seleccion] = ventana
        ventana.protocol("WM_DELETE_WINDOW", lambda: self.cerrar_ventana_reporte(seleccion))


            
    def abrir_editor_config(self):
        if self.ventana_config and self.ventana_config.winfo_exists():
            messagebox.showinfo("Ventana abierta", "Ya tienes abierta la ventana de configuración.")
            self.ventana_config.lift()
            self.ventana_config.focus_force()
            return

        self.ventana_config = VentanaConfiguracion(self.master, icono_path=ruta_icono)
        self.ventana_config.protocol("WM_DELETE_WINDOW", self.cerrar_editor_config)

        
    def cerrar_editor_config(self):
        if self.ventana_config:
            self.ventana_config.destroy()
            self.ventana_config = None

        
    
    def cerrar_ventana_reporte(self, seleccion):
        if seleccion in self.ventanas_abiertas:
            ventana = self.ventanas_abiertas.pop(seleccion)
            ventana.destroy()



    def cerrar_aplicacion(self):
        try:
            self.db.cerrar_conexion()
        except Exception as e:
            print(f"Error al cerrar: {e}")
        finally:
            self.master.destroy()
            
    
    def resetear_interfaz(self):
        self.area_var.set("Seleccionar reporte...")
        self.lista_tablas.delete(0, tk.END)
        self.tablas.clear()
        self.archivos_cargados = False
        self.reporte_seleccionado_pendiente = None
        self.ventana_reporte_abierta = False

        self.funciones.eliminar_todas_las_tablas()

            