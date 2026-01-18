import tkinter as tk
from tkinter import messagebox, ttk
import threading
from datetime import datetime
import calendar

from reportes.reporte7 import Reporte7

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]
MES_A_NUM = {m: i+1 for i, m in enumerate(MESES_ES)}

class Reporte7UI(tk.Toplevel):
    def __init__(self, master, funciones, icono_path=None):
        super().__init__(master)
        self.funciones = funciones
        self.reporte7 = Reporte7()
        self.icono_path = icono_path
        if icono_path:
            self.iconbitmap(icono_path)

        self.title("Reporte: Crédito y cobranza")
        self.geometry("560x350")
        self.configure(bg="white")

        # ---- variables de UI ----
        hoy = datetime.now()
        self.var_mes_actual = tk.BooleanVar(value=True)
        
        # Tipos de reporte (Terrenos / Construcciones)
        self.var_tipo_terrenos = tk.BooleanVar(value=True)
        self.var_tipo_construcciones = tk.BooleanVar(value=True)

        self.var_mes = tk.StringVar(value=MESES_ES[hoy.month - 1])
        self.var_anio = tk.StringVar(value=str(hoy.year))
        self.var_dia = tk.StringVar(value=str(hoy.day))

        # Empresas
        self.empresas_lista = list(self.reporte7.empresas) if hasattr(self.reporte7, "empresas") else []
        self.var_empresas_all = tk.BooleanVar(value=True)
        self.var_empresas = {e: tk.BooleanVar(value=True) for e in self.empresas_lista}

        self.crear_widgets()
        self._toggle_fecha_controls()   # desactivar si usa mes actual
        self._update_dias()             # cargar días correctos

    def crear_widgets(self):
        # ==== FRAME TIPO DE REPORTE ====
        frame_tipo = tk.LabelFrame(self, text="Tipo de reporte", bg="white")
        frame_tipo.pack(fill="x", padx=16, pady=(0, 12))

        tk.Checkbutton(
            frame_tipo, text="Terrenos", bg="white",
            variable=self.var_tipo_terrenos
        ).pack(side="left", padx=(12, 8), pady=6)

        tk.Checkbutton(
            frame_tipo, text="Construcciones", bg="white",
            variable=self.var_tipo_construcciones
        ).pack(side="left", padx=8, pady=6)


        # ==== FRAME PARAMETROS FECHA ====
        frame_params = tk.LabelFrame(self, text="Fecha de corte", bg="white")
        frame_params.pack(fill="x", padx=16, pady=(12, 8))

        chk = tk.Checkbutton(
            frame_params, text="Usar mes actual",
            variable=self.var_mes_actual, bg="white",
            command=self._on_toggle_mes_actual
        )
        chk.grid(row=0, column=0, columnspan=6, sticky="w", pady=(4, 8))

        # Mes
        tk.Label(frame_params, text="Mes:", bg="white").grid(row=1, column=0, sticky="e", padx=(0, 8))
        self.combo_mes = ttk.Combobox(frame_params, values=MESES_ES, textvariable=self.var_mes,
                                      state="readonly", width=16)
        self.combo_mes.grid(row=1, column=1, sticky="w", pady=2)
        self.combo_mes.bind("<<ComboboxSelected>>", lambda e: self._update_dias())

        # Año
        tk.Label(frame_params, text="Año:", bg="white").grid(row=1, column=2, sticky="e", padx=(16, 8))
        anios = [str(datetime.now().year + d) for d in range(1, -6, -1)]
        self.combo_anio = ttk.Combobox(frame_params, values=anios, textvariable=self.var_anio,
                                       state="readonly", width=10)
        self.combo_anio.grid(row=1, column=3, sticky="w", pady=2)
        self.combo_anio.bind("<<ComboboxSelected>>", lambda e: self._update_dias())

        # Día (solo se usa cuando NO es mes actual)
        tk.Label(frame_params, text="Día:", bg="white").grid(row=1, column=4, sticky="e", padx=(16, 8))
        self.combo_dia = ttk.Combobox(frame_params, values=[], textvariable=self.var_dia,
                                      state="readonly", width=6)
        self.combo_dia.grid(row=1, column=5, sticky="w", pady=2)

        for i in range(6):
            frame_params.columnconfigure(i, weight=1)

        # ==== FRAME EMPRESAS ====
        frame_emp = tk.LabelFrame(self, text="Empresas", bg="white")
        frame_emp.pack(fill="both", expand=False, padx=16, pady=(4, 12))

        # Seleccionar todas
        chk_all = tk.Checkbutton(frame_emp, text="Seleccionar todas",
                                 variable=self.var_empresas_all, bg="white",
                                 command=self._on_toggle_empresas_all)
        chk_all.grid(row=0, column=0, sticky="w", padx=4, pady=(6, 2))

        # Lista de checkbuttons (2 columnas)
        col = 0
        row = 1
        for idx, empresa in enumerate(self.empresas_lista):
            cb = tk.Checkbutton(frame_emp, text=str(empresa), bg="white",
                                variable=self.var_empresas[empresa],
                                command=self._on_empresas_changed)
            cb.grid(row=row, column=col, sticky="w", padx=12, pady=2)
            col += 1
            if col >= 2:  # 2 columnas
                col = 0
                row += 1

        for i in range(2):
            frame_emp.columnconfigure(i, weight=1)

        # ==== BOTONES DE ACCIÓN ====
        frame_botones = tk.Frame(self, bg="white")
        frame_botones.pack(pady=10)

        btn_generar = tk.Button(frame_botones, text="Generar reporte", width=20, command=self.generar_reporte)
        btn_generar.grid(row=0, column=0, padx=10)

        btn_reset = tk.Button(frame_botones, text="Resetear interfaz", width=20, command=self.resetear_interfaz)
        btn_reset.grid(row=0, column=1, padx=10)
        
        
    def _get_tipo_seleccionado(self) -> str:
        """Devuelve 'terrenos', 'construcciones' o 'ambos' según los checkbuttons.
        Si ninguno está marcado, por usabilidad activa ambos."""
        t = self.var_tipo_terrenos.get()
        c = self.var_tipo_construcciones.get()
        if t and c:
            return "ambos"
        if t:
            return "terrenos"
        if c:
            return "construcciones"
        # Si ninguno marcado, forzamos ambos
        self.var_tipo_terrenos.set(True)
        self.var_tipo_construcciones.set(True)
        return "ambos"


    # ============ LÓGICA DE FECHA ============
    def _on_toggle_mes_actual(self):
        self._toggle_fecha_controls()
        self._update_dias()

    def _toggle_fecha_controls(self):
        state = "disabled" if self.var_mes_actual.get() else "readonly"
        self.combo_mes.configure(state=state)
        self.combo_anio.configure(state=state)
        self.combo_dia.configure(state=state)

    def _update_dias(self):
        """Actualiza los días válidos para el Mes/Año seleccionados."""
        if self.var_mes_actual.get():
            # usar hoy
            hoy = datetime.now()
            dias_mes = calendar.monthrange(hoy.year, hoy.month)[1]
            self.combo_dia["values"] = [str(d) for d in range(1, dias_mes + 1)]
            self.var_mes.set(MESES_ES[hoy.month - 1])
            self.var_anio.set(str(hoy.year))
            self.var_dia.set(str(min(hoy.day, dias_mes)))
            return

        try:
            mes = MES_A_NUM.get(self.var_mes.get(), datetime.now().month)
            anio = int(self.var_anio.get())
        except Exception:
            mes = datetime.now().month
            anio = datetime.now().year

        dias_mes = calendar.monthrange(anio, mes)[1]
        valores = [str(d) for d in range(1, dias_mes + 1)]
        self.combo_dia["values"] = valores
        # si el día actual no está en rango, ajusta al último
        dia_sel = int(self.var_dia.get()) if self.var_dia.get().isdigit() else 1
        self.var_dia.set(str(min(max(1, dia_sel), dias_mes)))

    def _get_fecha_reporte(self) -> datetime:
        """Fecha de corte exacta (hoy si 'usar mes actual' está activo)."""
        if self.var_mes_actual.get():
            hoy = datetime.now()
            return datetime(hoy.year, hoy.month, hoy.day)
        try:
            mes = MES_A_NUM.get(self.var_mes.get(), datetime.now().month)
            anio = int(self.var_anio.get())
            dia = int(self.var_dia.get())
            # clamp por seguridad
            dia = max(1, min(dia, calendar.monthrange(anio, mes)[1]))
            return datetime(anio, mes, dia)
        except Exception:
            hoy = datetime.now()
            return datetime(hoy.year, hoy.month, hoy.day)

    # ============ LÓGICA EMPRESAS ============
    def _on_toggle_empresas_all(self):
        val = self.var_empresas_all.get()
        for v in self.var_empresas.values():
            v.set(val)

    def _on_empresas_changed(self):
        # si todas están marcadas, marca "todas"; si no, desmárcalo
        all_marked = all(v.get() for v in self.var_empresas.values()) if self.var_empresas else True
        self.var_empresas_all.set(all_marked)

    def _get_empresas_seleccionadas(self):
        seleccionadas = [e for e, v in self.var_empresas.items() if v.get()]
        # si ninguna seleccionada, por usabilidad toma todas
        if not seleccionadas:
            return list(self.empresas_lista)
        return seleccionadas

    # ============ GENERACIÓN ============
    def generar_reporte(self):
        ventana_carga = self.funciones.mostrar_ventana_cargando(
            self.master,
            title="Generando reporte",
            label_text="Tu reporte ya casi está listo...",
            icono_path=self.icono_path
        )
        self.master.update()

        def tarea():
            old_empresas = list(self.reporte7.empresas) if hasattr(self.reporte7, "empresas") else []
            try:
                fecha_corte = self._get_fecha_reporte()
                usar_mes_actual = self.var_mes_actual.get()
                empresas_sel = self._get_empresas_seleccionadas()

                # usar solo las empresas seleccionadas (temporalmente)
                if hasattr(self.reporte7, "empresas"):
                    self.reporte7.empresas = empresas_sel

                # 🔑 Pasa el flag 'usar_mes_actual' a tu lógica interna
                # 1) Generar las tablas finales (Terrenos y Construcciones) en SQLite y RAM
                res = self.reporte7.generar_tabla_credito_y_cobranza(
                    fecha_reporte=fecha_corte,
                    empresas_sel=empresas_sel,
                    usar_mes_actual=usar_mes_actual
                )
                # res es un dict: {"terrenos": df_terr, "construcciones": df_cons}

                # 2) Verifica si hay datos en lo elegido
                tipo = self._get_tipo_seleccionado()

                def _hay_datos(df):
                    try:
                        return df is not None and not df.empty
                    except:
                        return False

                if tipo == "terrenos" and not _hay_datos(res.get("terrenos")):
                    self.master.after(0, lambda: messagebox.showinfo(
                        "Sin datos", "No hay información para Terrenos en el periodo seleccionado.")
                    )
                    return

                if tipo == "construcciones" and not _hay_datos(res.get("construcciones")):
                    self.master.after(0, lambda: messagebox.showinfo(
                        "Sin datos", "No hay información para Construcciones en el periodo seleccionado.")
                    )
                    return

                if tipo == "ambos" and (not _hay_datos(res.get("terrenos")) and not _hay_datos(res.get("construcciones"))):
                    self.master.after(0, lambda: messagebox.showinfo(
                        "Sin datos", "No hay información ni en Terrenos ni en Construcciones para el periodo seleccionado.")
                    )
                    return

                # 3) Exportar (generará 1 o 2 PDFs/Excels según 'tipo')
                generado = self.reporte7.generar_pdf(
                    empresas=empresas_sel,
                    fecha_corte=fecha_corte,
                    usar_mes_actual=usar_mes_actual,
                    tipo=tipo
                )
                if generado:
                    self.master.after(0, lambda: messagebox.showinfo(
                        "Éxito", "✅ El/Los PDF(s) se generaron correctamente.")
                    )


            except Exception as e:
                print(f"❌ Error al generar el reporte 7: {e}")
                err_text = f"Ocurrió un error al generar el reporte:\n{e}"
                self.master.after(0, lambda: messagebox.showerror("Error", err_text))

            finally:
                # restaurar empresas originales
                if hasattr(self.reporte7, "empresas"):
                    self.reporte7.empresas = old_empresas
                self.master.after(0, ventana_carga.destroy)

        threading.Thread(target=tarea, daemon=True).start()

    def resetear_interfaz(self):
        self.var_tipo_terrenos.set(True)
        self.var_tipo_construcciones.set(True)
        hoy = datetime.now()
        self.var_mes_actual.set(True)
        self.var_mes.set(MESES_ES[hoy.month - 1])
        self.var_anio.set(str(hoy.year))
        self.var_dia.set(str(hoy.day))
        self._toggle_fecha_controls()
        self._update_dias()

        # empresas: marcar todas
        self.var_empresas_all.set(True)
        for v in self.var_empresas.values():
            v.set(True)

        self.funciones.resetear_widgets([])
        tablas_permanentes = ["admDocumentos", "admMovimientos", "admProductos"]
        self.funciones.eliminar_tablas_temporales(tablas_permanentes)
