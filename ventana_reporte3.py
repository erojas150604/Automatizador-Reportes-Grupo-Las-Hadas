import tkinter as tk
from tkinter import ttk, messagebox
import threading
from funciones import AppFunciones
from reportes.reporte3 import Reporte3
from reportes.config import cargar_json
from datetime import datetime

class Reporte3UI(tk.Toplevel):
    def __init__(self, master, icono_path=None):
        super().__init__(master)
        self.clientes_con_interes = cargar_json("clientes_con_interes.json")
        self.tasas_con_interes = cargar_json("tasas_interes.json")
        self.funciones = AppFunciones()
        self.title("Reporte: Estado de cuenta")
        self.geometry("600x400")
        
        self.icono_path = icono_path
        
        if icono_path:
            self.iconbitmap(icono_path)  # ✅ Aplica el ícono si se proporcionó
            
        
        
        self.reporte = Reporte3()

        # === Cliente ===
        self.label_cliente = tk.Label(self, text="Cliente:")
        self.label_cliente.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.lista_clientes = [c.upper() for c in self.reporte.obtener_clientes_unicos()]
        self.entry_cliente = AutocompleteCombobox(self, self.lista_clientes)
        self.entry_cliente.grid(row=0, column=1, padx=10, pady=10, sticky="we")
        self.entry_cliente.bind("<FocusOut>", self.on_cliente_focus_out)
        
        # Variable de empresa
        self.var_empresa = tk.StringVar()
        self.var_empresa.set("Selecciona una empresa")

        # === Empresa ===
        self.label_empresa = tk.Label(self, text="Empresa:")
        self.label_empresa.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        empresas_disponibles = self.reporte.empresas
        self.combo_empresa = ttk.Combobox(self, textvariable=self.var_empresa, values=empresas_disponibles, state="readonly")
        self.combo_empresa.grid(row=1, column=1, padx=10, pady=10, sticky="we")
        
        self.combo_empresa.bind("<<ComboboxSelected>>", self.cargar_lotes_cliente)


        # === Lote ===
        self.label_lote = tk.Label(self, text="Lote:")
        self.label_lote.grid(row=2, column=0, padx=10, pady=10, sticky="w")

        self.entry_lote = AutocompleteCombobox(self, [])  # lista vacía inicial
        self.entry_lote.grid(row=2, column=1, padx=10, pady=10, sticky="we")
        
        
        # === Filtros de forma de pago ===
        self.label_filtros = tk.Label(self, text="Forma de Pago:")
        self.label_filtros.grid(row=3, column=0, padx=10, pady=5, sticky="w")

        self.var_efectivo = tk.BooleanVar(value=True)
        self.var_transferencia = tk.BooleanVar(value=True)
        self.var_incorporacion = tk.BooleanVar(value=True)

        self.check_efectivo = tk.Checkbutton(self, text="Efectivo", variable=self.var_efectivo)
        self.check_transferencia = tk.Checkbutton(self, text="Transferencia", variable=self.var_transferencia)
        self.check_incorporacion = tk.Checkbutton(self, text="Cuenta de Incorporación", variable=self.var_incorporacion)

        self.check_efectivo.grid(row=3, column=1, sticky="w", padx=10)
        self.check_transferencia.grid(row=4, column=1, sticky="w", padx=10)
        self.check_incorporacion.grid(row=5, column=1, sticky="w", padx=10)
        
        
        vcmd_float = self.register(self.validar_float)
        vcmd_int = self.register(self.validar_entero)

        
        
        self.label_formula = tk.Label(self, text="Fórmula de interés:")
        self.combobox_formula = ttk.Combobox(self, values=["Saldo insoluto", "Saldo total"], state="readonly")
        self.combobox_formula.bind("<<ComboboxSelected>>", self.actualizar_estado_entry_enganche)

        vcmd_moneda = self.register(self.validar_moneda)

        self._formateando_enganche = False          # evita recursión del trace
        self.var_enganche = tk.StringVar()
        self.var_enganche.trace_add("write", self._formatear_enganche)

        self.label_enganche = tk.Label(self, text="Enganche:")
        self.entry_enganche = tk.Entry(self, textvariable=self.var_enganche,
                                    validate="key", validatecommand=(vcmd_moneda, '%P'))


        self.label_tasa = tk.Label(self, text="Tasa de interés:")
        self.combobox_tasa = ttk.Combobox(self, values=self.tasas_con_interes, state="readonly")

        self.label_meses = tk.Label(self, text="Número de meses:")
        self.entry_meses = tk.Entry(self, validate="key", validatecommand=(vcmd_int, '%P'))

        self.label_fecha_inicio = tk.Label(self, text="Fecha inicio (DD/MM/AAAA):")
        
        self.fecha_var = tk.StringVar()
        self.fecha_var.trace_add("write", self.formatear_fecha_dinamica)  # formateo en vivo

        self.entry_fecha_inicio = tk.Entry(self, textvariable=self.fecha_var)
        self.entry_fecha_inicio.bind("<FocusOut>", self.validar_fecha_valida)


    

        self.incluir_programa_var = tk.BooleanVar(value=False)
        self.check_programa = tk.Checkbutton(self, text="Incluir Programa de Pagos", variable=self.incluir_programa_var)
        self.check_programa.grid(row=7, column=0, sticky="w", padx=10)
        

        # === Botón Generar Tabla ===
        self.boton_generar = tk.Button(self, text="Generar Estado de Cuenta", command=self.generar_estado_cuenta)
        self.boton_generar.grid(row=8, column=0, columnspan=2, pady=20)
        
        self.boton_reset = tk.Button(self, text="Resetear interfaz", command=self.resetear_interfaz)
        self.boton_reset.grid(row=0, column=4, padx=5, pady=5)

        self.columnconfigure(1, weight=1)
        
    def on_cliente_focus_out(self, event):
        cliente = self.entry_cliente.get()
        self.mostrar_tasa_si_aplica(cliente)
        self.cargar_lotes_cliente(cliente)
    
        
    def mostrar_tasa_si_aplica(self, cliente):
        cliente_normalizado = cliente.strip().upper()
        
        tiene_interes = cliente_normalizado in [c.upper() for c in self.clientes_con_interes]

        if tiene_interes:
            
            self.label_formula.grid(row=6, column=0, sticky="w", padx=10)
            self.combobox_formula.grid(row=6, column=1, sticky="w", padx=10)
            
            self.label_enganche.grid(row=7, column=0, sticky="w", padx=10)
            self.entry_enganche.grid(row=7, column=1, sticky="we", padx=10)

            self.label_tasa.grid(row=8, column=0, sticky="w", padx=10, pady=(10, 0))
            self.combobox_tasa.grid(row=8, column=1, sticky="w", padx=10)
            
            self.label_meses.grid(row=9, column=0, sticky="w", padx=10)
            self.entry_meses.grid(row=9, column=1, sticky="we", padx=10)

            self.label_fecha_inicio.grid(row=10, column=0, sticky="w", padx=10)
            self.entry_fecha_inicio.grid(row=10, column=1, sticky="we", padx=10)

            # Reubicar botones
            self.check_programa.grid(row=11, column=0, sticky="w", padx=10)
            self.boton_generar.grid(row=12, column=0, columnspan=2, pady=20)
        
        else:
            # Ocultar todos
            for widget in [
                self.label_tasa, self.combobox_tasa,
                self.label_formula, self.combobox_formula,
                self.label_enganche, self.entry_enganche,
                self.label_meses, self.entry_meses,
                self.label_fecha_inicio, self.entry_fecha_inicio
            ]:
                if widget.winfo_exists():
                    widget.grid_remove()

            self.combobox_tasa.set("")
            self.combobox_formula.set("")
            self.entry_enganche.delete(0, tk.END)
            self.entry_meses.delete(0, tk.END)
            self.entry_fecha_inicio.delete(0, tk.END)
            
            # Reubicar botones si NO hay interés
            self.check_programa.grid(row=7, column=0, sticky="w", padx=10)
            self.boton_generar.grid(row=8, column=0, columnspan=2, pady=20)
            
    
    def actualizar_estado_entry_enganche(self, event=None):
        formula = self.combobox_formula.get().strip().lower()
        if formula == "saldo insoluto":
            self.entry_enganche.config(state="disabled", disabledbackground="#f0f0f0")
            self.entry_enganche.delete(0, tk.END)  # Limpia el valor si estaba escrito
        else:
            self.entry_enganche.config(state="normal")


            
    def validar_float(self, valor):
        if valor == "":
            return True
        try:
            float(valor)
            return True
        except ValueError:
            return False

    def validar_entero(self, valor):
        if valor == "":
            return True
        return valor.isdigit()
    
    def validar_moneda(self, valor):
        """Permite: vacío, dígitos, comas de miles y un solo punto decimal."""
        if valor == "":
            return True
        try:
            float(valor.replace(",", ""))
            return True
        except ValueError:
            return False
        
    def _formatear_enganche(self, *args):
        if self._formateando_enganche:
            return
        try:
            self._formateando_enganche = True

            txt = self.var_enganche.get()
            if not txt:
                return

            # Quitar comas actuales
            limpio = txt.replace(",", "").strip()

            # Permitir signo + decimales (aunque normalmente será positivo)
            # Separar entero y decimal
            if "." in limpio:
                entero, decimal = limpio.split(".", 1)
                if entero == "" and decimal == "":
                    self.var_enganche.set("")
                    return
                # Formatear entero si es numérico
                if entero == "" or entero == "+":
                    entero_fmt = entero  # evita fallo si usuario empieza con "."
                else:
                    entero_fmt = f"{int(entero):,}"
                nuevo = f"{entero_fmt}.{decimal}"
            else:
                if not limpio or limpio == "+":
                    self.var_enganche.set(limpio)
                    return
                nuevo = f"{int(limpio):,}"

            # Actualizar y llevar cursor al final (simple y estable)
            self.var_enganche.set(nuevo)
            if hasattr(self, "entry_enganche"):
                self.entry_enganche.icursor(tk.END)

        except Exception:
            # Si algo raro se teclea, no rompas el flujo
            pass
        finally:
            self._formateando_enganche = False




    def formatear_fecha_dinamica(self, *args):
        entry = self.entry_fecha_inicio
        texto_original = self.fecha_var.get()

        # Obtener la posición actual del cursor
        pos_cursor = entry.index(tk.INSERT)

        # Extraer solo dígitos
        digitos = ''.join(filter(str.isdigit, texto_original))[:8]

        # Formatear como DD/MM/AAAA
        nuevo_texto = ''
        if len(digitos) >= 2:
            nuevo_texto += digitos[:2] + '/'
        else:
            nuevo_texto += digitos
        if len(digitos) >= 4:
            nuevo_texto += digitos[2:4] + '/'
        elif len(digitos) > 2:
            nuevo_texto += digitos[2:]
        if len(digitos) > 4:
            nuevo_texto += digitos[4:]

        # Calcular nueva posición del cursor
        # Contar cuántos dígitos había antes del cursor
        digitos_antes = len([c for i, c in enumerate(texto_original[:pos_cursor]) if c.isdigit()])
        
        # Calcular nueva posición en texto formateado
        nueva_pos = digitos_antes
        if nueva_pos >= 2:
            nueva_pos += 1
        if nueva_pos >= 4:
            nueva_pos += 1

        # Aplicar texto y posicionar cursor
        self.fecha_var.set(nuevo_texto)
        self.entry_fecha_inicio.icursor(nueva_pos)




    def validar_fecha_valida(self, event):
        fecha_str = self.fecha_var.get()
        try:
            datetime.strptime(fecha_str, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Fecha inválida", "La fecha ingresada no es válida. Usa el formato DD/MM/AAAA.")
            self.entry_fecha_inicio.focus_set()






    def filtrar_clientes(self, event):
        texto = self.cliente_var.get().strip().upper()

        # Buscar sugerencias ignorando mayúsculas
        sugerencias = [c for c in self.lista_clientes if texto in c.upper()]
        self.combo_cliente['values'] = sugerencias

        # Restablecer el texto actual y dejar el cursor al final
        self.combo_cliente.delete(0, tk.END)
        self.combo_cliente.insert(0, texto)
        self.combo_cliente.icursor(tk.END)

        # ABRIR dropdown sin seleccionar nada
        if sugerencias:
            # Esto evita que seleccione el primer elemento
            self.combo_cliente.selection_clear()
            self.combo_cliente.event_generate("<Escape>")  # cerrar por si ya estaba
            self.after(1, lambda: self.combo_cliente.event_generate("<Down>"))  # abrir sin seleccionar

    def cliente_seleccionado(self, event):
        cliente = self.cliente_var.get().strip().upper()
        lotes = self.reporte.obtener_lotes_por_cliente(cliente)
        self.combo_lote['values'] = lotes
        if lotes:
            self.combo_lote.current(0)
            
    def cargar_lotes_cliente(self, event=None):
        cliente = self.entry_cliente.get().strip().upper()
        empresa = self.var_empresa.get()

        if not cliente or empresa == "Selecciona una empresa":
            return
        
        # 🧹 Eliminar tabla del programa de pagos anterior
        self.funciones.eliminar_tablas_especificas(prefijos_a_borrar=["EstadoCuenta"])

        # Obtener lotes filtrados por cliente y empresa
        lotes = self.reporte.obtener_lotes_por_cliente_y_empresa(cliente, empresa)

        print(f"📋 Lotes filtrados: {lotes}")

        # Actualizar lista de opciones del AutocompleteCombobox
        self.entry_lote.lista_opciones = lotes  # actualizar lista visible
        if lotes:
            self.entry_lote.delete(0, tk.END)
            self.entry_lote.insert(0, lotes[0])
        else:
            self.entry_lote.delete(0, tk.END)


    def generar_estado_cuenta(self):
        # 1. Mostrar ventana de carga
        ventana_carga = self.funciones.mostrar_ventana_cargando(self.master, title="Generando reporte", label_text="Tu reporte ya casi está listo...", icono_path=self.icono_path)
        self.master.update()

        def tarea():
            try:
                cliente = self.entry_cliente.get().strip().upper()
                lote = self.entry_lote.get().strip()

                empresa_seleccionada = self.var_empresa.get()
                if empresa_seleccionada == "Selecciona una empresa":
                    self.master.after(0, lambda: messagebox.showwarning("Falta empresa", "Por favor selecciona una empresa."))
                    return

                # Obtener filtros activos
                formas_pago_seleccionadas = []
                if self.var_efectivo.get():
                    formas_pago_seleccionadas.append("EFECTIVO")
                if self.var_transferencia.get():
                    formas_pago_seleccionadas.append("TRANSFERENCIA")
                if self.var_incorporacion.get():
                    formas_pago_seleccionadas.append("CUENTA DE INCORPORACIÓN")

                incluir_programa = self.incluir_programa_var.get()
                cliente_tiene_interes = cliente in self.clientes_con_interes

                tasa_str = self.combobox_tasa.get() if hasattr(self, "combobox_tasa") else None
                # tasa_float = None
                # meses = 0  # Valor por defecto si no aplica
                
                # Inicializar variables por defecto ANTES de cualquier if
                enganche = 0.0
                tasa_float = None
                meses = 0
                formula = None
                fecha_inicio = None

                if incluir_programa and cliente_tiene_interes:
                    formula = self.combobox_formula.get().strip()

                    if formula:  # Solo si se quiere generar tabla de amortización
                        if not tasa_str:
                            messagebox.showwarning("Tasa faltante", "Selecciona una tasa de interés.")
                            return
                        try:
                            tasa_float = float(tasa_str.replace("%", "")) / 100
                        except:
                            messagebox.showerror("Error", "Tasa inválida.")
                            return

                        # ✅ Validar meses
                        valor_meses = self.entry_meses.get().strip()
                        if not valor_meses.isdigit():
                            messagebox.showerror("Error", "Ingresa un número válido en el campo de meses.")
                            return
                        meses = int(valor_meses)

                        # ✅ Validar enganche y fórmula
                        tipo_formula = formula.lower()
                        valor = self.var_enganche.get().strip().replace(",", "")
                        if tipo_formula == "saldo insoluto":
                            enganche = 0.0
                        else:
                            try:
                                enganche = float(valor) if valor else 0.0
                            except ValueError:
                                messagebox.showerror("Error", "Ingresa un número válido en el campo de enganche.")
                                return


                        fecha_inicio = self.fecha_var.get()

                    else:
                        # Si no hay fórmula, no se genera tabla de amortización
                        tasa_float = None
                        meses = 0
                        enganche = 0.0
                        formula = None
                        fecha_inicio = None



                generado = self.reporte.generar_estado_cuenta_completo(
                    cliente=cliente,
                    lote=lote,
                    empresa_objetivo=empresa_seleccionada,
                    filtros_forma_pago=formas_pago_seleccionadas,
                    incluir_programa=incluir_programa,
                    tasa_interes=tasa_float,
                    enganche=enganche,
                    meses=meses,
                    formula=formula,
                    fecha_inicio=fecha_inicio
                )

                # generado = self.reporte.generar_estado_cuenta_completo(
                #     cliente=cliente,
                #     lote=lote,
                #     empresa_objetivo=empresa_seleccionada,
                #     filtros_forma_pago=formas_pago_seleccionadas,
                #     incluir_programa=incluir_programa,
                #     tasa_interes=tasa_float  # 👈 ¡Esto es lo que faltaba!
                # )

                if generado:
                    self.master.after(0, lambda: messagebox.showinfo("Éxito", "✅ El reporte PDF fue generado correctamente."))
                else:
                    print("ℹ️ El usuario canceló el guardado o no se generó el PDF.")

            except Exception as e:
                error_msg = str(e)
                self.master.after(0, lambda: messagebox.showerror("Error", f"Ocurrió un error al generar el reporte:\n{error_msg}"))


            finally:
                self.master.after(0, ventana_carga.destroy)

        # Ejecutar la lógica en segundo plano
        threading.Thread(target=tarea).start()


        
    
    def resetear_interfaz(self):
        self.var_empresa.set("Selecciona una empresa")
        self.incluir_programa_var.set(False)
        self.funciones.resetear_widgets([
            self.entry_cliente,
            self.combo_empresa,  # <--- Era 'combobox_empresa'
            self.entry_lote,
            self.check_efectivo,        # Era 'checkbutton_efectivo'
            self.check_transferencia,   # Era 'checkbutton_transferencia'
            self.check_incorporacion    # Era 'checkbutton_cuenta'
        ])
        
        # Limpiar entradas de interés
        if self.combobox_formula.winfo_exists():
            self.combobox_formula.set("")
        if self.entry_enganche.winfo_exists():
            self.entry_enganche.config(state="normal")
            self.var_enganche.set("")   # en lugar de delete; limpia y evita doble formateo
        if self.combobox_tasa.winfo_exists():
            self.combobox_tasa.set("")
        if self.entry_meses.winfo_exists():
            self.entry_meses.delete(0, tk.END)
        if self.entry_fecha_inicio.winfo_exists():
            self.entry_fecha_inicio.delete(0, tk.END)

        # Ocultar widgets de interés
        for widget in [
            self.label_formula, self.combobox_formula,
            self.label_enganche, self.entry_enganche,
            self.label_tasa, self.combobox_tasa,
            self.label_meses, self.entry_meses,
            self.label_fecha_inicio, self.entry_fecha_inicio
        ]:
            if widget.winfo_exists():
                widget.grid_remove()

        # Reacomodar elementos visibles como si no hubiera interés
        self.check_programa.grid(row=7, column=0, sticky="w", padx=10)
        self.boton_generar.grid(row=8, column=0, columnspan=2, pady=20)

        tablas_permanentes = ["admDocumentos", "admMovimientos", "admProductos"]

        self.funciones.eliminar_tablas_temporales(tablas_permanentes)

            
class AutocompleteCombobox(tk.Entry):
    def __init__(self, master, lista_opciones, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.lista_opciones = lista_opciones
        self.var = self["textvariable"] = tk.StringVar()
        self.var.trace("w", self.actualizar_sugerencias)

        self.listbox = None

        self.bind("<Down>", self.mover_abajo)
        self.bind("<Return>", self.seleccionar_sugerencia)
        self.bind("<Escape>", lambda e: self.ocultar_sugerencias())

    def actualizar_sugerencias(self, *args):
        texto = self.var.get().strip().upper()
        if texto == "":
            self.ocultar_sugerencias()
            return

        sugerencias = [op for op in self.lista_opciones if texto in op.upper()]
        if not sugerencias:
            self.ocultar_sugerencias()
            return

        if not self.listbox:
            self.listbox = tk.Listbox(self.master, height=min(8, len(sugerencias)))
            self.listbox.bind("<<ListboxSelect>>", self.seleccionar_sugerencia)

            # NUEVO: ajustar ubicación y ancho
            x = self.winfo_rootx() - self.master.winfo_rootx()
            y = self.winfo_y() + self.winfo_height()
            width = self.winfo_width()

            self.listbox.place(x=x, y=y, width=width)

        # Rellenar con sugerencias
        self.listbox.delete(0, tk.END)
        for s in sugerencias:
            self.listbox.insert(tk.END, s)

    def seleccionar_sugerencia(self, event=None):
        if self.listbox and self.listbox.curselection():
            seleccion = self.listbox.get(self.listbox.curselection())
            self.var.set(seleccion)
        self.ocultar_sugerencias()

    def ocultar_sugerencias(self):
        if self.listbox:
            self.listbox.destroy()
            self.listbox = None

    def mover_abajo(self, event):
        if self.listbox:
            self.listbox.focus_set()
            self.listbox.select_set(0)
    
