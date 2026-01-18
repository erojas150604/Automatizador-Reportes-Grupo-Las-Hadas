import re
import sqlite3
import pandas as pd
from datetime import datetime
from tkinter import messagebox, filedialog
from .pdf_utils import ReportePDF
from .config import name_db, tablas_por_empresa_reporte1, empresas, cargar_json


def _fecha_es(fecha):
    """Devuelve '30 de agosto de 2025' desde datetime/str."""
    try:
        # Si tienes un mapeo central usa ese primero:
        try:
            from reportes import config_pdf
            MESES = [config_pdf.MESES_ES.get(m, m) for m in
                     ["January","February","March","April","May","June",
                      "July","August","September","October","November","December"]]
        except Exception:
            MESES = ["enero","febrero","marzo","abril","mayo","junio",
                     "julio","agosto","septiembre","octubre","noviembre","diciembre"]

        if isinstance(fecha, str):
            fecha = pd.to_datetime(fecha, errors="coerce")
        if isinstance(fecha, pd.Timestamp):
            fecha = fecha.to_pydatetime()
        if isinstance(fecha, datetime):
            return f"{fecha.day} de {MESES[fecha.month-1]} de {fecha.year}"
    except Exception:
        pass
    return None


class Reporte7:
    def __init__(self):
        self.name_db = name_db
        self.tablas_por_empresa = tablas_por_empresa_reporte1
        self.empresas = empresas
        self.claves_construccion = cargar_json("claves_construccion.json")
        
        # Compilar regex (insensible a mayúsculas)
        patron = "|".join(map(re.escape, self.claves_construccion))
        self.regex_construccion = re.compile(patron, flags=re.IGNORECASE)

    @staticmethod
    def _sanitize_name(texto: str, maxlen: int = 48) -> str:
        base = re.sub(r'[^A-Za-z0-9]+', '_', str(texto).strip())
        return (base[:maxlen]).strip('_')

    def _append_total_row(self, df: pd.DataFrame, label="TOTAL") -> pd.DataFrame:
        if df.empty:
            return df
        total = {"CLIENTE": label}
        for col in df.columns:
            if col == "CLIENTE":
                continue
            serie = pd.to_numeric(df[col], errors="coerce")
            total[col] = float(serie.sum(skipna=True))
        return pd.concat([df, pd.DataFrame([total])], ignore_index=True)

    def _append_total_row_final(self, df: pd.DataFrame, label="TOTAL") -> pd.DataFrame:
        if df.empty:
            return df
        # No sumar el separador "CARTERA VENCIDA"
        mask = df["CLIENTE"].astype(str).str.upper() != "CARTERA VENCIDA"
        total = {"CLIENTE": label}
        for col in df.columns:
            if col == "CLIENTE":
                continue
            serie = pd.to_numeric(df.loc[mask, col], errors="coerce")
            total[col] = float(serie.sum(skipna=True))
        return pd.concat([df, pd.DataFrame([total])], ignore_index=True)

    def _leer_pagares_unificados(self) -> pd.DataFrame:
        frames = []
        for empresa in self.empresas:
            if empresa not in self.tablas_por_empresa:
                print(f"⚠️ Empresa '{empresa}' no tiene tablas cargadas. Se omite.")
                continue

            try:
                tablas = self.tablas_por_empresa[empresa]
                documentos = f'"{tablas["documentos"]}"'

                with sqlite3.connect(self.name_db) as conn:
                    # --- Pagarés (ID 15) ---
                    pag = pd.read_sql_query(f"""
                        SELECT CRAZONSOCIAL, CTOTAL, CPENDIENTE,
                            CIMPORTEEXTRA1, CCANCELADO, CDEVUELTO,
                            CFECHA, CFECHAVENCIMIENTO, CREFERENCIA,
                            CSERIEDOCUMENTO
                        FROM {documentos}
                        WHERE CIDDOCUMENTODE = 15
                        AND CCANCELADO = 0
                        AND CDEVUELTO = 0
                    """, conn)

                    # --- Facturas (ID 4) ---
                    fac = pd.read_sql_query(f"""
                        SELECT CSERIEDOCUMENTO, CFOLIO
                        FROM {documentos}
                        WHERE CIDDOCUMENTODE = 4
                    """, conn)

                if pag.empty:
                    print(f"ℹ️ Empresa '{empresa}' no tiene pagarés activos.")
                    continue

                # Normaliza fechas y montos
                pag["CFECHA"] = pd.to_datetime(pag.get("CFECHA"), errors="coerce")
                pag["CFECHAVENCIMIENTO"] = pd.to_datetime(pag.get("CFECHAVENCIMIENTO"), errors="coerce")
                pag["FECHA_REF"] = pag["CFECHAVENCIMIENTO"].fillna(pag["CFECHA"])
                for col in ["CTOTAL", "CPENDIENTE"]:
                    pag[col] = pd.to_numeric(pag[col], errors="coerce").fillna(0.0)
                pag["CIMPORTEEXTRA1"] = pd.to_numeric(pag.get("CIMPORTEEXTRA1"), errors="coerce").fillna(0).astype(int)
                pag["Pagado"] = pag["CTOTAL"] - pag["CPENDIENTE"]
                pag["Empresa"] = empresa

                # === Claves ===
                # Clave pagaré: a partir de CREFERENCIA (quitando espacios, mayúsculas, quitando prefijo F/R)
                pag["Clave"] = (
                    pag.get("CREFERENCIA", "").astype(str)
                    .str.replace(" ", "", regex=False)
                    .str.strip().str.upper()
                    .str.replace(r'^[FR]', '', regex=True)
                )

                # Clave factura: SERIE + FOLIO (sin espacios, mayúsculas, quitando prefijo F/R)
                fac = fac.dropna(subset=["CSERIEDOCUMENTO", "CFOLIO"]).copy()
                fac["Clave"] = (
                    fac["CSERIEDOCUMENTO"].astype(str).str.strip().str.upper()
                    + fac["CFOLIO"].astype(str).str.strip().str.upper()
                )
                fac["Clave"] = fac["Clave"].str.replace(r'^[FR]', '', regex=True)

                # Si hay facturas duplicadas por clave, nos quedamos con la primera (o la más antigua si quieres ordenar)
                fac = fac.drop_duplicates(subset=["Clave"], keep="first")

                # Merge: pagaré ← serie de FACTURA
                pag = pag.merge(
                    fac[["Clave", "CSERIEDOCUMENTO"]].rename(columns={"CSERIEDOCUMENTO": "CSERIEDOCUMENTO_FACT"}),
                    on="Clave", how="left"
                )
                
                pag = pag.rename(columns={"CSERIEDOCUMENTO": "CSERIEDOCUMENTO_PAGARE"})


                # Métricas de match
                tot = len(pag)
                con_match = pag["CSERIEDOCUMENTO_FACT"].notna().sum()
                print(f"🔗 {empresa}: pagarés emparejados con factura por Clave = {con_match}/{tot} ({(con_match/tot*100 if tot else 0):.1f}%)")

                # Listo para usar aguas abajo
                frames.append(
                    pag[[
                        "Empresa","CRAZONSOCIAL","CTOTAL","CPENDIENTE","Pagado",
                        "CIMPORTEEXTRA1","FECHA_REF", "CSERIEDOCUMENTO_PAGARE", "CSERIEDOCUMENTO_FACT","Clave"
                    ]]
                )

            except Exception as e:
                print(f"❌ Error al procesar empresa '{empresa}': {e}")

        if not frames:
            return pd.DataFrame(columns=[
                "Empresa","CRAZONSOCIAL","CTOTAL","CPENDIENTE","Pagado",
                "CIMPORTEEXTRA1","FECHA_REF", "CSERIEDOCUMENTO_PAGARE", "CSERIEDOCUMENTO_FACT","Clave"
            ])
        return pd.concat(frames, ignore_index=True)

    
    def _split_por_serie(self, df: pd.DataFrame):
        if df.empty:
            return df.copy(), df.copy()

        # Prioriza la serie de factura
        if "CSERIEDOCUMENTO_FACT" in df.columns:
            serie = df["CSERIEDOCUMENTO_FACT"].astype(str).str.upper()
        else:
            # Fallback (no ideal, pero por si acaso)
            serie = df.get("CSERIEDOCUMENTO", "").astype(str).str.upper()

        mask_construccion = serie.str.contains(self.regex_construccion, na=False)
        df_construccion = df[mask_construccion].copy()
        df_terreno = df[~mask_construccion].copy()
        return df_terreno, df_construccion




    def _guardar_tablas_por_cliente(self, df_pagares: pd.DataFrame) -> dict:
        clientes = sorted(df_pagares["CRAZONSOCIAL"].dropna().tolist())
        by_cliente = {}
        with sqlite3.connect(self.name_db) as conn:
            for cliente in clientes:
                try:
                    dfx = df_pagares[df_pagares["CRAZONSOCIAL"] == cliente].copy()
                    if dfx.empty:
                        continue
                    dfx.sort_values("FECHA_REF", inplace=True)
                    tname = f'Pagares_Cliente_{self._sanitize_name(cliente)}'
                    # guarda con la serie de FACTURA incluida
                    dfx.to_sql(tname, conn, index=False, if_exists="replace")
                    by_cliente[cliente] = dfx
                except Exception as e:
                    print(f"❌ Error guardando tabla de cliente '{cliente}': {e}")
        print(f"✅ Tablas por cliente generadas: {len(by_cliente)}")
        self._por_cliente = by_cliente
        return by_cliente

    
    def _renombrar_cols_cliente(self, df: pd.DataFrame) -> pd.DataFrame:
        """Renombra columnas para exportación a Excel."""
        mapping = {
            "CRAZONSOCIAL": "Cliente",
            "CTOTAL": "Pagaré del mes",
            "CPENDIENTE": "Pendiente",
            "FECHA_REF": "Fecha de vencimiento",
        }
        dfx = df.copy()
        dfx.rename(columns={k: v for k, v in mapping.items() if k in dfx.columns}, inplace=True)

        # Formatos opcionales
        if "Cliente" in dfx.columns:
            dfx["Cliente"] = dfx["Cliente"].astype(str).str.upper()
        if "Fecha de vencimiento" in dfx.columns:
            dfx["Fecha de vencimiento"] = pd.to_datetime(
                dfx["Fecha de vencimiento"], errors="coerce"
            ).dt.date  # fecha corta
        for c in ("Pagaré del mes", "Pendiente", "Pagado"):
            if c in dfx.columns:
                dfx[c] = pd.to_numeric(dfx[c], errors="coerce")

        # Orden sugerido
        orden = [c for c in [
            "Empresa", "Cliente", "Fecha de vencimiento", "Pagaré del mes", "Pendiente", "Pagado", "CIMPORTEEXTRA1"
        ] if c in dfx.columns]
        resto = [c for c in dfx.columns if c not in orden]
        return dfx[orden + resto]

    
    def _tablas_por_cliente_para_excel_por_tipo(self, tipo: str) -> dict:
        if not hasattr(self, "_por_cliente") or not self._por_cliente:
            return {}

        hojas, usados = {}, set()
        tipo = (tipo or "").lower().strip()
        sufijo = "_TERR" if tipo == "terrenos" else "_CONS"

        for cliente, df in self._por_cliente.items():
            if df is None or df.empty:
                continue

            col_ref = "CSERIEDOCUMENTO_FACT" if "CSERIEDOCUMENTO_FACT" in df.columns else "CSERIEDOCUMENTO"
            if col_ref not in df.columns:
                continue

            serie = df[col_ref].astype(str).str.upper()
            es_cons = serie.str.contains(self.regex_construccion, na=False)

            dfx = df[es_cons].copy() if tipo == "construcciones" else df[~es_cons].copy()
            if dfx.empty:
                continue

            dfx = self._renombrar_cols_cliente(dfx)

            base = (self._sanitize_name(cliente) or "CLIENTE")[:26]
            nombre = f"{base}{sufijo}"
            i = 1
            while nombre.upper() in usados or len(nombre) > 31:
                nombre = (f"{base}{sufijo}_{i}")[:31]
                i += 1
            usados.add(nombre.upper())
            hojas[nombre] = dfx

        return hojas



    # --------- Núcleo: construir tablas para un DF de UNA empresa ---------
    def _construir_tablas_por_df(self, df_emp: pd.DataFrame,
                                 inicio_mes: pd.Timestamp,
                                 fin_corte_excl: pd.Timestamp,
                                 usar_mes_actual: bool = False):
        """
        Reglas:
        - VENCIDO (interés=0): FECHA_REF < HOY (día de generación del reporte) y CPENDIENTE > 0. (Independiente del corte)
        - PAGARÉ DEL MES (interés=0):
            * usar_mes_actual=True  -> sumar TODO el mes actual [1°, 1° del siguiente)
            * usar_mes_actual=False -> sumar [fecha_corte, 1° del siguiente mes)
        - Cartera vencida (interés=1): totales sin filtro de fecha.
        """
        # Fechas base
        hoy = pd.Timestamp.today().normalize()
        base = hoy if usar_mes_actual else pd.to_datetime(fin_corte_excl).normalize()
        mes_ini = pd.Timestamp(year=base.year, month=base.month, day=1)
        mes_fin_excl = mes_ini + pd.offsets.MonthBegin(1)  # primer día del mes siguiente
        fecha_corte = base  # solo para el caso usar_mes_actual=False

        # Separar universos
        df_ok = df_emp[df_emp["CIMPORTEEXTRA1"] == 0].copy()  # interés = 0
        df_cv = df_emp[df_emp["CIMPORTEEXTRA1"] == 1].copy()  # cartera vencida interés = 1

        # Tipos / limpieza
        for d in (df_ok, df_cv):
            if "FECHA_REF" in d.columns and not pd.api.types.is_datetime64_any_dtype(d["FECHA_REF"]):
                d["FECHA_REF"] = pd.to_datetime(d["FECHA_REF"], errors="coerce")
            for col in ("CTOTAL", "CPENDIENTE"):
                if col in d.columns:
                    d[col] = pd.to_numeric(d[col], errors="coerce").fillna(0.0)

        # VENCIDO (independiente de la fecha de corte)
        vencidos = df_ok[(df_ok["FECHA_REF"] < hoy) & (df_ok["CPENDIENTE"] > 0)].copy()
        grp_venc = (
            vencidos.groupby("CRAZONSOCIAL", as_index=False)
            .agg(VENCIDO=("CPENDIENTE", "sum"))
            .rename(columns={"CRAZONSOCIAL": "CLIENTE"})
        )

        # PAGARÉ DEL MES
        if usar_mes_actual:
            # Todo el mes actual, sin importar el día
            del_mes = df_ok[(df_ok["FECHA_REF"] >= mes_ini) & (df_ok["FECHA_REF"] < mes_fin_excl)].copy()
        else:
            # Desde la fecha de corte manual hasta fin de mes
            del_mes = df_ok[(df_ok["FECHA_REF"] >= fecha_corte) & (df_ok["FECHA_REF"] < mes_fin_excl)].copy()

        if del_mes.empty:
            grp_mes = pd.DataFrame(columns=["CLIENTE", "PAGARÉ DEL MES", "PAGADO", "POR_COBRAR"])
        else:
            grp_mes = (
                del_mes.groupby("CRAZONSOCIAL", as_index=False)
                .agg(
                    **{
                        "PAGARÉ DEL MES": ("CTOTAL", "sum"),
                        "POR_COBRAR": ("CPENDIENTE", "sum"),
                    }
                )
                .rename(columns={"CRAZONSOCIAL": "CLIENTE"})
            )
            grp_mes["PAGADO"] = grp_mes["PAGARÉ DEL MES"] - grp_mes["POR_COBRAR"]

        # Unir corriente (mes) con vencido
        df_corriente = pd.merge(grp_mes, grp_venc, on="CLIENTE", how="outer").fillna(
            {"PAGARÉ DEL MES": 0.0, "PAGADO": 0.0, "POR_COBRAR": 0.0, "VENCIDO": 0.0}
        )
        if not df_corriente.empty:
            df_corriente["CLIENTE"] = df_corriente["CLIENTE"].astype(str).str.upper()
            df_corriente.sort_values("CLIENTE", inplace=True, ignore_index=True)

        # Cartera vencida (interés=1)
        if not df_cv.empty:
            df_cv["Pagado"] = df_cv["CTOTAL"] - df_cv["CPENDIENTE"]
            df_cv_agg = (
                df_cv.groupby("CRAZONSOCIAL", as_index=False)
                .agg(CUENTA_VENCIDA=("CTOTAL", "sum"),
                     COBRADO=("Pagado", "sum"))
                .rename(columns={"CRAZONSOCIAL": "CLIENTE"})
            )
            df_cv_agg["CLIENTE"] = df_cv_agg["CLIENTE"].astype(str).str.upper()
            df_cv_agg.sort_values("CLIENTE", inplace=True, ignore_index=True)
        else:
            df_cv_agg = pd.DataFrame(columns=["CLIENTE", "CUENTA_VENCIDA", "COBRADO"])

        # Construcción final con separador
        cols = ["CLIENTE", "CUENTA VENCIDA", "COBRADO", "PAGARÉ DEL MES", "PAGADO", "POR COBRAR", "VENCIDO"]

        a = (
            pd.DataFrame({
                "CLIENTE": df_corriente["CLIENTE"] if not df_corriente.empty else pd.Series(dtype=str),
                "CUENTA VENCIDA": pd.NA,
                "COBRADO": pd.NA,
                "PAGARÉ DEL MES": df_corriente.get("PAGARÉ DEL MES", pd.Series(dtype=float)),
                "PAGADO": df_corriente.get("PAGADO", pd.Series(dtype=float)),
                "POR COBRAR": df_corriente.get("POR_COBRAR", pd.Series(dtype=float)),
                "VENCIDO": df_corriente.get("VENCIDO", pd.Series(dtype=float)),
            })
            if not df_corriente.empty else pd.DataFrame(columns=cols)
        )

        sep = pd.DataFrame([{"CLIENTE": "CARTERA VENCIDA"}])

        c = (
            pd.DataFrame({
                "CLIENTE": df_cv_agg["CLIENTE"],
                "CUENTA VENCIDA": df_cv_agg.get("CUENTA_VENCIDA", 0.0),
                "COBRADO": df_cv_agg.get("COBRADO", 0.0),
                "PAGARÉ DEL MES": pd.NA, "PAGADO": pd.NA, "POR COBRAR": pd.NA, "VENCIDO": pd.NA,
            })
            if not df_cv_agg.empty else pd.DataFrame(columns=cols)
        )

        df_final = pd.concat([a, sep, c], ignore_index=True)[cols]

        # Total global al final
        df_final = self._append_total_row_final(df_final, label="TOTAL")

        return df_corriente, df_cv_agg, df_final

    # ---------- PASO 3: TABLA FINAL ----------
    def generar_tabla_credito_y_cobranza(self,
                                     fecha_reporte: datetime = None,
                                     empresas_sel=None,
                                     usar_mes_actual: bool = False):
        """
        Devuelve un dict con los DataFrames consolidados:
        {"terrenos": df_terr_consolidado, "construcciones": df_cons_consolidado}
        y además guarda en SQLite tablas separadas por tipo y empresa.
        """
        try:
            if fecha_reporte is None:
                fecha_reporte = datetime.now()

            inicio_mes = pd.Timestamp(year=fecha_reporte.year, month=fecha_reporte.month, day=1)
            fin_corte_excl = pd.Timestamp(year=fecha_reporte.year, month=fecha_reporte.month, day=fecha_reporte.day)

            df_pagares = self._leer_pagares_unificados()
            if df_pagares.empty:
                print("ℹ️ No se encontraron pagarés activos en ninguna empresa.")
                return {"terrenos": pd.DataFrame(), "construcciones": pd.DataFrame()}

            _ = self._guardar_tablas_por_cliente(df_pagares)  # auditoría por cliente (sin separar)

            empresas_a_usar = empresas_sel if empresas_sel else list(self.empresas)

            consol_terr, consol_cons = [], []
            with sqlite3.connect(self.name_db) as conn:
                for emp in empresas_a_usar:
                    df_emp = df_pagares[df_pagares["Empresa"] == emp].copy()
                    if df_emp.empty:
                        continue

                    df_terr, df_cons = self._split_por_serie(df_emp)

                    # --- TERRENOS ---
                    if not df_terr.empty:
                        corr_t, cv_t, fin_t = self._construir_tablas_por_df(
                            df_terr, inicio_mes, fin_corte_excl, usar_mes_actual=usar_mes_actual
                        )
                        for dfx in (corr_t, cv_t, fin_t):
                            if "CLIENTE" in dfx.columns:
                                dfx["CLIENTE"] = dfx["CLIENTE"].astype(str).str.upper()
                        suf = self._sanitize_name(emp)
                        try:
                            corr_t.to_sql(f"Clientes_Interes_Terrenos_{suf}", conn, index=False, if_exists="replace")
                            cv_t.to_sql(f"Clientes_Cartera_Terrenos_{suf}", conn, index=False, if_exists="replace")
                            fin_t.to_sql(f"Reporte7_CreditoCobranza_Terrenos_{suf}", conn, index=False, if_exists="replace")
                        except Exception as e:
                            print(f"❌ Error guardando tablas Terrenos '{emp}': {e}")
                        fin_t = fin_t.copy()
                        fin_t.insert(0, "EMPRESA", emp)
                        consol_terr.append(fin_t)

                    # --- CONSTRUCCIONES ---
                    if not df_cons.empty:
                        corr_c, cv_c, fin_c = self._construir_tablas_por_df(
                            df_cons, inicio_mes, fin_corte_excl, usar_mes_actual=usar_mes_actual
                        )
                        for dfx in (corr_c, cv_c, fin_c):
                            if "CLIENTE" in dfx.columns:
                                dfx["CLIENTE"] = dfx["CLIENTE"].astype(str).str.upper()
                        suf = self._sanitize_name(emp)
                        try:
                            corr_c.to_sql(f"Clientes_Interes_Construcciones_{suf}", conn, index=False, if_exists="replace")
                            cv_c.to_sql(f"Clientes_Cartera_Construcciones_{suf}", conn, index=False, if_exists="replace")
                            fin_c.to_sql(f"Reporte7_CreditoCobranza_Construcciones_{suf}", conn, index=False, if_exists="replace")
                        except Exception as e:
                            print(f"❌ Error guardando tablas Construcciones '{emp}': {e}")
                        fin_c = fin_c.copy()
                        fin_c.insert(0, "EMPRESA", emp)
                        consol_cons.append(fin_c)

                # Consolidado global por tipo
                if consol_terr:
                    df_terr_consol = pd.concat(consol_terr, ignore_index=True)
                else:
                    # Genera global por si no hubo por empresa (inusual)
                    df_terr, _ = self._split_por_serie(df_pagares)
                    _, _, df_terr_consol = self._construir_tablas_por_df(
                        df_terr, inicio_mes, fin_corte_excl, usar_mes_actual=usar_mes_actual
                    )

                if consol_cons:
                    df_cons_consol = pd.concat(consol_cons, ignore_index=True)
                else:
                    _, df_cons = self._split_por_serie(df_pagares)
                    _, _, df_cons_consol = self._construir_tablas_por_df(
                        df_cons, inicio_mes, fin_corte_excl, usar_mes_actual=usar_mes_actual
                    )

                try:
                    df_terr_consol.to_sql("Reporte7_CreditoCobranza_Terrenos", conn, index=False, if_exists="replace")
                    df_cons_consol.to_sql("Reporte7_CreditoCobranza_Construcciones", conn, index=False, if_exists="replace")
                except Exception as e:
                    print(f"❌ Error guardando consolidados Terrenos/Construcciones: {e}")

            print("✅ Reporte7 separados: 'Terrenos' y 'Construcciones' generados y guardados.")
            return {"terrenos": df_terr_consol, "construcciones": df_cons_consol}

        except Exception as e:
            print(f"❌ Error general en Reporte7 (split TyC): {e}")
            return {"terrenos": pd.DataFrame(), "construcciones": pd.DataFrame()}


    def generar_pdf(self, empresas=None, fecha_corte=None, usar_mes_actual: bool = False, tipo: str = "ambos"):
        """
        tipo: "terrenos" | "construcciones" | "ambos"
        Lee las tablas finales separadas en SQLite y genera 1 o 2 PDFs/Excels.
        """
        try:
            fecha_actual = datetime.now().strftime("%d-%m-%Y")

            # Encabezado fecha de corte
            try:
                texto_fecha = _fecha_es(fecha_corte)
                periodo_reporte = f"Fecha de corte: {texto_fecha}" if texto_fecha else "Fecha de corte: No especificada"
            except Exception as e:
                print(f"❌ Error determinando fecha de corte: {e}")
                periodo_reporte = "Fecha de corte: No especificada"

            
            # Helper local para no repetir
            def _emitir(tipo_nombre: str, tabla_sql: str, titulo_pdf: str):
                with sqlite3.connect(self.name_db) as conn:
                    try:
                        df_resumen = pd.read_sql(f"SELECT * FROM {tabla_sql}", conn)
                        if "EMPRESA" in df_resumen.columns:
                            df_resumen = df_resumen.drop(columns=["EMPRESA"])
                    except Exception as e:
                        print(f"❌ No se pudo leer '{tabla_sql}': {e}")
                        return False

                if df_resumen.empty:
                    messagebox.showinfo("Sin datos", f"No hay información para generar el PDF de {tipo_nombre}.")
                    return False

                try:
                    pdf = ReportePDF(
                        titulo_reporte=titulo_pdf,
                        periodo_reporte=periodo_reporte,
                        empresas=empresas
                    )
                except TypeError:
                    pdf = ReportePDF()
                    setattr(pdf, "titulo_reporte", titulo_pdf)
                    setattr(pdf, "periodo_reporte", periodo_reporte)
                    setattr(pdf, "empresas", empresas)

                pdf.add_page()
                pdf.add_subtitulo("Tabla de cobranza y cartera vencida")
                pdf.add_tabla(df_resumen)

                tipo_capital = tipo_nombre.capitalize()
                
                # Excel (solo resumen por tipo + opcional agregar hojas por cliente si te late filtrarlas)
                excel_dict = {f"Tabla crédito y cobranza_{tipo_capital}": df_resumen}

                
                try:
                    excel_dict.update(self._tablas_por_cliente_para_excel_por_tipo(tipo_nombre))
                except Exception as e:
                    print(f"⚠️ No se pudieron preparar hojas por cliente ({tipo_nombre}): {e}")
                
                # Guardados
                try:
                    # PDF
                    nombre_pdf = f"Reporte Crédito y Cobranza {tipo_nombre.capitalize()} {fecha_actual}.pdf"
                    ruta_pdf = filedialog.asksaveasfilename(
                        defaultextension=".pdf",
                        initialfile=nombre_pdf,
                        filetypes=[("Archivos PDF", "*.pdf")],
                        title=f"Guardar PDF ({tipo_nombre})"
                    )
                    if ruta_pdf:
                        pdf.guardar(ruta_pdf)
                        print(f"✅ PDF {tipo_nombre} guardado en: {ruta_pdf}")
                    else:
                        print(f"⚠️ Guardado del PDF {tipo_nombre} cancelado.")
                        return False

                    # Excel
                    nombre_xlsx = f"Tablas Crédito y Cobranza {tipo_nombre.capitalize()} {fecha_actual}.xlsx"
                    ruta_excel = filedialog.asksaveasfilename(
                        defaultextension=".xlsx",
                        initialfile=nombre_xlsx,
                        filetypes=[("Archivos Excel", "*.xlsx")],
                        title=f"Guardar Excel ({tipo_nombre})"
                    )
                    if ruta_excel:
                        try:
                            pdf.guardar_excel(ruta_excel, excel_dict)
                            print(f"📊 Excel {tipo_nombre} guardado en: {ruta_excel}")
                            messagebox.showinfo("Éxito", f"✅ El archivo Excel ({tipo_nombre}) fue guardado correctamente.")
                        except Exception as e:
                            print(f"❗ ERROR al guardar Excel ({tipo_nombre}): {e}")
                            messagebox.showerror("Error", f"❌ Ocurrió un error al guardar el Excel ({tipo_nombre}):\n{e}")
                            return False
                    return True

                except Exception as e:
                    print(f"❗ ERROR diálogos de guardado ({tipo_nombre}): {e}")
                    return False

            ok = True
            if tipo.lower() in ("terrenos", "ambos"):
                ok &= _emitir("terrenos", "Reporte7_CreditoCobranza_Terrenos", "Crédito y Cobranza - Terrenos")
            if tipo.lower() in ("construcciones", "ambos"):
                ok &= _emitir("construcciones", "Reporte7_CreditoCobranza_Construcciones", "Crédito y Cobranza - Construcciones")
            return bool(ok)

        except Exception as e:
            print(f"❌ Error general al generar PDF/Excel por tipo: {e}")
            return False

