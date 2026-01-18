import sqlite3
from matplotlib.ticker import FuncFormatter
import pandas as pd
from datetime import datetime
from fpdf import FPDF
import matplotlib.pyplot as plt
import os
from tkinter import filedialog, messagebox
from reportes.config import tablas_por_empresa_reporte1, name_db, empresas, cargar_json
from reportes.config_pdf import nombre_reporte
from reportes.pdf_utils import ReportePDF
import platform
import subprocess
import tempfile
import re
from openpyxl import load_workbook
from tkinter import messagebox


class Reporte1:
    def __init__(self):
        self.name_db = name_db
        self.empresas = empresas
        self.tablas_por_empresa = tablas_por_empresa_reporte1
        # Excluir "Otros" antes de usarlo
        proyectos_filtrados = {
            privada: proyectos
            for privada, proyectos in cargar_json("proyectos_por_privada.json").items()
            if privada != "Otros"
        }
        self.proyectos_por_privada = proyectos_filtrados
        self.tabla_maestra = None

    def generar_tabla_gastos_por_proyecto(self, años=None, meses=None, empresas=None, privadas=None, proyectos=None):
        try:
            print("🔍 Generando tabla de gastos por proyecto con parámetros seleccionados...")

            mapa_meses = {
                "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
                "05": "Mayo", "06": "Junio", "07": "Julio", "08": "Agosto",
                "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
            }

            años = años or []
            meses = meses or []
            empresas = empresas or list(self.empresas)
            privadas = privadas or list(self.proyectos_por_privada.keys())  # ← base en el nuevo diccionario

            # Construir el mapa proyecto → nombre usando las privadas seleccionadas
            proyectos_dict = {
                codigo: nombre
                for privada in privadas
                for codigo, nombre in self.proyectos_por_privada.get(privada, {}).items()
            }

            # Filtrar por proyectos seleccionados si aplica
            if proyectos:
                proyectos_validos = set(proyectos_dict.keys()) & set(proyectos)
            else:
                proyectos_validos = set(proyectos_dict.keys())

            resultados = []

            for empresa in empresas:
                
                try:
                    
                    print(f"🔍 Revisando empresa: {empresa}")
                    tablas = self.tablas_por_empresa[empresa]
                    movimientos = f'"{tablas["movimientos"]}"'
                    documentos = f'"{tablas["documentos"]}"'

                    with sqlite3.connect(self.name_db) as conn:
                        df_mov = pd.read_sql_query(f"""
                            SELECT m.CIDDOCUMENTO, m.CSCMOVTO, m.CTOTAL, m.CFECHA
                            FROM {movimientos} m
                            WHERE m.CIDDOCUMENTODE = 19
                        """, conn)

                        df_doc = pd.read_sql_query(f"""
                            SELECT d.CIDDOCUMENTO, d.CSERIEDOCUMENTO
                            FROM {documentos} d
                        """, conn)

                    df = pd.merge(df_mov, df_doc, on="CIDDOCUMENTO", how="left")
                    df["CFECHA"] = pd.to_datetime(df["CFECHA"], errors="coerce")
                    df = df[pd.to_numeric(df["CSCMOVTO"], errors="coerce").notna()]
                    df["CodigoProyecto"] = pd.to_numeric(df["CSCMOVTO"], errors="coerce")
                    df = df[df["CodigoProyecto"].notna()]
                    df["CodigoProyecto"] = df["CodigoProyecto"].astype(int)


                    # Filtrar por proyectos válidos
                    df = df[df["CodigoProyecto"].isin(proyectos_validos)]

                    # Mapear a nombre de proyecto
                    df["Proyecto"] = df["CodigoProyecto"].map(proyectos_dict)
                    df = df.dropna(subset=["Proyecto", "CFECHA"])

                    # Agregar columna Mes
                    df["Mes"] = df["CFECHA"].dt.year.astype(str) + "-" + df["CFECHA"].dt.strftime("%m").map(mapa_meses)

                    # Aplicar filtros de año y mes si vienen definidos
                    if años:
                        df = df[df["CFECHA"].dt.year.isin([int(a) for a in años])]
                    if meses:
                        df = df[df["Mes"].isin(meses)]

                    resultados.append(df[["Proyecto", "Mes", "CTOTAL"]])

                except Exception as e:
                    print(f"❌ Error procesando empresa {empresa}: {e}")

            if not resultados:
                print("⚠️ No se encontraron datos.")
                return pd.DataFrame()

            # Concatenar resultados
            df_total = pd.concat(resultados, ignore_index=True)

            # Agrupar y pivotear
            df_grouped = df_total.groupby(["Proyecto", "Mes"], as_index=False)["CTOTAL"].sum()
            df_pivot = df_grouped.pivot_table(index="Proyecto", columns="Mes", values="CTOTAL", aggfunc="sum").fillna(0)
            df_pivot = self.ordenar_y_completar_meses(df_pivot)
            df_pivot["Total"] = df_pivot.sum(axis=1)

            # Agregar fila TOTAL
            fila_total = pd.DataFrame([df_pivot.sum(numeric_only=True)], index=["TOTAL"])
            df_final = pd.concat([df_pivot, fila_total])
            df_final.index.name = "Proyecto"

            with sqlite3.connect(self.name_db) as conn:
                df_final.reset_index().to_sql("GastosXProyecto", conn, if_exists="replace", index=False)

            print("✅ Tabla GastosXProyecto generada correctamente.")
            return df_final
        except Exception as e:
            print(f"❌ Error al generar la tabla de gastos por proyecto: {e}")
            messagebox.showerror("Error", f"Error al generar la tabla de gastos por proyecto: {e}")
            return pd.DataFrame()



    def ordenar_columnas_meses_espanol(self, columnas):
        try:
        
            def extraer_orden(col):
                match = re.match(r"(\d{4})-(\d{2})", col)
                if match:
                    return (int(match.group(1)), int(match.group(2)))
                return (9999, 99)  # columnas que no coinciden van al final

            return sorted(columnas, key=extraer_orden)
        except Exception as e:
            print(f"❌ Error al ordenar columnas: {e}")
            return columnas
    
    
    def ordenar_y_completar_meses(self, df_pivot):
        try:

            # Filtramos solo las columnas válidas tipo "2024-Enero"
            columnas_validas = [
                col for col in df_pivot.columns
                if re.match(r"\d{4}-(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)", col)
            ]

            if not columnas_validas:
                return df_pivot

            def extraer_orden(col):
                meses_dict = {
                    "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, "Mayo": 5,
                    "Junio": 6, "Julio": 7, "Agosto": 8, "Septiembre": 9,
                    "Octubre": 10, "Noviembre": 11, "Diciembre": 12
                }
                match = re.match(r"(\d{4})-(\w+)", col)
                if match:
                    anio = int(match.group(1))
                    mes_nombre = match.group(2)
                    mes = meses_dict.get(mes_nombre)
                    return (anio, mes)
                return (9999, 99)

            ordenadas = sorted(columnas_validas, key=extraer_orden)

            primer_anio, primer_mes = extraer_orden(ordenadas[0])
            ultimo_anio, ultimo_mes = extraer_orden(ordenadas[-1])

            # Crear mapa inverso para convertir número → nombre
            meses_nombre = {
                1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo",
                6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre",
                10: "Octubre", 11: "Noviembre", 12: "Diciembre"
            }

            # Generar todos los meses del rango completo
            meses_completos = []
            for anio in range(primer_anio, ultimo_anio + 1):
                mes_inicio = primer_mes if anio == primer_anio else 1
                mes_fin = ultimo_mes if anio == ultimo_anio else 12
                for mes in range(mes_inicio, mes_fin + 1):
                    meses_completos.append(f"{anio}-{meses_nombre[mes]}")

            # Agregar columnas faltantes con cero
            for col in meses_completos:
                if col not in df_pivot.columns:
                    df_pivot[col] = 0

            # Reordenar columnas
            otras_columnas = [col for col in df_pivot.columns if col not in meses_completos]
            df_pivot = df_pivot[meses_completos + otras_columnas]

            return df_pivot
        except Exception as e:
            print(f"❌ Error al ordenar y completar meses: {e}")
            return df_pivot



    def verificar_gastos_proyecto_300(self):
        resultados = []

        for empresa in self.empresas:
            print(f"🔍 Verificando empresa: {empresa}")
            tablas = self.tablas_por_empresa[empresa]
            movimientos = f'"{tablas["movimientos"]}"'
            documentos  = f'"{tablas["documentos"]}"'

            query = f"""
                SELECT
                    m.CIDDOCUMENTO,
                    m.CSCMOVTO,
                    d.CFECHA,
                    m.CTOTAL
                FROM {movimientos} m
                INNER JOIN {documentos} d ON m.CIDDOCUMENTODE = d.CIDDOCUMENTODE
                WHERE d.CIDDOCUMENTODE = 19
                AND m.CSCMOVTO = 300
                AND strftime('%Y', d.CFECHA) = '2025'
            """

            with sqlite3.connect(self.name_db) as conn:
                df = pd.read_sql_query(query, conn)

            if not df.empty:
                df["Empresa"] = empresa
                df["CTOTAL"] = pd.to_numeric(df["CTOTAL"], errors="coerce").fillna(0)
                resultados.append(df)

        if resultados:
            df_total = pd.concat(resultados, ignore_index=True)
            total_general = df_total["CTOTAL"].sum()
            print(f"✅ Total real proyecto 300 (año 2025, todas las empresas): ${total_general:,.2f}")
            return df_total
        else:
            print("⚠️ No se encontraron registros para el proyecto 300 en 2025.")
            return pd.DataFrame()
        
        
    def generar_excel_auditoria_gastos(self, año="2025"):
        resultados = []

        for empresa in self.empresas:
            print(f"🔍 Procesando empresa: {empresa}")
            tablas = self.tablas_por_empresa[empresa]
            movimientos = f'"{tablas["movimientos"]}"'
            documentos  = f'"{tablas["documentos"]}"'

            query = f"""
                SELECT
                    m.CIDDOCUMENTO,
                    m.CSCMOVTO,
                    d.CFECHA,
                    m.CTOTAL
                FROM {movimientos} m
                INNER JOIN {documentos} d ON m.CIDDOCUMENTODE = d.CIDDOCUMENTODE
                WHERE d.CIDDOCUMENTODE = 19
                AND m.CSCMOVTO IS NOT NULL
                AND strftime('%Y', d.CFECHA) = '{año}'
            """

            with sqlite3.connect(self.name_db) as conn:
                df = pd.read_sql_query(query, conn)

            if not df.empty:
                df["Empresa"] = empresa
                df["CTOTAL"] = pd.to_numeric(df["CTOTAL"], errors="coerce").fillna(0)
                resultados.append(df)

        if resultados:
            df_auditoria = pd.concat(resultados, ignore_index=True)

            # Filtrar proyectos válidos
            df_auditoria = df_auditoria[
                df_auditoria["CSCMOVTO"].apply(lambda x: str(x).isdigit() and int(x) in self.proyectos)
            ].copy()

            df_auditoria["CSCMOVTO"] = df_auditoria["CSCMOVTO"].astype(int)
            df_auditoria["NombreProyecto"] = df_auditoria["CSCMOVTO"].apply(
                lambda x: self.proyectos.get(x, "Desconocido")
            )

            # Guardar Excel
            ruta = f"Auditoria_Gastos_{año}.xlsx"
            df_auditoria.to_excel(ruta, index=False)
            print(f"✅ Archivo de auditoría generado: {ruta}")
            return ruta

        else:
            print("⚠️ No se encontraron datos para auditoría.")
            return None
    

    def generar_tablas_gastos_por_proyecto_por_privada(self, años=None, meses=None, empresas=None, privadas=None, proyectos=None):
        try:
            print("🔍 Generando tabla de gastos por proyecto por privada para todos los parámetros...")

            mapa_meses = {
                "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
                "05": "Mayo", "06": "Junio", "07": "Julio", "08": "Agosto",
                "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
            }

            empresas = empresas or self.empresas
            privadas = privadas or self.privadas

            for privada in privadas:
                try:
                    resultados = []
                    print(f"➡️ Procesando privada: {privada}")

                    proyectos_dict = self.proyectos_por_privada.get(privada, {})
                    proyectos_validos = set(proyectos_dict.keys())

                    if proyectos:  # si el usuario filtró manualmente
                        proyectos_validos &= set(proyectos)

                    for empresa in empresas:
                        try: 
                            
                            print(f"🔍 Procesando empresa: {empresa}")
                            
                            tablas = self.tablas_por_empresa[empresa]
                            tabla_mov = tablas["movimientos"]
                            tabla_doc = tablas["documentos"]

                            with sqlite3.connect(self.name_db) as conn:
                                df_mov = pd.read_sql_query(f"""
                                    SELECT
                                        m.CIDDOCUMENTO,
                                        m.CSCMOVTO,
                                        m.CTOTAL,
                                        m.CFECHA
                                    FROM "{tabla_mov}" m
                                    WHERE m.CIDDOCUMENTODE = 19
                                """, conn)

                                df_doc = pd.read_sql_query(f"""
                                    SELECT
                                        d.CIDDOCUMENTO,
                                        d.CSERIEDOCUMENTO
                                    FROM "{tabla_doc}" d
                                """, conn)

                            # Unión de tablas
                            df = pd.merge(df_mov, df_doc, on="CIDDOCUMENTO", how="left")
                            df["CFECHA"] = pd.to_datetime(df["CFECHA"], errors="coerce")

                            # Filtrar por CSCMOVTO numérico y válido
                            df = df[pd.to_numeric(df["CSCMOVTO"], errors="coerce").notna()]
                            df["CodigoProyecto"] = df["CSCMOVTO"].astype(int)

                            # Filtrar por proyectos válidos
                            df = df[df["CodigoProyecto"].isin(proyectos_validos)]
                            if df.empty:
                                continue

                            # Mapear nombre de proyecto
                            df["Proyecto"] = df["CodigoProyecto"].map(proyectos_dict)
                            df = df.dropna(subset=["Proyecto", "CFECHA"])

                            df["CTOTAL"] = pd.to_numeric(df["CTOTAL"], errors="coerce").fillna(0)

                            # Filtro por año si aplica
                            if años:
                                df = df[df["CFECHA"].dt.year.isin([int(a) for a in años])]

                            # Columna mes con formato 2025-Julio
                            df["Mes"] = (
                                df["CFECHA"].dt.year.astype(str) + "-" +
                                df["CFECHA"].dt.strftime("%m").map(mapa_meses).fillna("Desconocido")
                            )

                            # Filtro por mes si aplica
                            if meses:
                                df = df[df["Mes"].isin(meses)]

                            # Agrupación
                            df_grouped = df.groupby(["Proyecto", "Mes"], as_index=False)["CTOTAL"].sum()
                            df_grouped.rename(columns={"CTOTAL": "Total"}, inplace=True)
                            resultados.append(df_grouped)


                        except Exception as e:
                            print(f"❌ Error procesando empresa {empresa} para privada {privada}: {e}")
                            continue

                    if resultados:
                        df_total = pd.concat(resultados, ignore_index=True)
                        
                        if df_total.empty:
                            print(f"⚠️ No se encontraron datos reales para la privada {privada}.")
                            continue  # Saltar esta privada sin guardar tabla vacía

                        df_pivot = df_total.pivot_table(index="Proyecto", columns="Mes", values="Total", aggfunc="sum").fillna(0)
                        df_pivot.columns.name = None
                        df_pivot = df_pivot.rename_axis(None, axis=1)
                        df_pivot = self.ordenar_y_completar_meses(df_pivot)
                        df_pivot["Total"] = df_pivot.sum(axis=1)
                        df_final = df_pivot.reset_index()

                        # Agregar fila TOTAL al final
                        fila_total = df_pivot.sum(numeric_only=True)
                        fila_total["Proyecto"] = "TOTAL"
                        df_final = pd.concat([df_final, pd.DataFrame([fila_total])], ignore_index=True)


                        table_name = f"GastosXProyecto_{privada}"
                        with sqlite3.connect(self.name_db) as conn:
                            df_final.to_sql(table_name, conn, if_exists="replace", index=False)

                        print(f"✅ Tabla '{table_name}' generada correctamente.")
                    else:
                        print(f"⚠️ No se encontraron datos para la privada {privada}.")
                
                except Exception as e:
                    print(f"❌ Error procesando privada {privada}: {e}")
                    continue
                
            return True
            
                
        except Exception as e:
            print(f"❌ Error al generar tablas de gastos por proyecto por privada: {e}")
            messagebox.showerror("Error", f"Error al generar tablas de gastos por proyecto por privada: {e}")
            return False

    
    def verificar_proyectos_en_multiples_privadas(self):
        print("🔎 Revisando proyectos en múltiples privadas...\n")

        resultados = []

        for empresa in self.empresas:
            tablas = self.tablas_por_empresa[empresa]
            tabla_mov = tablas["movimientos"]
            tabla_doc = tablas["documentos"]

            with sqlite3.connect(self.name_db) as conn:
                df_mov = pd.read_sql_query(f"""
                    SELECT
                        CIDDOCUMENTO,
                        CSCMOVTO AS CodigoProyecto
                    FROM "{tabla_mov}"
                    WHERE CIDDOCUMENTODE = 19
                    AND CSCMOVTO IS NOT NULL
                """, conn)

                df_doc = pd.read_sql_query(f"""
                    SELECT
                        CIDDOCUMENTO,
                        CSERIEDOCUMENTO
                    FROM "{tabla_doc}"
                """, conn)

            df_doc = df_doc.drop_duplicates(subset="CIDDOCUMENTO")
            df = pd.merge(df_mov, df_doc, on="CIDDOCUMENTO", how="left")
            df = df.dropna(subset=["CSERIEDOCUMENTO"])

            # Obtener privada desde la serie (removiendo la C inicial)
            df["Privada"] = df["CSERIEDOCUMENTO"].str.upper().str.replace("^C", "", regex=True)
            df["CodigoProyecto"] = pd.to_numeric(df["CodigoProyecto"], errors="coerce").astype("Int64")

            df = df.dropna(subset=["CodigoProyecto", "Privada"])

            resultados.append(df[["CodigoProyecto", "Privada"]])

        df_total = pd.concat(resultados, ignore_index=True)

        # Agrupar y contar cuántas privadas tiene cada proyecto
        agrupado = df_total.groupby("CodigoProyecto")["Privada"].nunique().reset_index()
        proyectos_multi_privada = agrupado[agrupado["Privada"] > 1]["CodigoProyecto"].tolist()

        if not proyectos_multi_privada:
            print("✅ Todos los proyectos están asignados a una sola privada.")
        else:
            print("⚠️ Proyectos que aparecen en más de una privada:")
            for codigo in proyectos_multi_privada:
                privadas = df_total[df_total["CodigoProyecto"] == codigo]["Privada"].unique()
                nombre = self.proyectos.get(int(codigo), "Desconocido")
                print(f"  - {codigo} ({nombre}): {', '.join(privadas)}")
            
    
    def mostrar_movimientos_proyecto_en_privadas(self, proyectos_a_buscar):
        print("\n📋 Buscando movimientos de proyectos en múltiples privadas...\n")

        resultados = []

        for empresa in self.empresas:
            tablas = self.tablas_por_empresa[empresa]
            tabla_mov = tablas["movimientos"]
            tabla_doc = tablas["documentos"]

            with sqlite3.connect(self.name_db) as conn:
                df_mov = pd.read_sql_query(f"""
                    SELECT
                        CIDDOCUMENTO,
                        CSCMOVTO AS CodigoProyecto,
                        CTOTAL,
                        CFECHA
                    FROM "{tabla_mov}"
                    WHERE CIDDOCUMENTODE = 19
                    AND CSCMOVTO IS NOT NULL
                """, conn)

                df_doc = pd.read_sql_query(f"""
                    SELECT
                        CIDDOCUMENTO,
                        CSERIEDOCUMENTO
                    FROM "{tabla_doc}"
                """, conn)

            df_doc = df_doc.drop_duplicates(subset="CIDDOCUMENTO")
            df = pd.merge(df_mov, df_doc, on="CIDDOCUMENTO", how="left")
            df = df.dropna(subset=["CSERIEDOCUMENTO"])

            df["Privada"] = df["CSERIEDOCUMENTO"].str.upper().str.replace("^C", "", regex=True)
            df["CodigoProyecto"] = pd.to_numeric(df["CodigoProyecto"], errors="coerce").astype("Int64")
            df = df.dropna(subset=["CodigoProyecto", "Privada"])

            df["NombreProyecto"] = df["CodigoProyecto"].map(self.proyectos)
            df["Empresa"] = empresa

            resultados.append(df)

        df_total = pd.concat(resultados, ignore_index=True)

        # Normalizar entrada del usuario
        codigos_a_buscar = set()

        for proyecto in proyectos_a_buscar:
            if isinstance(proyecto, int) or str(proyecto).isdigit():
                codigos_a_buscar.add(int(proyecto))
            else:
                # Buscar por nombre en el diccionario inverso
                for codigo, nombre in self.proyectos.items():
                    if nombre.strip().lower() == str(proyecto).strip().lower():
                        codigos_a_buscar.add(codigo)

        # Filtrar resultados
        df_filtrado = df_total[df_total["CodigoProyecto"].isin(codigos_a_buscar)]

        if df_filtrado.empty:
            print("❌ No se encontraron movimientos para los proyectos indicados.")
        else:
            print(f"✅ Movimientos encontrados para proyectos: {', '.join(str(c) for c in codigos_a_buscar)}\n")
            print(df_filtrado[["CIDDOCUMENTO", "Empresa", "Privada", "CodigoProyecto", "NombreProyecto", "CFECHA", "CTOTAL"]])
            
            # Opcional: Guardar como Excel para revisión más fácil
            ruta = "movimientos_problema.xlsx"
            df_filtrado.to_excel(ruta, index=False)
            print(f"\n📝 Archivo guardado como: {ruta}")


    def generar_tabla_historial_gastos_por_proyecto(self, años_disponibles=None, meses_seleccionados=None, empresas_seleccionadas=None, privadas_seleccionadas=None, proyectos_seleccionados=None, almacenes_seleccionados=None):
        try:

            mapa_meses = {
                "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
                "05": "Mayo", "06": "Junio", "07": "Julio", "08": "Agosto",
                "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
            }

            años_disponibles = años_disponibles or [str(datetime.now().year)]
            meses_seleccionados = meses_seleccionados or list(mapa_meses.values())
            empresas_seleccionadas = empresas_seleccionadas or self.empresas
            privadas_seleccionadas = privadas_seleccionadas or self.privadas
            proyectos_seleccionados = proyectos_seleccionados or list(self.proyectos.keys())

            mapa_inverso = {v: k for k, v in mapa_meses.items()}
            años_str = ', '.join([f"'{a}'" for a in años_disponibles])
            meses_str = ', '.join([f"'{mapa_inverso[m]}'" for m in meses_seleccionados if m in mapa_inverso])

            condicion_anio = f"strftime('%Y', d.CFECHA) IN ({años_str})"
            condicion_mes = f"strftime('%m', d.CFECHA) IN ({meses_str})"

            condicion_almacenes = ""
            if almacenes_seleccionados:
                almacenes_str = ', '.join([f"'{a}'" for a in almacenes_seleccionados])
                condicion_almacenes = f"AND a.CNOMBREALMACEN IN ({almacenes_str})"

            condicion_privadas = ""
            if privadas_seleccionadas and any(privadas_seleccionadas):
                privadas_like = ' OR '.join(
                    [f"UPPER(IFNULL(d.CSERIEDOCUMENTO, '')) LIKE '%{privada.upper()}%'" for privada in privadas_seleccionadas]
                )
                if privadas_like:
                    condicion_privadas = f"AND ({privadas_like})"

            resultados = []

            for empresa in empresas_seleccionadas:
                try:
                    tablas = self.tablas_por_empresa[empresa]
                    movimientos = f'"{tablas["movimientos"]}"'
                    documentos = f'"{tablas["documentos"]}"'
                    productos = f'"{tablas["productos"]}"'
                    almacenes = f'"{tablas["almacenes"]}"'

                    query = f"""
                        SELECT
                            m.CSCMOVTO AS ProyectoID,
                            d.CFECHA AS Fecha,
                            d.CTOTAL AS Total,
                            d.CSERIEDOCUMENTO AS Serie
                        FROM {movimientos} m
                        INNER JOIN {documentos} d ON m.CIDDOCUMENTODE = d.CIDDOCUMENTODE
                        INNER JOIN {productos} p ON m.CIDPRODUCTO = p.CIDPRODUCTO
                        INNER JOIN {almacenes} a ON m.CIDALMACEN = a.CIDALMACEN
                        WHERE {condicion_anio}
                        AND {condicion_mes}
                        AND d.CIDDOCUMENTODE = 19
                        {condicion_almacenes}
                        {condicion_privadas}
                    """

                    with sqlite3.connect(self.name_db) as conn:
                        df = pd.read_sql_query(query, conn)

                    if not df.empty:
                        df["Año"] = pd.to_datetime(df["Fecha"]).dt.strftime("%Y")
                        df["Mes"] = pd.to_datetime(df["Fecha"]).dt.strftime("%m").map(mapa_meses)
                        df["Proyecto"] = df["ProyectoID"].apply(lambda pid: self.proyectos.get(pid, "Desconocido"))
                        df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0)

                        df_grouped = df.groupby(["Proyecto", "Año", "Mes"], as_index=False)["Total"].sum()
                        resultados.append(df_grouped)
                except Exception as e:
                    print(f"❌ Error procesando empresa {empresa}: {e}")
                    continue
                    

            if resultados:
                df_final = pd.concat(resultados, ignore_index=True)
                with sqlite3.connect(self.name_db) as conn:
                    df_final.to_sql("HistorialGastosXProyecto", conn, if_exists="replace", index=False)

                print("✅ Tabla HistorialGastosXProyecto guardada.")
                return df_final
            else:
                print("⚠️ No se encontraron datos para el historial de gastos.")
                return pd.DataFrame()
        except Exception as e:
            print(f"❌ Error al generar la tabla de historial de gastos por proyecto: {e}")
            return pd.DataFrame()
        
    def verificar_duplicados(self, df):
        resumen = {}

        # Duplicados completos
        duplicados_completos = df[df.duplicated()]
        resumen['filas_duplicadas_completas'] = len(duplicados_completos)

        # Duplicados por Proyecto, Año, Mes
        if all(col in df.columns for col in ["Proyecto", "Año", "Mes"]):
            duplicados_parciales = df[df.duplicated(subset=["Proyecto", "Año", "Mes"])]
            resumen['duplicados_por_proyecto_año_mes'] = len(duplicados_parciales)
        else:
            resumen['duplicados_por_proyecto_año_mes'] = "No se encontró alguna columna requerida"

        # Primeras filas duplicadas completas como muestra
        if not duplicados_completos.empty:
            resumen['muestra'] = duplicados_completos.head()
        else:
            resumen['muestra'] = "No hay duplicados completos"

        return resumen
    
    def verificar_duplicados_movimientos(self):
        """
        Revisa si existen filas duplicadas exactas en las tablas de movimientos por empresa.
        Imprime resumen con número de duplicados encontrados por empresa.
        """
        print("🔍 Verificando duplicados en movimientos por empresa...")
        for empresa in self.empresas:
            nombre_tabla = self.tablas_por_empresa[empresa]["movimientos"]
            with sqlite3.connect(self.name_db) as conn:
                try:
                    df = pd.read_sql_query(f"SELECT * FROM '{nombre_tabla}'", conn)
                    duplicados = df[df.duplicated(keep=False)]  # Duplicados exactos (todas las columnas)

                    if not duplicados.empty:
                        print(f"⚠️ {empresa}: {len(duplicados)} filas duplicadas encontradas.")
                        print(duplicados.head(3))  # Muestra un ejemplo
                    else:
                        print(f"✅ {empresa}: Sin duplicados exactos.")
                except Exception as e:
                    print(f"❌ Error al revisar {empresa}: {e}")
    
    def verificar_duplicados_logicos(self):
        columnas_clave = [
            "CSERIEDOCUMENTO", "CFOLIO", "CIDPRODUCTO", 
            "CTOTAL", "CFECHA", "CSCMOVTO", "CNOMBREALMACEN"
        ]

        for empresa in self.empresas:
            tablas = self.tablas_por_empresa[empresa]
            movimientos = f'"{tablas["movimientos"]}"'
            documentos = f'"{tablas["documentos"]}"'
            productos = f'"{tablas["productos"]}"'
            almacenes = f'"{tablas["almacenes"]}"'

            query = f"""
                SELECT 
                    d.CSERIEDOCUMENTO, d.CFOLIO, p.CIDPRODUCTO, d.CTOTAL,
                    d.CFECHA, m.CSCMOVTO, a.CNOMBREALMACEN
                FROM {movimientos} m
                INNER JOIN {documentos} d ON m.CIDDOCUMENTODE = d.CIDDOCUMENTODE
                INNER JOIN {productos} p ON m.CIDPRODUCTO = p.CIDPRODUCTO
                INNER JOIN {almacenes} a ON m.CIDALMACEN = a.CIDALMACEN
                WHERE d.CIDDOCUMENTODE = 19
            """

            with sqlite3.connect(self.name_db) as conn:
                df = pd.read_sql_query(query, conn)

            duplicados = df.duplicated(subset=columnas_clave, keep=False)
            df_duplicados = df[duplicados]

            if not df_duplicados.empty:
                print(f"⚠️ Empresa '{empresa}': Se encontraron {df_duplicados.shape[0]} posibles duplicados lógicos.")
                print(df_duplicados.head(10))
            else:
                print(f"✅ Empresa '{empresa}': No se encontraron duplicados lógicos.")
    
    def obtener_tabla(self, nombre_tabla):
        with sqlite3.connect(self.name_db) as conn:
            df = pd.read_sql_query(f"SELECT * FROM {nombre_tabla}", conn)

        if nombre_tabla == "GastosXSolicitante":
            df.set_index(["Solicitante", "Proyecto"], inplace=True)
        elif nombre_tabla == "GastosXProyecto":
            df.set_index("Proyecto", inplace=True)
        elif nombre_tabla.startswith("GastosProyectos_"):
            df.set_index("Proyecto", inplace=True)
        elif nombre_tabla == "HistorialGastosXProyecto":
            df.set_index("Proyecto", inplace=True)

        return df
    
    def obtener_tablas(self):
        try:
            with sqlite3.connect(self.name_db) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
                tablas = [row[0] for row in cursor.fetchall()]
            return tablas
        except Exception as e:
            print(f"❌ Error al obtener tablas de la base de datos: {e}")
            return []
    
    def obtener_nombres_tablas(self):
        with sqlite3.connect(self.name_db) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tablas = [fila[0] for fila in cursor.fetchall()]
        return tablas
    
    def generar_pdf_reporte1(self, años=None, meses=None, empresas=None, privadas=None, proyectos=None, parent=None):
        try:
            with sqlite3.connect(self.name_db) as conn:
            
                def formatear_mes(mes_raw):
                    # Entrada esperada: "2025-Enero"
                    partes = mes_raw.split("-")
                    if len(partes) == 2:
                        año, mes = partes
                        return f"{mes} de {año}"
                    return mes_raw  # en caso de formato inválido

                # Asumiendo que recibes la lista ya ordenada:
                # meses = ["2025-Enero", "2025-Febrero", "2025-Marzo"]

                if meses:
                    if len(meses) == 1:
                        periodo_reporte = f"Periodo: {formatear_mes(meses[0])}"
                    else:
                        periodo_reporte = f"Periodo: {formatear_mes(meses[0])} - {formatear_mes(meses[-1])}"
                else:
                    periodo_reporte = "Periodo: No especificado"
                

                tablas_dict = {}
                df_gastos_proyecto = pd.read_sql_query("SELECT * FROM GastosXProyecto", conn)
                df_gastos_proyecto = df_gastos_proyecto.drop(columns=["index"], errors="ignore")
                tablas_dict['GastosXProyecto'] = df_gastos_proyecto.copy()

                privadas_tablas = [str(nombre) for nombre in self.obtener_tablas() if str(nombre).startswith("GastosXProyecto_")]

                # Crear PDF
                titulo = f"Gastos por proyecto"
                pdf = ReportePDF(
                    titulo_reporte=titulo,
                    periodo_reporte=periodo_reporte,
                    empresas=empresas,
                    privadas=privadas
                )

                pdf.add_page()
                
                # --- 1. Fila TOTAL ---
                pdf.add_subtitulo("Gasto total de los proyecto por mes")

                # Filtrar columnas usando el parámetro 'meses'
                columnas_filtradas = [col for col in df_gastos_proyecto.columns if col in meses]


                if columnas_filtradas and "TOTAL" in df_gastos_proyecto["Proyecto"].values:
                    fila_total = df_gastos_proyecto[df_gastos_proyecto["Proyecto"] == "TOTAL"][columnas_filtradas].round(2)
                    fila_total.index = ["TOTAL"]
                    pdf.add_tabla(fila_total, bloques_de_4=True)
                else:
                    pdf.add_texto("No se encontró la fila 'TOTAL' o no hay datos para los meses seleccionados.")


                # --- 2. Top proyectos por privada ---
                for privada in privadas_tablas:
                    df_privada = pd.read_sql_query(f"SELECT * FROM {privada}", conn)
                    df_privada = df_privada.drop(columns=["index"], errors="ignore")
                    tablas_dict[privada] = df_privada.copy()

                    # Quitar la fila TOTAL en todos los casos (si existe)
                    df_privada = df_privada[df_privada["Proyecto"].str.upper() != "TOTAL"]


                    columnas_validas = [col for col in meses if col in df_privada.columns]
                    if not columnas_validas:
                        continue

                    df_privada["Total Acumulado"] = df_privada[columnas_validas].sum(axis=1)

                    # Usar el último mes del rango seleccionado
                    ultimo_mes = meses[-1]
                    mes_formateado = formatear_mes(ultimo_mes)
                    
                    if ultimo_mes not in df_privada.columns:
                        print(f"⚠️ La privada '{privada}' no tiene columna para {ultimo_mes}, se omite.")
                        continue


                    # Verificar si hay al menos un gasto real (> 0) en el último mes
                    if df_privada[ultimo_mes].fillna(0).astype(float).sum() <= 0:
                        print(f"ℹ️ Se omite la privada '{privada}' porque no tiene gastos en {ultimo_mes}")
                        continue


                    # Calcular Top 5
                    df_top = df_privada.sort_values(ultimo_mes, ascending=False).head(5).copy()
                    df_top.reset_index(drop=True, inplace=True)
                    df_top["Total Acumulado"] = df_top[columnas_validas].sum(axis=1)
                    df_top = df_top[["Proyecto", ultimo_mes, "Total Acumulado"]]
                    df_top[ultimo_mes] = df_top[ultimo_mes].astype(int)
                    df_top["Total Acumulado"] = df_top["Total Acumulado"].astype(int)

                    # Agregar fila TOTAL del top 5
                    fila_total = {
                        "Proyecto": "TOTAL",
                        ultimo_mes: df_top[ultimo_mes].sum(),
                        "Total Acumulado": df_top["Total Acumulado"].sum()
                    }
                    df_top = pd.concat([df_top, pd.DataFrame([fila_total])], ignore_index=True)

                    nombre_privada = privada.replace("GastosXProyecto_", "")
                    pdf.add_subtitulo(f"Top 5 proyectos con más gastos - {nombre_privada} (Mes: {mes_formateado})")
                    pdf.add_tabla(df_top)


            # --- 3. Tabla de gastos por proyecto (según meses seleccionados) ---
            pdf.add_page(orientation='L')
            pdf.add_subtitulo("Gastos por proyecto")

            # 🔹 Renombrar columna Total → Total Acumulado (si existe)
            if "Total" in df_gastos_proyecto.columns:
                df_gastos_proyecto.rename(columns={"Total": "Total Acumulado"}, inplace=True)

            columnas_mes = [col for col in meses if col in df_gastos_proyecto.columns]
            if not columnas_mes:
                pdf.add_texto("No hay columnas de meses disponibles para mostrar.")
            else:
                columnas_fijas = ["Proyecto"]
                tiene_total = "Total Acumulado" in df_gastos_proyecto.columns  # Cambiado

                # Dividir en bloques de 6 meses
                for i in range(0, len(columnas_mes), 6):
                    bloque_meses = columnas_mes[i:i+6]
                    columnas_tabla = columnas_fijas + bloque_meses + (["Total Acumulado"] if tiene_total and i + 6 >= len(columnas_mes) else [])

                    df_bloque = df_gastos_proyecto[columnas_tabla].copy()

                    # Agregar nueva página solo a partir del segundo bloque
                    if i > 0:
                        pdf.add_page(orientation='L')

                    pdf.add_tabla(df_bloque)


            # --- 4. Historial de gastos por proyecto ---
            pdf.add_page(orientation="L")  # ✅ Nueva página horizontal
            pdf.add_subtitulo("Gráfica con los nueve proyectos con más gastos y total acumulado")

            # Filtrar columnas según los meses seleccionados
            columnas_filtradas = [col for col in meses if col in df_gastos_proyecto.columns]
            if not columnas_filtradas:
                print("⚠️ No hay columnas válidas para la gráfica.")
            else:
                # Preparar DataFrame
                df_para_grafica = df_gastos_proyecto.drop(index="TOTAL", errors="ignore").copy()
                df_para_grafica = df_para_grafica.reset_index()
                df_para_grafica.rename(columns={"Proyecto": "NombreProyecto"}, inplace=True)

                df_para_grafica = df_para_grafica[["NombreProyecto"] + columnas_filtradas].set_index("NombreProyecto")

                if df_para_grafica.empty:
                    print("⚠️ DataFrame vacío, no se puede graficar.")
                else:
                    # Mostrar solo los 10 proyectos con mayor gasto total en el periodo
                    df_para_grafica["TotalPeriodo"] = df_para_grafica.sum(axis=1)
                    df_para_grafica = df_para_grafica.sort_values(by="TotalPeriodo", ascending=False).head(10)
                    df_para_grafica = df_para_grafica.drop(columns="TotalPeriodo")

                    # Transponer para que los meses estén en el eje X
                    df_para_grafica_T = df_para_grafica.T

                    if not df_para_grafica_T.empty:
                        fig, ax = plt.subplots(figsize=(10, 6))
                        df_para_grafica_T.plot(ax=ax, marker='o', linewidth=1.5)

                        # Subtítulo con rango si hay varios meses
                        if len(columnas_filtradas) == 1:
                            titulo = f"Evolución de gastos por proyecto ({columnas_filtradas[0]})"
                        else:
                            titulo = f"Evolución de gastos por proyecto ({columnas_filtradas[0]} - {columnas_filtradas[-1]})"

                        ax.set_title(titulo)
                        ax.set_xlabel("Mes")
                        ax.set_ylabel("Monto ($)")
                        ax.grid(True)
                        ax.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f'${x/1e6:.1f}M'))
                        
                        ax.legend(loc="upper left", fontsize=6)
                        fig.tight_layout()

                        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                            fig.savefig(tmp.name)
                            imagen_path = tmp.name
                        plt.close(fig)

                        pdf.add_imagen_centrada_bajo_subtitulo(imagen_path)
                    else:
                        print("⚠️ La gráfica transpuesta está vacía, no se puede graficar.")



           # --- Guardar PDF y Excel ---
            try:
                fecha_actual = datetime.now().strftime("%Y-%m-%d")
                nombre_archivo_pdf = f"Reporte Gastos_{fecha_actual}.pdf"

                print("🧪 LLEGÓ AL MOMENTO DE GUARDAR EL PDF")
                
                # 1. Elegir ruta para guardar el PDF
                ruta_pdf = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    initialfile=nombre_archivo_pdf,
                    filetypes=[("Archivos PDF", "*.pdf")],
                    title="Guardar reporte PDF"
                )

                if not ruta_pdf:
                    print("⚠️ Guardado del PDF cancelado por el usuario.")
                    return False

                # 2. Guardar PDF
                pdf.guardar(ruta_pdf)
                print(f"✅ PDF guardado en: {ruta_pdf}")

                # # 3. Guardar imagen en la misma carpeta
                # if "imagen_path" in locals() and os.path.exists(imagen_path):
                #     imagen_destino = os.path.join(os.path.dirname(ruta_pdf), f"GraficaHistorialGastos_{fecha_actual}.png")
                #     os.replace(imagen_path, imagen_destino)
                #     print(f"✅ Gráfica guardada en: {imagen_destino}")
                # else:
                #     print("⚠️ La imagen de la gráfica no fue generada o no existe.")


                # 4. Guardar Excel (opcional)
                ruta_excel = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=f"Reporte Gastos_{fecha_actual}.xlsx",
                    filetypes=[("Archivos Excel", "*.xlsx")],
                    title="Guardar reporte Excel"
                )

                if ruta_excel:
                    try:
                        pdf.guardar_excel(ruta_excel, tablas_dict)
                        print(f"✅ Excel guardado en: {ruta_excel}")
                        messagebox.showinfo("Éxito", "✅ El archivo Excel fue guardado correctamente.")
                    except Exception as e:
                        print(f"❌ Error al guardar el archivo Excel: {e}")
                        messagebox.showerror("Error", f"❌ Ocurrió un error al guardar el archivo Excel:\n{e}")

                return True  # ✅ Indicar que todo fue exitoso

            except Exception as e:
                print(f"❗ ERROR al intentar abrir diálogo de guardado PDF: {e}")
                return False  # ❌ Algo falló
            
        except Exception as e:
            print(f"❌ Error general al generar el PDF completo: {e}")
            return False

            
    
    @staticmethod
    def ordenar_mes(mes_completo):
        meses_dict = {
            "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4,
            "Mayo": 5, "Junio": 6, "Julio": 7, "Agosto": 8,
            "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12
        }
        match = re.match(r"(\d{4})-(\w+)", mes_completo)
        if match:
            año = int(match.group(1))
            mes_nombre = match.group(2)
            mes_num = meses_dict.get(mes_nombre, 13)
            return (año, mes_num)
        return (9999, 99)
    
    @staticmethod
    def obtener_texto_privadas(privadas_seleccionadas, privadas_disponibles):
        if set(privadas_seleccionadas) == set(privadas_disponibles):
            return "Todas (" + ", ".join(privadas_disponibles) + ")"
        else:
            return ", ".join(privadas_seleccionadas)
