# Automatizador de Reportes Empresariales – Grupo Las Hadas

Aplicación de escritorio desarrollada en Python para automatizar la generación de reportes financieros, operativos y administrativos a partir de archivos contables reales.

Este sistema fue diseñado para optimizar procesos internos de análisis y consolidación de información dentro de Grupo Las Hadas, reduciendo tiempos manuales y errores humanos.

---

## Objetivo del sistema
Automatizar el procesamiento masivo de datos contables provenientes de múltiples empresas para generar reportes ejecutivos claros, estructurados y listos para toma de decisiones.

---

## Funcionalidades principales
- Lectura y consolidación automática de archivos Excel contables  
- Procesamiento y análisis de gastos por proyecto
- Cálculo de cuentas vencidas
- Creación de estados de cuenta
- Generación de reportes de materiales y servicios
- Cálculo de aportaciones de Seguro Social por empleado
- Reportes ejecutivos en PDF con diseño profesional  
- Exportación de tablas dinámicas a Excel  
- Base de datos SQLite intermedia para análisis eficiente  
- Sistema modular por tipo de reporte  

---

## Enfoque técnico
El sistema implementa un flujo automatizado:

1. Carga de archivos contables por empresa  
2. Limpieza y normalización de datos con Pandas  
3. Almacenamiento estructurado en SQLite  
4. Generación de tablas por proyecto, periodo y categoría  
5. Exportación automática a PDF y Excel  

Todo el procesamiento se realiza de forma local mediante una interfaz desarrollada en Tkinter.

---

## Tecnologías utilizadas
- Python  
- Pandas  
- SQLite  
- Tkinter  
- FPDF2  
- OpenPyXL / XlsxWriter  
- PyInstaller  

---

## Estructura general del proyecto

  automatizador-reportes/
  ├── ui/ # Interfaces de cada reporte
  ├── core/ # Procesamiento y lógica
  ├── database/ # Manejo de SQLite
  ├── pdf_utils/ # Generación de PDFs
  ├── config/ # Diccionarios dinámicos JSON
  ├── assets/ # Logos e imágenes
  └── main.py

---

## Tipos de reportes generados
- Gastos por proyecto
- Cuentas Vencidas
- Estados de Cuenta
- Gastos por encargado de obra  
- Reporte de materiales y servicios por proyecto    
- Reporte de costos por proyecto
- Reporte de Crédito y Cobranza
- Reporte de seguros (IMSS, SAR, INFONAVIT)  

---

## Aplicación empresarial real
El sistema fue desarrollado para uso operativo dentro de Grupo Las Hadas, permitiendo:

- Automatizar análisis financiero mensual  
- Detectar desviaciones de gasto  
- Consolidar información de múltiples empresas  
- Generar reportes ejecutivos listos para dirección  

---

Proyecto desarrollado como solución tecnológica aplicada a procesos empresariales reales mediante automatización avanzada en Python.


