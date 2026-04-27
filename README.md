Sistema de Gestión de Liquidaciones - Core VB6
Este repositorio contiene el núcleo del sistema de gestión de recursos humanos y liquidación de haberes desarrollado en Visual Basic 6.0. Es una solución integral diseñada para el manejo de nóminas, administración de legajos y procesamiento de salarios con salida de documentos para el empleado.

📋 Descripción del Proyecto
El sistema fue concebido para centralizar las operaciones administrativas de una organización, permitiendo un control exhaustivo sobre el historial laboral y financiero de los empleados. Su arquitectura está orientada a la estabilidad y eficiencia en el procesamiento de datos críticos sobre SQL Server.

Características Principales:
Gestión de Legajos: Administración completa de datos personales, CUIL, fechas de ingreso y categorías salariales.

Motor de Liquidación: Procesamiento de conceptos (Haberes y Retenciones) con cálculos automáticos basados en la configuración de la base de datos.

Generador de Recibos: Motor de exportación que transforma los datos de la base en documentos HTML estructurados, utilizando una nomenclatura estandarizada para su archivo: Recibo_LEG-{Legajo}_{Mes}_{Año}.html.

Interfaz de Escritorio: Diseño basado en formularios (MDI) para una navegación rápida y manejo de múltiples ventanas de gestión.

🛠️ Stack Tecnológico
Lenguaje: Visual Basic 6.0 (Legacy Core).

Base de Datos: Microsoft SQL Server.

Conectividad: ADO (ActiveX Data Objects) para persistencia y consultas.

Formato de Salida: HTML dinámico para recibos de sueldo.

🏗️ Estructura de Datos (SQL)
El proyecto se apoya en una estructura relacional sólida:

Tabla Empleados: Datos maestros y estado activo/inactivo.

LiquidacionCabecera: Resumen de totales (Neto, Haberes, Retenciones) y períodos (Mes/Año).

LiquidacionDetalle: Desglose de cada concepto liquidado, permitiendo la trazabilidad total de cada recibo.

🚀 Visión de Evolución y Modernización
Como parte de una estrategia de mejora continua, este proyecto ha sido diseñado para servir como base de datos centralizada.

Nota de migración: Existe un concepto de modernización paralelo orientado a llevar estas funcionalidades a una plataforma Web. La arquitectura de este sistema VB6 permite que nuevas interfaces (como ASP.NET Core Razor Pages) consuman la misma lógica de datos, facilitando una transición tecnológica sin pérdida de información histórica y permitiendo el acceso remoto a los recibos de sueldo generados por este motor.

(Próximamente se publicará el repositorio correspondiente a la versión Web del visualizador).
