# MINEX | Gestión de Relaciones Laborales

## Versión actual: V25.8 — 01/10/2026

Sistema web para la gestión de casos de Relaciones Laborales. La aplicación está preparada para publicarse en GitHub Pages, pero **GitHub no es la base de datos de MINEX**. La operación actual utiliza IndexedDB en el navegador.

## Cambio principal V25.8
- Se corrige la descarga de **Casos** cuando la aplicación se publica en GitHub Pages.
- La descarga ya no depende exclusivamente del archivo JavaScript con la plantilla embebida.
- MINEX intenta primero la plantilla embebida y, si no está disponible, carga el archivo XLSX real incluido en `assets/FORMATO_CARGA_CASOS_RELACIONES_LABORALES.xlsx`.
- La plantilla utilizada conserva la pestaña `VOLUMETRIAS`, fórmulas, estilos, tabla, hojas auxiliares y hojas ocultas del archivo corporativo.
- La exportación de casos utiliza los registros filtrados de MINEX y los coloca sobre la estructura de la plantilla, evitando una descarga plana.
- Se elimina de la interfaz de MINEX la configuración de GitHub como si fuera una base de datos compartida. GitHub queda como infraestructura de publicación del proyecto.
- Se conserva la detección y revisión de duplicados.

## Funcionalidades conservadas
- Dashboard.
- Casos: creación, consulta, edición, cierre y seguimiento.
- Seguimientos y pendientes.
- Informes e indicadores.
- Base de empleados.
- Importación y exportación.
- Filtros y consultas.
- Detección de posibles duplicados.
- Revisión de duplicados antes de incorporar registros al maestro.
- Detalle de la solicitud.
- IndexedDB como almacenamiento operativo actual.
- Corrección del scroll.
- Filtros de texto con escritura continua sin perder el foco.

## Plantilla oficial de Casos
La fuente de la plantilla es el libro Excel corporativo proporcionado para MINEX. La pestaña principal para Casos es `VOLUMETRIAS`. Se conservan las hojas auxiliares y estados de visibilidad del libro.

El archivo incluido en el proyecto es:
`assets/FORMATO_CARGA_CASOS_RELACIONES_LABORALES.xlsx`

## Regla crítica: duplicados
La detección de posibles duplicados y la cola de **Revisión de duplicados** son funcionalidades críticas. No deben eliminarse ni simplificarse en futuras ediciones.

## Principio de desarrollo
Cada nueva versión parte de la anterior y conserva los avances existentes. Este `README.md` es el único README oficial y se actualiza en cada edición. No crear `README_Vxx.md` adicionales.

## Historial

### V25.8
- Corrección robusta de carga de plantilla para descarga de casos en GitHub Pages.
- Uso de XLSX real como fallback.
- Separación conceptual entre GitHub (publicación) y MINEX (aplicación/datos).

### V25.7
- Plantilla oficial integrada en el proyecto para evitar dependencias de rutas externas.

### V25.6
- Descarga de casos con la estructura corporativa de `VOLUMETRIAS`.
- Conservación de fórmulas, tabla y formatos.

### V25.5
- Ajustes de exportación para evitar descargas planas.

### V25.4
- Uso de la plantilla Excel corporativa para la carga de casos.

### V25.3
- Consolidación de documentación en un único README.

### V25.2
- Mejora de filtros de texto para escritura continua.

### V25.1
- Corrección del desplazamiento vertical.

### V25.0
- Preparación del proyecto para GitHub Pages.
- Conservación de IndexedDB y del sistema de duplicados.

### V24
- Base funcional anterior: casos, seguimientos, empleados, informes, importación/exportación y revisión de duplicados.

## Publicación en GitHub Pages
1. Subir el contenido del proyecto al repositorio de GitHub.
2. En GitHub: **Settings → Pages**.
3. Seleccionar la rama `main` y la carpeta raíz `/`.
4. Abrir la URL de GitHub Pages generada.

GitHub se utiliza para alojar/publicar MINEX. No se debe pedir al usuario final un token de GitHub para operar la aplicación.

## Estructura principal
```text
index.html
app.js
styles.css
data_bootstrap_min.js
data/
  cases.json
  employees.json
  catalogs.json
data/case_template_base64.js
assets/
  FORMATO_CARGA_CASOS_RELACIONES_LABORALES.xlsx
README.md
```
