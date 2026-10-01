# MINEX | Gestión de Relaciones Laborales

## Versión actual
**V25.2**

Sistema web para la gestión de casos de Relaciones Laborales, desarrollado para funcionar con GitHub Pages.

## Funcionalidades
- Dashboard de gestión.
- Casos: creación, consulta, edición y seguimiento.
- Seguimientos y pendientes.
- Informes e indicadores.
- Base de empleados.
- Importación y exportación de datos.
- Filtros y consultas.
- Detección de posibles duplicados.
- Revisión de duplicados antes de incorporar registros al maestro.
- Almacenamiento local mediante IndexedDB.
- Preparación para trabajo con GitHub como repositorio de datos compartidos.

## Principio de desarrollo
Cada nueva versión debe partir de la versión anterior y **conservar los avances existentes**. No se deben eliminar funcionalidades sin una decisión explícita.

La funcionalidad de **detección y revisión de duplicados es crítica** y debe conservarse en todas las versiones.

## Historial de versiones

### V25.2
- Mejora de los filtros de texto para permitir escritura continua sin perder el foco.
- Correcciones de desplazamiento vertical.
- Conservación de las funcionalidades existentes.
- Consolidación de la documentación en este único README.

### V25.1
- Corrección del desplazamiento vertical de la aplicación.
- Ajustes de navegación y visualización.

### V25.0
- Preparación para GitHub/GitHub Pages.
- Estructura de sincronización de datos mediante GitHub.
- Conservación de IndexedDB como almacenamiento local.
- Conservación del sistema de duplicados y revisión.

### V24
- Base funcional anterior del sistema.
- Casos, seguimientos, empleados, informes, importación/exportación y revisión de duplicados.

## GitHub Pages
1. Crear o utilizar un repositorio de GitHub.
2. Subir el contenido de esta carpeta al repositorio.
3. Ir a **Settings → Pages**.
4. Seleccionar la rama `main` y la carpeta raíz `/`.
5. Guardar y abrir la URL generada por GitHub Pages.

## Regla para próximas ediciones
El README se mantiene como **único archivo de documentación principal**. En cada nueva versión se actualizará este mismo archivo agregando el cambio realizado al historial, sin crear `README_Vxx.md` adicionales.

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
README.md
```
