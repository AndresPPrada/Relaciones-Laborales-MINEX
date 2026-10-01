# MINEX | Gestión de Relaciones Laborales — V8

Aplicación frontend para GitHub Pages. Incluye dashboard ejecutivo, casos CRUD, seguimientos por vencimiento, volumetrías independientes, carga/descarga Excel/CSV, detección de duplicados, empleados CRUD, informes por área/empresa/requerimiento/responsable/mes y control de calidad de datos.

## Publicación
1. Sube todo el contenido de esta carpeta al repositorio de GitHub Pages.
2. Conserva la carpeta `data/` junto a `index.html`.
3. En GitHub: Settings → Pages → Deploy from branch → `/root`.
4. Recarga con Ctrl+F5.

## Datos y privacidad
Esta versión usa IndexedDB local. Los archivos JSON incluidos contienen la estructura/datos iniciales del Excel original. No publicar en un repositorio público si contienen información laboral real.

## Carga
- Casos: Excel/CSV, validación y duplicados.
- Volumetrías: Excel/CSV, validación y duplicados.
- Empleados: Excel/CSV, validación por documento.

## Exportación
Los estados 0/1 se convierten a texto y las fechas se exportan como fechas legibles. Los informes pueden descargarse en Excel o imprimirse/guardarse como PDF desde el navegador.


## V12
- Los datos existentes se conservan al recargar y actualizar la aplicación mediante IndexedDB.
- Se incorporan acciones separadas para eliminar casos, empleados, volumetrías o toda la información.
- Restaurar datos iniciales es una acción independiente y explícita.


## V23
Optimizada para GitHub Pages: Chart.js y SheetJS se cargan bajo demanda, el respaldo de datos se carga solo si falla la lectura de `/data/`, IndexedDB mantiene la información existente y las cargas/eliminaciones masivas usan una sola transacción por operación.
