# MINEX | Gestión de Relaciones Laborales — V25 GitHub

## Arquitectura

- Frontend estático: GitHub Pages.
- Almacenamiento local: IndexedDB.
- Almacenamiento compartido opcional: GitHub Repository Contents API.
- Sin Supabase, Firebase, PostgreSQL ni backend propio.
- Los registros compartidos se guardan como archivos JSON independientes por registro para reducir conflictos entre usuarios.

## Colaboración

Cada usuario configura en **Configuración → GitHub · datos compartidos**:

- Propietario/organización.
- Repositorio.
- Rama.
- Ruta de datos.
- Token GitHub de sesión.

El token **no se incluye en el código** ni se publica en el repositorio; se mantiene únicamente en `sessionStorage` del navegador.

Para el repositorio corporativo se recomienda un repositorio privado y un token con el mínimo permiso requerido para Contents.

## Sincronización

- **Descargar cambios:** incorpora registros publicados por otros usuarios.
- **Publicar cambios:** publica los datos locales.
- Las modificaciones y eliminaciones de casos, empleados y revisiones de duplicados intentan publicarse automáticamente cuando GitHub está conectado.
- Cada registro tiene `_uid` para evitar depender del `_id` autoincremental local.
- Las eliminaciones se representan mediante archivos de tipo tombstone para que otro navegador no vuelva a recuperar el registro eliminado.

## Duplicados

Se conserva el flujo existente de detección y revisión de duplicados:

1. Detección durante importación.
2. Envío a `duplicateReviews`.
3. Edición.
4. Eliminación.
5. Incorporación a la base maestra.

La clave continúa considerando fecha, empresa, requerimiento, identificación/empleado y detalle de la solicitud.

## Excel de SharePoint

La integración bidireccional con Excel de SharePoint queda como una capa independiente. No se introduce un backend para resolverla. Para activarla de forma segura se requiere definir el libro `.xlsx`, la tabla de Excel y el método de autenticación Microsoft que utilizará la organización.
