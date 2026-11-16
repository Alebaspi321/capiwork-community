# Capiwork Community

Versión pública y sanitizada de **Capiwork**, una aplicación de monitoreo,
evaluación y aprendizaje construida con Google Apps Script y Google Sheets.

Este repositorio contiene únicamente código y recursos visuales incorporados.
No contiene bases de datos, correos, contraseñas ni identificadores reales de
Google Drive.

## Funciones incluidas

- Inicio de sesión y administración de sesiones.
- Roles `ADMIN`, `EDITOR` y `VIEWER`.
- Dashboard de personal, proyectos, publicaciones, becarios, eventos e
  investigaciones.
- Tema claro/oscuro, gráficos y exportación a PDF/Excel.
- Registro de auditoría y recuperación de contraseña.
- Recursos visuales integrados para evitar dependencias privadas.

## Estructura

- `src/Code.gs`: backend de Google Apps Script.
- `src/Index.html`: interfaz del sistema.
- `src/Assets.gs`: recursos gráficos incorporados.
- `src/appsscript.json`: manifiesto de Apps Script.
- `docs/ROADMAP.md`: propuesta de evolución del producto.

## Configuración segura

1. Cree un proyecto nuevo de Google Apps Script.
2. Copie los archivos de `src/` o use [clasp](https://github.com/google/clasp).
3. En **Configuración del proyecto > Propiedades del script**, defina:

   - `ADMIN_EMAIL`: correo del administrador inicial.
   - `ADMIN_NAME`: nombre visible del administrador.
   - `ADMIN_TEMP_PASSWORD`: contraseña temporal robusta.
   - `SUPER_ADMIN_EMAILS`: correos separados por comas.
   - `SOURCE_SPREADSHEET_ID`: ID de la base principal.
   - `RESEARCH_SPREADSHEET_ID`: ID de la base de investigaciones.
   - `PROJECT_SPREADSHEET_ID`: ID del banco de proyectos.

4. Ejecute `configurarSistema()` desde el editor de Apps Script.
5. Implemente el proyecto como aplicación web.

Nunca publique valores reales de estas propiedades en GitHub. El archivo
`src/Code.gs` usa datos ficticios como valores de demostración y deja vacíos
todos los identificadores externos.

## Credenciales de demostración

Los valores de ejemplo son `admin@capiwork.example` y
`CambiarEstaClave2026!`. Cámbielos mediante Propiedades del script antes de
usar el sistema fuera de un entorno de prueba.

## Desarrollo

El proyecto conserva compatibilidad con Google Apps Script para facilitar una
primera publicación. La evolución recomendada está descrita en
`docs/ROADMAP.md`, incluyendo separación de API, base de datos, autenticación
y pruebas automatizadas.

## Licencia

MIT. Consulte `LICENSE`.
