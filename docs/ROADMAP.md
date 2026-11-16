# Hoja de ruta

## Etapa 1 — Base pública segura

- Código sin datos personales ni identificadores privados.
- Configuración mediante Propiedades del script.
- Documentación de instalación y despliegue.
- Separación básica entre backend, interfaz y recursos.

## Etapa 2 — Calidad y mantenimiento

- Dividir `Code.gs` por dominios: autenticación, personal, proyectos,
  publicaciones, becarios, eventos e investigaciones.
- Dividir `Index.html` en componentes y módulos JavaScript.
- Añadir validación centralizada, manejo uniforme de errores y pruebas.
- Incorporar CI en GitHub para análisis estático y revisión de secretos.

## Etapa 3 — Aplicación escalable

- API independiente con Node.js/TypeScript.
- Base de datos PostgreSQL con migraciones y respaldos.
- Autenticación con Google Workspace/OAuth y control de acceso por rol.
- Frontend moderno y accesible, instalable como PWA.
- Historial de cambios, notificaciones y paneles configurables.

## Etapa 4 — Producto avanzado

- Indicadores y tableros configurables por programa y país.
- Flujos de aprobación y seguimiento de compromisos.
- Importación validada desde hojas de cálculo.
- Exportaciones programadas y API de integración.
- Pruebas de seguridad, observabilidad y despliegues por ambiente.
