# Mendiak

Diario de montaña: registrar salidas y buscar cuándo y con quién se subió cada monte.

- **App**: `index.html` (PWA, sin dependencias). Se publica con GitHub Pages.
- **Datos**: una hoja de Google Sheets con las pestañas `Salidas` y `Listas`.
- **Back**: `Code.gs`, pegado en Extensiones → Apps Script de esa hoja e implementado como aplicación web.

La app se configura sola al abrir el enlace de instalación `…/mendiak/#u=<URL /exec>&k=<clave>`. La clave no está en este repositorio.
