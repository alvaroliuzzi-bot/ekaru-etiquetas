# e-Karú Etiquetas

Programa de escritorio (Python + customtkinter) que genera e imprime las etiquetas de productos de Glacé en la impresora térmica **HPRT HT400 (203 dpi)**: nombre, código de barras y tabla nutricional. Imprime en RAW/TSPL o por GDI y muestra una vista previa idéntica a lo impreso.

Los baldes **no** se etiquetan aquí: los imprime la balanza con su propio software (Scale Manager). Ver el documento de procesos en [`ekaru-procesos`](https://github.com/alvaroliuzzi-bot/ekaru-procesos).

## Uso (Windows)

1. Instalar Python 3 desde python.org (marcar «Add to PATH»).
2. Doble clic en `Iniciar_Ekaru.bat`. La primera vez instala `customtkinter`.
3. `Crear_Acceso_Directo.bat` deja un acceso directo en el escritorio. `Iniciar_Ekaru_DEBUG.bat` abre con consola para ver errores.

## Archivos

| Archivo | Para qué |
|---|---|
| `app.py` | La aplicación (catálogo de productos, diseño e impresión) |
| `ekaru_config.json` | Impresora, densidad, velocidad, márgenes y grilla de la etiqueta |
| `cantidades_impresion.json` | Últimas cantidades por producto |
| `codigos_cambios.csv` | Cambios de códigos del catálogo |
| `ekaru_ventas.html` | Prototipo de la pantalla de ventas de e-Karú |
| `tools/` | Scripts de diagnóstico de la impresora |
| `docs/` | Catálogo de códigos, documento para los desarrolladores y diseños de etiquetas de baldes |
| `encuesta_img/` | Imágenes usadas en la encuesta al personal de caja |
