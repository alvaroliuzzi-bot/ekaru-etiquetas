# Cambios

## 2.2.0 · 13/11/2026 · Etiko

- **Nuevo nombre: Etiko** (antes «E-Karú Etiquetas»; e-Karú es la caja). Icono nuevo (`etiko.ico`); al abrir la 2.2 se crea «Etiko» en el escritorio y se quita el acceso viejo «E-Karu».
- **Diseño nuevo**: grafito con acento coral, títulos en Bahnschrift, tarjetas y botones redondeados, barra lateral ordenada (Trabajo / Herramientas / Impresora). En el catálogo solo la fila elegida lleva el botón «Imprimir» lleno; las demás, uno tonal.
- **Accesibilidad (WCAG 2.2 AA)**: contraste ≥ 4,5:1 en textos y botones, botones − / + más grandes, borde coral en el campo con el foco, la fila elegida se marca también con una franja, texto chico a 12 px.
- El buscador vuelve a mostrar su texto de ayuda.
- Sin cambios en códigos, impresión ni archivos de configuración.

## 2.1.0 · 12/11/2026 · Numeración v2 del Informe final

### Nuevo
- **Numeración v2** (Informe final, pág. 3). Productos: `22` + Nº de producto (4) + fecha de elaboración AAMMDD (6) + control, por ejemplo `22 0001 261112 8`. Cada familia tiene su prefijo: `23` Promo Pack, `24` ID de balde (plan B), `25` insumo interno; `29` (grupos PADRE) y `20` (baldes de la balanza) se mantienen.
- **Fecha de elaboración** editable sobre la cantidad. Arranca en hoy y pasa sola al día nuevo después de medianoche; acepta `12/11/2026`, `12-11-26` o `12112026`; no puede ser futura ni de hace más de 400 días. «Hoy» vuelve a la fecha actual.
- **Etiquetas internas** (barra lateral): insumo (25, con fecha de recepción), Promo Pack (23) e ID de balde (24). Numeración correlativa en `contadores.json`: el contador avanza solo si el envío a la impresora salió bien y nunca hacia atrás. Lista de insumos en `insumos.json` (se agregan desde la misma ventana; números provisorios).
- **Leer un código**: aplica las 7 reglas del Informe en el mismo orden que e-Karú (correcto · código viejo · rechazado · código externo), con historial de la sesión.
- **Juego de prueba**: una etiqueta de cada familia (22, 29, 25, 23 y 24) con la marca «PRUEBA».
- **Tabla de transición** (`docs/tabla_transicion_codigos.csv`, también desde Diagnóstico y desde «Leer un código»): 47 productos y 9 grupos PADRE con el código del catálogo de e-Karú, el de la 2.0 y el nuevo.
- **Búsqueda por código escaneado** en el catálogo: encuentra el producto de una etiqueta de cualquier día y de las numeraciones anteriores.
- **Aviso de primera vez** al imprimir un producto (Sí = nueva · No = anterior · Cancelar = no imprimir) y un interruptor permanente en Diagnóstico → Ajustes.
- `tools/prueba_numeracion.py` (27 pruebas) y `tools/_tk_shim.py`.
- Prototipo `ekaru_ventas.html` con la numeración v2: códigos `22` con la fecha, escáner con las 7 reglas, avisos de «reetiquetar» y la pestaña «Probar un código».

### Cambia
- Los códigos de producto que se imprimen pasan de `2` + producto (6) + AADDD + control a `22` + Nº + AAMMDD + control. Es la configuración por defecto; en Ajustes se puede volver a la numeración anterior.
- La barra lateral muestra qué numeración está activa y se compacta en pantallas chicas o con zoom del 200 %.

### Sin cambios
- Grupos PADRE (`2 9000NN 80008`) y baldes de la balanza (`20…`).
- Impresión por driver (GDI) y RAW/TSPL, grilla del rollo, márgenes y demás ajustes: la 2.1 lee `ekaru_config.json` y `cantidades_impresion.json` tal cual. Con la numeración anterior, la salida a la impresora es idéntica a la de la 2.0.

### Para tener en cuenta
- La caja (e-Karú) tiene que conocer los códigos nuevos antes de vender con ellos. Las etiquetas viejas ya pegadas se reconocen como «código viejo».
- `docs/catalogo_ekaru_codigos.pdf`, `docs/e-Karu_Etiquetas_documento_desarrolladores.pdf` y `codigos_cambios.csv` describen la numeración 2.0; para la v2 vale la tabla de transición.

## 2.0.0 · 2/11/2026

- Impresión RAW directa por ctypes (sin pywin32) y por el driver de Windows (GDI); diagnóstico de impresión y registro.
- Catálogo de 66 productos con lista virtualizada (sin parpadeo), vista previa idéntica a lo impreso y códigos de barras dibujados desde los mismos bits verificados.
- Corrección del dígito de control de los códigos del catálogo original (`codigos_cambios.csv`) y códigos con fecha de elaboración (`2` + producto + AADDD).
