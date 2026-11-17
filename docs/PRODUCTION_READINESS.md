# Mini Word: evaluación para beta y producción

Fecha de auditoría: 2026-11-15

## Veredicto ejecutivo

| Objetivo | Estado | Estimación |
|---|---|---|
| Beta privada para DOCX simples/medios | **Cerca, pero todavía no responsable** | 2–4 semanas |
| Beta pública | **No todavía** | 6–10 semanas |
| Producción 1.0 con alcance DOCX limitado | **No todavía** | 4–7 meses |
| Compatibilidad alta con Word | **Lejos** | 12–24+ meses |
| Paridad funcional completa con Word | **No realista a corto plazo** | Años / equipo mayor |

Mini Word ya es un editor propio funcional, no un visor: pagina, edita, guarda DOCX, maneja tablas, imágenes flotantes, secciones, columnas, notas, comentarios, hipervínculos, formas y revisiones ortográficas. Para documentos sencillos y medianos puede ser útil hoy.

Para llamarlo **beta** hace falta demostrar que no pierde datos y que no se bloquea con documentos reales. Para llamarlo **producción** hace falta bastante más que «abre y guarda DOCX»: preservación de lo no soportado, guardado atómico, recuperación, encabezados/pies correctos, campos, estilos, accesibilidad, instalación firmada y pruebas con corpus reales.

## Puntaje orientativo

Escala 0–10. No es una medición formal de conformidad con el estándar.

| Área | Puntaje | Comentario |
|---|---:|---|
| Edición de texto y párrafos | 8 | Motor propio, selección, formato, listas y paginación |
| DOCX simple | 8 | Texto, fuentes, párrafos, imágenes, tablas básicas |
| DOCX medio | 6 | Secciones, columnas, notas, comentarios, objetos flotantes |
| DOCX avanzado | 2–3 | Campos, contenido estructurado, estilos/temas, gráficos |
| Word 97-2003 (`.doc`) | 2–3 | Solo importación parcial; no se guarda `.doc` |
| Fidelidad visual | 5–6 | Buena en documentos propios; variable en documentos de Word |
| Integridad de datos al guardar | 3–4 | Reescribe el paquete y descarta partes no modeladas |
| Estabilidad | 5 | Buenos casos simples; se encontró y corrigió un bloqueo real |
| Rendimiento | 6 | Bueno en DOCX de 1,4 MB; falta probar documentos grandes |
| Seguridad de archivos | 4 | Sin ejecución de OLE, pero sin fuzz ni límites formales |
| Accesibilidad | 2–3 | Control GDI+ propio, sin árbol de accesibilidad dedicado |
| Instalación y soporte | 2 | Un `.exe`; sin instalador, firma, telemetría ni actualizador |
| Pruebas y CI | 4 | 47 autopruebas y sondas manuales; sin CI ni corpus continuo |

## Compatibilidad funcional con Word

### Nivel alto

- Texto Unicode, fuentes, tamaños, negrita, cursiva, subrayado, tachado, color, resaltado, superíndice y subíndice.
- Alineación, sangrías, espaciado, interlineado, tabulaciones y control de página.
- Páginas, márgenes, orientación, zoom, impresión y vista previa.
- Tablas con celdas combinadas, estilo básico, bordes y sombreado.
- Imágenes, cuadros de texto y formas flotantes con selección, movimiento y redimensionado.
- Listas con viñeta y numeración simple.
- Hipervínculos externos.

### Nivel medio

- Secciones y columnas periodísticas.
- Encabezados y pies de texto.
- Notas al pie/finales y comentarios de lectura.
- Estilos de tabla conocidos.
- Ajuste cuadrado de texto alrededor de objetos.
- Objetos OLE conservados con vista previa, sin edición.

### Nivel bajo o ausente

- Encabezados/pies independientes por sección y con imágenes al guardar.
- Campos complejos: `PAGE`, `NUMPAGES`, referencias, índices y tablas de contenido.
- Controles de contenido `w:sdt` al guardar.
- Estilos y temas completos de Word.
- Control de cambios, marcas de revisión, autores y aceptar/rechazar.
- Comentarios editables y anclados al texto.
- Gráficos, SmartArt, ecuaciones, diagramas y objetos incrustados ejecutables.
- Formas con rotación, grupos, conectores, geometría personalizada y ajuste `tight`/`through` real.
- Saltos de sección continuos y equilibrio de columnas.
- Protección de documento, firmas digitales, macros VBA y formularios.
- Reconstrucción completa de OfficeArt/Escher y guardado `.doc`.

## Cobertura aproximada

| Tipo de documento | Cobertura realista |
|---|---:|
| DOCX simple creado por Mini Word | 85–95% |
| DOCX simple/medio de Word | 60–75% |
| DOCX con campos, índices, controles y temas | 30–50% |
| DOCX avanzado con gráficos, revisiones y macros | 10–25% |
| DOC 97-2003 importado | 20–35% |
| Word global, contando toda su superficie | 30–40% |

La cifra de «Word global» no debe usarse para marketing. Un editor puede ser excelente para el 80% de documentos cotidianos y seguir estando muy lejos de Word en el 20% restante.

## Hallazgo crítico corregido durante esta auditoría

Se encontró un **bucle infinito** en `DocumentSection.Clone()` cuando una sección tenía imágenes de encabezado o pie:

- `MemberwiseClone()` compartía la lista de imágenes.
- El bucle añadía imágenes a la misma lista que estaba recorriendo.
- El documento real `guia-presentacion-tesis-trabajos-grado.docx` (1,4 MB) se quedaba bloqueado al guardar.

Corrección aplicada: crear listas nuevas antes de copiar las imágenes.

Resultado después de la corrección:

```text
read_ms=18
save_ms=53
paras=342
tables=2
sections=12
images=1
notes=2
copy_paras=342
copy_tables=2
copy_images=1
```

Esto demuestra la importancia de una beta con documentos reales: el caso simple pasaba, pero un encabezado con imagen podía colgar el guardado.

## Lo que falta para una beta

Una beta puede tener funciones incompletas, pero no debe destruir documentos ni bloquearse. Antes de una beta pública:

1. **Guardado atómico y copia de seguridad** — implementado en `AtomicFile`
   - Escritura en archivo temporal.
   - Validación de que el archivo generado no está vacío.
   - Sustitución con `File.Replace` y respaldo `.bak`.
   - Si la sustitución no está disponible, se copia el original y luego el temporal.

2. **Preservación de contenido no soportado**
   - No reconstruir todo el paquete desde cero.
   - Conservar partes y XML desconocidos, o avisar exactamente qué se perderá.
   - Casos críticos: campos, controles de contenido, temas, gráficos, comentarios, revisiones y macros.

3. **Recuperación automática al iniciar** — implementada
   - El autoguardado de documentos sin guardar crea una copia con metadatos de ruta original y fecha.
   - Al iniciar, si existe una copia de una sesión interrumpida, se ofrece recuperarla.
   - La copia se elimina después de recuperarla o descartarla.
   - El guardado normal también elimina la copia de recuperación.

4. **Encabezados y pies por sección**
   - Un `header1.xml`/`footer1.xml` único no representa documentos con primeras páginas distintas o secciones diferentes.
   - Escribir partes separadas, relaciones y referencias por sección.
   - Incluir imágenes de encabezado/pie.

5. **Campos y numeración**
   - Leer/escribir campos `w:fldSimple` y `w:instrText`.
   - Actualizar o al menos conservar `PAGE`, `NUMPAGES`, TOC y referencias.

6. **Corpus de regresión**
   - 100–300 DOCX reales de Word, LibreOffice, Google Docs y generadores.
   - Pruebas automáticas de lectura, guardado, reapertura y comparación de estructura.
   - Casos de documentos grandes, imágenes grandes, encabezados, tablas y XML desconocido.

7. **Crash reporting y registro**
   - Capturar excepciones no controladas.
   - Escribir un log con versión, archivo y operación.
   - No mostrar solo un mensaje genérico.

8. **Instalador, firma y versión**
   - Instalador MSI o MSIX.
   - Firma de código para evitar SmartScreen.
   - Versión visible y canal beta/estable.

9. **Límites y seguridad**
   - Límites de tamaño de ZIP y XML.
   - Detección de zip bombs.
   - Fuzz del lector DOCX/DOC/ODT.
   - Pruebas con archivos malformados.

10. **Accesibilidad básica**
    - Navegación completa por teclado.
    - Nombres accesibles y árbol UI Automation.
    - Contraste y escalado.

Con estas diez condiciones, una beta privada o una beta pública limitada sería defendible.

## Lo que falta para producción

Además de todo lo anterior:

- Comentarios editables y control de cambios con aceptar/rechazar.
- Índices, tablas de contenido y actualización de campos.
- Estilos y temas de Word con herencia real.
- Gráficos, SmartArt, ecuaciones y diagramas: al menos conservarlos o renderizarlos.
- Formas con rotación, grupos, conectores y geometría personalizada.
- Notas al pie con reserva de espacio y continuación entre páginas.
- Secciones continuas y columnas equilibradas.
- Guardado `.doc` o una decisión de producto explícita de no soportarlo.
- Importación `.doc` con advertencia de fidelidad y sin sobrescritura.
- Protección de documentos, firmas y tratamiento de macros.
- Accesibilidad verificable.
- Telemetría de errores, soporte y actualizaciones.
- Documentación de usuario y política de compatibilidad.

## Estimación de esfuerzo

Estimaciones de orden de magnitud, con un equipo pequeño y un alcance DOCX explícito:

| Hito | Tiempo | Esfuerzo aproximado |
|---|---:|---:|
| Correcciones críticas de integridad y encabezados | 2–4 semanas | 1,5–2,5 personas-mes |
| Beta privada | 3–5 semanas | 2–4 personas-mes |
| Beta pública | 6–10 semanas | 3–6 personas-mes |
| Producción 1.0 acotada a DOCX simple/medio | 4–7 meses | 8–14 personas-mes |
| Compatibilidad alta con Word | 12–24+ meses | 5–8 personas |
| Paridad funcional completa | 3+ años | equipo amplio |

Si el equipo es una sola persona, los plazos calendario deben multiplicarse. La compatibilidad de Word no es un problema de «leer XML»: es un problema de preservar semántica, estilos, campos, objetos y revisiones sin pérdida.

## Recomendación de lanzamiento

### Ahora

No lanzar como producción. Se puede usar para documentos propios y pruebas controladas, con copia de seguridad.

### Beta privada recomendada

Alcance:

- DOCX simple/medio.
- `.doc` solo importación y conversión a `.docx`.
- Sin promesa de paridad con Word.
- Aviso de funciones no soportadas.
- Copia de seguridad automática antes de guardar.

### Beta pública

Solo después de:

- preservación de campos, TOC y controles de contenido;
- estrategia de preservación de XML/partes desconocidas;
- encabezados/pies de primera página y página par;
- corpus de 100+ documentos y CI;
- instalador firmado y accesibilidad básica.

Ya están cubiertos el guardado atómico, la recuperación al iniciar, los encabezados/pies predeterminados por sección y el log local de fallos.

### Producción 1.0

Solo después de añadir campos/TOC, estilos/temas, comentarios/revisiones y una estrategia clara para contenido no soportado. Si no se implementan gráficos y SmartArt, la documentación debe decirlo y el producto debe posicionarse como editor DOCX ligero, no como sustituto universal de Word.

## Checklist beta numerada

Estado a 2026-11-15:

1. **Guardado atómico y copia `.bak`** — implementado y probado.
2. **Recuperación automática al iniciar** — implementada y probada con `RestoreProbe`.
3. **Encabezados/pies por sección con imágenes** — implementado para encabezado/pie predeterminado por sección; primera página y página par siguen pendientes.
4. **Campos, TOC y controles de contenido** — pendiente de preservación real; por ahora se detectan y se avisa antes de guardar.
5. **XML/partes desconocidas** — parcial: se conserva el original en `.bak` y se avisa, pero no se reconstruyen las partes no modeladas.
6. **Corpus y CI** — parcial: 11 DOCX reales y fuzz local; faltan 100+ documentos y CI continuo.
7. **Crash reporting y límites** — implementado el log local y los límites; falta integración con un canal de soporte.
8. **Instalador firmado, versión y actualizaciones** — pendiente.
9. **Accesibilidad y navegación completa por teclado** — pendiente.
10. **Tablas complejas** — parcial: fila completa, filas/columnas simples y bloqueo seguro de operaciones sobre celdas combinadas; fusiones anidadas pendientes.

Para beta privada controlada faltan sobre todo **4, 5, 6, 8 y 9**. Para beta pública deben cerrarse además los casos de primera página/página par y las tablas anidadas.

## Próximos cinco pasos

1. Preservar o reconstruir campos, TOC y controles de contenido.
2. Añadir encabezados/pies de primera página y página par.
3. Ampliar el corpus a 100+ DOCX y montar CI continuo.
4. Instalador firmado, versión y canal de actualización.
5. Accesibilidad básica y pruebas de teclado completas.

## Evidencia

- Autoprueba de formatos: **47/47**.
- DOCX real de 1,4 MB: lectura **18 ms** y guardado **53 ms**; la copia conservó 342 párrafos, 2 tablas y 1 imagen, con 12 secciones.
- Inserción de 8 formas y 5 imágenes sin bloqueo.
- Guardado de formas con color y contorno: `COLORS=True`.
- Proyecto: 39 archivos C#, ~20.000 líneas; ejecutable de ~402 KB.
- Se detectaron 62 bloques `catch`, 36 de ellos vacíos; los críticos de apertura, guardado, recuperación y autoguardado ya registran en el log.
- El guardado de usuario ya usa `AtomicFile` y deja `.bak`; la escritura ZIP interna directa solo se usa sobre el temporal.
- Límites añadidos: paquete ZIP de 64 MB, 4096 entradas, 128 MB por entrada, 256 MB descomprimidos, XML de 64 MB, imagen de 32 MB y 50 megapíxeles.
- El control de texto flotante incluye fuente, tamaño, negrita, cursiva, subrayado y color.
- La regla es una banda fija de la interfaz, no un objeto de la hoja.
- `RestoreProbe`: recuperación automática al iniciar con el texto `texto recuperado`; el autoguardado genera `.docx` y `.meta` incluso si el autoguardado programado está desactivado.
- `HeaderProbe`: **2 encabezados**, **2 pies**, **2 relaciones de encabezado**, con texto e imágenes independientes por sección.
- Corpus: **101/101** documentos (11 DOCX, 6 ODT de LibreOffice, 3 DOC y 3 reales) abiertos, guardados y reabiertos sin fallos.
- `FuzzProbe`: **60/60** entradas aleatorias rechazadas; DTD externa y ruta ZIP `../` bloqueadas.
- `LogProbe`: archivo de log creado con contexto y excepción.
- `UnsupportedProbe`: detección de controles de contenido y campos; el guardado pide confirmación y conserva el original en `.bak`.
- `OpenFile`: 101/101 documentos (DOCX, ODT y DOC) cargados desde la interfaz; la ventana abre maximizada.
- Regla en un control fijo independiente del editor: scroll sin duplicados ni parpadeo de la banda.
- Cuadro de texto: color de fondo editable, previsualizado y conservado en el DOCX.
- Corrector: palabra larga y botón «Añadir al diccionario» sin recortes; la marca 20 aparece en la regla.
- Encabezados/pies: formatos de fuente, color y alineación por sección; se guardan, se releen y se dibujan con banda clara.
- `ShapeFormatDialog`: valor de grosor y `px` visibles.
- `SizeProbe`: el tamaño cambia sin mover el viewport (`SCROLL=500/500`) y conserva negrita, cursiva, fuente y color de cada tramo.
- `ShapeRoundtrip`: posición, tamaño, tipo, relleno y ajuste de la forma se conservan al guardar y reabrir (`X=2400`, `Y=1800`).
- `ShapeLayoutProbe`: la posición de diseño es idéntica antes y después de reabrir (`X=218,9`, `Y=152,4667`).
- `ContextFormatProbe`: aplicar formato dentro de una tabla no mueve el scroll (`SCROLL=300/300`), conserva la selección y no toca el cuerpo; los cuadros de texto reciben el formato (`TEXTBOX=True`).
- `TableFormatRoundtrip` y `TextBoxFormatProbe`: fuente, tamaño, color y negrita se conservan al guardar y reabrir.
- `TableNavProbe`: navegación por flechas dentro de celdas (`A → B → A`) y por la selección de tabla (`SELECTION=1,0`).
- `TableExitProbe`: salida por los bordes sin mouse (`LEFT_PARA=0`, `RIGHT_PARA=2`).
- `TableTopExitProbe`: tabla al inicio sin párrafo superior; sale al párrafo siguiente sin insertar líneas (`PARA=1`, `COUNT=2/2`).
- `MergedNavProbe`: navegación correcta desde una celda fusionada (`DOWN=2:D`, `UP=0`), con salida por el borde superior al párrafo siguiente.
- `BodyFormatProbe`: tamaño/fuente solo cambian el párrafo seleccionado; los párrafos vecinos conservan su formato.
- `FinalFormatProbe`: solo la celda editada cambia (`CELL_ONLY=True`), su fila crece inmediatamente (`ROW_H=56,8`), la selección se conserva y el scroll no se mueve (`SCROLL=300/300`).
- `TwoTableProbe`: con dos tablas, el formato no salta a la primera tabla (`FIRST=True`, `SECOND=True`, `CELL=0,0`).
- `RoundtripScopeProbe`: al guardar/reabrir solo el fragmento formateado cambia; los párrafos vecinos conservan su formato original.
- `BulletProbe`: texto de viñeta que ocupa 3 líneas usa solo 1 viñeta; el párrafo numerado usa solo 1 número.
- `BulletUserTest`: 3 páginas y 5 viñetas; al cambiar una viñeta a 30 pt, únicamente esa recibe `60` medios puntos, las demás permanecen en `22`, el documento conserva **3 páginas** y las viñetas largas no generan marcas fantasma. El punto de viñeta conserva el tamaño base y no se agranda.
- Sangría: iconos corregidos; prueba `SCROLL_BEFORE=900`, `SCROLL_AFTER=900`, `INDENT=360`.
