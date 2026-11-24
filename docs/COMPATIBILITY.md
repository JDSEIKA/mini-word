# Estado de compatibilidad y visualización

Fecha de la última auditoría: 2026-11-13. Actualización funcional: **2026-11-23, 0.0.8 Beta**.

> **Estado 0.0.8:** ya existen campos PAGE/NUMPAGES, índice, control de cambios, comparación, estilos con nombre, respuestas de comentarios y tramos en celdas. La lista de pendientes de este documento es histórica y debe leerse con esa actualización.

## Resumen

Mini Word ya implementa una capa avanzada sobre su motor nativo: secciones con columnas periodísticas, objetos flotantes, cuadros de texto, formas, ajuste de texto cuadrado, notas al pie/finales, comentarios, estilos de tabla y lectura parcial de imágenes/encabezados en Word 97-2003.

Todavía no es un clon completo de Word. La diferencia principal está en la fidelidad de los objetos OLE, el ajuste de texto de contorno y la reconstrucción total de formas binarias `.doc`.

## Pruebas con archivos reales

Se descargaron documentos de calibre, python-docx y Apache POI.

| Archivo | Formato | Páginas | Párrafos | Secciones | Notas | Comentarios | Objetos flotantes | Imágenes |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| demo-calibre.docx | DOCX | 6 | 95 | 1 | 0 | 0 | 0 | 4 |
| shapes-with-text.docx | DOCX | 1 | 1 | 1 | 0 | 0 | 7 | 0 |
| various-pictures.docx | DOCX | 1 | 7 | 1 | 0 | 0 | 0 | 5 |
| footnotes.docx | DOCX | 1 | 1 | 1 | 1 | 0 | 0 | 0 |
| endnotes.docx | DOCX | 1 | 12 | 1 | 1 | 0 | 0 | 0 |
| comment.docx | DOCX | 1 | 1 | 1 | 0 | 1 | 0 | 0 |
| test-comment.docx | DOCX | 1 | 1 | 1 | 0 | 1 | 0 | 0 |
| header-pic.docx | DOCX | 1 | 1 | 1 | 0 | 0 | 0 | 1 en encabezado |
| embedded.docx | DOCX | 1 | 3 | 1 | 0 | 0 | 1 | vista previa |
| two-images.doc | DOC 97 | 1 | 5 | 1 | 0 | 0 | 0 | 2 |
| header-image.doc | DOC 97 | 1 | 11 | 1 | 0 | 0 | 0 | 2 + encabezado |
| seven-columns.doc | DOC 97 | 1 | 2 | 1 | 0 | 0 | 0 | 0 |
| art-shapes.doc | DOC 97 | 1 | 1 | 1 | 0 | 0 | 0 | 0 |

También se probaron `PageSpecificHeadFoot.docx`, `DiffFirstPageHeadFoot.docx`, `EmptyDocumentWithHeaderFooter.docx`, `form_footnotes.docx`, `table_footnotes.docx`, `png-picture.doc`, `endingnote.doc`, `comment.doc` y `EmbeddedDocument.docx`.

Fixtures:

- https://calibre-ebook.com/downloads/demos/demo.docx
- https://github.com/python-openxml/python-docx/tree/master/tests/test_files
- https://github.com/apache/poi/tree/trunk/test-data/document

## Funciones avanzadas implementadas

### Secciones y columnas

- Modelo `DocumentSection` con tamaño, orientación, márgenes, distancia de encabezado/pie, tipo de salto y columnas.
- Lectura de `w:sectPr` por párrafo y del `sectPr` final del cuerpo.
- Escritura de `sectPr` en el último párrafo de cada sección y en el cuerpo.
- Flujo periodístico de 1 a 4 columnas, con separación y línea opcional.
- Diálogo **Disposición → Columnas** y comando **Salto de sección**.

Limitación: los saltos continuos se tratan como salto de página y las columnas no se equilibran al final de la sección.

### Objetos flotantes y ajuste de texto

- `FloatingObject` para imagen, cuadro de texto, forma y objeto incrustado.
- Lectura de `wp:anchor` (posición, extensión, `wrapSquare`, `wrapTopAndBottom`, `wrapNone`, `behindDoc`).
- Lectura de cuadros de texto `wps:txbx` y VML `v:textbox`, incluso dentro de `mc:AlternateContent`.
- Cuadros de texto y formas insertables desde **Insertar → Objetos**.
- Ajuste cuadrado: el texto se desplaza al lado del objeto; ajuste superior/inferior reserva altura.
- Escritura de imágenes flotantes como `wp:anchor` y de cuadros/formas como VML.

Limitaciones: el ajuste cuadrado se aplica por bandas de párrafo, no con contorno exacto (`wrapTight` se aproxima a cuadrado); no hay edición de vértices ni rotación.

### Notas y comentarios

- Hipervínculos externos: lectura y escritura de `w:hyperlink` con relaciones `TargetMode="External"`.
- Lectura de `footnotes.xml`, `endnotes.xml` y `comments.xml`.
- Referencias `w:footnoteReference`, `w:endnoteReference` y `w:commentReference`.
- Marcadores `[n]` y `[C]` en el texto; notas dibujadas al pie de la página.
- Diálogo **Revisar → Comentarios**.
- Escritura de los tres archivos y sus relaciones al guardar DOCX.

Limitaciones: el área de notas no reserva espacio en el flujo principal; los comentarios se muestran en un panel, no se crean ni editan desde el texto.

### Estilos de tabla

- Lectura de `w:tblStyle`, `w:tblBorders`, `w:tblShd`, `w:tcBorders`, `w:shd` y fusión.
- Estilos conocidos con encabezado, bandas y borde de cuadrícula.
- Dibujo de sombreado, borde por celda, negrita/cursiva de celda.
- Escritura de estilo, bordes, sombreado y celdas.

Limitación: no se resuelve la jerarquía completa de estilos condicionales de Word.

### Objetos incrustados

- Detección de `w:object`, `o:OLEObject` y vista previa `v:imagedata`.
- Dibujo de la vista previa.
- Lectura del flujo OLE (`ProgID`, nombre y bytes) y conservación del binario al guardar en `word/embeddings/`.
- Prueba: `embedded.docx` conserva un objeto Excel de **18.432 bytes** después de ida y vuelta.

Limitación: el objeto no se ejecuta ni se edita; se conserva y se muestra su vista previa.

### Word 97-2003

- Encabezados: lectura del subdocumento `ccpHdd` y dibujo como texto.
- Imágenes: extracción de imágenes PNG/JPEG del flujo `Data` y colocación en el cuerpo.
- Cuadros de texto: intento de lectura del subdocumento `ccpTxbx`/`ccpHdrTxbx`.
- Secciones: geometría básica de la primera sección.

Limitaciones: las imágenes se detectan por firma en `Data` y no siempre se asocian a su ancla; los cuadros de texto se extraen del subdocumento `ccpTxbx` y se muestran como texto, pero no se reconstruye su posición flotante; las formas Escher sin texto no se reconstruyen completamente.

## Comparación con un editor avanzado

**Cubierto**

- Páginas reales, zoom, selección, edición y guardado.
- Tablas, fusiones, estilos básicos y edición directa.
- Imágenes inline y flotantes con ajuste cuadrado.
- Secciones, columnas, encabezados/pies de texto.
- Notas, comentarios, cuadros de texto, formas.
- Lectura parcial de DOC 97-2003 con imágenes y encabezados.

**Pendiente para una compatibilidad cercana a Word**

- Formas con geometría personalizada, rotación, conectores, grupos y texto en trayectoria.
- Ajuste de texto `tight`/`through` con contorno real.
- Notas al pie con reserva de espacio y continuación automática.
- Comentarios editables, control de cambios, marcas de revisión y autores.
- Ejecución/edición de objetos OLE, gráficos, ecuaciones, SmartArt y diagramas.
- Secciones continuas, equilibrio de columnas, notas al final por sección.
- Estilos condicionales completos y herencia de estilos de tabla.
- Reconstrucción completa de OfficeArt/Escher en `.doc`.

## Pruebas automatizadas finales

- Autoprueba de formatos: **47/47**.
- `AdvancedProbe`: **8/8**.
  - secciones y columnas DOCX;
  - cuadro de texto y forma flotantes;
  - notas y comentarios DOCX;
  - estilo de tabla;
  - cuadro de texto DOC 97-2003;
  - dos imágenes DOC 97-2003;
  - encabezado DOC 97-2003;
  - objeto OLE incrustado con 18.432 bytes conservados.
- Compilación con `/warn:4`: sin errores ni advertencias.
- `InteractionProbe`:
  - imagen insertada como objeto flotante y seleccionada;
  - movimiento y redimensionado de la imagen;
  - dos llamadas rápidas a insertar cuadro de texto producen **un solo cuadro**;
  - cuadro de texto seleccionado con ocho tiradores.
- Cinta con **8 pestañas**: Archivo, Inicio, Insertar, Tablas, Disposición, Revisar, Ver y Herramientas. Insertar ya no contiene Tabla ni Navegación.
- La inserción de cuadros de texto se protege con un intervalo mínimo y devuelve el foco al editor.
- `FocusProbe`: después de insertar un cuadro de texto y pulsar **Enter**, el documento conserva **1** cuadro (`TEXTBOX_AFTER_ENTER=1`); la causa era que el salto de párrafo copiaba los objetos flotantes.
- `ShapeProbe`: ocho formas (rectángulo, círculo, triángulo, estrella, flecha derecha, flecha izquierda, rombo y corazón) se dibujan con relleno azul claro y contorno azul oscuro, y sobreviven a la ida y vuelta DOCX.
- `ImageEnter` confirma que una imagen tampoco se duplica al crear un párrafo después de insertarla.
- `WrapProbe`: texto alrededor de una imagen flotante con **0 solapamientos** entre segmentos de texto y el rectángulo de la imagen.
- La viñeta se dibuja como un círculo negro relleno y anti-aliased, no como un glifo de fuente.
- La lista numerada usa un icono independiente junto a la viñeta; al desactivar una lista se limpian la sangría izquierda y la sangría francesa del párrafo.
- `NumberProbe`: **BULLET_CLEARED=True** y **NUMBERED=3** tras crear, convertir y volver a leer tres elementos numerados.
- `ParagraphProbe`: **APPLIED=True** y **ROUNDTRIP=True** para alineación, sangrías, espaciado, interlineado, paginación y control de viudas.
- `DialogProbe`: captura del diálogo **Párrafo** y del **Corrector ortográfico** con sugerencias reales (`fadsqsfaf` → `farsesca`, `farista`, etc.).
- `ShapeProbe`: documentos nuevos en **Arial**, interlineado inicial **1,5**, idioma **es-MX**, corazón con el mismo formato azul de las demás formas y selector gráfico con **Flecha derecha** y **Flecha izquierda**.
- `SpellProbe`: palabra visible completa, idioma `es-MX` y selección del rango `0:0-0:11` al revisar una palabra del documento.
- Las formas se seleccionan con un diálogo previo y **Color de forma** permite cambiar relleno, contorno y grosor; el corazón conserva el formato estándar de las formas.
- `AuditProbe`: **SHAPES=8**, **READ_SHAPES=8**, **COLORS=True**; las ocho formas se insertan, personalizan y guardan sin bloqueo (~3,7 s incluyendo esperas antiduplicado).
- `ImageAuditProbe`: **IMAGES=5**, **READ_IMAGES=5**; cinco imágenes se insertan y vuelven a leerse sin bloqueo (~1,5 s).
- `TableProbe`: una fila completa de 5 columnas se fusiona correctamente en ambos sentidos de selección; añadir y quitar columnas mantiene la tabla operativa.
- `RulerProbe`: la regla superior muestra centímetros y ancho de hoja; `IndentProbe`: **UP=360**, **DOWN=0**, **FLOOR=0**.
- `SpellLongProbe`: una palabra de 48 caracteres se muestra completa en dos líneas.
- `SecurityProbe`: **ATOMIC=True**, **NEW=True**, **BACKUP=True**; el guardado conserva la versión anterior en `.bak`.
- La regla es una banda fija de la ventana y permanece visible al desplazarse.
- Los cuadros de texto permiten editar fuente, tamaño, negrita, cursiva, subrayado y color desde su diálogo.
- `RecoveryProbe`: el autoguardado de un documento sin guardar crea `recovery-*.docx` y su archivo `.meta`; al iniciar se ofrece la recuperación.
- `RowProbe`: **ROWS_ADD=3**, **ROWS_REMOVE=2**, **RULER=True**; la regla está activa por defecto.
- La cinta tiene una pestaña independiente **Tablas**; Insertar ya no contiene Tabla ni Navegación, y se eliminó **Editar celda**.
- `HeaderProbe`: 2 encabezados y 2 pies independientes por sección, con imágenes y relaciones propias.
- Corpus público: 11/11 DOCX de python-docx y Apache POI abiertos, guardados y reabiertos.
- `FuzzProbe`: 60/60 entradas aleatorias rechazadas; DTD externa y ruta ZIP `../` bloqueadas.
- `UnsupportedProbe`: detección de campos y controles de contenido antes de guardar.
- `CrashLog`: log local con contexto y excepción.
- `OpenFile`: 11/11 documentos del corpus y 3 documentos reales abiertos desde la interfaz, con ventana maximizada.
- Regla en control fijo separado del área desplazable, sin parpadeos ni marcas duplicadas.
- `TextBox`: fondo de cuadro de texto editable y conservado en DOCX.
- `SpellCheckDialog`: palabra larga y botón «Añadir al diccionario» completos.
- Encabezados/pies: formato de tamaño, color y alineación editable, conservado en DOCX y visible en la vista de diseño con banda diferenciada.
- `ShapeFormatDialog`: grosor visible y con unidad `px`.
- `SizeProbe`: cambiar tamaño conserva la posición (`SCROLL=500/500`) y el resto del formato de cada tramo (negrita, cursiva, familia y color).
- `ShapeRoundtrip`: una forma en `OffsetXTwips=2400`, `OffsetYTwips=1800` conserva posición, tamaño, forma, relleno y modo de ajuste al guardar/reabrir (`X=2400`, `Y=1800`).
- `ShapeLayoutProbe`: la posición calculada no cambia antes/después de reabrir (`X=218,9`, `Y=152,4667`) y `TopAndBottom` se conserva.
- `ContextFormatProbe`: formato en tabla, celda en edición y cuadro de texto sin perder selección (`SCROLL=300/300`, `KEEP_SELECTION=True`, `TEXTBOX=True`).
- `TableFormatRoundtrip`: una celda con Arial 18 pt y color rojo conserva formato al guardar/reabrir.
- `TextBoxFormatProbe`: Arial 14 pt, color y negrita se conservan en el cuadro de texto.
- `TableNavProbe`: flecha derecha de A a B, flecha izquierda de B a A y flecha abajo en la selección de tabla (`SELECTION=1,0`).
- `TableExitProbe`: al llegar al borde, la flecha izquierda sale antes de la tabla (`LEFT_PARA=0`) y la derecha sale después (`RIGHT_PARA=2`), sin editor de celda activo.
- `TableTopExitProbe`: una tabla al inicio sale al párrafo siguiente sin crear párrafos nuevos (`PARA=1`, `COUNT=2` antes y después).
- `MergedNavProbe`: desde una celda fusionada, abajo pasa a la fila 2 (`DOWN=2:D`), arriba vuelve a la fila 0 y arriba en el borde superior sale de la tabla (`TOP_PARA=1`, `COUNT=2/2`).
- `BodyFormatProbe`: cambiar tamaño/fuente en un párrafo solo altera ese párrafo (`P0=Arial/22`, `P1=Verdana/40`, `P2=Courier New/24`).
- `FinalFormatProbe`: formato corporal aislado (`BODY=True`), una sola celda afectada (`CELL_ONLY=True`), alto de fila recalculado al instante (`ROW_H=56,8`), selección conservada y scroll estable (`SCROLL=300/300`).
- `TwoTableProbe`: con dos tablas, formatear una celda de la segunda no toca la primera (`FIRST=True`, `SECOND=True`) y mantiene el cursor en `0,0` de esa tabla.
- `RoundtripScopeProbe`: tras guardar y reabrir, solo el fragmento seleccionado conserva el tamaño nuevo; los párrafos vecinos mantienen `Times New Roman/22`.
- `BulletProbe`: un párrafo de viñeta de 3 líneas y uno numerado de 3 líneas producen exactamente **1 viñeta** y **1 número**, no una marca por línea.
- `BulletUserTest`: documento de **3 páginas** con 5 viñetas; al poner la primera viñeta a 30 pt solo esa queda en `60` medios puntos, las otras cuatro siguen en `22`; las líneas largas conservan **1 viñeta** y tras reabrir se mantiene `60/22`. La marca de la viñeta mantiene el tamaño base de 11 pt y no crece con el texto.
- `MergedClickProbe`: pulsar una celda fusionada conserva el viewport (`BEFORE=431`, `AFTER=431`) y activa el editor de celda.
- La pulsación de una celda fusionada conserva la posición del viewport (`_tableClickScroll`).
- Sangría: iconos con flechas en el sentido correcto y viewport conservado en la primera sangría (`SCROLL_BEFORE=900`, `SCROLL_AFTER=900`, `INDENT=360`).

## Diagnóstico

El proyecto ya está en una categoría intermedia-alta para visualización nativa de DOCX simple y medio: **paginación, secciones, columnas, tablas con estilo, imágenes flotantes, notas y comentarios**. Para competir con Word en documentos complejos falta especialmente el motor de formas/OLE y el ajuste de texto fino.
