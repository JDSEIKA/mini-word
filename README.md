# Mini Word

Procesador de textos de escritorio para Windows con un **motor tipográfico propio**, edición nativa de DOCX y una interfaz inspirada en Microsoft Word. No usa Word, Office, LibreOffice, RichEdit en la vista principal, NuGet ni bibliotecas externas en tiempo de ejecución.

El ejecutable es un único `MiniWord.exe` de unos **451 KB** y solo depende de .NET Framework 4.0 o posterior.

Versión: **0.0.6 Beta**. Estado: **beta pública inicial**. El alcance y las limitaciones están en [`docs/BETA_SCOPE.md`](docs/BETA_SCOPE.md). La vista y la edición usan un **motor tipográfico propio**; la ruta principal no crea ningún control RichEdit.

## Ejecutable incluido en el repositorio

El ejecutable compilado también está versionado en la raíz del repositorio:

- [MiniWord.exe](https://github.com/JDSEIKA/mini-word/blob/main/MiniWord.exe) — 451 KB
- SHA-256: 13B1DAFB81C9ADFBF078408A6B9BA7768F449791A3D1E4016A636F42CC0B3DA1

No es un archivo generado por una GitHub Action: está incluido para que pueda descargarse directamente desde el código fuente. La release contiene además el paquete beta y el instalador portátil.

## Probar y reportar errores

La beta pública inicial está en [Releases](https://github.com/JDSEIKA/mini-word/releases/tag/v0.0.6-beta).

- [Ejecutable directo](https://github.com/JDSEIKA/mini-word/releases/download/v0.0.6-beta/MiniWord.exe)
- [Paquete beta](https://github.com/JDSEIKA/mini-word/releases/download/v0.0.6-beta/MiniWord-Beta.zip)
- [Instalador portátil](https://github.com/JDSEIKA/mini-word/releases/download/v0.0.6-beta/MiniWord-Installer.zip)

Para reportar un fallo usa el [formulario de error](https://github.com/JDSEIKA/mini-word/issues/new?template=bug_report.yml). Incluye un DOCX mínimo sin datos privados, la versión de Windows y los pasos exactos. Más detalles en [CONTRIBUTING.md](CONTRIBUTING.md).

## Formatos

| Formato | Abrir | Guardar | Implementación |
|---|---:|---:|---|
| Office Open XML `.docx` | Sí | Sí | Lector/escritor WordprocessingML nativo |
| Word 97-2003 `.doc` / `.dot` | Sí | Se convierte a `.docx` | Lector OLE2/CFB + MS-DOC nativo |
| OpenDocument Text `.odt` | Sí | Sí | Lector/escritor ODF 1.2 nativo |
| Rich Text Format `.rtf` | Sí | Sí | Lector/escritor RTF nativo |
| Texto plano `.txt` | Sí | Sí | UTF-8, UTF-16, ANSI |

Los documentos `.doc` se abren en modo de importación. Al guardar, Mini Word avisa y crea un `.docx`; nunca sobrescribe el `.doc` original.

## Funciones

- **Cinta de iconos** con pestaña Archivo en azul y pestañas Inicio, Insertar, Disposición, Revisar, Ver y **Herramientas**; iconos vectoriales propios, grupos de comandos, botones grandes, tooltips y adaptación del ancho. No hay barra de menús desplegable: todos los comandos están en la cinta y en sus atajos de teclado.
- **Herramientas**: Ortografía, Autoguardado, Ayuda y Acerca de, con iconos propios en una pestaña dedicada.
- **Motor tipográfico nativo**: cada párrafo se divide en líneas medidas con GDI+, las líneas fluyen a páginas y el cursor, la selección y el clic se resuelven contra esas cajas de línea. No se rasterizan páginas ni se usa RichEdit para el documento.
- **Fuentes del sistema**: el selector de fuente carga todas las familias instaladas en Windows, no una lista fija.
- **Tipografía multilingüe**: normalización Unicode NFC, sustitución de fuente por script (latino, cirílico, griego, árabe, hebreo, índico, tailandés y CJK) y compilación con página de códigos UTF-8 para que los textos y la interfaz no pierdan acentos ni signos.
- **Páginas reales apiladas**: todas las hojas se dibujan en vivo, con márgenes, encabezado, pie, numeración y huecos entre páginas; la barra vertical y la rueda recorren el documento de forma continua.
- La vista de trabajo y la impresión comparten el mismo modelo de página, márgenes y saltos, de modo que el corte de página coincide con la salida impresa.
- Formato de carácter: fuente, tamaño, negrita, cursiva, subrayado, tachado, color, resaltado, superíndice y subíndice.
- Formato de párrafo: alineación, sangrías con iconos para aumentar/disminuir, espaciado, interlineado, tabulaciones y viñetas.
- **Corrección ortográfica del sistema**: usa el corrector de Windows (Windows 8+) por COM, sin diccionarios empaquetados; subraya palabras dudosas y ofrece sugerencias, omitir y agregar al diccionario al pulsar con el botón derecho. El control de idioma está en Herramientas.
- **Tablas**: pestaña propia con insertar/editar tabla, **Agregar fila**, **Eliminar fila**, fusionar/dividir celdas, añadir/quitar columnas y edición directa de celdas; conservación en DOCX (`gridSpan`/`vMerge`).
- **Párrafo**: diálogo en Disposición con alineación, nivel de esquema, sangrías, especial/francesa, espaciado anterior/posterior, interlineado, paginación y control de viudas.
- **Revisión ortográfica**: Herramientas incluye el control de activación, el idioma ortográfico detectado del sistema y el botón Ortografía, que abre el diálogo con palabra dudosa, sugerencias, cambiar, ignorar, ignorar todas y añadir al diccionario.
- **Imágenes**: inserción desde archivo como objeto flotante, ancho máximo de 2/3 de la página calculado antes de importar, relación de aspecto bloqueada, selección con clic, ocho tiradores para redimensionar y arrastre directo para recolocarlas en la hoja.
- **Cuadros de texto y formas**: se insertan seleccionados, con ocho tiradores para redimensionar y arrastre para moverlos; el doble clic o **Formato de texto** permite editar fuente, tamaño, negrita, cursiva, subrayado y color. El selector gráfico permite rectángulo, círculo, triángulo, estrella, flecha derecha, flecha izquierda, rombo y corazón. Color de forma permite cambiar relleno, contorno y grosor.
- **Fecha**: inserta la fecha actual con el formato largo del idioma de Windows.
- **Autoguardado opcional**: el botón Autoguardado abre un diálogo para activarlo/desactivarlo y elegir el intervalo de 1 a 120 minutos; se guarda en el documento o en una copia de recuperación.
- **Recuperación automática**: si la aplicación se cierra inesperadamente, al iniciar ofrece recuperar el último autoguardado y conserva la ruta original.
- **Aviso de compatibilidad**: si el documento contiene campos, TOC, controles de contenido, revisiones o gráficos no modelados, se avisa antes de guardar y se conserva el original en `.bak`.
- **Guardado atómico**: cada guardado escribe un temporal, sustituye el original y conserva una copia `.bak`; los paquetes ZIP, XML e imágenes tienen límites de tamaño.
- **Ctrl + rueda del ratón** cambia el zoom de la página de edición.
- **Secciones y columnas periodísticas**: `DocumentSection` conserva tamaño, orientación, márgenes y 1-4 columnas; Disposición → Columnas y Salto de sección.
- **Objetos flotantes**: imágenes, cuadros de texto y formas con posición, ajuste cuadrado/superior-inferior y orden detrás/delante del texto.
- **Encabezados y pies por sección**: cada sección conserva su propio texto e imágenes, con relaciones DOCX independientes.
- **Notas y comentarios**: notas al pie/finales, referencias en el texto y panel de comentarios en Revisar.
- **Estilos de tabla**: estilo, bordes, sombreado, encabezado y celdas combinadas.
- **Objetos incrustados**: detección, vista previa y conservación del binario OLE en `word/embeddings/`; no se ejecuta el objeto.
- **Regla fija activa por defecto**: Ver → Regla muestra una banda superior de la interfaz en centímetros; permanece visible al desplazarse por todas las hojas.
- **Word 97-2003**: encabezados y extracción de imágenes del flujo `Data`; los cuadros de texto Escher siguen siendo una limitación.
- Apertura y conservación de tablas e imágenes en DOCX; las tablas se dibujan con su rejilla y las imágenes se escalan en la línea.
- Encabezados y pies editables con `{PAGE}` y `{NUMPAGES}`; se guardan en DOCX y se dibujan en pantalla, impresión y vista previa.
- Crear página con `Ctrl+Enter`, saltos de página, buscar/reemplazar, deshacer/rehacer, cortar/copiar/pegar e impresión.
- **Compatibilidad de objetos**: formas y cuadros de texto se exportan como VML para Word 2007/2010+ y LibreOffice; las imágenes flotantes usan anclas relativas a página. Esto evita que cambien de página al abrirlas.
- **Hipervínculos**: el texto con URL se dibuja en azul y subrayado, incluso si el formato original era negro. Con `Ctrl+Click` se muestra el destino y se pide autorización antes de abrirlo. El botón **Hipervínculo** permite crear o editar enlaces; también se admiten enlaces por celda completa en tablas.
- **Orden Z y rotación**: **Traer al frente**, **Enviar al fondo** y **Rotar 90°** para formas, imágenes y cuadros de texto.
- **Archivo → Recientes**: lista desplegable con los últimos 5 documentos abiertos o guardados.
- **Apertura segura**: rechazo de ZIP anidados malformados, bombas de descompresión, rutas `../`, entradas duplicadas, DTD/entidades XML, RTF excesivamente anidado y BOM inválidos. Macros VBA, ActiveX y scripts ODT se avisan y no se ejecutan.
- **Enlaces en tablas**: enlace por celda completa, con color azul, subrayado, clic y conservación en DOCX.
- **Pegar sin formato**: `Ctrl+Shift+V` o el botón de la cinta inserta solo el texto, con el formato base del documento.
- **Portapapeles con formato**: al pegar desde el navegador o desde Word conserva tamaño, negrita, cursiva, subrayado y color; el texto plano se pega con el formato base del documento, no con el tamaño seleccionado en la barra.
- **Corrección de codificación al pegar**: repara texto UTF-8 que Windows entrega como Windows-1252, por ejemplo `tambiÃÂ©n` → `también`.
- **Vista previa navegable con la rueda del ratón**: cada giro avanza o retrocede una página de la vista previa de impresión.
- Lector OOXML compatible con paquetes *Transitional* y *Strict* (normalización de espacios de nombres).
- Lector MS-DOC con tabla de piezas, estilos, CHPX/PAPX, secciones, listas, tablas y campos.
- Informe de compatibilidad y brechas: [`docs/COMPATIBILITY.md`](docs/COMPATIBILITY.md).
- Registro de fallos y auditorías por fecha: [`docs/FALLOS_2026-11-17.md`](docs/FALLOS_2026-11-17.md).
- Funciones avanzadas: [`docs/FUNCIONES_AVANZADAS.md`](docs/FUNCIONES_AVANZADAS.md).
- Análisis del procesador externo: [`docs/ANALISIS_EXTERNO_PROCESADOR_TEXTO.md`](docs/ANALISIS_EXTERNO_PROCESADOR_TEXTO.md).
- Auditoría RichEdit: [`docs/AUDITORIA_RICHEDIT.md`](docs/AUDITORIA_RICHEDIT.md).
- Dictado y lectura en voz alta: [`docs/VOZ_DICTADO.md`](docs/VOZ_DICTADO.md).
- Accesibilidad de comandos: [`docs/ACCESIBILIDAD_FUNCIONES.md`](docs/ACCESIBILIDAD_FUNCIONES.md).
- Auditoría de seguridad y recuento consolidado: [`docs/FALLOS_2026-11-18.md`](docs/FALLOS_2026-11-18.md).

## Compilar

Requisitos: Windows 7 o posterior y .NET Framework 4.x (incluye `csc.exe`).

```bat
build.cmd
```

El resultado queda en `dist\MiniWord.exe`. También se puede abrir `MiniWord.csproj` con Visual Studio/MSBuild.

Compilar y ejecutar todas las pruebas:

```bat
build.cmd -RunSelfTest
```

La autoprueba es headless y cubre RTF, DOCX, ODT, TXT, imágenes, tablas, Unicode y un `.doc` real de Word 97 generado con LibreOffice. La interfaz nativa se prueba abriendo DOCX, escribiendo, moviendo el cursor entre páginas, recorriendo las pestañas de la cinta y guardando de nuevo.

## Estructura

```
src/
  Program.cs                 Punto de entrada y --selftest
  NativeEditorForm.cs        Interfaz nativa, comandos de archivo y formato
  Controls/
    NativeDocumentEditor.cs  Layout de líneas, paginación, caret, selección y pintura
    FontResolver.cs          Selección de fuente por script/Unicode
    SystemSpellChecker.cs    Corrector ortográfico de Windows por COM/MTA
    RibbonGroup.cs           Grupos y render de la cinta
    RibbonIcons.cs           Iconos vectoriales de la cinta
  Dialogs/                   Buscar, página, tabla, columnas, comentarios, autoguardado y encabezado/pie
    NativeFindDialog.cs      Buscar/reemplazar del editor nativo
    NativePageSetupDialog.cs Tamaño, orientación y márgenes
    TableEditorDialog.cs     Inserción y edición de tablas
    AutoSaveDialog.cs        Intervalo y activación del autoguardado
    ColumnsDialog.cs         Número y separación de columnas
    ShapePickerDialog.cs     Selector de forma antes de insertar
    FloatingTextDialog.cs    Edición del texto de un cuadro o forma
    HelpDialog.cs            Guía rápida de funcionamiento
    ParagraphDialog.cs       Sangría, espaciado, interlineado y paginación
    ShapeFormatDialog.cs     Relleno, contorno y grosor de las formas
    SpellCheckDialog.cs      Revisión ortográfica con sugerencias
    CommentsDialog.cs        Panel de comentarios del documento
   HeaderFooterDialog.cs    Editor de encabezado, pie y distancias
  Formats/
    CompoundFile.cs          Lector OLE2/CFB (Word 97-2003)
    DocFormat.cs             Lector MS-DOC: FIB, CLX, CHPX/PAPX, estilos, listas
    DocxFormat.cs            Lector/escritor Office Open XML
    OdtFormat.cs             Lector/escritor OpenDocument
    RtfReader.cs             Lector RTF (incluye tablas)
    RtfWriter.cs             Escritor RTF
    ZipArchiveLite.cs        ZIP propio (usa DeflateStream de la BCL, no ZipArchive)
    TextFile.cs              Codificaciones de texto
test-assets/
  sample.doc                 Fixture Word 97 con formato, listas, tabla e imagen
  ascii.doc                  Fixture Word 97 de texto ANSI simple
  sample.html                Fuente reproducible del fixture principal
assets/
  MiniWord.ico               Icono multirresolución
  app.manifest               DPI, compatibilidad y ejecución como usuario normal
```

## Compatibilidad OOXML

Se cubre el subconjunto editable que utiliza un procesador de textos ligero:

- párrafos, secciones, estilos de párrafo/carácter, numeración y listas;
- runs con formato, colores, resaltado, superíndice/subíndice, tabulaciones y saltos;
- tablas, imágenes DrawingML y VML y propiedades principales;
- tamaño de página, orientación, márgenes, encabezados y pies con campos de página;
- la vista dibuja las hojas apiladas; rango de página, márgenes y encabezados coinciden con la impresión.

Limitaciones conocidas: la edición directa de celdas usa un campo de texto superpuesto y no conserva formato rico por celda; las tablas se dividen en páginas por filas completas y las fusiones verticales complejas se simplifican. Las columnas se aproximan por bandas y los saltos continuos se tratan como salto de página. El ajuste de texto es cuadrado/superior-inferior, no contorno exacto. Las notas no reservan espacio en el flujo y los comentarios se consultan en un panel; no se editan desde el texto. Los objetos OLE se muestran mediante su vista previa, no se ejecutan ni se conserva su binario. En `.doc` se importan encabezados e imágenes del flujo Data, pero las formas Escher, los cuadros de texto y la asociación exacta con su ancla siguen siendo limitados. La corrección ortográfica requiere el corrector de Windows (Windows 8 o posterior) y su idioma debe estar instalado.

## Motor tipográfico

La vista sigue una separación estricta entre **modelo**, **layout** y **pintura**:

- el `DocumentModel` de DOCX/ODT/RTF/DOC es la única fuente de verdad; no hay conversión a RTF para editar;
- `NativeDocumentEditor` convierte cada párrafo en tokens, mide las palabras con `Graphics.MeasureString`, aplica sangrías, tabulaciones, alineación y viñetas, y construye cajas de línea;
- las líneas se colocan en páginas con los márgenes del documento; el pie de página y el encabezado se dibujan en cada hoja;
- el cursor se calcula midiendo el prefijo de la línea; el clic se resuelve contra las mismas cajas de línea; la selección se pinta por intervalos de caracteres;
- la barra de desplazamiento es la del control de páginas, no la de un editor de texto continuo, así que no hay saltos a la página 1 al escribir o navegar.

Referencia de diseño: [WordCraft](https://github.com/storytold/wordcraft), en particular su [arquitectura de layout](https://github.com/storytold/wordcraft/blob/main/docs/architecture.md), el [modelo de páginas](https://github.com/storytold/wordcraft/blob/main/crates/layout/src/page.rs) y el [canvas de la UI](https://github.com/storytold/wordcraft/blob/main/crates/ui-egui/src/canvas.rs). Mini Word adapta esas ideas a C#/WinForms/GDI+ y a su modelo DOCX compacto.

## Privacidad

Todo el procesamiento es local. Mini Word no accede a la red, no usa telemetría y no envía documentos a ningún servicio.
