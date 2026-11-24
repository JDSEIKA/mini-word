# Auditoría de funciones y accesibilidad

Fecha: **2026-11-18**  
Versión: **0.0.8 Beta**

Este documento revisa los comandos públicos del editor y comprueba si tienen un acceso visible en la cinta, un diálogo, el ratón o un atajo de teclado. No se consideran “inaccesibles” las funciones internas de maquetación, impresión o edición de celdas.

## Accesibles mediante iconos o botones

| Función | Acceso |
|---|---|
| Nuevo, Abrir, Guardar, Guardar como, Cerrar | Archivo |
| Recientes, Imprimir, Vista previa, Autoguardado | Archivo |
| Pegar, Pegar sin formato, Cortar, Copiar, Deshacer, Rehacer | Inicio |
| Fuente, tamaño, negrita, cursiva, subrayado, color | Inicio |
| Alineación, sangría, viñetas, numeración | Inicio |
| Encabezado y pie, Salto de página, Hipervínculo | Inicio |
| Imagen, Fecha, Ajuste de texto | Insertar → Imagen y fecha |
| Cuadro de texto, Formato de texto, Forma, Color de forma | Insertar → Objetos |
| Traer al frente, Enviar al fondo, Rotar 90° | Insertar → Objetos |
| Insertar/editar tabla, filas, columnas, fusionar/dividir | Tablas |
| Configurar página, Párrafo, Columnas, Salto de sección, Navegación | Disposición |
| Activar ortografía, idioma, Ortografía, Comentarios, Buscar y reemplazar | Revisar |
| Regla, Zoom, Página anterior/siguiente | Ver |
| Ayuda, Acerca de | Ayuda |

## Accesibles por teclado o diálogo, sin icono exclusivo

| Función | Acceso alternativo |
|---|---|
| Seleccionar todo | `Ctrl+A` |
| Buscar siguiente / reemplazar | Diálogo Buscar y reemplazar |
| Revisar ortografía de nuevo | Se ejecuta al abrir/actualizar el diálogo de ortografía |
| Insertar fecha | Botón Fecha; usa `InsertText` |
| Editar tabla completa | Botón Editar tabla → diálogo; incluye eliminar y reemplazar tabla |
| Selección múltiple de celdas | Arrastrar con el ratón sobre la tabla |
| Editar una celda | Doble clic o clic en la celda |
| Cambiar formato de una selección | Botones de fuente/párrafo o diálogos |
| Configurar encabezado/pie por variantes | Diálogo Encabezado y pie |
| Configurar página | Diálogo Configurar página |
| Imprimir / vista previa | Iconos de Archivo |

## Funciones de API o internas

Estas funciones son deliberadamente programáticas y no necesitan icono:

- `InsertText`: inserción usada por Fecha y pegado.
- `SelectRange`, `PlaceCaret`: API de diálogos y pruebas.
- `RecheckSpelling`, `GetSpellParagraphs`: integración con el diálogo de ortografía.
- `ClearSpellingHighlight`, `SelectSpellingRange`: resaltado del diálogo.
- `ReplaceTable`, `DeleteTable`, `GetTableRows`: las usa el editor de tablas.
- `TableParagraphAt`, `SelectTableCells`: ratón y edición de celdas.
- `EditSelectedTableCell`, `FocusCellEditor`, `SetCellEditorCommitSuppressed`, `CellEditorFormat`: edición interna de celdas.
- `ApplySelectionFormat`: aplicada por los diálogos de formato.
- `SetPageSize`, `SetMargins`: las usa Configurar página.
- `SetCurrentHeaderFooter`: lo usa el diálogo de encabezado/pie.
- `DrawPageForPrint`, `GetRulerMetrics`: impresión y regla.
- `GetImageMaxWidthTwips`: cálculo interno de inserción.
- `RefreshDocument`: normalización interna tras cambios de layout.

## Conclusión

No se detectó ninguna función de usuario importante sin acceso. Las únicas funciones sin **icono dedicado** son:

1. **Seleccionar todo** (`Ctrl+A`).
2. **Revisar de nuevo la ortografía**, que se inicia desde el diálogo.
3. **Insertar fecha**, que tiene botón propio y usa la API `InsertText`.
4. Las operaciones internas de tabla y celda, accesibles mediante los diálogos y el ratón.

El salto de sección ya tiene botón en **Disposición → Sección**.
