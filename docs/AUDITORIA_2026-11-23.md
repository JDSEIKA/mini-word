# Auditoría estricta — 2026-11-23

> **Documento histórico.** La versión actual del proyecto es **0.0.16 Beta**; las versiones citadas pertenecen al alcance de este informe.

Alcance: Mini Word 0.0.8 Beta, cambios de las Fases A/B y release 0.0.7.

## Método

- Compilación completa con `csc /warn:4`.
- Regresión completa: 101 documentos, UI, portapapeles, funciones, caret y seguridad.
- Revisión manual de los cambios recientes.
- Pruebas reproducibles de tablas fusionadas, estilos, revisiones individuales y campos multipágina.
- Revisión de documentación y artefactos publicados.

## Resultado de la línea base

```text
CORPUS=101 OK=101 FAIL=0
UI_SMOKE=PASS
CLIPBOARD_SMOKE=PASS
FEATURE_AUDIT=PASS
CARET_SMOKE=PASS
SECURITY_SMOKE=PASS
REGRESSION=PASS
```

La línea base pasa, pero **no cubría varios defectos nuevos**. La auditoría encontró y corrigió cuatro defectos reales.

## Hallazgos

| ID | Severidad | Hallazgo | Estado |
|---|---|---|---|
| A-01 | Crítica | Una celda con `vMerge` escribía el texto del propietario otra vez en la celda de continuación. El DOCX podía duplicar el contenido al abrirse en Word/LibreOffice. | Corregido |
| A-02 | Alta | Abrir el editor de una celda con varios tramos borraba los tramos inmediatamente, aunque el usuario no escribiera nada. | Corregido |
| A-03 | Alta | Aplicar estilo, cambiar nivel de lista, alternar encabezado de tabla o aceptar/rechazar un cambio no apilaba undo. | Corregido |
| A-04 | Alta | Los índices TOC guardaban números de párrafo; insertar o eliminar párrafos antes del destino dejaba el índice apuntando a otro párrafo. | Corregido en las rutas principales |
| A-05 | Alta | El binario y la documentación publicados quedaron desactualizados después de los cambios de las Fases A/B. | Corregido en 0.0.8 |
| A-06 | Media | Las definiciones de estilo (`w:style`) son básicas: nombre y `basedOn`, sin una definición completa de `pPr/rPr`. El formato directo sí se conserva. | Pendiente |
| A-07 | Media | `Repetir encabezado` exporta `w:tblHeader`, pero la vista previa nativa no divide tablas largas. | Limitación declarada |
| A-08 | Media | El estado “resuelto” de comentarios es local y no se serializa en DOCX. | Limitación declarada |
| A-09 | Media | La comparación de formato solo emite un aviso; no crea revisiones de formato. | Limitación declarada |
| A-10 | Media | `RefreshFieldValues` recorre todas las líneas por cada campo; puede degradarse en documentos con muchos campos. | Pendiente |
| A-11 | Baja | El lector marca `tblHeader` si aparece en cualquier fila, no solo en la primera. | Pendiente |
| A-12 | Baja | Quedan rutas de edición estructural no cubiertas por pruebas de ajuste TOC. | Pendiente |

## Evidencia de los defectos corregidos

### A-01: fusión vertical duplicada

Antes:

```xml
<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:vMerge/><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
```

Después:

```xml
<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:vMerge/><w:p /></w:tc></w:tr>
```

Se añadió una prueba de regresión que verifica que `A` aparece una sola vez.

### A-02: pérdida de tramos al abrir una celda

`StartCellEditor` hacía `Runs.Clear()` antes de que el usuario editara. Ahora:

- abrir la celda no borra los tramos;
- solo se aplanan cuando cambia el texto;
- cancelar o salir sin escribir conserva el formato rico.

### A-03: undo

Se añadió `PushUndo()` antes de:

- aplicar estilo;
- cambiar nivel de lista;
- alternar encabezado de tabla;
- aceptar/rechazar un cambio.

### A-04: índices TOC

Se añadieron `AdjustTocOnInsert` y `AdjustTocOnRemove` en:

- división de párrafo;
- unión de párrafos;
- eliminación de selección multipárrafo;
- inserción/eliminación de tablas;
- salto de página.

## Seguridad

Revisión estática:

- No hay llamadas de red (`HttpClient`, `WebClient`, sockets).
- XML con `DtdProcessing.Prohibit`.
- Límites de ZIP, XML, imágenes y RTF.
- Los hipervínculos solo abren `http`, `https` o `mailto` tras confirmación.
- No se ejecutan macros, ActiveX, scripts ODT ni OLE.
- `ShellExecute`/`Process.Start` solo se usan para el enlace confirmado.

Riesgo residual:

- El parser DOCX/ODT sigue siendo código propio y requiere fuzzing continuo.
- El portapapeles y las rutas de archivos dependen de Windows.
- No hay firma de código; SmartScreen puede advertir.

## Documentación

Se corrigieron en el README:

- pestañas reales: Inicio, Insertar, Tablas, Disposición, Revisar, Ver y Ayuda;
- funciones nuevas de campos, TOC, cambios, comparación, estilos y listas;
- enlaces por tramo en tablas;
- limitación real de tablas largas;
- contradicción sobre OLE.

Siguen siendo históricos y no deben tomarse como estado actual:

- `docs/COMPATIBILITY.md` (auditoría 2026-11-13);
- `docs/BETA_SCOPE.md` (criterios anteriores a las Fases A/B);
- `docs/ANALISIS_EXTERNO_PROCESADOR_TEXTO.md` (análisis previo a 0.0.6).

## Veredicto

**No está listo para producción.**

Está en condiciones de **beta técnica** con estas reservas:

1. Firma de código.
2. Matriz real Word 2007–365/LibreOffice.
3. Pruebas de pérdida de datos con documentos reales.
4. Cierre de tablas largas y estilos completos.
5. Fuzzing continuo de DOCX/ODT/RTF.

La regresión automatizada es útil, pero no demuestra compatibilidad total con Word. La afirmación correcta es **“beta funcional con limitaciones documentadas”**, no “compatible con Word”.
