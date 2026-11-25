# Funciones avanzadas

Fecha: **2026-11-22**

Estas funciones se inspiraron en el análisis del procesador de texto externo
`procesador-texto-main` y se implementaron sobre el modelo nativo de Mini Word.

## Campos

### Insertar

**Insertar → Campos e índice → Campo**

Campos disponibles:

- Número de página (`PAGE`).
- Total de páginas (`NUMPAGES`).
- Fecha (`DATE`).
- Hora (`TIME`).
- Título del documento (`TITLE`).
- Autor (`AUTHOR`).

El campo se guarda con instrucción WordprocessingML y se exporta como
`w:fldSimple`, por lo que Word puede reconocerlo y actualizarlo.

### Límites

- El valor mostrado se calcula al insertar.
- `PAGE` no cambia automáticamente cuando el campo se mueve a otra página.
- Para actualizarlo, vuelva a insertarlo o edite el documento.
- Los campos complejos que Mini Word no modela se conservan como XML original.

## Índice (TOC)

### Insertar

**Insertar → Campos e índice → Índice**

Genera:

- Un título **Índice**.
- Una entrada por cada párrafo con nivel de esquema.
- Niveles 1 a 3.
- Número de página calculado con el layout actual.
- Sangría por nivel.
- Tabulador derecho con puntos.

### Actualizar

**Insertar → Campos e índice → Actualizar índice**

Vuelve a calcular las entradas y los números de página.

### Requisito

Los párrafos deben tener **nivel de esquema** (`Outlining level`), que se
establece desde **Disposición → Párrafo**. Los títulos de Word traen ese nivel.

### Límites

- El índice es estático en el modelo nativo; no es un campo TOC complejo.
- Al abrir un DOCX con TOC de Word, el XML original se conserva.
- El índice no se actualiza solo al cambiar el documento; use **Actualizar índice**.

## Control de cambios

### Activar

**Revisar → Cambios → Control de cambios**

Al activarlo:

- Lo que se escribe se marca como inserción.
- Lo que se borra se marca como eliminación y permanece visible.
- Las inserciones se dibujan en verde subrayado.
- Las eliminaciones se dibujan en rojo tachado.

### Aceptar y rechazar

**Revisar → Cambios → Aceptar todo**  
**Revisar → Cambios → Rechazar todo**

También se conservan en DOCX como:

- `w:ins`
- `w:del`
- `w:delText`

### Límites

- Los cambios se registran por tramo de texto.
- Unir o dividir párrafos no se registra como cambio.
- Los cambios en tablas se registran por el texto de la celda.
- Las marcas de revisión de párrafo no se modelan.

## Comparar documentos

**Revisar → Cambios → Comparar documentos**

1. Seleccione el documento anterior.
2. Mini Word alinea los párrafos con LCS.
3. Las diferencias se muestran con control de cambios:
   - Texto eliminado en rojo.
   - Texto nuevo en verde.
4. Puede aceptar o rechazar los cambios.
5. Guarde el resultado como DOCX.

### Límites

- Compara el cuerpo del documento.
- No compara encabezados, pies, notas ni comentarios.
- No detecta cambios que sean solo de formato.
- Los párrafos muy distintos se muestran como bloques eliminados e insertados.

## Compatibilidad

- Los campos simples y las revisiones se exportan a DOCX.
- Word y LibreOffice pueden abrir el resultado.
- Los campos desconocidos y los TOC complejos se preservan, no se reconstruyen.
- El guardado conserva `.bak` y las medidas de seguridad existentes.

## Pruebas

Las pruebas cubren:

- Round-trip de campos.
- Round-trip de inserciones y eliminaciones.
- Comparación de documentos y generación de revisiones.
- Construcción de índice.
- Aceptar/rechazar cambios.

## Actualización funcional acumulada hasta 0.0.16

### Campos dinámicos

- `PAGE` y `NUMPAGES` se recalculan después del layout.
- `DATE` y `TIME` se actualizan al repaginar.
- El resto de campos conserva su valor calculado.

### Índice

- Los números de página del índice se refrescan al repaginar.
- **Actualizar índice** sigue siendo la operación que reconstruye las entradas.

### Cambios individuales

- **Aceptar cambio** y **Rechazar cambio** actúan sobre el cambio bajo el cursor.
- **Aceptar todo** y **Rechazar todo** siguen disponibles.

### Tablas por fragmento

- Una celda puede contener varios tramos con formato y enlace propios.
- Se conservan al abrir y guardar DOCX.
- Al editar una celda, Mini Word la convierte en un tramo único; es una limitación conocida.
- Los enlaces por fragmento funcionan al hacer clic dentro de la celda.

### Estilos con nombre

- **Inicio → Estilos** aplica: Normal, Título, Subtítulo, Título 1, Título 2, Título 3 y Cita.
- Se exportan como `w:pStyle` y se conservan al reabrir.
- No hay editor completo de estilos ni herencia de estilos de Word.

### Listas multinivel

- Hasta 9 niveles.
- **Inicio → Lista** permite aumentar o disminuir el nivel.
- Se exporta `w:ilvl` y la definición de numeración de 9 niveles.
- No hay listas de verificación ni numeración de esquema completo.

### Repetir encabezado de tabla

- **Tablas → Repetir encabezado** marca la primera fila como encabezado.
- Se exporta como `w:tblHeader` y Word/LibreOffice la repiten al paginar.
- **Limitación:** la vista previa de Mini Word todavía coloca la tabla completa en una página; no divide tablas largas.

### Comparación con formato

- Además del texto, la comparación detecta cambios de formato.
- Se muestran como un aviso con el número de párrafos afectados.
- Los cambios de formato no se representan como marcas de revisión.

### Comentarios con respuestas

- **Revisar → Comentarios** permite responder y marcar como resuelto.
- Las respuestas se guardan como párrafos adicionales en `comments.xml`.
- El estado resuelto es local; Word no lo conserva como estado nativo.
### Alineación y formato por palabra en celdas

- Cada celda conserva alineación izquierda, centrada, derecha o justificada.
- Al editar una celda, los botones de alineación afectan a esa celda.
- Se puede seleccionar parte del texto de una celda y aplicar fuente, tamaño, negrita, cursiva, subrayado o color.
- El formato se guarda por tramos en DOCX; la alineación se guarda como `w:jc` en el párrafo de la celda.
- **Vista previa:** el editor de celda usa un `RichTextBox` superpuesto; los cambios de fuente, tamaño, color, negrita y alineación se ven mientras se edita. La tabla de fondo también se actualiza. Escribir texto nuevo aplana los tramos del modelo hasta el siguiente formateo.
### Varios tamaños en una celda

- La altura de la fila se calcula con el mayor tamaño de letra de todos los tramos, no solo con el primer tramo.
- Los tramos de distinto tamaño se alinean por la línea base, no por el borde superior.
- Cambiar el tamaño de la última palabra también amplía la celda.
### Ajuste y recorte en celdas

- El texto de una celda se recorta a los límites de la celda: no invade columnas vecinas.
- Las palabras largas se parten por caracteres cuando no caben.
- La altura de la fila cuenta las líneas necesarias, no solo una.
- Los tramos de distinto tamaño comparten línea base.
### Ortografía en tablas

- La revisión ortográfica incluye el texto de cada celda.
- Las palabras dudosas se subrayan también dentro de la tabla.
- El diálogo puede seleccionar una celda y reemplazar la palabra conservando el formato del tramo.
### Creación de tablas

- El diálogo de inserción numera las filas y las columnas.
- El contador inferior muestra cuántas filas y columnas se van a insertar.
- Los números se actualizan al añadir o quitar filas/columnas.
### Color de celda

- **Tablas → Color de celda** abre el selector de color y aplica el fondo a la celda o al rango seleccionado.
- **Sin color** quita el relleno.
- El color se conserva al guardar en DOCX.
### Menú contextual de tabla

Al hacer clic derecho sobre una celda o un rango seleccionado aparece un menú con:

- Fusionar celdas.
- Dividir celda.
- Color de celda.
- Sin color.
- Añadir/eliminar fila.
- Añadir/eliminar columna.