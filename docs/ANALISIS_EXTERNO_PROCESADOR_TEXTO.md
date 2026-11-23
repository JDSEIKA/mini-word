# Análisis del procesador de texto externo

Fecha: **2026-11-22**  
Origen analizado: `C:\Users\JD\Desktop\procesador-texto-main.zip`  
Proyecto: **Procesador de texto** de `jceronch1`  
Licencia: **MIT**  
Pruebas ejecutadas: **110 correctas, 0 fallos** (`node tests/run.js`)

## Resumen ejecutivo

Es una aplicación web completa, sin dependencias, escrita en JavaScript puro. Tiene un modelo de documento propio, paginación en DOM, lectura/escritura DOCX amplia y varias funciones que Mini Word todavía no tiene.

**Sí conviene reutilizar su lógica conceptual**, sobre todo:

- Campos de Word.
- Tabla de contenido (TOC).
- Control de cambios.
- Comparación de documentos.
- Revisor de estilo y accesibilidad.
- Conversión HTML/Markdown/RTF.
- Numeración de página, marca de agua y números de línea.

**No conviene trasplantar su código directamente.** Es JavaScript, depende del DOM y su capa ZIP/XML no tiene los límites de seguridad de Mini Word.

## Arquitectura del proyecto externo

| Parte | Implementación |
|---|---|
| Interfaz | HTML, CSS y JavaScript clásico |
| Editor | `contenteditable` con modelo propio |
| Modelo | Tramos, marcas, párrafos, tablas, estilos, listas e historias |
| Paginación | DOM: espaciadores y hojas dibujadas detrás |
| DOCX | ZIP propio + XML propio + WordprocessingML |
| Almacenamiento | IndexedDB y localStorage |
| Pruebas | Node, 110 casos |
| Dependencias | Ninguna |

Tamaño aproximado:

```text
22 archivos JS
12.322 líneas
786 KB de JavaScript
```

## Qué hace mejor que Mini Word

| Función | Procesador externo | Mini Word actual |
|---|---|---|
| Campos PAGE, NUMPAGES, DATE, TITLE | Sí | Parcial: `{PAGE}` y `{NUMPAGES}` en encabezados |
| TOC automático | Sí | No modelado |
| Control de cambios | Sí, aceptar/rechazar | No editable; se avisa/preserva |
| Comentarios con respuestas | Sí | Panel de comentarios básico |
| Comparar documentos | Sí, LCS + diferencias por palabra | No |
| Revisor de estilo en español | Sí, sin conexión | No |
| Comprobador de accesibilidad | Sí | No |
| HTML | Abrir/guardar | No |
| Markdown | Abrir/guardar | No |
| Pegar desde Word/Google Docs/web | Sí, conserva formato | Parcial |
| Listas multinivel/checklist | Sí | Básicas |
| Números de línea | Sí, en pantalla y DOCX | No |
| Marca de agua | Sí | No |
| Filas de encabezado repetidas al paginar | Sí | Limitado |
| Plantillas, biblioteca e historial | Sí | No |
| Variables y combinación de correspondencia | Sí | No |
| Paleta de comandos | Sí | No |
| Asistente de IA/voz | Sí, opcional | No |

## Qué hace mejor Mini Word

| Función | Mini Word | Procesador externo |
|---|---|---|
| Aplicación nativa Windows | Sí | Navegador |
| Secciones múltiples | Sí | No |
| Columnas periodísticas | Sí | No |
| Formas y cuadros de texto | Sí, VML | No |
| DOC 97-2003 | Importación | No |
| ODT | Sí | No |
| Seguridad de ZIP/XML/RTF | Endurecida | Básica |
| Macros/ActiveX | No se ejecutan | No aplica |
| Impresión nativa | GDI+ | Navegador |
| Offline sin navegador | Sí | Sí, pero en navegador |

## Funciones reutilizables

### 1. Campos y TOC

**Valor:** muy alto. Es una de las mayores brechas de Mini Word.

El proyecto externo ya implementa:

- `w:fldSimple`.
- `w:fldChar` begin/separate/end.
- `w:instrText`.
- Campos PAGE, NUMPAGES, DATE, TITLE, DOCVARIABLE.
- TOC con `w:sdt`, `w:hyperlink`, bookmarks y campo `TOC`.
- Sustitución del índice por un bloque propio al importar.

**Portabilidad:** alta como especificación, media como código.

Habría que añadir a Mini Word:

- Un átomo `DocumentField` con instrucción, resultado y bookmarks.
- Renderizado de campos en layout.
- Actualización de PAGE/NUMPAGES al paginar.
- Exportación de campos complejos.
- Un bloque TOC que recorra títulos y genere entradas.

No conviene copiar el TOC tal cual porque el externo calcula las páginas en su layout DOM. Mini Word debe usar su propio motor.

### 2. Control de cambios

**Valor:** alto.

El modelo externo guarda en cada tramo:

```text
ins  → autor|fecha de inserción
del  → autor|fecha de borrado
```

Y exporta:

- `w:ins`
- `w:del`
- `w:delText`

Los límites conocidos del externo son:

- No registra unir/dividir párrafos como cambio.
- Las respuestas a comentarios se guardan como párrafos del comentario.

**Portabilidad:** media-alta.

Mini Word necesitaría:

- Metadatos de revisión por tramo.
- Aceptar/rechazar uno o todos.
- Vista de cambios.
- Round-trip DOCX.

### 3. Comparar documentos

**Valor:** alto y coste relativamente bajo.

`js/model/compare.js` tiene solo unas 160 líneas:

- LCS por bloques.
- LCS por palabras.
- Alineación aproximada para documentos grandes.
- Generación de cambios controlados.

**Portabilidad:** alta. Se puede reescribir en C# sin depender del DOM.

### 4. Revisor de estilo y accesibilidad

**Valor:** alto como diferenciador.

`js/ui/review.js` incluye reglas para:

- Oraciones largas.
- Voz pasiva.
- Adverbios en `-mente`.
- Repeticiones cercanas.
- Muletillas.
- Dequeísmo/queísmo.
- Signos de apertura.
- Legibilidad Fernández Huerta.
- Imágenes sin texto alternativo.
- Saltos de nivel de títulos.
- Tablas sin encabezado.
- Contraste.
- Enlaces “aquí”.
- Texto pequeño o en mayúsculas.

**Portabilidad:** media. Son reglas y expresiones regulares; se pueden portar a C# y adaptar al modelo nativo.

### 5. HTML, Markdown y portapapeles

**Valor:** alto para interoperabilidad.

`js/io/convert.js` cubre:

- HTML.
- Markdown.
- RTF.
- Texto.
- Pegado desde Word/Google Docs/web.

**Portabilidad:** media.

Mini Word ya tiene RTF; faltaría HTML/Markdown y un pegado enriquecido más completo.

### 6. Numeración, agua y líneas

**Valor:** medio-alto.

El externo exporta:

- `w:pgNumType` para el inicio de numeración.
- `w:lnNumType` para números de línea.
- Encabezados/pies.
- Marca de agua.
- `w:tblHeader` para repetir encabezados de tabla.

**Portabilidad:** media. La serialización DOCX es directamente aprovechable como referencia; el dibujo debe adaptarse al layout nativo.

### 7. Plantillas, bloques, variables y correspondencia

**Valor:** medio.

Son funciones de productividad:

- Plantillas.
- Historial de versiones.
- Bloques reutilizables.
- Variables de documento.
- Combinación de correspondencia desde CSV.

**Portabilidad:** media. Es sobre todo modelo y UI, no algoritmos complejos.

### 8. Búsqueda y formato automático

**Valor:** medio.

El externo tiene:

- Buscar con mayúsculas.
- Palabra completa.
- Ignorar tildes.
- Expresiones regulares.
- Reemplazar todo.
- Markdown automático.
- Comillas tipográficas.
- Correcciones `(c)`, `--`, `...`.

**Portabilidad:** alta. Son funciones pequeñas y bien acotadas.

## Qué no conviene reutilizar

### El layout DOM

`js/view/layout.js` pagina insertando espaciadores en el DOM. Mini Word tiene un motor GDI+ con páginas, líneas y objetos propios. No es trasplantable.

### El editor `contenteditable`

El externo conserva el corrector, IME y accesibilidad del navegador. Mini Word usa un control nativo. Copiar esa capa sería un cambio de arquitectura, no una integración.

### ZIP y XML propios

`js/io/zip.js` y `js/io/xml.js` **no tienen los límites de seguridad de Mini Word**:

- Sin límite de tamaño descomprimido.
- Sin control de bombas ZIP.
- Sin validación CRC estricta.
- Sin límite de profundidad XML.
- Sin política DTD/entidades equiparable.

**No se deben sustituir** `ZipArchiveLite` ni el lector XML endurecido de Mini Word.

### Almacenamiento e IA

- IndexedDB/localStorage no encajan con la aplicación nativa.
- El asistente de IA envía texto a servicios externos; debe ser opcional, explícito y con consentimiento.
- Voz/dictado depende del navegador.

## Licencia

El proyecto usa **MIT**:

```text
Copyright (c) 2026 jceronch1
```

MIT permite usar, copiar, modificar y distribuir, siempre que se conserve el aviso de copyright y la licencia.

Si se porta código sustancial:

1. Añadir `THIRD_PARTY_NOTICES.md`.
2. Conservar el aviso MIT.
3. Indicar qué partes se adaptaron.
4. No atribuir a Mini Word código no portado.

Si solo se reutilizan ideas, algoritmos o especificaciones, también es recomendable dejar constancia en la documentación.

## Recomendación

No integrar el proyecto completo ni incrustarlo en WebView2.

**Sí usar su lógica como referencia y portar selectivamente.**

### Prioridad 1

1. Campos PAGE/NUMPAGES y numeración de página.
2. TOC automático.
3. Control de cambios.
4. HTML/Markdown.
5. Comparar documentos.

### Prioridad 2

6. Revisor de estilo.
7. Comprobador de accesibilidad.
8. Búsqueda con expresiones regulares.
9. Números de línea.
10. Filas de encabezado repetidas.

### Prioridad 3

11. Plantillas e historial.
12. Bloques reutilizables.
13. Variables y combinación de correspondencia.
14. Paleta de comandos.

## Conclusión

El proyecto externo es una **buena fuente de especificación y algoritmos** para las funciones que a Mini Word le faltan. Lo más valioso es su tratamiento de campos, TOC, control de cambios, comparación y revisión.

No es un reemplazo directo de Mini Word: no tiene secciones, columnas, formas, DOC/ODT y su seguridad de paquetes es básica. La integración correcta es **portar funciones al modelo C# nativo**, manteniendo el motor de páginas y la seguridad actuales.
