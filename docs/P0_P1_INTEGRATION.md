# Integración P0/P1 — estado 2026-11-16

## Completado y verificado

### Formas exportables

- VML ahora escribe `coordsize` y `path` para rectángulo, elipse, triángulo, estrella, flecha, flecha izquierda, rombo y corazón.
- Una forma sin texto ya no genera un `v:textbox` vacío que Word/LibreOffice dibujaba como rectángulo extra.
- El lector acepta formas VML sin `v:textbox`.
- Render externo de un triángulo en LibreOffice: **correcto**, con relleno, borde y posición.

### Encabezados de primera página y páginas pares

- Modelo: `FirstPageHeaderText`, `FirstPageFooterText`, `EvenPageHeaderText`, `EvenPageFooterText`, `TitlePage`, `DifferentOddEven`.
- DOCX: referencias `first` y `even`, `<w:titlePg/>` y `<w:evenAndOddHeaders/>` en `settings.xml`.
- Lector: recupera primero, par, normal y banderas.
- Render de Mini Word: elige primera página, páginas pares o normal.
- Diálogo: pestañas **Principal**, **Primera página** y **Páginas pares**, con casillas para activarlas.
- Render externo LibreOffice: página 1 usa encabezado/pie de primera; página 2 usa encabezado/pie par.

#### Límite conocido

- Formatos, alineación e imágenes son independientes por variante y se conservan en DOCX y render externo.

### Partes desconocidas del paquete

- `DocumentModel.PreservedParts` conserva `word/theme/*`, `word/glossary/*`, `word/charts/*`, `word/diagrams/*`, `word/activeX/*`, `customXml/*` y `word/fontTable.xml`.
- Se copian al guardar si no colisionan con una parte generada.
- Prueba: 5/5 partes conservadas en `python-docx-test.docx`; el DOCX resultante abre en LibreOffice.

#### Alcance
- Se reconstruyen las relaciones de `word/_rels/document.xml.rels` que apuntan a esas partes y sus `Override`/`Default` de `[Content_Types].xml`.
- Las relaciones internas de las propias partes preservadas no se fusionan.

### Pruebas de regresión

- `tests/run-regression.ps1`: build, `--selftest`, compilación del smoke test y corpus de 101 documentos.
- Resultado local: `REGRESSION=PASS`.
- El corpus cubre DOCX, ODT de LibreOffice, DOC y documentos reales.

### Campos, TOC y controles de contenido

- Los párrafos sin editar que contienen `fldSimple`, `fldChar`, `instrText` o `sdt` conservan su XML original.
- Al guardar se reescribe el XML original mientras el texto del párrafo no haya cambiado.
- Si el usuario edita el texto **o cambia el formato** del párrafo, se guarda el texto modelado y se pierde el código del campo; así no se conserva un campo con formato obsoleto.
- Prueba: `poi-field-codes.docx` conserva 12 marcas `fldChar` sin cambios; al modificar el formato de un párrafo con campo quedan 6, correspondientes al otro párrafo no tocado.
- El DOCX sin cambios abre en LibreOffice mostrando el resultado.

### Accesibilidad básica

- El editor expone nombre, descripción y rol `Document`.
- La ventana expone nombre y descripción.

## Pendiente

### P0

- Campos, TOC y controles de contenido: **parcial**. Se conserva el XML de párrafos no editados; al editar el texto se pierde el código.
- Relaciones y tipos de contenido de partes desconocidas: **cubiertos para el paquete raíz**. Las relaciones internas de cada parte no se fusionan.
- Suite de interfaz/CI: **parcial**; existe `tests/run-regression.ps1`, pero no cubre toda la interfaz.
- Corpus de 100+: **cumplido**; 101 documentos reales.

### P1

- Fidelidad ODT/DOC: **parcial**; lectura y texto, no geometría/bordes/estilos avanzados.
- Fusiones anidadas: **pendiente**.
- Instalador portátil y versionado: **completado sin firma**; actualizaciones automáticas y accesibilidad completa siguen pendientes.
- Edición visual independiente de formatos, alineación e imágenes por variante: **completada** (añadir/eliminar imágenes desde cada pestaña).

## Pruebas ejecutadas

```text
--selftest                 47/47
corpus DOCX/ODT/DOC        101/101
diálogo encabezados        PASS
roundtrip first/even       PASS
render forma LibreOffice   PASS
render headers LibreOffice PASS\nrender final DOCX         triángulo, tabla, viñetas y encabezado par correctos\nfirst/even format+image  FIRST=Arial 18pt Center IMG=1; EVEN=Arial 15pt Right\ndiálogo por variante    FIRST=36 Center; EVEN=30 Right; imagen por variante=1
partes preservadas         5/5
campos conservados         poi-field-codes: 12 marcas intactas; 6 tras formatear un párrafo\nUI smoke                   PASS\ncorpus ampliado            101/101
```

## Veredicto

La beta privada mejora de forma importante: las formas ya sobreviven a Word/LibreOffice y los encabezados especiales se conservan y se pueden editar.

La beta pública sigue bloqueada por campos/TOC, relaciones de partes desconocidas, fusión anidada, corpus/CI y distribución. No se debe declarar compatibilidad completa con Word todavía.

### Instalador y versión

- Versión de ensamblado: 0.0.1 Beta.
- installer/install.ps1: instalación portátil en %LOCALAPPDATA%\\MiniWord, acceso directo y VERSION.txt.
- **No está firmado digitalmente**; no se declara instalador firmado.
- No hay actualizador automático todavía.
