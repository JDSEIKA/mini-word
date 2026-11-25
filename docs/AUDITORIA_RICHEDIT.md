# Auditoría: código RichEdit / “Rift Edit”

Fecha: **2026-11-18**  
Versión: **0.0.18 Beta**

## Búsqueda realizada

Se buscó en todo el proyecto, sin distinguir mayúsculas:

```text
rift
riftedit
RichEdit
RichTextBox
EM_FORMATRANGE
PARAFORMAT
NativeMethods
```

## Resultado

No existe ningún identificador `RiftEdit` ni `Rift Edit` en el código.

El hallazgo real es **RichEdit**, el control de texto enriquecido de Windows.

## Qué había

Archivo:

```text
src/Util/NativeMethods.cs
```

Contenía:

- Constantes `EM_*` de RichEdit.
- `PARAFORMAT2`.
- `FORMATRANGE`.
- Envío de `EM_FORMATRANGE`.
- Funciones de paginación RichEdit.
- P/Invoke a `user32.dll`.

## ¿Se usaba en producción?

**No.**

La aplicación utiliza un editor propio:

```text
src/Controls/NativeDocumentEditor.cs
```

Ese editor dibuja páginas con GDI+, tiene su propio layout y no crea ningún `RichTextBox`.

### Única referencia

`NativeMethods` solo se usaba en:

```text
src/Formats/SelfTest.cs
```

En una prueba que enviaba `EM_FORMATRANGE` a un `RichTextBox` de prueba y lo llamaba “Page pagination”.

Esa prueba no validaba el editor real de Mini Word. Era una prueba del control RichEdit, no del motor de páginas de la aplicación.

## Corrección aplicada

1. Se eliminó `src/Util/NativeMethods.cs`.
2. Se sustituyó la prueba falsa de paginación por una prueba real:

```csharp
NativeDocumentEditor pageEditor = new NativeDocumentEditor();
pageEditor.Document = longDoc;
measured = pageEditor.PageCount > 1;
```

3. El test ahora comprueba el mismo motor que usa la aplicación.

## Qué queda de RichEdit

Solo quedan referencias documentales o de prueba:

| Ubicación | Uso |
|---|---|
| `SelfTest.cs` | `RichTextBox` como fixture para validar RTF |
| `RtfReader.cs` | Comentario sobre el comportamiento de RichEdit |
| `RtfSanitizer.cs` | Comentario sobre RTF aceptado por RichEdit |
| `RtfWriter.cs` | Comentario sobre el dialecto RTF |
| `NativeDocumentEditor.cs` | `RichTextBox` superpuesto para editar celdas de tabla desde 0.0.12 |

El motor principal del documento no crea un control RichEdit. El editor de celda de tabla sí usa un `RichTextBox` superpuesto para mostrar el formato mientras se edita; no interviene en la paginación ni en el motor tipográfico.

## Verificación

```text
RTF load text                    PASS
RTF paragraph count              PASS
RTF bullet                       PASS
RTF center                       PASS
RTF bold                         PASS
RTF superscript                  PASS
Page pagination                  PASS
RTF table round trip             PASS
RichEdit table round trip        PASS
RESULT: PASS
```

## Conclusión

El código RichEdit heredado **no se usaba en el motor principal**. Era una prueba engañosa y se eliminó de forma estricta. La edición de celdas de tabla usa un `RichTextBox` superpuesto de forma deliberada desde 0.0.12; es la única excepción y no afecta a la vista principal ni a la impresión.
