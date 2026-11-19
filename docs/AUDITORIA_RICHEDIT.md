# Auditoría: código RichEdit / “Rift Edit”

Fecha: **2026-11-18**  
Versión: **0.0.4 Beta**

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

Ninguna de esas referencias crea un control RichEdit en la aplicación.

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

El código RichEdit **no se usaba en producción**. Era código heredado de una etapa anterior y una prueba engañosa. Se eliminó de forma estricta.
