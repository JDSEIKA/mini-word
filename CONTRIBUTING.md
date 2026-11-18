# Probar Mini Word 0.0.3 Beta

Gracias por dedicar tiempo a probar una versión beta.

## Descarga

- [MiniWord.exe](https://github.com/JDSEIKA/mini-word/releases/download/v0.0.3-beta/MiniWord.exe)
- [MiniWord-Beta.zip](https://github.com/JDSEIKA/mini-word/releases/download/v0.0.3-beta/MiniWord-Beta.zip)
- [MiniWord-Installer.zip](https://github.com/JDSEIKA/mini-word/releases/download/v0.0.3-beta/MiniWord-Installer.zip)

No hace falta instalar Office ni LibreOffice. Se requiere Windows 7 SP1 o posterior y .NET Framework 4.0 o posterior.

## Pruebas de 10 minutos

1. Abre un DOCX con texto, una imagen y una tabla.
2. Cambia tamaño y color en un fragmento de texto y guarda.
3. Crea una tabla, escribe en una celda y cambia el tamaño de la fuente.
4. Inserta una forma y un cuadro de texto; arrástralos a otra página.
5. Guarda, cierra y vuelve a abrir.
6. Abre el mismo documento en Word o LibreOffice y comprueba posición de formas, cuadros y tabla.
7. Revisa ortografía y confirma que la palabra revisada queda seleccionada.

## Cómo reportar un error

Usa [el formulario de error](https://github.com/JDSEIKA/mini-word/issues/new?template=bug_report.yml) e incluye:

- Versión de Mini Word y de Windows.
- Pasos exactos.
- Resultado esperado y real.
- Un DOCX mínimo sin datos privados.
- Captura si el problema es visual.
- Ruta del `CrashLog` si apareció una excepción.

**No adjuntes documentos confidenciales.** Un ejemplo mínimo vale más que un documento completo.

## Qué información es más útil

- Formato de origen: DOCX, DOC, ODT, RTF.
- Si el fallo aparece al abrir, editar, guardar, reabrir o imprimir.
- Si ocurre con una tabla, forma, cuadro de texto, encabezado o columna.
- Si Word y LibreOffice muestran el mismo resultado.

## Código

El proyecto se compila con:

```bat
build.cmd
```

Pruebas:

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\run-regression.ps1
```

No se aceptan documentos de terceros que no puedan compartirse legalmente.
