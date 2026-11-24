# Auditoría de remediación — 0.0.9 Beta

Fecha: 2026-11-23  
Origen: auditoría externa sobre 0.0.8 y revisión interna posterior.

## Alcance

Se revisaron los hallazgos de la auditoría externa y se aplicaron parches en:

- profundidad XML de DOCX y ODT;
- autoguardado;
- recuperación automática;
- lector OLE2/CFB;
- relaciones externas y macros de glosario;
- escritura atómica;
- registro de fallos;
- límites de imágenes;
- mensajes de ZIP;
- memoria del lector DOC;
- limpieza de código.

## Correcciones aplicadas

| Hallazgo | Corrección | Prueba |
|---|---|---|
| Pila por XML anidado | Límite de 128 niveles en DOCX y ODT | `SecuritySmoke` crea DOCX/ODT con 200 niveles y exige rechazo |
| Autoguardado destructivo | El autoguardado escribe siempre en recuperación; nunca sobrescribe el original | Revisión de `AutoSaveNow` |
| OLE2 sin límites | Sectores DIFAT/FAT visitados, contadores, tamaño de flujo y raíz validados | Corpus `.doc` y lectura de CFB |
| Recuperación borra copias | No se borran tras recuperar; se cuarentena al rechazar; mutex por sesión | Revisión de código |
| Relaciones externas | Se descartan las externas no-hipervínculo | `SecuritySmoke` con relación `file:///` |
| VBA en glosario | Se detecta y no se conserva `/vbaProject.bin` | Revisión de `DetectSecurityRisks` |
| Fallback de `File.Replace` | Se aparta el original y se restaura si el intercambio falla | Revisión de `AtomicFile` |
| Logs sin rotación | Rotación a 2 MB y validación de ruta | Revisión de `CrashLog` |
| Imágenes enormes | Se rechazan dimensiones no válidas y se limita el render a 25 MP | `SecuritySmoke` con PNG negativo |
| Mensajes ZIP | 64/128 MB coherentes con el código | Revisión |
| `DocChar` por carácter | Pasa de clase a `struct` | Corpus `.doc` |
| Ternario muerto | Eliminado | Compilación |

## Verificación

```text
CORPUS=101
OK=101
FAIL=0
UI_SMOKE=PASS
CLIPBOARD_SMOKE=PASS
FEATURE_AUDIT=PASS
CARET_SMOKE=PASS
AUTOSAVE_AUDIT=PASS
SECURITY_SMOKE=PASS
REGRESSION=PASS
SELFTEST=PASS
```

Compilación sin errores ni advertencias con `csc /warn:4`.

## Límites que siguen abiertos

| Tema | Estado |
|---|---|
| `--selftest` en binario de producción | Se mantiene como diagnóstico local; escribe en el directorio indicado por el usuario |
| Tablas largas en vista previa | No se dividen; el DOCX exporta `w:tblHeader` |
| Estilos `w:style` | Definiciones básicas; el formato directo sí se conserva |
| Comentarios resueltos | Estado local |
| Comparación de formato | Solo aviso |
| Firma de código y matriz Word | Pendiente |
| Cobertura de `NativeDocumentEditor` | Auditoría externa no lo revisó; la interna sí revisó las rutas modificadas |

## Riesgo residual

- El lector OLE2 sigue siendo código propio; se endureció, pero requiere fuzzing continuo.
- El autoguardado ya no destruye el original, pero el guardado manual conserva un solo `.bak`.
- Los parsers DOCX/ODT/RTF necesitan pruebas con documentos reales y entradas malformadas.
- No hay validación visual en Word 365 ni firma Authenticode.

## Veredicto

Los hallazgos principales de la auditoría externa quedaron corregidos o mitigados.

La aplicación sigue siendo **beta técnica**, no producción:

- no hay firma de código;
- no hay matriz completa Word 365/LibreOffice;
- no hay fuzzing continuo;
- la compatibilidad avanzada sigue siendo parcial.
