# Re-auditoría 0.0.9 — correcciones

Fecha: 2026-11-24  
Origen: re-auditoría externa sobre 0.0.9.

## Problemas nuevos corregidos

### A. Recuperación desactivada por defecto

`AutoSaveNow` solo escribía recuperación si `_autoSaveEnabled` era verdadero. Ahora:

- la recuperación se escribe siempre que el documento esté sucio;
- `_autoSaveEnabled` solo controla el intervalo y el texto de estado;
- el original nunca se sobrescribe.

`AUTOSAVE_AUDIT` prueba ambos casos: activado y desactivado.

### B. Copias de trabajo descartado

Se restauró el borrado de la copia de recuperación cuando el usuario descarta cambios en:

- nuevo documento;
- abrir otro documento;
- cerrar la aplicación.

La cuarentena `Skipped` se purga automáticamente después de 7 días.

### C. Mutex y segunda instancia

Se añadió un candado por archivo de recuperación:

- `recovery-*.docx.lock`;
- abierto con `FileShare.None` por la instancia propietaria;
- otra instancia no puede ofrecer ni cuarentenar una copia viva;
- el bloqueo se libera al guardar, borrar o cerrar.

El mutex global se mantiene para serializar el arranque, pero ya no es la única protección.

### D. Versión duplicada

Se añadió `AppInfo.Version` y `AppInfo.VersionTag`; `NativeEditorForm` usa la constante en título, Acerca de y ventana. `AssemblyInfo.cs` conserva los valores numéricos.

## Hallazgos de la re-auditoría

| # | Estado anterior | Estado actual |
|---|---|---|
| 1 Pila XML | Corregido | Corregido |
| 2 Autoguardado | Regresión A | Corregido |
| 3 OLE2 | Corregido | Corregido |
| 4 Recuperación | Parcial B/C | Corregido |
| 5 Relaciones y partes | Parcial | Corregido en lo principal: se sanean `.rels` preservados y se descartan destinos externos |
| 6 Clic de enlace | Abierto | Documentación corregida; el comportamiento es clic simple |
| 7 AtomicFile | Corregido | Corregido; puede quedar `.swap` huérfano si el proceso muere entre los dos movimientos |
| 8 Variables de entorno | Parcial | Rutas normalizadas; sigue siendo una variable local de configuración |
| 9 Imágenes | Corregido | Límite reducido a 25 MP; WMF/EMF/TIFF dependen de GDI+ |
| 10 Mensajes ZIP | Corregido | Corregido |
| 11 Calidad/selftest | Abierto | `--selftest` se mantiene como diagnóstico local |

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

Compilación con `csc /warn:4` sin errores.

## Riesgo residual

- No se ha ejecutado una matriz Word 2007–365.
- No hay firma Authenticode.
- Los parsers DOCX/ODT/RTF necesitan fuzzing continuo.
- `--selftest` sigue en el binario de producción.
- La vista previa no divide tablas largas.
- Los estilos `w:style` son básicos.

## Veredicto

Los problemas nuevos A, B, C y D quedaron corregidos.

La aplicación sigue siendo **beta técnica**, no producción.
