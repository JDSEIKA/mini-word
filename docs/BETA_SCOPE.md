# Mini Word — alcance de beta

Versión de evaluación: 0.0.8 Beta
Fecha: 2026-11-15

> **Actualización 0.0.8:** campos simples, índice, control de cambios, comparación, estilos y comentarios con respuestas están implementados. Los campos/TOC complejos de Word siguen siendo parciales.

## Qué incluye esta beta

- Edición nativa de documentos DOCX simples y medios.
- Apertura de `.doc` 97-2003 en modo importación; se guarda como `.docx`.
- Apertura y guardado de ODT, RTF y TXT.
- Texto, formatos de carácter y párrafo, listas, tablas, imágenes, formas, cuadros de texto, secciones, columnas, notas, comentarios, hipervínculos y objetos OLE con vista previa.
- Guardado atómico con respaldo `.bak`.
- Recuperación automática al iniciar para sesiones interrumpidas, incluso con el autoguardado programado desactivado.
- Aviso explícito cuando el documento contiene campos, controles de contenido, control de cambios, gráficos, ecuaciones o partes que la beta no conserva por completo.
- Límites de seguridad para ZIP, XML, imágenes y documentos.
- Regla fija activa por defecto.
- Cinta con pestaña **Tablas** independiente.

## Qué no promete esta beta

- No es un sustituto completo de Microsoft Word.
- No ejecuta macros, controles ActiveX, OLE ni contenido remoto.
- Conserva y modela campos simples, índice estático, control de cambios por tramo y revisiones; no garantiza campos/TOC complejos, temas, gráficos, SmartArt ni ecuaciones.
- No guarda `.doc`; solo lo importa.
- No incluye instalador firmado ni actualizaciones automáticas.
- No sustituye una revisión visual en Word/LibreOffice de documentos críticos.

## Seguridad y recuperación

- Antes de sobrescribir un archivo se crea un temporal y una copia `.bak`.
- Las copias de recuperación están en `%TEMP%\MiniWord\Recovery`.
- Los registros de fallos están en `%TEMP%\MiniWord\Logs`.
- El guardado no se considera una validación visual: si el documento usa funciones avanzadas, guarde una copia y compruébela en Word.

## Pruebas de esta beta

- Autoprueba: **47/47**.
- Apertura desde la interfaz: **23/23** documentos (DOCX, ODT, DOC); la ventana abre maximizada.
- Corpus de 23 documentos (DOCX de python-docx/Apache POI, ODT de LibreOffice, DOC y documentos reales): **23/23** abiertos, guardados y reabiertos.
- Encabezados/pies independientes por sección, con imágenes: probado.
- Fuzz de 60 entradas inválidas: ninguna ejecutó ni colapsó el lector.
- DTD externa y ruta ZIP `../`: bloqueadas.
- Guardado atómico y recuperación: probados.
- 60 entradas aleatorias, 11 documentos reales y pruebas de tablas.

## Criterios que aún deben cerrarse para beta pública

1. Preservación real de campos, TOC y controles de contenido.
2. Encabezados/pies de primera página y página par.
3. Corpus de 100+ documentos y CI continuo.
4. Crash reporting conectado a un canal de soporte (el log local ya existe).
5. Instalador firmado y estrategia de actualización.
6. Accesibilidad básica y navegación completa por teclado.
7. Cierre de tablas con fusiones anidadas.

## Recomendación

Usar como **beta pública inicial** para DOCX cotidianos. Mantener el original y el `.bak`, y validar en Word los documentos con campos, índices, gráficos o revisiones antes de usarlos en producción.
