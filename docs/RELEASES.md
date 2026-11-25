# Versiones y notas de publicación

## Versión actual

**Mini Word 0.0.21 Beta**

- Ejecutable: 467 KB.
- Release: https://github.com/JDSEIKA/mini-word/releases/tag/v0.0.21-beta
- Atajos: [`ATAJOS_TECLADO.md`](ATAJOS_TECLADO.md).
- Alcance: [`BETA_SCOPE.md`](BETA_SCOPE.md).
- Compatibilidad: [`COMPATIBILITY.md`](COMPATIBILITY.md).

## Notas históricas

Los archivos `release-notes-0.0.x.md` son el registro de cada publicación. No describen la versión actual y no deben actualizarse: conservan los cambios, las pruebas y los hashes de esa versión.

Las auditorías con versión en el nombre (`AUDITORIA_0.0.9.md`, `REAUDITORIA_0.0.9.md`, etc.) también son documentos históricos.

## Criterio de versión

- `src/AppInfo.cs` es la fuente visible de la versión.
- `src/Properties/AssemblyInfo.cs` fija la versión del ensamblado.
- `README.md` y `CONTRIBUTING.md` deben apuntar a la release actual.
- El instalador escribe `VERSION.txt` con la versión actual.
- Las notas nuevas se crean como `release-notes-X.Y.Z.md`.
