# Instalación portátil

Copie `MiniWord.exe` junto a `install.ps1` y ejecute:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

Instala en `%LOCALAPPDATA%\MiniWord`, crea un acceso directo en el menú Inicio
y escribe `VERSION.txt`.

El instalador es portátil y **no está firmado digitalmente**. No se debe
presentar como instalador firmado ni como actualizador automático.

Para desinstalar, elimine la carpeta `%LOCALAPPDATA%\MiniWord` y el acceso
directo `Mini Word.lnk`.
