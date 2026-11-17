# Regression tests

Run from the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\run-regression.ps1
```

The script:

1. Builds `dist\MiniWord.exe`.
2. Runs `--selftest` with the local DOC fixtures.
3. Compiles `CorpusSmoke.cs` with the application sources.
4. Opens, saves and reopens the available corpus (97 fixtures in this repository).

Expected final line:

```text
REGRESSION=PASS
```

The corpus includes 11 DOCX, 84 LibreOffice ODT, 3 DOC and 3 project documents. External visual checks (LibreOffice Kit) remain separate because they need a renderer.
