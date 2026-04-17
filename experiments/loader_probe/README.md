# Loader Probe

This directory isolates small experiments for the document parsing stack used by the
main app.

It generates sample `.docx` and `.xlsx` files, parses them with the lower-level
libraries used under the business code, and writes JSON snapshots under `outputs/`.

Run with:

```powershell
.\.venv\Scripts\python.exe .\experiments\loader_probe\run_probe.py
```
