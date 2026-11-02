@echo off
cd /d "%~dp0"
powershell -NoProfile -Command "$html = Get-ChildItem -Filter '*.html' | Select-Object -First 1; if ($html) { Start-Process $html.FullName }"
