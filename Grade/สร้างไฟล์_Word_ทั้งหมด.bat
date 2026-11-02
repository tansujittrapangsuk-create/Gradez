@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$f = Get-ChildItem -Filter '*.ps1' | Where-Object { $_.Name -like '*Word*' } | Select-Object -First 1; if ($f) { $b = [System.IO.File]::ReadAllBytes($f.FullName); if ($b.Length -lt 3 -or $b[0] -ne 0xEF -or $b[1] -ne 0xBB -or $b[2] -ne 0xBF) { [System.IO.File]::WriteAllBytes($f.FullName, [byte[]]@(0xEF,0xBB,0xBF) + $b) }; & $f.FullName } else { Write-Error 'Script not found' }"
pause
