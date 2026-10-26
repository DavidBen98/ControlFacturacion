@echo off
setlocal

set "PROGRAMA=C:\Users\Servidor\Downloads\ControlFacturacion\ControlFacturacion\bin\Debug\net8.0-windows\ControlFacturacion.exe"

tasklist /FI "IMAGENAME eq ControlFacturacion.exe" | find /I "ControlFacturacion.exe" >nul

if errorlevel 1 (
    if exist "%PROGRAMA%" (
        start "" "%PROGRAMA%" --auto13
    ) else (
        echo ERROR: No se encuentra el ejecutable:
        echo %PROGRAMA%
        pause
    )
) else (
    echo El programa ya esta en ejecucion.
)

endlocal