# ControlFacturacion con Excel

Requiere Windows y .NET 8 SDK. La ruta de Excel configurada es:
`C:\Users\Lalá\Downloads\Siesta.xlsx`

## Preparación
1. Colocá `spa.traineddata` en `tessdata\spa.traineddata` dentro de esta carpeta.
2. Cerrá Siesta.xlsx en Excel antes de guardar.
3. Abrí PowerShell en esta carpeta y ejecutá:
   `dotnet restore`
   `dotnet run`
4. Seleccioná una vez el área del número azul FACTURADO. Mantené abierta la aplicación y dejá servidor.exe visible en las horas programadas.

## Prueba manual
- Elegí `Leer importe`.
- Si el número detectado es correcto, podés pulsar `Guardar como 13:00` para escribir en la columna C o `Guardar como 16:00` para la columna D de la fila de la fecha de hoy.
- Probá primero con una copia de Siesta.xlsx para confirmar que las celdas son las correctas.

## Programación
La aplicación intenta leer y guardar a las 13:00 y 16:00, hora local de Windows, mientras está abierta. Debe estar visible en pantalla el dato que se quiere capturar y la región seleccionada debe seguir apuntando al importe. No puede recuperar una captura si la computadora está apagada, suspendida o la aplicación está cerrada.


## Modo automático (sin dejar abierta la ventana)

Esta versión guarda el rectángulo de lectura al seleccionar el área. Primero abrí ControlFacturacion manualmente, seleccioná el área FACTURADO y cerrá la app. El área queda guardada en `%LOCALAPPDATA%\\ControlFacturacion\\region.txt`.

Después de compilar/publicar la aplicación en Windows, creá dos tareas en el Programador de tareas. En cada tarea elegí **Ejecutar tanto si el usuario inició sesión como si no** NO es recomendable: la captura de pantalla necesita una sesión interactiva iniciada. Elegí **Ejecutar solo cuando el usuario haya iniciado sesión**.

- Tarea de las 13:00: iniciar el ejecutable `ControlFacturacion.exe`; en **Agregar argumentos**, poner `--auto13`.
- Tarea de las 16:00: iniciar el mismo ejecutable; en **Agregar argumentos**, poner `--auto16`.
- Configurá ambas para ejecutarse diariamente y activá “Despertar el equipo para ejecutar esta tarea” si aparece esa opción. La PC debe estar encendida y la sesión de Windows iniciada; no bloquees la pantalla mientras se captura.

El registro de cada ejecución queda en `%LOCALAPPDATA%\\ControlFacturacion\\automatizacion.log`. También se guarda `ultima_captura.png` junto al ejecutable para revisar la imagen leída. Si el OCR tiene poca confianza, no se modifica Excel.

**Importante:** la región se guarda en coordenadas de pantalla. No cambies la resolución ni muevas el contenido de `servidor.exe` a otra zona de la pantalla. Antes de depender del proceso, probá ambos argumentos manualmente y comprobá que `Siesta.xlsx` se actualiza en la columna correcta.
