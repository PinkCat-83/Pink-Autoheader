@echo off
:: Limpiamos la variable mensaje por seguridad
set "mensaje="

set /p mensaje="Introduce el mensaje del commit (pulsa Enter para usar fecha/hora): "

:: Si el usuario no escribe nada, usamos la fecha y hora
if not defined mensaje (
    set mensaje=Actualizacion automatica: %date% %time%
)

:: 'git add -A' asegura que se incluyan los archivos BORRADOS, nuevos y modificados
git add -A

:: Realizar el commit con el mensaje (o la fecha si estaba vacio)
git commit -m "%mensaje%"

:: Empujar los cambios
git push

echo.
echo ---------------------------------------
echo Tarea completada con exito.
echo Mensaje usado: "%mensaje%"
echo ---------------------------------------
pause