@echo off
echo Obteniendo cambios del repositorio remoto...
echo.

:: Traer los cambios sin fusionar
git fetch

echo.
echo Cambios obtenidos. Mostrando estado...
echo.
git status

echo.
set /p fusion="¿Deseas fusionar los cambios? (s/n): "

if /i "%fusion%"=="s" (
    git merge
    echo.
    echo Fusion completada.
) else (
    echo.
    echo Fusion cancelada. Los cambios estan descargados pero no fusionados.
)

echo.
echo ---------------------------------------
echo Tarea completada con exito.
echo ---------------------------------------
pause