@echo off
chcp 65001 >nul
echo ====================================
echo Сборка EmailSender в EXE
echo ====================================
echo.

REM Очистка предыдущей сборки
echo [1/3] Очистка предыдущих файлов сборки...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "*.spec" del /q "*.spec"
echo.

REM Копирование spec файла
echo [2/3] Копирование spec файла...
copy /Y "build_config\email_sender.spec" "email_sender.spec" >nul
echo.

REM Сборка
echo [3/3] Запуск PyInstaller...
echo Это может занять несколько минут...
echo.

python -m PyInstaller --clean email_sender.spec

echo.
echo ====================================
if exist "dist\EmailSender.exe" (
    echo УСПЕШНО! EXE файл создан:
    echo %CD%\dist\EmailSender.exe
) else (
    echo ОШИБКА! EXE файл не создан.
    echo Проверьте логи выше.
)
echo ====================================
echo.
pause
