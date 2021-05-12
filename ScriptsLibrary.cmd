:: Этот скрипт открывает одноименный solution в Visual Studio, используя конфигурацию IPS SDK из
:: файла ConfigureIPSSDK.ps1 в папке с файлом solution.

:: Внимание!!! Не редактировать!!!
:: Содержимое этого файла создается автоматически и может быть изменено в любой момент.

@echo off

:: передаем управление основному powershell-скрипту
set __SCRIPTDIR=%~dp0
set __SCRIPTPATH=%~dpn0.ps1
powershell -Version 3 -ExecutionPolicy RemoteSigned -Command "& '%__SCRIPTPATH%'" %*
if errorlevel 1 pause
