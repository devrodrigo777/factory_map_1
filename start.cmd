@echo off
setlocal
chcp 65001 >nul
title EasyLOTOTO for Knauf v2 - Servidor Local
cd /d "%~dp0"

echo ===================================================================
echo                 EasyLOTOTO for Knauf v2
echo        Gestao Visual de Bloqueios e Etiquetagens LOTOTO
echo ===================================================================
echo.

where node >nul 2>nul
if errorlevel 1 goto no_node

if not exist "client\dist\index.html" goto do_build

:start_server
set PORT=3001

echo [*] Iniciando o servidor local...
echo.
echo ===================================================================
echo   [OK] O sistema esta pronto para uso!
echo.
echo   -^> Link de acesso:  http://localhost:%PORT%
echo.
echo   (Seu navegador padrao sera aberto automaticamente em instantes)
echo.
echo   Dica: Mantenha esta janela aberta enquanto estiver usando o sistema.
echo         Para desligar, tecle Ctrl + C ou feche esta janela.
echo ===================================================================
echo.

start "" cmd /c "ping -n 3 127.0.0.1 >nul && start http://localhost:%PORT%"

node server/src/index.js

echo.
echo ===================================================================
echo   [Servidor Encerrado]
echo   Pressione qualquer tecla para fechar esta janela...
echo ===================================================================
pause >nul
goto end

:no_node
echo.
echo ===================================================================
echo   [ERRO] O Node.js nao foi encontrado neste computador!
echo   Para executar o EasyLOTOTO for Knauf, instale o Node.js v18 ou superior.
echo   Download oficial: https://nodejs.org/
echo ===================================================================
echo.
echo   Pressione qualquer tecla para sair...
pause >nul
goto end

:do_build
echo [*] Compilando a interface do sistema pela primeira vez...
call npm run build
echo.
goto start_server

:end
