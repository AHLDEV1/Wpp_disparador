@echo off

REM Vai para o diretório do script, mesmo se executado de outro local
cd /d "%~dp0"

echo ==============================================
echo Bem-vindo(a) ao setup do disparador de mensagens!
echo Este processo pode levar alguns minutos.
echo Por favor, nao feche esta janela ate a conclusao.
echo ==============================================
echo.

REM Cria ambiente virtual (se já existir, pode substituir)
py -m venv venv

REM Ativa o ambiente virtual
call venv\Scripts\activate.bat

REM Atualiza pip
py -m pip install --upgrade pip

REM Instala as dependências necessárias
py -m pip install colorama selenium pandas webdriver_manager pyinstaller openpyxl pygetwindow

REM Remove builds e dist antigas para evitar conflitos
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist whatsapp_disparador.spec del whatsapp_disparador.spec

REM Gera o executável "Disparador.exe" usando PyInstaller com ícone relativo
py -m PyInstaller --onefile --name Disparador --icon="%cd%\Py e drivers\icone.ico" whatsapp_disparador.py

echo.
echo ==============================================
echo Setup concluido com sucesso!
echo Executavel criado na pasta dist como Disparador.exe
echo IMPORTANTE: Copie seu arquivo contacts.xlsx para a pasta dist antes de executar.
echo Para rodar, abra o arquivo Disparador.exe dentro da pasta dist.
echo Pressione qualquer tecla para sair.
echo ==============================================

pause
