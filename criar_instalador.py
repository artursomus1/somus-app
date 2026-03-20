"""
Somus Capital - Criador de Instalador v5.3
Gera um ZIP instalador que sempre instala a versao mais recente do GitHub.
"""

import os
import zipfile
import subprocess

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Ler versao atual do version.py
def _get_version():
    vpath = os.path.join(BASE_DIR, "version.py")
    with open(vpath, encoding="utf-8") as f:
        for line in f:
            if line.startswith("VERSION"):
                return line.split("=")[1].strip().strip('"').strip("'")
    return "1.0.0"

def _get_desktop():
    try:
        result = subprocess.run(
            ['reg', 'query',
             r'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders',
             '/v', 'Desktop'],
            capture_output=True, text=True, timeout=5
        )
        if result.returncode == 0:
            for line in result.stdout.splitlines():
                if 'Desktop' in line and 'REG_' in line:
                    path = line.split('REG_EXPAND_SZ')[-1].strip().split('REG_SZ')[-1].strip()
                    path = os.path.expandvars(path)
                    if os.path.exists(path):
                        return path
    except Exception:
        pass
    for candidate in [
        os.path.join(os.environ.get("USERPROFILE", ""), "Desktop"),
        os.path.join(os.environ.get("USERPROFILE", ""), "Area de Trabalho"),
    ]:
        if os.path.exists(candidate):
            return candidate
    return os.path.join(os.environ.get("USERPROFILE", r"C:\Users\User"), "Desktop")

OUTPUT_ZIP = os.path.join(_get_desktop(), "SomusCapital_Instalador.zip")

EXCLUIR_DIRS  = {"__pycache__", ".git", ".vscode", ".idea", ".mypy_cache", ".pytest_cache", "node_modules", "historico"}
EXCLUIR_ARQUIVOS = {"criar_instalador.py", ".env", "email_preview.html", "test_preview_A38675.html"}
EXCLUIR_EXT   = {".pyc", ".pyo", ".log", ".tmp"}
EXCLUIR_PREFIXOS = {"Codigo_"}
PASTAS_SAIDA  = {
    os.path.join("Mesa Produtos", "Fluxo RF", "PDFs"),
    os.path.join("Mesa Produtos", "Organizador", "ORGANIZADO"),
    os.path.join("Mesa Produtos", "Consolidador", "PDFs"),
    os.path.join("Mesa Produtos", "Envio Saldos", "SAIDA SALDOS"),
    os.path.join("Corporate", "Simulador", "PDFs"),
}
PASTAS_VAZIAS = [
    os.path.join("Mesa Produtos", "Fluxo RF", "PDFs"),
    os.path.join("Mesa Produtos", "Fluxo RF", "ENTRADA"),
    os.path.join("Mesa Produtos", "Organizador", "ORGANIZADO"),
    os.path.join("Mesa Produtos", "Consolidador", "PDFs"),
    os.path.join("Mesa Produtos", "Envio Saldos", "SAIDA SALDOS"),
    os.path.join("Mesa Produtos", "Envio Saldos", "BASE"),
    os.path.join("Mesa Produtos", "Envio Aniversarios"),
    os.path.join("Mesa Produtos", "Envio Mesa"),
    os.path.join("Mesa Produtos", "Envio Ordens"),
    os.path.join("Mesa Produtos", "Dashboard"),
    os.path.join("Mesa Produtos", "Informativo"),
    os.path.join("Mesa Produtos", "Tarefas"),
    os.path.join("Mesa Produtos", "Info Agio", "Agio"),
    os.path.join("Mesa Produtos", "Ctrl Receita"),
    os.path.join("Corporate", "Dashboard"),
    os.path.join("Corporate", "Simulador", "PDFs"),
    "BASE",
    "assets",
    os.path.join("Seguros", "Renovacoes Anuais"),
]

def deve_incluir(rel_path):
    partes = rel_path.replace(os.sep, "/").split("/")
    nome = partes[-1]
    if nome in EXCLUIR_ARQUIVOS:
        return False
    for p in partes:
        if p in EXCLUIR_DIRS:
            return False
    _, ext = os.path.splitext(nome)
    if ext.lower() in EXCLUIR_EXT:
        return False
    for pref in EXCLUIR_PREFIXOS:
        if nome.startswith(pref):
            return False
    for pasta_saida in PASTAS_SAIDA:
        if rel_path.replace(os.sep, "/").startswith(pasta_saida.replace(os.sep, "/") + "/"):
            return False
    return True


# ===========================================================================
# INSTALAR.bat — versao injetada dinamicamente pelo criar_instalador.py
# ===========================================================================
INSTALAR_BAT_TEMPLATE = r"""@echo off
:: Auto-relaunch em modo /k para janela nao fechar em erros
if "%~1"=="--keep" goto :main
cmd /k ""%~f0" --keep"
exit /b
:main

chcp 65001 >nul 2>&1
color 0A

set "BUNDLE_VERSION=__VERSION__"
set "GITHUB_REPO=artursomus1/somus-app"
set "GITHUB_API=https://api.github.com/repos/%GITHUB_REPO%/releases/latest"
set "UPDATE_ZIP=%TEMP%\somus_latest.zip"
set "UPDATE_DIR=%TEMP%\somus_latest"
set "INSTALL_SOURCE=%~dp0APP SOMUS"

echo.
echo  ============================================
echo     SOMUS CAPITAL - INSTALADOR DO APP
echo     Versao do pacote: v%BUNDLE_VERSION%
echo  ============================================
echo.

:: -----------------------------------------
:: 1. Verificar Python
:: -----------------------------------------
echo  [1/4] Verificando Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  [ERRO] Python nao encontrado!
    echo  Instale o Python 3.10+ em https://www.python.org/downloads/
    echo  IMPORTANTE: Marque "Add Python to PATH"
    echo.
    pause
    exit /b 1
)
for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set PYTHON_VERSION=%%v
echo         Python %PYTHON_VERSION% encontrado.
echo.

:: -----------------------------------------
:: 2. Verificar versao mais recente no GitHub
:: -----------------------------------------
echo  [2/4] Verificando atualizacoes no GitHub...

if exist "%UPDATE_ZIP%"  del "%UPDATE_ZIP%"  >nul 2>&1
if exist "%UPDATE_DIR%"  rd /s /q "%UPDATE_DIR%" >nul 2>&1
if exist "%TEMP%\somus_latest_version.txt" del "%TEMP%\somus_latest_version.txt" >nul 2>&1

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { ^
     $api = Invoke-RestMethod -Uri '%GITHUB_API%' -UseBasicParsing -TimeoutSec 12; ^
     $latest = $api.tag_name.TrimStart('v'); ^
     $local  = '%BUNDLE_VERSION%'; ^
     if ([version]$latest -gt [version]$local) { ^
       $url = ($api.assets | Where-Object { $_.name -like '*.zip' } | Select-Object -First 1).browser_download_url; ^
       Write-Host '         Nova versao encontrada: v'$latest' (este pacote: v'$local')'; ^
       Write-Host '         Baixando do GitHub...'; ^
       Invoke-WebRequest -Uri $url -OutFile '%UPDATE_ZIP%' -UseBasicParsing; ^
       Expand-Archive -Path '%UPDATE_ZIP%' -DestinationPath '%UPDATE_DIR%' -Force; ^
       Set-Content -Path '%TEMP%\somus_latest_version.txt' -Value $latest ^
     } else { ^
       Write-Host '         Pacote ja esta na versao mais recente (v'$local').' ^
     } ^
   } catch { ^
     Write-Host '         GitHub indisponivel. Instalando versao local (v%BUNDLE_VERSION%).' ^
   }" 2>&1

:: Usar arquivos do GitHub se foram baixados
if exist "%TEMP%\somus_latest_version.txt" (
    set /p LATEST_VERSION=<"%TEMP%\somus_latest_version.txt"
    del "%TEMP%\somus_latest_version.txt" >nul 2>&1
    set "INSTALL_SOURCE=%UPDATE_DIR%\SomusCapital_Instalador\APP SOMUS"
    echo         Instalando versao v%LATEST_VERSION% baixada do GitHub.
) else (
    echo         Instalando versao local v%BUNDLE_VERSION%.
)
echo.

:: Verificar source
if not exist "%INSTALL_SOURCE%\" (
    echo  [ERRO] Pasta de instalacao nao encontrada: %INSTALL_SOURCE%
    echo         Extraia o ZIP completo antes de executar este arquivo.
    echo.
    pause
    exit /b 1
)

:: -----------------------------------------
:: 3. Copiar arquivos
:: -----------------------------------------
set "DEST=%USERPROFILE%\Documents\APP SOMUS"
set "ENV_BACKUP=%TEMP%\somus_env_backup.tmp"

if exist "%DEST%\executar.py" (
    echo  [AVISO] Versao anterior encontrada. Sera substituida.
    echo.
    choice /C SN /M "  Deseja continuar? (S/N)"
    if errorlevel 2 goto :fim
    echo.

    if exist "%DEST%\.env" (
        copy /Y "%DEST%\.env" "%ENV_BACKUP%" >nul 2>&1
        echo         Configuracoes ^(.env^) salvas.
    )

    echo         Removendo versao anterior...
    rd /s /q "%DEST%" >nul 2>&1
    if exist "%DEST%" (
        echo  [ERRO] Nao foi possivel remover a pasta anterior.
        echo         Feche o aplicativo e tente novamente.
        pause
        exit /b 1
    )
)

echo  [3/4] Instalando arquivos em %DEST%...
mkdir "%DEST%"
xcopy /E /I /Y /Q "%INSTALL_SOURCE%\*" "%DEST%\"
if %errorlevel% neq 0 (
    echo  [ERRO] Falha ao copiar arquivos!
    pause
    exit /b 1
)

if exist "%ENV_BACKUP%" (
    copy /Y "%ENV_BACKUP%" "%DEST%\.env" >nul 2>&1
    del "%ENV_BACKUP%" >nul 2>&1
    echo         Configuracoes ^(.env^) restauradas.
)

:: Limpar temporarios do GitHub
if exist "%UPDATE_ZIP%"  del "%UPDATE_ZIP%"  >nul 2>&1
if exist "%UPDATE_DIR%"  rd /s /q "%UPDATE_DIR%" >nul 2>&1

echo         Arquivos instalados com sucesso!
echo.

:: -----------------------------------------
:: 4. Dependencias + Atalhos
:: -----------------------------------------
echo  [4/4] Instalando dependencias e criando atalhos...
echo.

python -m pip install --quiet --upgrade pip >nul 2>&1
pip install --quiet "customtkinter>=5.2.0" "openpyxl>=3.1.0" "Pillow>=10.0.0" "matplotlib>=3.8.0" "pywin32>=306" "fpdf2>=2.7.0" "requests>=2.31.0" >nul 2>&1

if exist "%DEST%\requirements.txt" (
    pip install --quiet -r "%DEST%\requirements.txt" >nul 2>&1
)
echo         Dependencias instaladas!

for /f "delims=" %%p in ('python -c "import sys,os; print(os.path.join(os.path.dirname(sys.executable), 'pythonw.exe'))"') do set PYTHONW=%%p
if not exist "%PYTHONW%" (
    for /f "delims=" %%p in ('python -c "import sys; print(sys.executable)"') do set PYTHONW=%%p
)

set "TARGET=%DEST%\executar.py"
set "ICON=%DEST%\assets\icon_somus.ico"
set "PS_SCRIPT=%TEMP%\somus_atalho.ps1"

(
echo $pythonw  = '%PYTHONW%'
echo $target   = '%TARGET%'
echo $workdir  = '%DEST%'
echo $ico      = '%ICON%'
echo $appName  = 'Somus Capital'
echo $desc     = 'Somus Capital - Mesa de Produtos'
echo $desktop   = [Environment]::GetFolderPath('Desktop'^)
echo $startMenu = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Somus Capital'
echo foreach ($name in @('Somus Capital.lnk','Somus Capital.bat','APP Final.lnk','APP SOMUS.lnk'^)^) {
echo     $p = Join-Path $desktop $name
echo     if (Test-Path $p^) { Remove-Item $p -Force }
echo }
echo if (Test-Path $startMenu^) { Remove-Item $startMenu -Recurse -Force }
echo New-Item -ItemType Directory -Path $startMenu -Force ^| Out-Null
echo function New-Shortcut($lnkPath, $exePath, $exeArgs, $workdir, $ico, $desc^) {
echo     $ws = New-Object -ComObject WScript.Shell
echo     $sc = $ws.CreateShortcut($lnkPath^)
echo     $sc.TargetPath = $exePath
echo     if ($exeArgs^) { $sc.Arguments = $exeArgs }
echo     $sc.WorkingDirectory = $workdir
echo     if ($ico -and (Test-Path $ico^)^) { $sc.IconLocation = $ico }
echo     $sc.Description = $desc
echo     $sc.Save(^)
echo }
echo $exeArgs = "`"$target`""
echo $lnkDesktop = Join-Path $desktop "$appName.lnk"
echo New-Shortcut $lnkDesktop $pythonw $exeArgs $workdir $ico $desc
echo $lnkMenu = Join-Path $startMenu "$appName.lnk"
echo New-Shortcut $lnkMenu $pythonw $exeArgs $workdir $ico $desc
echo if (Test-Path $lnkDesktop^) { Write-Host '        Atalho Desktop: OK' }
echo else                         { Write-Host '        Atalho Desktop: ERRO' }
echo if (-not (Test-Path $lnkDesktop^)^) { exit 1 }
) > "%PS_SCRIPT%"

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
set PS_RESULT=%errorlevel%
del "%PS_SCRIPT%" >nul 2>&1

if %PS_RESULT% neq 0 (
    python "%DEST%\criar_atalho.py"
)

echo.
echo  ============================================
echo.
echo     INSTALACAO CONCLUIDA COM SUCESSO!
echo     Pasta: %DEST%
echo     Atalho "Somus Capital" criado no Desktop
echo  ============================================
echo.

:fim
pause
"""


def criar_zip():
    version = _get_version()

    print("=" * 60)
    print(f"  SOMUS CAPITAL - CRIADOR DE INSTALADOR v5.3")
    print(f"  Versao do app: {version}")
    print("=" * 60)
    print()

    if not os.path.exists(os.path.join(BASE_DIR, "executar.py")):
        print("  ERRO: executar.py nao encontrado.")
        return None

    if os.path.exists(OUTPUT_ZIP):
        os.remove(OUTPUT_ZIP)

    total_arquivos = 0
    arcnames_vistos = set()

    # Injetar versao no BAT
    bat_content = INSTALAR_BAT_TEMPLATE.replace("__VERSION__", version)
    bat_content = bat_content.lstrip("\n").replace("\r\n", "\n").replace("\n", "\r\n")

    with zipfile.ZipFile(OUTPUT_ZIP, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:

        print("[1/3] Adicionando INSTALAR.bat...")
        zf.writestr("SomusCapital_Instalador/INSTALAR.bat", bat_content)

        print("[2/3] Adicionando arquivos do APP SOMUS...")
        for root, dirs, files in os.walk(BASE_DIR):
            dirs[:] = [d for d in dirs if d not in EXCLUIR_DIRS]
            for f in files:
                abs_path = os.path.join(root, f)
                rel_path = os.path.relpath(abs_path, BASE_DIR)
                if not deve_incluir(rel_path):
                    continue
                arcname = os.path.join("SomusCapital_Instalador", "APP SOMUS", rel_path)
                if arcname in arcnames_vistos:
                    continue
                arcnames_vistos.add(arcname)
                zf.write(abs_path, arcname)
                size = os.path.getsize(abs_path)
                total_arquivos += 1
                if size > 100_000:
                    print(f"       + {rel_path} ({size/1024:.1f} KB)")

        print("[3/3] Criando estrutura de pastas...")
        for pasta in PASTAS_VAZIAS:
            arcname = os.path.join("SomusCapital_Instalador", "APP SOMUS", pasta, ".gitkeep")
            if arcname not in arcnames_vistos:
                zf.writestr(arcname, "")
                arcnames_vistos.add(arcname)

    size_mb = os.path.getsize(OUTPUT_ZIP) / (1024 * 1024)
    print()
    print(f"  ZIP CRIADO: {OUTPUT_ZIP}")
    print(f"  Versao    : v{version}")
    print(f"  Tamanho   : {size_mb:.1f} MB ({total_arquivos} arquivos)")
    return OUTPUT_ZIP


if __name__ == "__main__":
    resultado = criar_zip()
    print(f"\n  Arquivo pronto: {resultado}" if resultado else "\n  ERRO ao criar.")
    input("\n  Pressione ENTER para fechar...")
