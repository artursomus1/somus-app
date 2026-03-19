"""Cria o atalho 'Somus Capital' na Area de Trabalho do usuario."""
import os
import sys
import subprocess


def get_desktop_path():
    """Detecta o caminho real da Area de Trabalho via registro do Windows."""
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

    # Fallbacks
    for candidate in [
        os.path.join(os.environ.get("USERPROFILE", ""), "Desktop"),
        os.path.join(os.environ.get("USERPROFILE", ""), "Area de Trabalho"),
    ]:
        if os.path.exists(candidate):
            return candidate

    # OneDrive
    for var in ["OneDrive", "OneDriveConsumer", "OneDriveCommercial"]:
        od = os.environ.get(var, "")
        if od:
            for sub in ["Desktop", "Area de Trabalho"]:
                p = os.path.join(od, sub)
                if os.path.exists(p):
                    return p

    return os.path.join(os.environ.get("USERPROFILE", r"C:\Users\User"), "Desktop")


def criar_atalho():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ICON_PATH = os.path.join(BASE_DIR, "assets", "icon_somus.ico")
    TARGET_SCRIPT = os.path.join(BASE_DIR, "executar.py")

    # Buscar pythonw.exe
    python_dir = os.path.dirname(sys.executable)
    pythonw = os.path.join(python_dir, "pythonw.exe")
    if not os.path.exists(pythonw):
        pythonw = sys.executable

    desktop = get_desktop_path()
    shortcut_path = os.path.join(desktop, "Somus Capital.lnk")

    print(f"  Desktop: {desktop}")
    print(f"  Target:  {TARGET_SCRIPT}")
    print(f"  pythonw: {pythonw}")

    # Remover atalhos antigos
    for old_name in ["Somus Capital.lnk", "Somus Capital.bat", "APP Final.lnk"]:
        old_path = os.path.join(desktop, old_name)
        if os.path.exists(old_path):
            os.remove(old_path)
            print(f"  Removido atalho antigo: {old_name}")

    # Metodo 1: pywin32 COM
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = pythonw
        shortcut.Arguments = f'"{TARGET_SCRIPT}"'
        shortcut.WorkingDirectory = BASE_DIR
        if os.path.exists(ICON_PATH):
            shortcut.IconLocation = ICON_PATH
        shortcut.Description = "Somus Capital - Mesa de Produtos"
        shortcut.save()

        if os.path.exists(shortcut_path):
            print(f"  Atalho criado: {shortcut_path}")
            return True
    except ImportError:
        print("  pywin32 nao disponivel, tentando PowerShell...")
    except Exception as e:
        print(f"  Erro COM: {e}")

    # Metodo 2: PowerShell
    try:
        ps_script = f"""
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut('{shortcut_path}')
$sc.TargetPath = '{pythonw}'
$sc.Arguments = '"{TARGET_SCRIPT}"'
$sc.WorkingDirectory = '{BASE_DIR}'
$ico = '{ICON_PATH}'
if (Test-Path $ico) {{ $sc.IconLocation = $ico }}
$sc.Description = 'Somus Capital - Mesa de Produtos'
$sc.Save()
"""
        result = subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps_script],
            capture_output=True, text=True, timeout=15
        )
        if os.path.exists(shortcut_path):
            print(f"  Atalho criado via PowerShell: {shortcut_path}")
            return True
    except Exception as e:
        print(f"  Erro PowerShell: {e}")

    print("  FALHA: Nenhum metodo de criacao de atalho funcionou.")
    return False


if __name__ == "__main__":
    criar_atalho()
