"""
Somus Capital - Auto Updater
Verifica atualizacoes no GitHub Releases e aplica automaticamente.
"""

import os
import sys
import zipfile
import tempfile
import threading
import subprocess

# ── Configuracao ──────────────────────────────────────────────────────────────
GITHUB_REPO   = "artursomus1/somus-app"
GITHUB_API    = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
ZIP_PREFIX    = "SomusCapital_Instalador/APP SOMUS/"  # prefixo dentro do ZIP
APP_DIR       = os.path.dirname(os.path.abspath(__file__))

# Arquivos/pastas que nunca sao sobrescritos no update (dados do usuario)
PRESERVAR = {".env"}
# ──────────────────────────────────────────────────────────────────────────────


def _versao_mais_nova(remota: str, local: str) -> bool:
    """Retorna True se a versao remota for maior que a local."""
    try:
        r = tuple(int(x) for x in remota.lstrip("v").split("."))
        l = tuple(int(x) for x in local.lstrip("v").split("."))
        return r > l
    except Exception:
        return False


def verificar_atualizacao(callback_disponivel, callback_erro=None):
    """
    Roda em background. Se houver update disponivel, chama:
        callback_disponivel(versao_str, download_url)
    """
    def _check():
        try:
            import urllib.request, json
            req = urllib.request.Request(
                GITHUB_API,
                headers={"User-Agent": "SomusCapital-Updater/1.0"}
            )
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = json.loads(resp.read().decode())

            versao_remota = data.get("tag_name", "").lstrip("v")
            if not versao_remota:
                return

            from version import VERSION
            if not _versao_mais_nova(versao_remota, VERSION):
                return

            # Encontra o asset ZIP na release
            url_zip = None
            for asset in data.get("assets", []):
                if asset["name"].endswith(".zip"):
                    url_zip = asset["browser_download_url"]
                    break

            if url_zip:
                callback_disponivel(versao_remota, url_zip)

        except Exception:
            pass  # Sem internet ou GitHub fora — silencioso

    threading.Thread(target=_check, daemon=True).start()


def baixar_e_instalar(url_zip, callback_progresso=None, callback_concluido=None, callback_erro=None):
    """
    Baixa o ZIP da Release e extrai sobre a pasta do app.
    Preserva arquivos em PRESERVAR (.env, etc).
    Chama callback_concluido() ao terminar ou callback_erro(msg) em caso de falha.
    """
    def _install():
        temp_zip = None
        backups = {}
        try:
            # 1. Backup dos arquivos preservados
            for nome in PRESERVAR:
                caminho = os.path.join(APP_DIR, nome)
                if os.path.exists(caminho):
                    with open(caminho, "rb") as f:
                        backups[nome] = f.read()

            # 2. Download do ZIP
            import urllib.request
            temp_zip = os.path.join(tempfile.gettempdir(), "somus_update.zip")

            if callback_progresso:
                callback_progresso(5, "Baixando atualizacao...")

            def _reporthook(block_num, block_size, total_size):
                if callback_progresso and total_size > 0:
                    pct = min(90, int(block_num * block_size / total_size * 85) + 5)
                    callback_progresso(pct, "Baixando atualizacao...")

            urllib.request.urlretrieve(url_zip, temp_zip, reporthook=_reporthook)

            if callback_progresso:
                callback_progresso(92, "Extraindo arquivos...")

            # 3. Extrair ZIP (substitui tudo, exceto arquivos preservados)
            with zipfile.ZipFile(temp_zip, "r") as zf:
                membros = [m for m in zf.namelist() if m.startswith(ZIP_PREFIX)]
                for membro in membros:
                    rel = membro[len(ZIP_PREFIX):]
                    if not rel or rel.endswith(".gitkeep"):
                        continue

                    nome_arquivo = os.path.basename(rel)
                    if nome_arquivo in PRESERVAR:
                        continue

                    destino = os.path.join(APP_DIR, rel.replace("/", os.sep))

                    if membro.endswith("/"):
                        os.makedirs(destino, exist_ok=True)
                    else:
                        os.makedirs(os.path.dirname(destino), exist_ok=True)
                        with zf.open(membro) as src, open(destino, "wb") as dst:
                            dst.write(src.read())

            # 4. Restaurar arquivos preservados
            for nome, conteudo in backups.items():
                with open(os.path.join(APP_DIR, nome), "wb") as f:
                    f.write(conteudo)

            if callback_progresso:
                callback_progresso(100, "Concluido!")

            if callback_concluido:
                callback_concluido()

        except Exception as e:
            if callback_erro:
                callback_erro(str(e))
        finally:
            if temp_zip and os.path.exists(temp_zip):
                try:
                    os.remove(temp_zip)
                except Exception:
                    pass

    threading.Thread(target=_install, daemon=True).start()


def reiniciar_app():
    """Reinicia o aplicativo."""
    python = sys.executable
    # Usa pythonw se disponivel (sem janela cmd)
    pythonw = os.path.join(os.path.dirname(python), "pythonw.exe")
    exe = pythonw if os.path.exists(pythonw) else python
    script = os.path.join(APP_DIR, "executar.py")
    subprocess.Popen([exe, script], cwd=APP_DIR)
    sys.exit(0)
