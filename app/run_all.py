"""
Orchestrator runner — executa múltiplos scripts Python na ordem definida.
Cada arquivo é executado em namespace isolado (via runpy).
- Gera logs detalhados com sucesso, erro e tempo de execução.
- Salva os logs em um arquivo local `orchestrator.log`.
"""

import runpy
import os
import sys
from pathlib import Path
import time
import traceback
from datetime import datetime
import logging

# ==============================
# CONFIGURAÇÃO DE LOGS
# ==============================
APP_DIR = Path(__file__).resolve().parent
LOG_FILE = APP_DIR / "orchestrator.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, mode="a", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

# ==============================
# ORDEM DOS SCRIPTS A EXECUTAR
# ==============================
order = [
    "itau-main.py",
    "itau-pandas.py",
    "caixa-main.py",
    "caixa-pandas.py",
    "consolidado-pandas.py",
]

# ==============================
# EXECUÇÃO DE UM SCRIPT
# ==============================
def run_script_rel(path: str):
    """Executa um script Python isoladamente e retorna o status."""
    full = APP_DIR / path
    if not full.exists():
        logging.warning(f"[SKIP] Arquivo não encontrado: {full}")
        return {"path": path, "success": False, "error": "Arquivo não encontrado", "duration": None}

    logging.info(f"\n--- Iniciando execução: {path} ---")

    result = {
        "path": path,
        "start": time.time(),
        "end": None,
        "duration": None,
        "success": False,
        "error": None,
    }

    try:
        # Executa o script em namespace isolado
        ns = runpy.run_path(str(full), run_name="__main__")

        # Se o script tiver função main(), chama explicitamente
        if "main" in ns and callable(ns["main"]):
            logging.info(f"Chamando main() de {path}")
            ns["main"]()

        result["success"] = True
        logging.info(f"✅ {path} finalizado com sucesso.")

    except Exception as e:
        tb = traceback.format_exc()
        result["error"] = tb
        logging.error(f"❌ Erro ao executar {path}: {e}")
        logging.error(tb)

    finally:
        result["end"] = time.time()
        result["duration"] = result["end"] - result["start"]
        dur = f"{result['duration']:.2f}s" if result["duration"] else "n/a"
        logging.info(f"⏱️ Duração de {path}: {dur}")

    return result


# ==============================
# FUNÇÃO PRINCIPAL
# ==============================
def main():
    os.chdir(APP_DIR)
    results = []
    overall_start = time.time()
    logging.info("🚀 Iniciando execução completa do Orchestrator...")

    for p in order:
        r = run_script_rel(p)
        results.append(r)
        logging.info("-" * 60)

    overall_end = time.time()
    total_duration = overall_end - overall_start
    logging.info(f"🏁 Execução concluída em {total_duration:.2f}s")

    # Resumo final
    logging.info("\nResumo da execução:")
    any_error = False
    for r in results:
        status = "OK" if r.get("success") else "ERRO"
        dur = f"{r.get('duration'):.2f}s" if r.get("duration") else "n/a"
        logging.info(f" - {r['path']}: {status} ({dur})")
        if not r.get("success"):
            any_error = True

    resumo = (
        f"\n{'='*60}\n"
        f"🕒 Início: {time.ctime(overall_start)}\n"
        f"🏁 Fim: {time.ctime(overall_end)}\n"
        f"⏱️ Duração total: {total_duration:.2f}s\n"
        f"Status geral: {'ERROS DETECTADOS' if any_error else 'SUCESSO TOTAL'}\n"
        f"Log salvo em: {LOG_FILE}\n{'='*60}\n"
    )
    logging.info(resumo)

    print(resumo)  # exibe também no console


if __name__ == "__main__":
    main()
