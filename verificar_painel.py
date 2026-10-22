#!/usr/bin/env python3
"""
Confere se o painel do Tableau Public foi republicado desde a última extração.

  python verificar_painel.py --checar     # escreve mudou=true|false em $GITHUB_OUTPUT
  python verificar_painel.py --registrar  # grava a data da versão extraída

A data vem do campo lastUpdateDate da API pública de perfil do Tableau.
"""
import argparse
import csv
import json
import os
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

URL = "https://public.tableau.com/profile/api/workbook/MonitoramentodasDDA"
CONTROLE = Path("dados/controle/painel_tableau.json")
ATUALIZACAO = Path("dados/abas/atualizacao.csv")
BRT = timezone(timedelta(hours=-3))


def versao_publicada():
    r = requests.get(URL, timeout=60, headers={"User-Agent": "Mozilla/5.0"})
    r.raise_for_status()
    ms = r.json()["lastUpdateDate"]
    return ms, datetime.fromtimestamp(ms / 1000, BRT)


def saida(chave, valor):
    print(f"{chave}={valor}")
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a", encoding="utf-8") as f:
            f.write(f"{chave}={valor}\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--checar", action="store_true")
    ap.add_argument("--registrar", action="store_true")
    a = ap.parse_args()

    ms, quando = versao_publicada()
    print(f"painel publicado em {quando:%d/%m/%Y %H:%M} (Brasília)")

    if a.checar:
        anterior = json.loads(CONTROLE.read_text())["lastUpdateDate"] if CONTROLE.exists() else None
        saida("mudou", "true" if anterior != ms else "false")
    elif a.registrar:
        CONTROLE.parent.mkdir(parents=True, exist_ok=True)
        CONTROLE.write_text(json.dumps({"lastUpdateDate": ms,
                                        "publicado_em": quando.isoformat()}, indent=1))
        agora = datetime.now(BRT)
        ATUALIZACAO.parent.mkdir(parents=True, exist_ok=True)
        with ATUALIZACAO.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["DATA_EXTRACAO", "TEXTO_EXTRACAO", "PAINEL_TABLEAU_PUBLICADO_EM"])
            w.writerow([agora.isoformat(timespec="seconds"),
                        f"{agora:%d/%m/%Y} às {agora:%H:%M} hrs",
                        quando.isoformat(timespec="seconds")])
    else:
        sys.exit("use --checar ou --registrar")


if __name__ == "__main__":
    main()
