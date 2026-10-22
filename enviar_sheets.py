#!/usr/bin/env python3
"""
Envia os CSVs de dados/abas/ para a planilha do Google que alimenta o painel,
substituindo o conteúdo de cada aba NO MESMO arquivo (o link não muda).

Credencial (a primeira que existir):
  1. variável GOOGLE_SERVICE_ACCOUNT_JSON  (conteúdo do JSON da conta de serviço)
  2. variável GOOGLE_TOKEN_JSON            (conteúdo do token_sheets.json OAuth)
  3. arquivo service_account.json ou token_sheets.json na pasta

Uso:  python enviar_sheets.py [--pasta dados/abas] [--planilha ID]
"""
import argparse
import json
import os
import re
import sys
import time
from pathlib import Path

import pandas as pd
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "19AbG5ogeBWFjwC1y-DWHPRMbt_7ZqmTWoYrxWiyq_XQ")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
ABAS = ["aba12", "aba15", "aba16", "aba17", "aba23", "aba24", "aba25", "aba26", "municipiosAlerta"]
LINHAS_POR_ENVIO = 5000


def credenciais():
    sa = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa and Path("service_account.json").exists():
        sa = Path("service_account.json").read_text(encoding="utf-8")
    if sa:
        print("credencial: conta de serviço")
        return service_account.Credentials.from_service_account_info(json.loads(sa), scopes=SCOPES)

    tk = os.environ.get("GOOGLE_TOKEN_JSON")
    if not tk and Path("token_sheets.json").exists():
        tk = Path("token_sheets.json").read_text(encoding="utf-8")
    if tk:
        print("credencial: token OAuth")
        creds = Credentials.from_authorized_user_info(json.loads(tk), SCOPES)
        if not creds.valid:
            creds.refresh(Request())
        return creds
    sys.exit("Nenhuma credencial do Google encontrada (veja docs/configuracao.md).")


def valor(v):
    """Texto do CSV -> número quando for número (JSON não depende do idioma da planilha)."""
    if v == "":
        return ""
    if re.fullmatch(r"-?\d+", v):
        return int(v)
    if re.fullmatch(r"-?\d+\.\d+", v):
        f = float(v)
        return int(f) if f.is_integer() else f
    return v


def executar(req, tentativas=5):
    for i in range(tentativas):
        try:
            return req.execute()
        except HttpError as e:
            if e.resp.status in (429, 500, 502, 503) and i < tentativas - 1:
                time.sleep(2 ** i * 5)
                continue
            raise


def garantir_tamanho(svc, abas_info, nome, linhas, colunas):
    p = abas_info[nome]
    grade = p["gridProperties"]
    pedidos = []
    if grade["rowCount"] < linhas or grade.get("columnCount", 26) < colunas:
        pedidos.append({"updateSheetProperties": {
            "properties": {"sheetId": p["sheetId"], "gridProperties": {
                "rowCount": max(grade["rowCount"], linhas + 100),
                "columnCount": max(grade.get("columnCount", 26), colunas)}},
            "fields": "gridProperties(rowCount,columnCount)"}})
    if pedidos:
        executar(svc.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": pedidos}))


def coluna_letra(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def main():
    global SPREADSHEET_ID
    ap = argparse.ArgumentParser()
    ap.add_argument("--pasta", default="dados/abas")
    ap.add_argument("--planilha", default=SPREADSHEET_ID)
    a = ap.parse_args()
    SPREADSHEET_ID = a.planilha

    svc = build("sheets", "v4", credentials=credenciais(), cache_discovery=False)
    meta = executar(svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID,
                                            fields="sheets.properties"))
    abas_info = {s["properties"]["title"]: s["properties"] for s in meta["sheets"]}

    faltando = [n for n in ABAS if n not in abas_info]
    if faltando:
        sys.exit(f"abas não existem na planilha: {faltando}")

    for nome in ABAS:
        arq = Path(a.pasta) / f"{nome}.csv"
        df = pd.read_csv(arq, dtype=str, keep_default_na=False)
        linhas = [list(df.columns)] + [[valor(v) for v in r] for r in df.itertuples(index=False, name=None)]
        ncol = len(df.columns)
        garantir_tamanho(svc, abas_info, nome, len(linhas), ncol)

        # Limpa da linha 2 para baixo (como o processo atual) e regrava tudo.
        executar(svc.spreadsheets().values().clear(spreadsheetId=SPREADSHEET_ID, range=f"{nome}!A2:ZZ"))
        for ini in range(0, len(linhas), LINHAS_POR_ENVIO):
            bloco = linhas[ini:ini + LINHAS_POR_ENVIO]
            faixa = f"{nome}!A{ini + 1}:{coluna_letra(ncol)}{ini + len(bloco)}"
            executar(svc.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID, range=faixa,
                valueInputOption="USER_ENTERED", body={"values": bloco}))
        print(f"ok {nome}: {len(df)} linhas")
    print("Planilha atualizada.")


if __name__ == "__main__":
    main()
