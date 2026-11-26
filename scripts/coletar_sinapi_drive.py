import io
import json
import os
import re
import time
import unicodedata
import zipfile

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
import pandas as pd
import requests


def normalizar_texto(texto):
  if pd.isna(texto):
    return ""
  texto = str(texto).strip().upper()
  texto = (
      unicodedata.normalize("NFKD", texto)
      .encode("ASCII", "ignore")
      .decode("utf-8")
  )
  return re.sub(r"\s+", " ", texto)


def limpar_valor(val):
  if pd.isna(val):
    return 0.0
  if isinstance(val, (int, float)):
    return float(val)
  s = str(val).replace("R$", "").replace(".", "").replace(",", ".").strip()
  try:
    return float(s)
  except ValueError:
    return 0.0


def obter_servico_drive():
  cred_json = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON")
  if not cred_json:
    raise RuntimeError(
        "Variável GDRIVE_SERVICE_ACCOUNT_JSON não encontrada no ambiente."
    )
  info = json.loads(cred_json)
  creds = service_account.Credentials.from_service_account_info(
      info, scopes=["https://www.googleapis.com/auth/drive"]
  )
  return build("drive", "v3", credentials=creds)


def localizar_e_baixar_zip(service, folder_id, ano_mes):
  ano = ano_mes[:4]
  mes = ano_mes[4:6]
  nome_padrao = f"SINAPI-{ano}-{mes}-formato-xlsx.zip"

  query = f"'{folder_id}' in parents and trashed = false"
  res = (
      service.files()
      .list(
          q=query,
          fields="files(id, name, mimeType)",
          pageSize=50,
          supportsAllDrives=True,
          includeItemsFromAllDrives=True,
      )
      .execute()
  )

  todos_arquivos = res.get("files", [])
  print("Arquivos encontrados na pasta do Drive:", [a["name"] for a in todos_arquivos])

  arquivos_zip = [
      a for a in todos_arquivos if str(a.get("name", "")).lower().endswith(".zip")
  ]

  escolhido = None
  for arq in arquivos_zip:
    if arq["name"] == nome_padrao:
      escolhido = arq
      break

  if not escolhido and arquivos_zip:
    for arq in arquivos_zip:
      if (
          f"{ano}-{mes}" in arq["name"]
          or f"{ano}_{mes}" in arq["name"]
          or ano_mes in arq["name"]
      ):
        escolhido = arq
        break

  if not escolhido and arquivos_zip:
    escolhido = arquivos_zip[0]

  if not escolhido:
    raise FileNotFoundError(
        f"Nenhum arquivo .zip encontrado na pasta do Drive (esperado algo"
        f" como {nome_padrao}). Envie o .zip para a pasta 01_BASE_SINAPI_JSON"
        " antes de rodar."
    )

  print(f"ZIP localizado no Drive: {escolhido['name']}")
  request = service.files().get_media(
      fileId=escolhido["id"], supportsAllDrives=True
  )
  buffer = io.BytesIO()
  downloader = MediaIoBaseDownload(buffer, request)
  done = False
  while not done:
    _, done = downloader.next_chunk()
  return buffer.getvalue()


def baixar_zip_caixa(url, tentativas=2, espera_inicial=3):
  session = requests.Session()
  headers = {
      "User-Agent": (
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
          "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
      ),
      "Accept": "application/octet-stream,*/*;q=0.8",
      "Accept-Language": "pt-BR,pt;q=0.9",
      "Referer": "https://www.caixa.gov.br/",
  }
  for tentativa in range(1, tentativas + 1):
    resp = session.get(url, headers=headers, timeout=60)
    if resp.status_code == 200:
      return resp.content
    if resp.status_code in (429, 500, 502, 503, 504) and tentativa < tentativas:
      time.sleep(espera_inicial * tentativa)
      continue
    resp.raise_for_status()


def enviar_ou_atualizar_drive(service, caminho_local, folder_id):
  nome_arquivo = os.path.basename(caminho_local)
  media = MediaFileUpload(
      caminho_local, mimetype="application/json", resumable=True
  )

  query = (
      f"name = '{nome_arquivo}' and '{folder_id}' in parents "
      "and trashed = false"
  )
  res = (
      service.files()
      .list(
          q=query,
          fields="files(id, name)",
          pageSize=10,
          supportsAllDrives=True,
          includeItemsFromAllDrives=True,
      )
      .execute()
  )
  arquivos = res.get("files", [])

  if arquivos:
    file_id = arquivos[0]["id"]
    atualizado = (
        service.files()
        .update(
            fileId=file_id,
            media_body=media,
            fields="id, name",
            supportsAllDrives=True,
        )
        .execute()
    )
    print(f"Atualizado no Drive: {atualizado['name']}")
  else:
    metadata = {"name": nome_arquivo, "parents": [folder_id]}
    criado = (
        service.files()
        .create(
            body=metadata,
            media_body=media,
            fields="id, name",
            supportsAllDrives=True,
        )
        .execute()
    )
    print(f"Enviado para o Drive: {criado['name']}")


def extrair_registros_de_planilha(file_obj, uf):
  excel_file = pd.ExcelFile(file_obj)
  registros = []

  for aba in excel_file.sheet_names:
    df_raw = excel_file.parse(aba, header=None)
    if df_raw.empty:
      continue

    header_idx = None
    for idx, row in df_raw.iterrows():
      linha_str = " ".join([normalizar_texto(x) for x in row.values])
      if ("CODIGO" in linha_str or "COD" in linha_str) and (
          "DESCRICAO" in linha_str or "UNIDADE" in linha_str
      ):
        header_idx = idx
        break

    if header_idx is None:
      continue

    df = df_raw.iloc[header_idx + 1 :].copy()
    df.columns = [normalizar_texto(c) for c in df_raw.iloc[header_idx]]

    col_cod = next(
        (c for c in df.columns if "CODIGO" in c or c.startswith("COD")),
        None,
    )
    col_desc = next((c for c in df.columns if "DESCRICAO" in c), None)
    col_unid = next((c for c in df.columns if "UNIDADE" in c), None)

    col_preco = next(
        (
            c
            for c in df.columns
            if uf in c and ("PRECO" in c or "CUSTO" in c or "R$" in c)
        ),
        None,
    )
    if not col_preco:
      col_preco = next(
          (
              c
              for c in df.columns
              if "PRECO" in c or "CUSTO" in c or "VALOR" in c
          ),
          None,
      )

    if not all([col_cod, col_desc, col_unid, col_preco]):
      continue

    sub = df[[col_cod, col_desc, col_unid, col_preco]].dropna(
        subset=[col_cod, col_desc]
    )
    for _, r in sub.iterrows():
      cod = re.sub(r"\.0$", "", str(r[col_cod])).strip()
      if not cod.isdigit():
        continue
      desc = str(r[col_desc]).strip()
      unid = str(r[col_unid]).strip().upper()
      preco = limpar_valor(r[col_preco])

      registros.append({
          "aba": aba,
          "codigo": cod,
          "descricao": desc,
          "descricao_busca": normalizar_texto(desc),
          "unidade": unid,
          "preco_unitario": preco,
      })

  return registros


def processar_e_enviar():
  uf = os.environ.get("UF", "RJ").strip().upper()
  ano_mes = os.environ.get("ANO_MES", "202610").strip()
  desonerado = os.environ.get("DESONERADO", "false").strip().lower() == "true"
  folder_id = os.environ.get("GDRIVE_SINAPI_FOLDER_ID", "").strip()

  if not folder_id:
    raise RuntimeError("Variável GDRIVE_SINAPI_FOLDER_ID não configurada.")

  drive_service = obter_servico_drive()

  try:
    conteudo_zip = localizar_e_baixar_zip(drive_service, folder_id, ano_mes)
  except FileNotFoundError:
    tipo_encargo = "Desonerado" if desonerado else "NaoDesonerado"
    url = (
        f"https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-{uf.lower()}/"
        f"SINAPI_ref_Insumos_Composicoes_{uf}_{ano_mes}_{tipo_encargo}.zip"
    )
    print("ZIP não encontrado no Drive. Tentando download direto da Caixa...")
    conteudo_zip = baixar_zip_caixa(url)

  pasta_tmp = "./tmp_sinapi"
  os.makedirs(pasta_tmp, exist_ok=True)

  with zipfile.ZipFile(io.BytesIO(conteudo_zip)) as z:
    for nome_arq in z.namelist():
      if not nome_arq.endswith((".xlsx", ".xls")):
        continue

      print(f"Lendo planilha dentro do ZIP: {nome_arq}")
      with z.open(nome_arq) as f:
        dados = extrair_registros_de_planilha(io.BytesIO(f.read()), uf)

      if not dados:
        print(f"Sem registros tabulares padrão em {nome_arq}, pulando.")
        continue

      base_limpa = re.sub(
          r"[^A-Za-z0-9_-]+",
          "_",
          os.path.splitext(os.path.basename(nome_arq))[0],
      )
      caminho_json = os.path.join(pasta_tmp, f"{base_limpa}_{uf}.json")

      with open(caminho_json, "w", encoding="utf-8") as jf:
        json.dump(dados, jf, ensure_ascii=False, indent=2)

      print(f"Gerado {os.path.basename(caminho_json)} com {len(dados)} itens.")
      enviar_ou_atualizar_drive(drive_service, caminho_json, folder_id)


if __name__ == "__main__":
  processar_e_enviar()
