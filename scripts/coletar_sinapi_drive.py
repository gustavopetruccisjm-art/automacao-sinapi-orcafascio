import io
import json
import os
import re
import unicodedata
import zipfile

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
import pandas as pd


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

  query = (
      f"'{folder_id}' in parents and mimeType = 'application/zip' "
      "and trashed = false"
  )
  res = (
      service.files()
      .list(q=query, fields="files(id, name)", pageSize=50)
      .execute()
  )
  arquivos = res.get("files", [])

  escolhido = None
  for arq in arquivos:
    if arq["name"] == nome_padrao:
      escolhido = arq
      break

  if not escolhido and arquivos:
    for arq in arquivos:
      if f"{ano}-{mes}" in arq["name"] or f"{ano}_{mes}" in arq["name"] or ano_mes in arq["name"]:
        escolhido = arq
        break

  if not escolhido and arquivos:
    escolhido = arquivos[0]

  if not escolhido:
    raise FileNotFoundError(
        f"Nenhum arquivo .zip encontrado na pasta do Drive (esperado algo como {nome_padrao}). "
        "Envie o .zip para a pasta 01_BASE_SINAPI_JSON antes de rodar."
    )

  print(f"ZIP localizado no Drive: {escolhido['name']}")
  request = service.files().get_media(fileId=escolhido["id"])
  buffer = io.BytesIO()
  downloader = MediaIoBaseDownload(buffer, request)
  done = False
  while not done:
    _, done = downloader.next_chunk()
  return buffer.getvalue()


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
      .list(q=query, fields="files(id, name)", pageSize=10)
      .execute()
  )
  arquivos = res.get("files", [])

  if arquivos:
    file_id = arquivos[0]["id"]
    atualizado = (
        service.files()
        .update(fileId=file_id, media_body=media, fields="id, name")
        .execute()
    )
    print(f"Atualizado no Drive: {atualizado['name']}")
  else:
    metadata = {"name": nome_arquivo, "parents": [folder_id]}
    criado = (
        service.files()
        .create(body=metadata, media_body=media, fields="id, name")
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
        (
            c
            for c in df.columns
            if "CODIGO" in c or c.startswith("COD")
        ),
        None,
    )
    col_desc = next(
        (c for c in df.columns if "DESCRICAO" in c), None
    )
    col_unid = next((c for c in df.columns if "UNIDADE" in c), None)

    # Procura coluna de preço da UF ou genérica
    col_preco = next(
        (c for c in df.columns if uf in c and ("PRECO" in c or "CUSTO" in c or "R$" in c)),
        None,
    )
    if not col_preco:
      col_preco = next(
          (c for c in df.columns if "PRECO" in c or "CUSTO" in c or "VALOR" in c),
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
  folder_id = os.environ.get("GDRIVE_SINAPI_FOLDER_ID")

  if not folder_id:
    raise RuntimeError("Variável GDRIVE_SINAPI_FOLDER_ID não configurada.")

  drive_service = obter_servico_drive()
  conteudo_zip = localizar_e_baixar_zip(drive_service, folder_id, ano_mes)

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

      base_limpa = re.sub(r"[^A-Za-z0-9_-]+", "_", os.path.splitext(os.path.basename(nome_arq))[0])
      caminho_json = os.path.join(pasta_tmp, f"{base_limpa}_{uf}.json")

      with open(caminho_json, "w", encoding="utf-8") as jf:
        json.dump(dados, jf, ensure_ascii=False, indent=2)

      print(f"Gerado {os.path.basename(caminho_json)} com {len(dados)} itens.")
      enviar_ou_atualizar_drive(drive_service, caminho_json, folder_id)


if __name__ == "__main__":
  processar_e_enviar()
