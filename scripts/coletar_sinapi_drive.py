import io
import json
import os
import re
import time
import unicodedata
import zipfile

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
import pandas as pd
import requests


def normalizar_texto(texto):
  if texto is None or (isinstance(texto, float) and pd.isna(texto)):
    return ""
  texto = str(texto).strip().upper()
  if texto in ("NAN", "NONE", "NAT"):
    return ""
  texto = (
      unicodedata.normalize("NFKD", texto)
      .encode("ASCII", "ignore")
      .decode("utf-8")
  )
  return re.sub(r"\s+", " ", texto)


def limpar_valor(val):
  if isinstance(val, pd.Series):
    val = val.iloc[0] if not val.empty else 0.0
  if val is None or (isinstance(val, float) and pd.isna(val)):
    return 0.0
  if isinstance(val, (int, float)):
    return float(val)
  s = str(val).replace("R$", "").replace(".", "").replace(",", ".").strip()
  if not s or s in ("-", "—", "N/A", "NAN", "NONE"):
    return 0.0
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
  print(
      "Arquivos encontrados na pasta do Drive:",
      [a["name"] for a in todos_arquivos],
  )

  arquivos_zip = [
      a
      for a in todos_arquivos
      if str(a.get("name", "")).lower().endswith(".zip")
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
        f" como {nome_padrao})."
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
      caminho_local, mimetype="application/json", resumable=False
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
    print(f"Atualizado no Drive com sucesso: {atualizado['name']}")
  else:
    try:
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
      print(f"Criado no Drive: {criado['name']}")
    except HttpError as e:
      if "storageQuotaExceeded" in str(e):
        raise RuntimeError(
            f"O arquivo '{nome_arquivo}' ainda não existe na pasta do Drive. "
            "Como Contas de Serviço não têm cota própria, crie um arquivo "
            f"vazio chamado '{nome_arquivo}' na pasta 01_BASE_SINAPI_JSON "
            "para que o robô possa atualizá-lo."
        ) from e
      raise


def encontrar_indice_preco(df_raw, header_idx, headers_norm, uf):
  for i, h in enumerate(headers_norm):
    if h == uf:
      return i

  for i, h in enumerate(headers_norm):
    if "ORIGEM" in h:
      continue
    if re.search(rf"\b{re.escape(uf)}\b", h):
      return i
    for offset in (1, 2, 3):
      if header_idx - offset >= 0:
        acima = normalizar_texto(df_raw.iloc[header_idx - offset, i])
        if acima == uf or re.search(rf"\b{re.escape(uf)}\b", acima):
          return i

  for i, h in enumerate(headers_norm):
    if "ORIGEM" in h:
      continue
    if "PRECO" in h or "CUSTO" in h or "VALOR" in h:
      return i

  return None


def aba_corresponde_regime(aba_norm, desonerado):
  if "ANALITICO" in aba_norm:
    return False
  if desonerado:
    if aba_norm in ("ISD", "CSD") or "SEM DESONERACAO" in aba_norm:
      return False
  else:
    if aba_norm in ("ICD", "CCD") or "COM DESONERACAO" in aba_norm:
      return False
  return True


def extrair_abas_de_planilha(file_obj, uf, desonerado):
  excel_file = pd.ExcelFile(file_obj)
  print("  Abas encontradas:", excel_file.sheet_names)
  resultados_por_aba = {}

  for aba in excel_file.sheet_names:
    aba_norm = normalizar_texto(aba)
    if not aba_corresponde_regime(aba_norm, desonerado):
      print(f"  Pulando aba '{aba}' (fora do regime selecionado).")
      continue

    df_raw = pd.read_excel(excel_file, sheet_name=aba, header=None)
    if df_raw.empty:
      continue

    header_idx = None
    max_scan = min(35, len(df_raw))
    for idx in range(max_scan):
      linha_str = " ".join([normalizar_texto(x) for x in df_raw.iloc[idx].values])
      if ("CODIGO" in linha_str or "COD" in linha_str) and (
          "DESCRICAO" in linha_str or "UNIDADE" in linha_str
      ):
        header_idx = idx
        break

    if header_idx is None:
      continue

    headers_norm = [normalizar_texto(c) for c in df_raw.iloc[header_idx].values]

    idx_cod = next(
        (
            i
            for i, h in enumerate(headers_norm)
            if "CODIGO" in h or h.startswith("COD")
        ),
        None,
    )
    idx_desc = next(
        (i for i, h in enumerate(headers_norm) if "DESCRICAO" in h), None
    )
    idx_unid = next(
        (i for i, h in enumerate(headers_norm) if "UNIDADE" in h), None
    )
    idx_preco = encontrar_indice_preco(df_raw, header_idx, headers_norm, uf)

    if any(x is None for x in (idx_cod, idx_desc, idx_unid, idx_preco)):
      print(
          f"  Aba '{aba}': colunas insuficientes "
          f"(cod={idx_cod}, desc={idx_desc}, unid={idx_unid}, preco={idx_preco})."
      )
      continue

    print(
        f"  Extraindo aba '{aba}' -> cod[{idx_cod}], desc[{idx_desc}], "
        f"unid[{idx_unid}], preco_{uf}[{idx_preco}] ({headers_norm[idx_preco]})"
    )

    df_dados = df_raw.iloc[header_idx + 1 :]
    registros = []

    for _, row in df_dados.iterrows():
      cod_raw = row.iloc[idx_cod]
      desc_raw = row.iloc[idx_desc]
      if pd.isna(cod_raw) or pd.isna(desc_raw):
        continue

      cod = re.sub(r"\.0$", "", str(cod_raw)).strip()
      if not cod.isdigit():
        continue

      desc = str(desc_raw).strip()
      unid = str(row.iloc[idx_unid]).strip().upper()
      preco = limpar_valor(row.iloc[idx_preco])

      registros.append({
          "aba": aba,
          "codigo": cod,
          "descricao": desc,
          "descricao_busca": normalizar_texto(desc),
          "unidade": unid,
          "preco_unitario": preco,
      })

    if registros:
      resultados_por_aba[aba] = registros

  return resultados_por_aba


def definir_nome_arquivo_saida(nome_arq_xlsx, aba, uf, ano_mes, tipo_encargo):
  aba_norm = normalizar_texto(aba)
  if aba_norm in ("ISD", "ICD") or "INSUMO" in aba_norm:
    return f"sinapi_insumos_{uf}_{ano_mes}_{tipo_encargo}.json"
  if aba_norm in ("CSD", "CCD") or "COMPOSIC" in aba_norm:
    return f"sinapi_composicoes_{uf}_{ano_mes}_{tipo_encargo}.json"

  base_limpa = re.sub(
      r"[^A-Za-z0-9_-]+", "_", os.path.splitext(os.path.basename(nome_arq_xlsx))[0]
  )
  aba_limpa = re.sub(r"[^A-Za-z0-9_-]+", "_", aba_norm)
  return f"{base_limpa}_{aba_limpa}_{uf}.json"


def processar_e_enviar():
  uf = os.environ.get("UF", "RJ").strip().upper()
  ano_mes = os.environ.get("ANO_MES", "202610").strip()
  desonerado = os.environ.get("DESONERADO", "false").strip().lower() == "true"
  folder_id = os.environ.get("GDRIVE_SINAPI_FOLDER_ID", "").strip()

  if not folder_id:
    raise RuntimeError("Variável GDRIVE_SINAPI_FOLDER_ID não configurada.")

  tipo_encargo = "Desonerado" if desonerado else "NaoDesonerado"
  drive_service = obter_servico_drive()

  try:
    conteudo_zip = localizar_e_baixar_zip(drive_service, folder_id, ano_mes)
  except FileNotFoundError:
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

      # Foca apenas na planilha SINAPI_Referência (preços de insumos e composições)
      if "REFER" not in normalizar_texto(nome_arq):
        print(f"Ignorando planilha auxiliar: {nome_arq}")
        continue

      print(f"Lendo planilha principal: {nome_arq}")
      with z.open(nome_arq) as f:
        abas_extraidas = extrair_abas_de_planilha(
            io.BytesIO(f.read()), uf, desonerado
        )

      if not abas_extraidas:
        print(f"Sem registros tabulares padrão em {nome_arq}, pulando.")
        continue

      for aba, dados in abas_extraidas.items():
        nome_json = definir_nome_arquivo_saida(
            nome_arq, aba, uf, ano_mes, tipo_encargo
        )
        caminho_json = os.path.join(pasta_tmp, nome_json)

        with open(caminho_json, "w", encoding="utf-8") as jf:
          json.dump(dados, jf, ensure_ascii=False, indent=2)

        print(f"Gerado {nome_json} com {len(dados)} itens.")
        enviar_ou_atualizar_drive(drive_service, caminho_json, folder_id)


if __name__ == "__main__":
  processar_e_enviar()
