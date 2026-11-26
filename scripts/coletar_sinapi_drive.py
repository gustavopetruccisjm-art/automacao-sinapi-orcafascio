import io
import json
import os
import re
import unicodedata
import zipfile

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
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
      info, scopes=["https://www.googleapis.com/auth/drive.file"]
  )
  return build("drive", "v3", credentials=creds)


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


def processar_e_enviar():
  uf = os.environ.get("UF", "RJ").strip().upper()
  ano_mes = os.environ.get("ANO_MES", "202610").strip()
  desonerado = os.environ.get("DESONERADO", "false").strip().lower() == "true"
  folder_id = os.environ.get("DRIVE_FOLDER_ID")

  if not folder_id:
    raise RuntimeError("Variável DRIVE_FOLDER_ID não configurada.")

  tipo_encargo = "Desonerado" if desonerado else "NaoDesonerado"
  url = (
      f"https://www.caixa.gov.br/Downloads/sinapi-a-partir-jul-2009-{uf.lower()}/"
      f"SINAPI_ref_Insumos_Composicoes_{uf}_{ano_mes}_{tipo_encargo}.zip"
  )

  print(f"Baixando SINAPI {uf} | {ano_mes} | {tipo_encargo}...")
  resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=90)
  resp.raise_for_status()

  pasta_tmp = "./tmp_sinapi"
  os.makedirs(pasta_tmp, exist_ok=True)
  drive_service = obter_servico_drive()

  with zipfile.ZipFile(io.BytesIO(resp.content)) as z:
    for nome_arq in z.namelist():
      if not nome_arq.endswith((".xlsx", ".xls")):
        continue

      is_insumo = "Insumos" in nome_arq and "Analitico" not in nome_arq
      is_comp_sintetica = (
          "Composicoes" in nome_arq and "Sintetico" in nome_arq
      )
      if not (is_insumo or is_comp_sintetica):
        continue

      with z.open(nome_arq) as f:
        df_raw = pd.read_excel(f, header=None)

      header_idx = None
      for idx, row in df_raw.iterrows():
        linha_str = " ".join([normalizar_texto(x) for x in row.values])
        if "CODIGO" in linha_str and (
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
              if "CODIGO" in c
              and ("INSUMO" in c or "COMPOSICAO" in c or c == "CODIGO")
          ),
          None,
      )
      col_desc = next(
          (c for c in df.columns if "DESCRICAO" in c), None
      )
      col_unid = next((c for c in df.columns if "UNIDADE" in c), None)
      col_preco = next(
          (c for c in df.columns if "PRECO" in c or "CUSTO" in c), None
      )

      if not all([col_cod, col_desc, col_unid, col_preco]):
        continue

      df_limpo = df[[col_cod, col_desc, col_unid, col_preco]].dropna(
          subset=[col_cod, col_desc]
      )
      df_limpo.columns = ["codigo", "descricao", "unidade", "preco_unitario"]
      df_limpo["codigo"] = (
          df_limpo["codigo"]
          .astype(str)
          .str.replace(r"\.0$", "", regex=True)
          .str.strip()
      )
      df_limpo = df_limpo[df_limpo["codigo"].str.isdigit()]
      df_limpo["descricao_busca"] = df_limpo["descricao"].apply(normalizar_texto)
      df_limpo["unidade"] = (
          df_limpo["unidade"].astype(str).str.strip().str.upper()
      )
      df_limpo["preco_unitario"] = df_limpo["preco_unitario"].apply(limpar_valor)

      tipo_saida = "insumos" if is_insumo else "composicoes"
      caminho_json = os.path.join(
          pasta_tmp, f"sinapi_{tipo_saida}_{uf}_{ano_mes}_{tipo_encargo}.json"
      )

      with open(caminho_json, "w", encoding="utf-8") as jf:
        json.dump(
            df_limpo.to_dict(orient="records"), jf, ensure_ascii=False, indent=2
        )

      enviar_ou_atualizar_drive(drive_service, caminho_json, folder_id)


if __name__ == "__main__":
  processar_e_enviar()
