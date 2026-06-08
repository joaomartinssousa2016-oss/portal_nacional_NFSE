import os
import re
import time
import requests
from requests_pkcs12 import Pkcs12Adapter
import xml.etree.ElementTree as ET
import pandas as pd

# ----------------- CONFIGURAÇÕES -----------------

BASE_URL = "https://adn.nfse.gov.br"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DANFSE_URL_TEMPLATE = BASE_URL + "/danfse/{chave}"

XML_DIR = os.path.join(BASE_DIR, "xmlspdf")          # pasta dos XMLs de entrada
CERT_DIR = os.path.join(BASE_DIR, "certificados")    # certificados/{cnpj}.pfx
OUTPUT_DIR = os.path.join(BASE_DIR, "notaspdf")      # pasta dos PDFs de saída
EMPRESAS_XLSX = os.path.join(BASE_DIR, "empresas.xlsx")

DELAY_SECONDS = 5            # delay entre requisições

NFSE_NS = {"n": "http://www.sped.fazenda.gov.br/nfse"}
# -------------------------------------------------


def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def only_digits(s: str) -> str:
    if not s:
        return ""
    return re.sub(r"\D", "", str(s))


def carregar_empresas():
    """
    Lê empresas.xlsx e retorna:
    - dict_cnpj_senha: {cnpj_digits: senha}
    """
    if not os.path.exists(EMPRESAS_XLSX):
        raise FileNotFoundError(f"Planilha de empresas não encontrada: {EMPRESAS_XLSX}")

    df = pd.read_excel(EMPRESAS_XLSX, dtype=str)

    # Garante colunas
    colunas = [c.lower() for c in df.columns]
    if "cnpj" not in colunas or "senha" not in colunas:
        raise ValueError("A planilha empresas.xlsx deve conter as colunas 'cnpj' e 'senha'.")

    # Normaliza nomes de coluna
    df.columns = [c.lower() for c in df.columns]

    dict_cnpj_senha = {}
    for _, row in df.iterrows():
        cnpj = only_digits(row["cnpj"])
        senha = (row["senha"] or "").strip()
        if cnpj:
            dict_cnpj_senha[cnpj] = senha

    return dict_cnpj_senha


def montar_sessao_pkcs12(cert_path: str, senha_cert: str) -> requests.Session:
    """
    Cria sessão Requests com certificado PFX informado e senha da planilha.
    """
    if not os.path.exists(cert_path):
        raise FileNotFoundError(f"Certificado não encontrado: {cert_path}")

    s = requests.Session()
    s.mount(
        BASE_URL,
        Pkcs12Adapter(
            pkcs12_filename=cert_path,
            pkcs12_password=senha_cert
        )
    )
    return s


def extrair_dados_xml(xml_path: str):
    """
    Lê o XML e retorna (chave_acesso, cnpj_prestador, cnpj_tomador).
    - chave: vem do atributo Id de <infNFSe>, ex: "NFS5219..." -> só dígitos.
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Chave de acesso
    chave = ""
    inf_nfse = root.find(".//n:infNFSe", NFSE_NS)
    if inf_nfse is not None:
        id_attr = inf_nfse.get("Id", "")
        chave = only_digits(id_attr)

    # fallback (caso em algum layout não venha Id)
    if not chave:
        ndfse_el = root.find(".//n:nDFSe", NFSE_NS)
        if ndfse_el is not None:
            chave = only_digits(ndfse_el.text)

    # Prestador: primeiro tenta <emit><CNPJ>, se não achar usa <prest><CNPJ>
    prest_cnpj_el = root.find(".//n:emit/n:CNPJ", NFSE_NS)
    if prest_cnpj_el is None:
        prest_cnpj_el = root.find(".//n:prest/n:CNPJ", NFSE_NS)
    prest_cnpj = only_digits(prest_cnpj_el.text) if prest_cnpj_el is not None else ""

    # Tomador: <toma><CNPJ>
    toma_cnpj_el = root.find(".//n:toma/n:CNPJ", NFSE_NS)
    toma_cnpj = only_digits(toma_cnpj_el.text) if toma_cnpj_el is not None else ""

    return chave, prest_cnpj, toma_cnpj


def baixar_danfse_pdf(sess: requests.Session, chave: str) -> str | None:
    """
    Faz o GET em /danfse/{chave} e salva o PDF em OUTPUT_DIR.
    Retorna o caminho do arquivo salvo ou None em caso de erro.
    """
    chave = only_digits(chave)
    if not chave:
        print("  [AVISO] Chave vazia/inválida, pulando...")
        return None

    url = DANFSE_URL_TEMPLATE.format(chave=chave)
    print(f"  Baixando DANFSe da chave {chave} -> {url}")

    try:
        resp = sess.get(url, timeout=60)
    except Exception as e:
        print(f"  [ERRO] Falha na requisição para {chave}: {e}")
        return None

    if resp.status_code != 200:
        print(f"  [ERRO] HTTP {resp.status_code} ao baixar chave {chave}")
        return None

    content_type = resp.headers.get("Content-Type", "")
    if "pdf" not in content_type.lower():
        print(f"  [AVISO] Conteúdo não parece ser PDF (Content-Type={content_type}) para chave {chave}")

    ensure_dir(OUTPUT_DIR)
    file_path = os.path.join(OUTPUT_DIR, f"{chave}.pdf")
    try:
        with open(file_path, "wb") as f:
            f.write(resp.content)
        print(f"  [OK] PDF salvo em: {file_path}")
        return file_path
    except Exception as e:
        print(f"  [ERRO] Falha ao salvar PDF em {file_path}: {e}")
        return None


def main():
    # Carrega mapa CNPJ -> senha a partir da planilha
    try:
        dict_cnpj_senha = carregar_empresas()
    except Exception as e:
        print(f"[ERRO] Ao carregar empresas.xlsx: {e}")
        return

    if not os.path.exists(XML_DIR):
        print(f"Pasta de XMLs não encontrada: {XML_DIR}")
        return

    xml_files = [f for f in os.listdir(XML_DIR) if f.lower().endswith(".xml")]
    if not xml_files:
        print(f"Nenhum XML encontrado em: {XML_DIR}")
        return

    print(f"Total de XMLs encontrados: {len(xml_files)}")

    total_ok = 0
    total_err = 0

    for idx, filename in enumerate(sorted(xml_files)):
        xml_path = os.path.join(XML_DIR, filename)
        print(f"\nXML {idx} - Arquivo: {filename}")

        try:
            chave, prest_cnpj, toma_cnpj = extrair_dados_xml(xml_path)
        except Exception as e:
            print(f"  [ERRO] Falha ao ler/parsing do XML: {e}")
            total_err += 1
            continue

        print(f"  Chave: {chave}")
        print(f"  Prestador: {prest_cnpj} | Tomador: {toma_cnpj}")

        if not chave:
            print("  [ERRO] Não foi possível extrair a chave de acesso, pulando XML.")
            total_err += 1
            continue

        # ordem de tentativa: prestador -> tomador
        cnpjs_para_tentar = []
        if prest_cnpj:
            cnpjs_para_tentar.append(prest_cnpj)
        if toma_cnpj and toma_cnpj != prest_cnpj:
            cnpjs_para_tentar.append(toma_cnpj)

        if not cnpjs_para_tentar:
            print("  [ERRO] Nenhum CNPJ de prestador/tomador encontrado no XML.")
            total_err += 1
            continue

        pdf_ok = False
        for cnpj in cnpjs_para_tentar:
            cnpj_digits = only_digits(cnpj)

            if cnpj_digits not in dict_cnpj_senha:
                print(f"  [AVISO] CNPJ {cnpj_digits} não encontrado na planilha empresas.xlsx; pulando este CNPJ.")
                continue

            senha_cert = dict_cnpj_senha[cnpj_digits]
            if not senha_cert:
                print(f"  [AVISO] Senha vazia para CNPJ {cnpj_digits} na planilha; pulando este CNPJ.")
                continue

            cert_path = os.path.join(CERT_DIR, f"{cnpj_digits}.pfx")
            if not os.path.exists(cert_path):
                print(f"  [AVISO] Certificado não encontrado para CNPJ {cnpj_digits}: {cert_path}")
                continue

            print(f"  Tentando com certificado do CNPJ {cnpj_digits}...")
            try:
                sess = montar_sessao_pkcs12(cert_path, senha_cert)
                resultado = baixar_danfse_pdf(sess, chave)
            except Exception as e:
                print(f"  [ERRO] Falha usando certificado {cnpj_digits}: {e}")
                resultado = None

            if resultado:
                pdf_ok = True
                break  # já deu certo com um dos certificados

        if pdf_ok:
            total_ok += 1
        else:
            total_err += 1

        print(f"  Aguardando {DELAY_SECONDS} segundos para o próximo XML...")
        time.sleep(DELAY_SECONDS)

    print("\n=== Resumo ===")
    print(f"Sucessos: {total_ok}")
    print(f"Erros:    {total_err}")


if __name__ == "__main__":
    main()
