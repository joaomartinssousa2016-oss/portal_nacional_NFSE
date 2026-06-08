import os
import re
import json
import base64
import gzip
import time
from io import BytesIO
from datetime import datetime

import pandas as pd  # precisa ser instalado via pip
import requests  # precisa ser instalado via pip
from requests_pkcs12 import Pkcs12Adapter  # precisa ser instalado via pip
from requests.exceptions import RequestException, Timeout
import xml.etree.ElementTree as ET
from pathlib import Path

BASE_URL = "https://adn.nfse.gov.br"
DFe_URL_TEMPLATE = BASE_URL + "/contribuintes/DFe/{nsu}"
DANFSE_URL_TEMPLATE = BASE_URL + "/danfse/{chave}"

EXCEL_PATH = "empresas.xlsx"  # <-- ajuste o caminho do seu Excel aqui
CERT_DIR = "certificados"
OUTPUT_BASE_DIR = "notasfiscais"

# Configurações para repetição automática
MAX_RODADAS_POR_EMPRESA = 1000         # máximo de vezes que vamos "insistir" na empresa
INTERVALO_SEGUNDOS_ENTRE_RODADAS = 5  # tempo entre rodadas, em segundos
INTERVALO_SEGUNDOS_ENTRE_PDFS = 1.2
TIMEOUT_REQUISICAO_SEGUNDOS = 120
MAX_TENTATIVAS_REQUISICAO = 4
BACKOFF_BASE_SEGUNDOS = 2
HTTP_STATUS_RETRY = {429, 500, 502, 503, 504}


def only_digits(s: str) -> str:
    if not s:
        return ""
    return re.sub(r"\D", "", s)


def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def decode_arquivo_xml(arquivo_xml_b64_gzip: str) -> str:
    """
    Decodifica campo ArquivoXml (gzip + base64) para string XML.
    """
    raw = base64.b64decode(arquivo_xml_b64_gzip)
    xml_bytes = gzip.decompress(raw)
    return xml_bytes.decode("utf-8", errors="ignore")


def get_com_retry(sess: requests.Session, url: str, timeout: int = TIMEOUT_REQUISICAO_SEGUNDOS):
    """
    Faz GET com tentativas e backoff exponencial para erros transitórios de rede/timeout.
    """
    ultima_exc = None
    for tentativa in range(1, MAX_TENTATIVAS_REQUISICAO + 1):
        try:
            return sess.get(url, timeout=timeout)
        except (Timeout, RequestException) as e:
            ultima_exc = e
            if tentativa == MAX_TENTATIVAS_REQUISICAO:
                break
            espera = BACKOFF_BASE_SEGUNDOS ** (tentativa - 1)
            print(
                f"    [AVISO] Falha de rede (tentativa {tentativa}/{MAX_TENTATIVAS_REQUISICAO}): {e}. "
                f"Nova tentativa em {espera}s..."
            )
            time.sleep(espera)
    raise ultima_exc


def extrair_cnpj_prestador(xml_str: str) -> str | None:
    """
    Extrai o CNPJ do emitente da NFSe.
    Considera apenas o padrão:
        <emit>
            <CNPJ>...</CNPJ>
        </emit>

    Se não existir <emit><CNPJ>, retorna None.
    """
    try:
        root = ET.fromstring(xml_str)
    except ET.ParseError:
        return None

    def local_name(tag):
        return tag.split("}")[-1].lower()

    # Busca somente: <emit><CNPJ>...</CNPJ></emit>
    for elem in root.iter():
        if local_name(elem.tag) == "emit":
            for child in elem:
                if local_name(child.tag) == "cnpj" and child.text:
                    return only_digits(child.text)

    return None


def extrair_chave_acesso_xml(xml_str: str) -> str | None:
    """
    Extrai a chave de acesso do XML (do atributo Id de <infNFSe>).
    """
    try:
        root = ET.fromstring(xml_str)
    except ET.ParseError:
        return None

    def local_name(tag):
        return tag.split("}")[-1].lower()

    # Busca infNFSe e pega o atributo Id
    for elem in root.iter():
        if local_name(elem.tag) == "infnfse":
            id_attr = elem.get("Id", "")
            chave = only_digits(id_attr)
            if chave:
                return chave

    # Fallback: procura por <nDFSe>
    for elem in root.iter():
        if local_name(elem.tag) == "ndfse" and elem.text:
            return only_digits(elem.text)

    return None


def baixar_danfse_pdf(sess: requests.Session, chave: str, arquivo_path: str) -> bool:
    """
    Baixa o PDF via API /danfse/{chave} e salva no arquivo_path.
    Retorna True se sucesso, False em caso de erro.
    """
    chave = only_digits(chave)
    if not chave:
        print("    [AVISO] Chave vazia/inválida, pulando PDF...")
        return False

    url = DANFSE_URL_TEMPLATE.format(chave=chave)

    resp = None
    for tentativa in range(1, MAX_TENTATIVAS_REQUISICAO + 1):
        try:
            resp = get_com_retry(sess, url)
        except Exception as e:
            print(f"    [ERRO] Falha ao baixar PDF (chave={chave}, tentativa {tentativa}): {e}")
            if tentativa < MAX_TENTATIVAS_REQUISICAO:
                espera = BACKOFF_BASE_SEGUNDOS ** (tentativa - 1)
                time.sleep(espera)
            continue

        if resp.status_code == 200:
            break

        if resp.status_code in HTTP_STATUS_RETRY and tentativa < MAX_TENTATIVAS_REQUISICAO:
            retry_after = resp.headers.get("Retry-After", "")
            if retry_after.isdigit():
                espera = int(retry_after)
            else:
                espera = BACKOFF_BASE_SEGUNDOS ** tentativa
            print(
                f"    [AVISO] HTTP {resp.status_code} ao baixar PDF para chave {chave} "
                f"(tentativa {tentativa}/{MAX_TENTATIVAS_REQUISICAO}). Nova tentativa em {espera}s..."
            )
            time.sleep(espera)
            continue

        print(f"    [AVISO] HTTP {resp.status_code} ao baixar PDF para chave {chave}")
        return False

    if resp is None or resp.status_code != 200:
        print(f"    [ERRO] Não foi possível baixar PDF da chave {chave} após tentativas.")
        return False

    try:
        with open(arquivo_path, "wb") as f:
            f.write(resp.content)
        print(f"    [OK] PDF salvo: {arquivo_path}")
        return True
    except Exception as e:
        print(f"    [ERRO] Falha ao salvar PDF em {arquivo_path}: {e}")
        return False


def format_mes_ano(data_str: str) -> str:
    """
    Recebe algo como '2023-09-27T08:28:28.377'
    e devolve '2023-09' para usar na pasta MesAno.
    """
    try:
        # Tenta parsear com microssegundos
        dt = datetime.fromisoformat(data_str)
    except ValueError:
        # fallback: pega só os 10 primeiros caracteres (YYYY-MM-DD)
        try:
            dt = datetime.strptime(data_str[:10], "%Y-%m-%d")
        except Exception:
            return "desconhecido"
    return f"{dt.year:04d}-{dt.month:02d}"


def sanitize_folder_name(name: str) -> str:
    """
    Limpa nome de empresa para uso em path de pasta.
    """
    name = name.strip()
    # troca qualquer caractere estranho por underscore
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = re.sub(r"\s+", " ", name)
    return name


def parse_data_hora(data_str: str):
    """
    Parseia string de data/hora retornada pela API (ex: '2023-09-27T08:28:28.377').
    Retorna datetime ou None se não conseguir parsear.
    """
    if not data_str:
        return None
    try:
        return datetime.fromisoformat(data_str)
    except ValueError:
        try:
            return datetime.strptime(data_str[:10], "%Y-%m-%d")
        except Exception:
            return None


def parse_data_usuario(data_val) -> datetime:
    """
    Parseia data informada pelo usuário no Excel.
    Aceita: datetime, DD/MM/YYYY, YYYY-MM-DD, DD-MM-YYYY.
    """
    if isinstance(data_val, datetime):
        return data_val
    s = str(data_val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    raise ValueError(f"Formato de data não reconhecido: '{s}'. Use DD/MM/YYYY ou YYYY-MM-DD.")


def _parse_selected_rows_env(value):
    if not value:
        return None

    selected_rows = set()
    for part in str(value).split(","):
        part = part.strip()
        if not part:
            continue
        try:
            selected_rows.add(int(part))
        except ValueError:
            continue

    return selected_rows


def encontrar_certificado_pkcs12(cnpj: str) -> str:
    """
    Procura um certificado PKCS#12 para o CNPJ dentro de CERT_DIR.
    Aceita .pfx e .p12 (case-insensitive).
    Prioriza .pfx se ambos existirem.
    """
    base = Path(CERT_DIR) / cnpj

    candidatos = [
        base.with_suffix(".pfx"),
        base.with_suffix(".p12"),
        base.with_suffix(".PFX"),
        base.with_suffix(".P12"),
    ]

    for p in candidatos:
        if p.exists():
            return str(p)

    # fallback: caso o arquivo tenha nome diferente, mas contenha o CNPJ no nome
    # (opcional — comente se você só usa {cnpj}.ext)
    fallback = list(Path(CERT_DIR).glob(f"*{cnpj}*.p12")) + list(Path(CERT_DIR).glob(f"*{cnpj}*.pfx"))
    if fallback:
        return str(fallback[0])

    raise FileNotFoundError(
        f"Certificado não encontrado para CNPJ {cnpj}. "
        f"Esperado: {base.with_suffix('.pfx')} ou {base.with_suffix('.p12')}"
    )


def montar_sessao_pkcs12(cnpj: str, senha: str | None) -> requests.Session:
    """
    Cria sessão Requests com certificado PKCS#12 (.pfx ou .p12) via Pkcs12Adapter.
    """
    cert_path = encontrar_certificado_pkcs12(cnpj)

    # algumas empresas podem ter senha vazia; normalize para string ou None
    if senha is not None:
        senha = str(senha)
        if senha.strip().lower() == "nan":
            senha = ""
    else:
        senha = ""

    s = requests.Session()
    s.mount(
        BASE_URL,
        Pkcs12Adapter(
            pkcs12_filename=cert_path,
            pkcs12_password=senha
        )
    )
    return s


def processar_empresa(
    row,
    data_inicio: datetime,
    data_fim: datetime,
    baixar_xml: bool = True,
    baixar_pdf: bool = False,
    pdf_tipo: str = "todos",
) -> None:
    """
    Processa uma empresa (linha do DataFrame), baixando NFSe dentro do período
    [data_inicio, data_fim].

    Pagina pela API usando NSU (começa em 0) até esgotar os documentos.
    Documentos fora do intervalo de datas são ignorados (não salvos).
    
    Se baixar_pdf=True, também baixa o PDF de cada NFSe.
    """
    nome_empresa = str(row["Nome empresa"])
    cnpj_excel = only_digits(str(row["cnpj"]))
    senha = str(row["senha"])

    # data_fim cobre até o último instante do dia
    data_fim_eod = data_fim.replace(hour=23, minute=59, second=59)

    print(f"\n=== Empresa: {nome_empresa} | CNPJ: {cnpj_excel} | Período: {data_inicio.date()} a {data_fim.date()} ===")

    sess = montar_sessao_pkcs12(cnpj_excel, senha)
    empresa_folder_name = sanitize_folder_name(nome_empresa)
    nsu_atual = 0
    ultimo_nsu = 0

    for rodada in range(1, MAX_RODADAS_POR_EMPRESA + 1):
        url = DFe_URL_TEMPLATE.format(nsu=nsu_atual)
        print(f"  Rodada {rodada} | NSU={nsu_atual} -> {url}")

        try:
            resp = get_com_retry(sess, url)
        except Exception as e:
            print(f"  [ERRO] Falha de conexão após tentativas para NSU={nsu_atual}: {e}")
            print("  [ERRO] Encerrando empresa por instabilidade de rede/portal.")
            break

        if resp.status_code != 200:
            print(f"  [ERRO] HTTP {resp.status_code} para NSU={nsu_atual}. Encerrando empresa.")
            break

        try:
            data = resp.json()
        except json.JSONDecodeError:
            print(f"  [ERRO] Resposta não é JSON para NSU={nsu_atual}. Encerrando empresa.")
            break

        status = data.get("StatusProcessamento", "")
        lote = data.get("LoteDFe") or []

        if status != "DOCUMENTOS_LOCALIZADOS" or not lote:
            print(f"  StatusProcessamento='{status}'. Nenhum documento localizado. Fim da empresa.")
            break

        # Exibe faixa de datas do lote para acompanhamento
        datas_lote = [
            parse_data_hora(d.get("DataHoraGeracao", ""))
            for d in lote
            if d.get("DataHoraGeracao")
        ]
        datas_validas = [dt for dt in datas_lote if dt is not None]
        if datas_validas:
            dt_min = min(datas_validas).strftime("%d/%m/%Y")
            dt_max = max(datas_validas).strftime("%d/%m/%Y")
            print(f"  {len(lote)} documento(s) retornado(s). Datas do lote: {dt_min} a {dt_max}")
        else:
            print(f"  {len(lote)} documento(s) retornado(s).")

        # Se o documento mais recente do lote ainda é anterior ao início, pula tudo
        if datas_validas and max(datas_validas) < data_inicio:
            print(f"  Lote ainda antes do período desejado ({data_inicio.date()}). Avançando...")
            # Atualiza ultimo_nsu para continuar paginando
            for doc in lote:
                try:
                    doc_nsu = int(doc.get("NSU", 0))
                except (TypeError, ValueError):
                    doc_nsu = 0
                if doc_nsu > ultimo_nsu:
                    ultimo_nsu = doc_nsu
            nsu_atual = ultimo_nsu
            time.sleep(1)  # pausa curta ao pular lotes históricos
            continue

        passou_data_fim = False

        for doc in lote:
            try:
                doc_nsu = int(doc.get("NSU", 0))
            except (TypeError, ValueError):
                doc_nsu = 0

            if doc_nsu > ultimo_nsu:
                ultimo_nsu = doc_nsu

            data_hora_geracao = doc.get("DataHoraGeracao", "")
            doc_dt = parse_data_hora(data_hora_geracao)

            # Filtra pelo período informado pelo usuário
            if doc_dt is not None:
                if doc_dt > data_fim_eod:
                    passou_data_fim = True
                    continue  # além do fim, não salva
                if doc_dt < data_inicio:
                    continue  # antes do início, não salva

            chave_acesso = doc.get("ChaveAcesso", "").strip()
            arquivo_xml_b64_gzip = doc.get("ArquivoXml", "")

            if not chave_acesso:
                continue

            xml_str = ""
            cnpj_prestador_digits = ""
            if arquivo_xml_b64_gzip:
                try:
                    xml_str = decode_arquivo_xml(arquivo_xml_b64_gzip)
                except Exception as e:
                    print(f"    [ERRO] Falha ao decodificar ArquivoXml (NSU={doc_nsu}, Chave={chave_acesso}): {e}")

            if xml_str:
                cnpj_prestador = extrair_cnpj_prestador(xml_str)
                cnpj_prestador_digits = only_digits(cnpj_prestador) if cnpj_prestador else ""

            if cnpj_prestador_digits:
                if cnpj_prestador_digits == cnpj_excel:
                    tipo_pasta = "prestados"
                else:
                    tipo_pasta = "tomados"
            else:
                tipo_pasta = "desconhecido"

            mes_ano = format_mes_ano(data_hora_geracao)
            empresa_dir = os.path.join(
                OUTPUT_BASE_DIR,
                empresa_folder_name,
                mes_ano,
                tipo_pasta
            )
            ensure_dir(empresa_dir)

            if baixar_xml and xml_str:
                file_path = os.path.join(empresa_dir, f"{chave_acesso}.xml")
                try:
                    with open(file_path, "w", encoding="utf-8") as f:
                        f.write(xml_str)
                    print(f"    [OK] Salvo XML: {file_path} ({tipo_pasta}, NSU={doc_nsu})")
                except Exception as e:
                    print(f"    [ERRO] Falha ao salvar XML em {file_path}: {e}")
            elif baixar_xml:
                print(f"    [AVISO] XML ausente para NSU={doc_nsu}, chave={chave_acesso}.")

            baixar_pdf_doc = False
            if baixar_pdf:
                if pdf_tipo == "todos":
                    baixar_pdf_doc = True
                elif pdf_tipo == "emitidas" and tipo_pasta == "prestados":
                    baixar_pdf_doc = True
                elif pdf_tipo == "recebidas" and tipo_pasta == "tomados":
                    baixar_pdf_doc = True

            # Se solicitado, baixa o PDF conforme filtro selecionado
            if baixar_pdf_doc:
                pdf_path = os.path.join(empresa_dir, f"{chave_acesso}.pdf")
                baixar_danfse_pdf(sess, chave_acesso, pdf_path)
                time.sleep(INTERVALO_SEGUNDOS_ENTRE_PDFS)

        if passou_data_fim:
            print(f"  Documentos ultrapassaram data_fim ({data_fim.date()}). Encerrando empresa.")
            break

        if ultimo_nsu == nsu_atual:
            print("  Nenhum NSU novo encontrado. Fim da empresa.")
            break

        nsu_atual = ultimo_nsu
        time.sleep(INTERVALO_SEGUNDOS_ENTRE_RODADAS)

    print(f"=== Fim da empresa {nome_empresa}. ===")


def main(
    baixar_xml: bool = True,
    baixar_pdf: bool = False,
    pdf_tipo: str = "todos",
    selected_rows=None,
):
    pdf_tipo = (pdf_tipo or "todos").strip().lower()
    if pdf_tipo not in {"todos", "emitidas", "recebidas"}:
        print(f"[AVISO] Tipo de PDF inválido '{pdf_tipo}'. Usando 'todos'.")
        pdf_tipo = "todos"

    if selected_rows is None:
        selected_rows = _parse_selected_rows_env(os.environ.get("NOVA_BUSCA_LINHAS_SELECIONADAS"))
    else:
        selected_rows = {int(idx) for idx in selected_rows}

    print(
        f"[INFO] Modos selecionados: baixar_xml={baixar_xml} | "
        f"baixar_pdf={baixar_pdf} | pdf_tipo={pdf_tipo}"
    )

    if not os.path.exists(EXCEL_PATH):
        print(f"Arquivo Excel não encontrado: {EXCEL_PATH}")
        return

    df = pd.read_excel(EXCEL_PATH, dtype={"cnpj": str})

    required_cols = {"Nome empresa", "cnpj", "senha", "data_inicio", "data_fim"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Colunas ausentes no Excel: {missing}")

    if selected_rows is not None:
        selected_rows = {idx for idx in selected_rows if idx >= 0}
        valid_rows = [idx for idx in sorted(selected_rows) if idx < len(df)]
        if not valid_rows:
            print("[AVISO] Nenhuma linha selecionada para processamento. Encerrando.")
            return
        if len(valid_rows) != len(selected_rows):
            print("[AVISO] Algumas linhas selecionadas estavam fora do intervalo e foram ignoradas.")
        print(f"[INFO] Processando somente as linhas selecionadas: {valid_rows}")
        df = df.iloc[valid_rows]

    for idx, row in df.iterrows():
        print("\n" + "#" * 80)
        print(f"Processando linha {idx} - empresa: {row['Nome empresa']}")
        print("#" * 80)

        try:
            data_inicio = parse_data_usuario(row["data_inicio"])
            data_fim = parse_data_usuario(row["data_fim"])
        except ValueError as e:
            print(f"[ERRO] Data inválida na linha {idx}: {e}. Pulando empresa.")
            continue

        try:
            processar_empresa(
                row,
                data_inicio,
                data_fim,
                baixar_xml=baixar_xml,
                baixar_pdf=baixar_pdf,
                pdf_tipo=pdf_tipo,
            )
        except Exception as e:
            print(f"\n[ERRO] Falha ao processar empresa na linha {idx}: {e}\n")

    print("\nProcessamento concluído.")


if __name__ == "__main__":
    main()
