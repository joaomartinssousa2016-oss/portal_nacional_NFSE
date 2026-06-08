import contextlib
import io
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from NovaBusca import main as worker_main

APP_DIR = Path(__file__).resolve().parent
EXCEL_PATH = APP_DIR / "empresas.xlsx"
OUTPUT_DIR = APP_DIR / "notasfiscais"
COLUMNS = ["Nome empresa", "cnpj", "senha", "data_inicio", "data_fim"]


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=COLUMNS)

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""

    out = df[COLUMNS].copy()
    out = out.fillna("").astype(str)
    return out


def load_excel() -> pd.DataFrame:
    if not EXCEL_PATH.exists():
        return pd.DataFrame(columns=COLUMNS)
    try:
        df = pd.read_excel(EXCEL_PATH, dtype=str)
        return normalize_df(df)
    except Exception as exc:
        st.error(f"Falha ao carregar Excel: {exc}")
        return pd.DataFrame(columns=COLUMNS)


def save_excel(df: pd.DataFrame) -> None:
    out = normalize_df(df)
    out.to_excel(EXCEL_PATH, index=False)


def validate_date(value: str) -> bool:
    value = str(value).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            datetime.strptime(value, fmt)
            return True
        except ValueError:
            continue
    return False


def validate_rows(df: pd.DataFrame) -> list[str]:
    errors = []
    for idx, row in df.iterrows():
        nome = str(row.get("Nome empresa", "")).strip()
        cnpj = str(row.get("cnpj", "")).strip()
        senha = str(row.get("senha", "")).strip()
        data_inicio = str(row.get("data_inicio", "")).strip()
        data_fim = str(row.get("data_fim", "")).strip()

        if not nome or not cnpj or not senha:
            errors.append(f"Linha {idx}: preencha Nome empresa, cnpj e senha.")

        if not validate_date(data_inicio):
            errors.append(f"Linha {idx}: data_inicio invalida ({data_inicio}).")

        if not validate_date(data_fim):
            errors.append(f"Linha {idx}: data_fim invalida ({data_fim}).")

    return errors


def run_worker(baixar_xml: bool, baixar_pdf: bool, pdf_tipo: str, selected_rows: list[int] | None) -> str:
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer), contextlib.redirect_stderr(buffer):
        try:
            worker_main(
                baixar_xml=baixar_xml,
                baixar_pdf=baixar_pdf,
                pdf_tipo=pdf_tipo,
                selected_rows=selected_rows,
            )
        except Exception as exc:
            print(f"[ERRO] Falha fatal na execucao: {exc}")
    return buffer.getvalue()


def build_report_dataframe() -> pd.DataFrame:
    if not OUTPUT_DIR.exists():
        return pd.DataFrame(columns=["empresa", "mes_ano", "tipo", "quantidade_xml"])

    records = []
    for path in OUTPUT_DIR.rglob("*.xml"):
        try:
            rel = path.relative_to(OUTPUT_DIR)
            parts = rel.parts
            if len(parts) < 4:
                continue
            records.append((parts[0], parts[1], parts[2], 1))
        except Exception:
            continue

    if not records:
        return pd.DataFrame(columns=["empresa", "mes_ano", "tipo", "quantidade_xml"])

    df = pd.DataFrame(records, columns=["empresa", "mes_ano", "tipo", "qtd"])
    return (
        df.groupby(["empresa", "mes_ano", "tipo"], as_index=False)["qtd"]
        .sum()
        .rename(columns={"qtd": "quantidade_xml"})
        .sort_values(["empresa", "mes_ano", "tipo"])
    )


def main() -> None:
    st.set_page_config(page_title="NovaBusca Streamlit", layout="wide")
    st.title("NovaBusca - Portal Nacional NFSe")
    st.caption("Cadastro em Excel e execucao da busca com log no navegador.")

    if "log" not in st.session_state:
        st.session_state.log = ""

    if "excel_df" not in st.session_state:
        st.session_state.excel_df = load_excel()

    st.subheader("Cadastro de empresas")
    edited_df = st.data_editor(
        st.session_state.excel_df,
        num_rows="dynamic",
        use_container_width=True,
        key="empresas_editor",
    )
    edited_df = normalize_df(edited_df)

    st.subheader("Configuracoes")
    cfg_col1, cfg_col2, cfg_col3, cfg_col4 = st.columns([1, 1, 1, 2])
    with cfg_col1:
        baixar_xml = st.checkbox("Baixar XML", value=True)
    with cfg_col2:
        baixar_pdf = st.checkbox("Baixar PDF", value=False)
    with cfg_col3:
        pdf_tipo = st.selectbox("Tipo PDF", ["todos", "emitidas", "recebidas"], index=0)
    with cfg_col4:
        somente_selecionadas = st.checkbox("Processar somente linhas selecionadas", value=False)

    selected_rows = None
    if somente_selecionadas and not edited_df.empty:
        options = list(range(len(edited_df)))
        labels = [f"{idx} - {edited_df.iloc[idx]['Nome empresa']}" for idx in options]
        selected_labels = st.multiselect("Linhas para processar", options=labels)
        selected_rows = [int(label.split(" - ", 1)[0]) for label in selected_labels]

    action_col1, action_col2, action_col3 = st.columns([1, 1, 2])

    with action_col1:
        if st.button("Salvar Excel", use_container_width=True):
            errors = validate_rows(edited_df)
            if errors:
                st.error("Corrija os dados antes de salvar.")
                for err in errors[:10]:
                    st.write(f"- {err}")
            else:
                save_excel(edited_df)
                st.session_state.excel_df = edited_df
                st.success(f"Excel salvo em: {EXCEL_PATH}")

    with action_col2:
        if st.button("Recarregar Excel", use_container_width=True):
            st.session_state.excel_df = load_excel()
            st.success("Excel recarregado.")
            st.rerun()

    with action_col3:
        if st.button("Rodar busca", type="primary", use_container_width=True):
            if edited_df.empty:
                st.warning("Adicione pelo menos uma empresa para executar.")
            else:
                errors = validate_rows(edited_df)
                if errors:
                    st.error("Existem dados invalidos. Corrija antes de executar.")
                    for err in errors[:10]:
                        st.write(f"- {err}")
                elif somente_selecionadas and not selected_rows:
                    st.warning("Selecione ao menos uma linha para processar.")
                else:
                    save_excel(edited_df)
                    st.session_state.excel_df = edited_df
                    with st.spinner("Executando busca... isso pode levar alguns minutos."):
                        output = run_worker(baixar_xml, baixar_pdf, pdf_tipo, selected_rows)
                    st.session_state.log = output
                    st.success("Processamento concluido.")

    st.subheader("Relatorio")
    if st.button("Gerar relatorio de XML"):
        resumo = build_report_dataframe()
        if resumo.empty:
            st.info("Nenhum XML encontrado para gerar relatorio.")
        else:
            st.dataframe(resumo, use_container_width=True)
            out_buffer = io.BytesIO()
            resumo.to_excel(out_buffer, index=False)
            out_buffer.seek(0)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="Baixar relatorio.xlsx",
                data=out_buffer,
                file_name=f"relatorio_execucao_{stamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    st.subheader("Log")
    st.text_area("Saida do processamento", value=st.session_state.log, height=300)


if __name__ == "__main__":
    main()
