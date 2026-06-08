import contextlib
import io
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from NovaBusca import main as worker_main

APP_DIR = Path(__file__).resolve().parent
EXCEL_PATH = APP_DIR / "empresas.xlsx"
OUTPUT_DIR = APP_DIR / "notasfiscais"
CERT_DIR = APP_DIR / "certificados"
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


def save_uploaded_certificates(uploaded_files) -> tuple[list[str], list[str]]:
    CERT_DIR.mkdir(parents=True, exist_ok=True)
    saved = []
    skipped = []

    for up in uploaded_files:
        name = (up.name or "").strip()
        lower = name.lower()
        if not (lower.endswith(".pfx") or lower.endswith(".p12")):
            skipped.append(name or "(sem nome)")
            continue

        target = CERT_DIR / name
        target.write_bytes(up.getvalue())
        saved.append(name)

    return saved, skipped


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


def build_notes_zip() -> tuple[io.BytesIO | None, int]:
    if not OUTPUT_DIR.exists():
        return None, 0

    files = [p for p in OUTPUT_DIR.rglob("*") if p.is_file()]
    if not files:
        return None, 0

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in files:
            arcname = file_path.relative_to(OUTPUT_DIR)
            zf.write(file_path, arcname=str(arcname))

    zip_buffer.seek(0)
    return zip_buffer, len(files)


def build_full_backup_zip(log_text: str) -> tuple[io.BytesIO, int]:
    backup_buffer = io.BytesIO()
    count = 0

    with zipfile.ZipFile(backup_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        if EXCEL_PATH.exists():
            zf.write(EXCEL_PATH, arcname="empresas.xlsx")
            count += 1

        if CERT_DIR.exists():
            for cert_file in CERT_DIR.rglob("*"):
                if not cert_file.is_file():
                    continue
                arcname = Path("certificados") / cert_file.relative_to(CERT_DIR)
                zf.write(cert_file, arcname=str(arcname))
                count += 1

        if OUTPUT_DIR.exists():
            for out_file in OUTPUT_DIR.rglob("*"):
                if not out_file.is_file():
                    continue
                arcname = Path("notasfiscais") / out_file.relative_to(OUTPUT_DIR)
                zf.write(out_file, arcname=str(arcname))
                count += 1

        if log_text.strip():
            zf.writestr("logs/ultimo_processamento.log", log_text.encode("utf-8", errors="ignore"))
            count += 1

    backup_buffer.seek(0)
    return backup_buffer, count


def restore_from_backup_zip(uploaded_zip) -> tuple[int, int]:
    restored = 0
    ignored = 0

    data = uploaded_zip.getvalue()
    with zipfile.ZipFile(io.BytesIO(data), mode="r") as zf:
        for name in zf.namelist():
            normalized = name.replace("\\", "/").strip("/")
            if not normalized or normalized.endswith("/"):
                continue

            if normalized == "empresas.xlsx":
                target = EXCEL_PATH
            elif normalized.startswith("certificados/"):
                rel = Path(normalized).relative_to("certificados")
                target = CERT_DIR / rel
            elif normalized.startswith("notasfiscais/"):
                rel = Path(normalized).relative_to("notasfiscais")
                target = OUTPUT_DIR / rel
            else:
                ignored += 1
                continue

            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(name, "r") as src, open(target, "wb") as dst:
                shutil.copyfileobj(src, dst)
            restored += 1

    return restored, ignored


def main() -> None:
    st.set_page_config(page_title="NovaBusca Streamlit", layout="wide")
    st.title("NovaBusca - Portal Nacional NFSe")
    st.caption("Cadastro em Excel e execucao da busca com log no navegador.")

    if "log" not in st.session_state:
        st.session_state.log = ""

    if "excel_df" not in st.session_state:
        st.session_state.excel_df = load_excel()

    with st.expander("Modo cloud (teste): upload de planilha e certificados"):
        st.info(
            "Use esta area quando o app estiver no Streamlit Cloud. "
            "Os arquivos podem ser perdidos em reinicio de servidor, pois e um modo de teste."
        )

        up_col1, up_col2 = st.columns(2)

        with up_col1:
            uploaded_excel = st.file_uploader(
                "Importar empresas.xlsx",
                type=["xlsx"],
                accept_multiple_files=False,
                key="upload_excel_cloud",
            )
            if st.button("Aplicar planilha enviada", use_container_width=True):
                if uploaded_excel is None:
                    st.warning("Selecione um arquivo .xlsx antes de aplicar.")
                else:
                    try:
                        imported_df = pd.read_excel(uploaded_excel, dtype=str)
                        imported_df = normalize_df(imported_df)
                        st.session_state.excel_df = imported_df
                        save_excel(imported_df)
                        st.success("Planilha carregada no app e salva no servidor.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Falha ao importar planilha: {exc}")

        with up_col2:
            uploaded_certs = st.file_uploader(
                "Enviar certificados (.pfx/.p12)",
                type=["pfx", "p12"],
                accept_multiple_files=True,
                key="upload_cert_cloud",
            )
            if st.button("Salvar certificados enviados", use_container_width=True):
                files = uploaded_certs or []
                if not files:
                    st.warning("Selecione ao menos um certificado para salvar.")
                else:
                    saved, skipped = save_uploaded_certificates(files)
                    if saved:
                        st.success(f"Certificados salvos: {', '.join(saved)}")
                    if skipped:
                        st.warning(f"Ignorados (extensao invalida): {', '.join(skipped)}")

        if CERT_DIR.exists():
            cert_files = sorted([p.name for p in CERT_DIR.glob("*.pfx")]) + sorted([p.name for p in CERT_DIR.glob("*.p12")])
            if cert_files:
                st.caption("Certificados atualmente disponiveis no servidor:")
                st.write("\n".join(f"- {name}" for name in cert_files))
            else:
                st.caption("Nenhum certificado carregado no servidor.")

        template_df = pd.DataFrame(columns=COLUMNS)
        template_buffer = io.BytesIO()
        template_df.to_excel(template_buffer, index=False)
        template_buffer.seek(0)
        st.download_button(
            label="Baixar modelo empresas.xlsx",
            data=template_buffer,
            file_name="empresas_modelo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.divider()
        st.caption("Persistencia sem banco externo: use backup ZIP para salvar e restaurar tudo no cloud.")

        bk_col1, bk_col2 = st.columns(2)
        with bk_col1:
            backup_data, backup_count = build_full_backup_zip(st.session_state.log)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label=f"Baixar backup completo ({backup_count} arquivo(s))",
                data=backup_data,
                file_name=f"novabusca_backup_{stamp}.zip",
                mime="application/zip",
            )

        with bk_col2:
            uploaded_backup = st.file_uploader(
                "Restaurar backup completo (.zip)",
                type=["zip"],
                accept_multiple_files=False,
                key="upload_backup_cloud",
            )
            if st.button("Aplicar backup", use_container_width=True):
                if uploaded_backup is None:
                    st.warning("Selecione um arquivo .zip de backup antes de aplicar.")
                else:
                    try:
                        restored, ignored = restore_from_backup_zip(uploaded_backup)
                        st.session_state.excel_df = load_excel()
                        st.success(
                            f"Backup aplicado com sucesso. Restaurados: {restored}. Ignorados: {ignored}."
                        )
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Falha ao restaurar backup: {exc}")

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

    st.subheader("Exportar notas")
    zip_data, total_files = build_notes_zip()
    if zip_data is None:
        st.info("Nenhum arquivo encontrado em notasfiscais para exportar.")
    else:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.caption(f"Arquivos prontos para exportacao: {total_files}")
        st.download_button(
            label="Baixar notasfiscais.zip",
            data=zip_data,
            file_name=f"notasfiscais_{stamp}.zip",
            mime="application/zip",
        )

    st.subheader("Log")
    st.text_area("Saida do processamento", value=st.session_state.log, height=300)


if __name__ == "__main__":
    main()
