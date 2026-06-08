# NovaBusca NFSe (Streamlit)

Aplicacao para cadastro de empresas e processamento de NFSe do Portal Nacional, com interface web via Streamlit.

## Funcionalidades

- Cadastro de empresas em tabela (nome, CNPJ, senha, periodo)
- Salvamento e recarga de planilha de empresas
- Execucao da busca com XML e PDF (filtro: todos, emitidas, recebidas)
- Relatorio de XML por empresa/mes/tipo
- Log de processamento na interface

## Estrutura principal

- `app.py` - entrada padrao para deploy Streamlit
- `NovaBuscaStreamlit.py` - interface web
- `NovaBusca.py` - motor de processamento
- `requirements.txt` - dependencias Python

## Requisitos

- Python 3.10+
- Certificados PKCS#12 na pasta `certificados` (nome: CNPJ.pfx ou CNPJ.p12)

## Execucao local

1. Instalar dependencias:

```bat
Instalar_Dependencias.bat
```

2. Rodar Streamlit:

```bat
Iniciar_NovaBusca_Streamlit.bat
```

## Deploy no Streamlit Community Cloud

1. Suba este projeto no GitHub.
2. No Streamlit Cloud, clique em New app.
3. Selecione o repositorio e branch.
4. Defina Main file path como `app.py`.
5. Clique em Deploy.

## Seguranca

- Nao subir para GitHub: certificados, notasfiscais e empresas.xlsx.
- Estes itens ja estao no `.gitignore`.

## Observacao importante

Para execucao real da busca em ambiente cloud, voce precisara disponibilizar certificados e planilha de forma segura no ambiente de deploy (ex.: secrets/volumes privados). Em ambiente local, basta manter os arquivos nas pastas esperadas.
