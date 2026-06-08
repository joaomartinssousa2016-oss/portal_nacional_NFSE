# 🚀 NovaBusca - Guia de Instalação Rápida

## ⚠️ IMPORTANTE - Primeira Execução

Antes de usar a aplicação pela primeira vez, você **DEVE** instalar as dependências Python.

### Opção 1: Instalação Automática ✨ (Recomendado)

1. **Clique 2x em:** `Instalar_Dependencias_Automatico.bat`
2. Aguarde a instalação terminar
3. Pronto! Agora pode usar a aplicação

### Opção 2: Instalação Manual

Abra PowerShell ou CMD na pasta do projeto e execute:

```bash
pip install -r requirements.txt
```

---

## ▶️ Como Usar

Após instalar as dependências, **clique 2x em:**

```
Abrir_NovaBusca_UI.bat
```

Isso abrirá a interface gráfica para cadastrar empresas e baixar NFSe em XML ou PDF.

---

## 📋 Requisitos

- **Python 3.10+** instalado e no PATH
- Conexão com internet
- Certificados PKCS#12 na pasta `certificados/`

---

## 🎯 Fluxo de Uso

1. **Abra:** `Abrir_NovaBusca_UI.bat`
2. **Cadastre empresas:** Nome, CNPJ, Senha, datas
3. **Escolha formato:** ☑️ XML | ☑️ PDF
4. **Clique:** "Rodar busca"
5. **Resultado:** Documentos em `notasfiscais/`

---

## 📦 O que foi instalado

- `pandas` - Leitura de Excel
- `requests` - Requisições HTTP
- `requests-pkcs12` - Autenticação com certificado
- `openpyxl` - Manipulação de Excel

---

## ❓ Problemas?

**"ModuleNotFoundError: No module named..."**
→ Rode novamente: `Instalar_Dependencias_Automatico.bat`

**"Certificado não encontrado"**
→ Coloque o certificado (CNPJ.pfx) na pasta `certificados/`

**Interface não abre**
→ Verifique se Python está instalado: `python --version`

---

**Versão:** 2.0 com suporte a XML e PDF
**Data:** 2026-06-03
