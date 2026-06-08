# Publicacao GitHub e Deploy

## 1) Preparar repositorio

No terminal, dentro da pasta do projeto:

```powershell
git init
git add .
git commit -m "feat: streamlit app and deploy files"
```

## 2) Criar repositorio remoto no GitHub

Crie um repositorio novo no GitHub e copie a URL.

```powershell
git branch -M main
git remote add origin <URL_DO_REPOSITORIO>
git push -u origin main
```

## 3) Deploy no Streamlit Community Cloud

1. Acesse https://share.streamlit.io/
2. Clique em New app
3. Selecione o repositorio e branch `main`
4. Main file path: `app.py`
5. Clique em Deploy

## 4) Pos-deploy

- Verifique logs no painel do Streamlit.
- Se houver erro de dependencia, confirme `requirements.txt`.

## Nota sobre dados sensiveis

- `certificados/`, `notasfiscais/` e `empresas.xlsx` estao ignorados no `.gitignore`.
- Para processamento real em cloud, configure estrategia segura para certificados e cadastro.
