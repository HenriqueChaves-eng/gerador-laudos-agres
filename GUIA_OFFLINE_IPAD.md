# Coleta Offline no iPad

## Arquitetura correta

- `docs/`: PWA de coleta 100% offline para iPad.
- `app.py`: app Streamlit online para importar o pacote e gerar o relatório Word.

O Streamlit não é usado para preencher offline. Ele só gera o relatório quando a internet voltar.

## Publicar a PWA offline

1. Suba a pasta `docs/` para o GitHub junto com o projeto.
2. No GitHub, abra `Settings > Pages`.
3. Em `Build and deployment`, selecione:
   - Source: `Deploy from a branch`
   - Branch: `main`
   - Folder: `/docs`
4. Salve.
5. O GitHub vai gerar um link HTTPS parecido com:
   `https://usuario.github.io/repositorio/`

## Instalar no iPad

1. Abra o link HTTPS do GitHub Pages no Safari do iPad.
2. Toque em `Compartilhar`.
3. Toque em `Adicionar à Tela de Início`.
4. Abra pelo ícone criado.
5. Faça um teste: coloque o iPad em modo avião e abra pelo ícone.

## Sincronizar com o Streamlit

Quando voltar internet:

1. Na PWA, toque em `Copiar pacote para colar no app online`.
2. Abra o app Streamlit online.
3. Cole no campo `Ou cole aqui o pacote copiado no modo offline`.
4. Clique em `Importar pacote para este rascunho`.
5. Gere o relatório.

Também é possível usar `Compartilhar pacote JSON` ou `Exportar pacote offline`, mas copiar e colar costuma ser mais simples no iPad.
