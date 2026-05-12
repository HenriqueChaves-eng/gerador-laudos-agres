# Coleta offline no iPad

## Arquitetura recomendada

- `docs/`: PWA de coleta para iPad, instalada pela tela de inicio e preparada para uso sem internet.
- `app.py`: app Streamlit online usado apenas para importar o pacote offline, conferir os dados e gerar o Word + fotos em ZIP.

O preenchimento de campo deve ser feito pela PWA offline. O Streamlit fica como central de geracao do relatorio quando a internet voltar.

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
3. Toque em `Adicionar a Tela de Inicio`.
4. Abra pelo icone criado, nao pelo navegador.
5. Faca um teste: coloque o iPad em modo aviao e abra pelo icone.

## Sincronizar com o Streamlit

Quando voltar internet:

1. Na PWA, toque em `Copiar pacote para colar no app online` ou exporte o arquivo JSON.
2. Abra o app Streamlit online.
3. Use `Carregar arquivo JSON` ou `Colar pacote copiado`.
4. Clique em `Carregar arquivo JSON` ou `Importar texto copiado`.
5. Confira o pacote importado e clique em `Gerar relatorio tecnico`.

O ZIP final contem o arquivo Word e uma pasta unica `FOTOS DO ATENDIMENTO` com todas as fotos do atendimento.
