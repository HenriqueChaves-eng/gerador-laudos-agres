# Publicar e instalar a coleta offline no iPhone/iPad

## Arquitetura recomendada

- `docs/`: PWA de coleta para iPhone/iPad, instalada pela Tela de Início e preparada para uso sem internet.
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

## Instalar no iPhone ou iPad

1. Abra o link HTTPS do GitHub Pages no Safari.
2. Toque em `Compartilhar`.
3. Toque em `Adicionar à Tela de Início`.
4. Abra pelo ícone criado, não por uma aba antiga do navegador.
5. Autorize câmera, microfone e localização.
6. Faça um teste: coloque o aparelho em modo avião e abra pelo ícone.

## Sincronizar com o Streamlit

Quando voltar internet:

1. Na PWA, toque em `Exportar Pacote Relatorio Offline`.
2. Abra o app Streamlit online.
3. Selecione o pacote exportado. No iPhone/iPad ele pode sair como ZIP para aparecer melhor no WhatsApp, OneDrive e Arquivos.
4. Confira o pacote importado e clique em `Gerar relatório técnico agora`.

O ZIP final contem o arquivo Word e uma pasta unica `FOTOS DO ATENDIMENTO` com todas as fotos do atendimento.
