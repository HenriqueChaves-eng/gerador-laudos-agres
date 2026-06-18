# Auditoria Final de Campo - Agres

Data: 18/06/2026

Versao offline: 2026.06.18.26  
Versao APK Android: 1.0.17 (versionCode 18)

## Escopo Revalidado

- Coleta offline PWA para iPhone e iPad.
- Coleta offline embarcada no APK Android.
- Persistencia local de textos, audios, fotos, legendas, localizacao e assinaturas.
- Exportacao do Pacote Relatorio Offline.
- Compartilhamento no iPhone/iPad como ZIP para melhorar compatibilidade com WhatsApp, OneDrive e Arquivos.
- Compartilhamento no iPhone/iPad somente com arquivo anexado, sem texto/titulo adicional que gere `.txt` separado.
- Importacao online de pacote JSON ou ZIP.
- Geracao do Word e do ZIP final com pasta unica de fotos.
- Atualizacao de cache offline para evitar versoes antigas no iPad/iPhone.

## Correcoes Desta Revisao

- A PWA passou a gerar um ZIP contendo o JSON original quando usar o compartilhamento/download web.
- O gerador online agora aceita `.json` e `.zip`.
- O ZIP importado e validado pelo gerador extrai somente o JSON interno da coleta offline.
- Foram adicionadas validacoes de tamanho e formato para ZIP corrompido, ZIP sem JSON e JSON interno grande demais.
- Textos do online foram ajustados de "JSON" para "pacote JSON ou ZIP" onde o tecnico interage com a tela.
- Removido texto/titulo do Web Share API no iPhone/iPad; o compartilhamento agora envia apenas `{ files: [...] }`.
- Cache da PWA atualizado:
  - GitHub Pages: `agres-pages-offline-v52`
  - Streamlit static offline: `agres-offline-v56`
  - Streamlit static coleta: `agres-coleta-v49`
- APK Android atualizado para `1.0.17`.
- Guia do iPad atualizado para explicar que o pacote pode sair como ZIP.

## Testes Executados Nesta Revalidacao

- PWA aberta por servidor local limpo em `http://127.0.0.1:8766/index.html`.
- Validado carregamento sem erros de console da PWA atual.
- Validada presenca do botao `Exportar Pacote Relatorio Offline`.
- Validada obrigatoriedade do `Tecnico Responsavel Agres` antes da exportacao.
- Validado fallback de exportacao gerando arquivo `.zip`:
  - Exemplo: `20260615_RELATORIO_ATIVIDADES_INSTALACAO_HENRIQUE.zip`
  - Link exibido: `baixar pacote ZIP`
- Validada ausencia da frase `Pacote JSON da coleta offline Agres.` nos arquivos ativos.
- Validado que nao ha `EXTRA_TEXT` no compartilhamento Android nem `text/title` no compartilhamento iPhone/iPad.
- Compilacao sintatica de `app.py` e `build_release_packages.py` concluida sem erros.
- Quatro copias do HTML offline sincronizadas com o mesmo hash:
  - `docs/index.html`
  - `static/offline/index.html`
  - `static/coleta/index.html`
  - `android-apk/app/src/main/assets/offline/index.html`
- Service workers atualizados.
- Pacotes finais recriados e testados por integridade ZIP.
- Conferido que `secrets.toml` nao entrou em nenhum pacote final.

## Pacotes de Entrega Gerados

- `arquivos_finais_github.zip`: projeto completo para GitHub/Streamlit.
- `agres_offline_android_project.zip`: projeto Android para gerar APK pelo GitHub Actions.
- `agres_offline_iphone_ipad_pwa.zip`: PWA iPhone/iPad para publicar pelo GitHub Pages.
- `MANUAL_UTILIZACAO_RELATORIOS_TECNICOS_AGRES_ABNT_1.0.17.docx`: manual de uso.

## Observacoes Importantes

- No iPhone/iPad, o menu de compartilhamento depende dos aplicativos instalados e das extensoes aceitas por cada app. Usar ZIP aumenta a chance de aparecer WhatsApp, OneDrive e Arquivos.
- O pacote ZIP gerado pelo iPad/iPhone pode ser importado diretamente no gerador online.
- O JSON original continua dentro do ZIP, preservando compatibilidade e rastreabilidade.
- Apos subir no GitHub, abrir a PWA uma vez com internet para atualizar o cache. Se o iPad ainda abrir a versao antiga, remover o icone antigo da Tela de Inicio e adicionar novamente pelo link publicado.

## Recomendacao

Liberar esta versao para piloto controlado com tecnicos em Android APK e iPhone/iPad PWA, mantendo a regra: nao limpar a coleta antes de confirmar que o pacote foi importado no online e que o Word final foi gerado.
