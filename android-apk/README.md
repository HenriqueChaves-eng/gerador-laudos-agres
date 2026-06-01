# Agres Offline APK

Projeto Android nativo que embala a coleta offline Agres em um WebView.

## Recursos incluídos

- Funciona sem internet, carregando `app/src/main/assets/offline/index.html`.
- Câmera e galeria via campo de upload do HTML.
- Microfone para gravação de áudio.
- Localização via geolocalização do Android/WebView.
- Salvamento automático das fotos tiradas pela câmera na galeria do Android, na pasta `Pictures/Agres Relatorios`.
- Exportação do pacote JSON pelo próprio app offline.

## Como gerar o APK pelo Android Studio

1. Instale o Android Studio.
2. Abra a pasta `android-apk`.
3. Aguarde o Gradle sincronizar o projeto.
4. Clique em `Build > Build Bundle(s) / APK(s) > Build APK(s)`.
5. O APK de debug será gerado em:

```text
android-apk/app/build/outputs/apk/debug/app-debug.apk
```

## Como instalar no Android para teste

1. Copie o `app-debug.apk` para o celular Android.
2. Abra o arquivo no celular.
3. Permita instalar apps de origem externa, se o Android solicitar.
4. Abra o app `Agres Offline`.

## Observações importantes

- Este projeto não depende do Streamlit para coletar offline.
- O online continua necessário para importar o JSON e gerar o relatório Word/ZIP.
- O salvamento direto na galeria é recurso nativo do APK Android. No iPhone/iPad, o navegador/PWA não permite salvar automaticamente na galeria sem ação do usuário.
