# Gerar APK pelo GitHub

Este projeto já inclui o workflow `.github/workflows/build-android-apk.yml`.

## Como gerar o APK

1. Suba para o GitHub as pastas e arquivos atualizados:
   - `.github`
   - `android-apk`
   - `docs`
   - `static`
   - `app.py`
   - demais arquivos do projeto

2. No GitHub, abra o repositório.

3. Clique em `Actions`.

4. Clique em `Build Android APK`.

5. Clique em `Run workflow`.

6. Aguarde finalizar.

7. Abra a execução finalizada e baixe o artefato:

```text
agres-offline-debug-apk
```

Dentro dele estará:

```text
app-debug.apk
```

## Instalar no Android

1. Envie o `app-debug.apk` para o celular.
2. Abra o arquivo no Android.
3. Permita instalação de origem externa, se solicitado.
4. Abra o app `Agres Offline`.

## Observação

Esse APK é de teste/debug. Para distribuir oficialmente para todos os técnicos, o próximo passo é gerar um APK/AAB assinado com chave da empresa.
