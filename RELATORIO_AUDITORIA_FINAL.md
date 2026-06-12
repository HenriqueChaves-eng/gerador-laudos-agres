# Auditoria Final de Campo - Agres

Data: 12/06/2026

Versão offline: 2026.06.12.22  
Versão APK Android: 1.0.12 (versionCode 13)

## Escopo validado

- Coleta offline no navegador/PWA para iPhone e iPad.
- Coleta offline no APK Android.
- Persistência local de textos, áudios, fotos, legendas, localização e assinaturas.
- Exportação do pacote JSON.
- Importação do JSON no gerador online.
- Processamento pela IA, geração do Word e pacote ZIP com fotos.
- Formatação, paginação, nomenclatura e integridade dos arquivos gerados.

## Correções críticas implementadas

- Gravações do rascunho offline passaram a ser serializadas para evitar operações concorrentes e perda de dados.
- Importação online sempre inicia um atendimento limpo, sem misturar arquivos do pacote anterior.
- Manifesto online passou a usar gravação atômica e arquivo de recuperação.
- Arquivos órfãos de importações anteriores são removidos com segurança.
- Pacotes Base64 corrompidos são rejeitados em vez de serem aceitos parcialmente.
- Foram adicionados limites de segurança para pacote, imagem, áudio e resolução.
- O cache offline só é atualizado quando todos os arquivos essenciais foram armazenados.
- Erros ao finalizar gravações de áudio agora são exibidos e tratados.
- O modo antigo de gravação WAV possui limite para evitar estouro de memória.
- A data usada nos nomes passa a ser a data final cronologicamente mais recente.
- A biblioteca Gemini obsoleta foi substituída por `google-genai`.
- Dependências foram fixadas em versões verificadas.
- Backup do APK foi desabilitado e permissões do WebView foram restringidas à origem local.
- Arquivos temporários antigos do APK são limpos e gravações incompletas no Android são removidas.

## Testes executados

- Compilação sintática do Python.
- Compatibilidade das dependências com `pip check`.
- Análise sintática dos arquivos Java do APK.
- Validação dos manifests JSON/XML e do modelo Word.
- Validação de IDs únicos e associação de labels no HTML.
- Teste de importação com mais de 30 fotos, áudio, cabeçalho e assinatura.
- Teste de troca de atendimento sem permanência de fotos anteriores.
- Teste de recuperação de manifesto corrompido.
- Teste de rejeição de Base64 inválido.
- Geração real de Word e ZIP, com verificação interna dos arquivos.
- Verificação de início de página para Fotos, Configurações, Outros Registros e Assinaturas.
- Verificação de blocos completos de figura com título, imagem, fonte e legenda.
- Testes responsivos em 360, 390, 820 e 1280 px, sem rolagem horizontal.
- Teste de persistência offline após recarregar a página.
- Inicialização do gerador Streamlit sem erros ou avisos de dependências obsoletas.

## Limites operacionais

- O pacote JSON possui limite de 100 MB; imagens individuais, 25 MB; áudios individuais, 100 MB.
- O gerador online e a IA dependem de internet e da disponibilidade da API Gemini.
- iPhone/iPad exigem que a versão HTTPS seja aberta uma vez com internet e adicionada à Tela de Início para uso offline confiável.
- Câmera, microfone e GPS dependem das permissões concedidas no aparelho.
- O APK gerado pelo workflow atual é de depuração. Para distribuição oficial, deve ser assinado com uma chave mantida pela empresa.

## Recomendação de liberação

Realizar um piloto curto com dois técnicos em aparelhos diferentes antes da entrega para toda a equipe. Após o piloto, distribuir exatamente o mesmo pacote aprovado, sem alterações intermediárias.
