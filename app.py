import streamlit as st
import google.generativeai as genai
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import json
import os
from datetime import date

# ==========================================
# 1. Configurações Iniciais e Layout
# ==========================================
st.set_page_config(page_title="Agres | Laudos Técnicos", page_icon="🚜", layout="centered")

# Variável de memória para permitir múltiplos gravadores na tela
if 'qtd_gravacoes' not in st.session_state:
    st.session_state.qtd_gravacoes = 1

st.markdown("""
    <style>
        #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
        div.stButton > button:first-child {
            background-color: #0b4f2c; color: white; border-radius: 8px; height: 55px;
            font-size: 18px; font-weight: bold; border: none; width: 100%; transition: all 0.3s ease;
        }
        div.stButton > button:first-child:hover { background-color: #157343; transform: translateY(-2px); }
        .titulo-app { text-align: center; color: #1e293b; font-weight: 800; margin-bottom: 5px; }
        .subtitulo-app { text-align: center; color: #64748b; font-size: 16px; margin-bottom: 30px; }
        .btn-adicionar > div > button { height: 40px !important; background-color: #475569 !important; font-size: 14px !important;}
        .btn-adicionar > div > button:hover { background-color: #334155 !important;}
    </style>
""", unsafe_allow_html=True)

GOOGLE_API_KEY = "AIzaSyBr7Et5pOlGuFxzQWr8xfaIk0nvpexiB2I"
genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel('models/gemini-2.0-flash-lite')

# ==========================================
# 2. Funções de Processamento
# ==========================================
def processar_atendimento_completo(arquivos_audio_temp):
    materiais_para_ia = []
    arquivos_api = [] 
    
    for audio in arquivos_audio_temp:
        temp_file = genai.upload_file(path=audio)
        materiais_para_ia.append(temp_file)
        arquivos_api.append(temp_file)
        
    prompt = f"""
    Analise os áudios de atendimento anexados e extraia os dados para um relatório técnico da Agres.
    
    DIRETRIZES:
    1. Gramática: Corrija para o português formal técnico (Norma Culta).
    2. Relato: Una todos os áudios em um laudo coeso e profissional. Evite repetições.
    3. Fotos: Extraia as legendas que o técnico mencionar no áudio.
    
    Retorne APENAS um JSON válido com estas exatas chaves:
    {{
        "suporte": "", "instalacao": "", "treinamento": "",
        "data_visita": "{date.today().strftime('%d/%m/%Y')}",
        "tecnicos": "Henrique Chaves",
        "cliente_local": "", "equipamentos": "", "maquinas": "",
        "objetivos": "", "configuracoes": "", "calibracoes": "",
        "acompanhantes": "",
        "relato": "Seu texto completo aqui.",
        "legendas_fotos": ["legenda 1", "legenda 2"]
    }}
    """
    
    try:
        resposta = model.generate_content([prompt] + materiais_para_ia)
        texto_json = resposta.text.strip().replace('```json', '').replace('```', '')
        dados = json.loads(texto_json)
    except Exception as e:
        raise Exception(f"Falha de comunicação com a IA. Erro: {str(e)}")
    finally:
        for f in arquivos_api:
            try: genai.delete_file(f.name)
            except: pass
                
    return dados

def gerar_docx(dados_json, caminhos_fotos_temp):
    doc = DocxTemplate("modelo_tags.docx")
    lista_fotos = []
    legendas = dados_json.get("legendas_fotos", [])
    
    for i, foto_path in enumerate(caminhos_fotos_temp):
        desc = legendas[i] if i < len(legendas) else "Registro fotográfico da manutenção"
        titulo = f"Figura {i+1} – {desc.capitalize()}."
        fonte = f"Fonte: O autor ({date.today().year})."
        imagem_word = InlineImage(doc, foto_path, width=Mm(150))
        lista_fotos.append({"titulo": titulo, "imagem": imagem_word, "fonte": fonte})
    
    dados_json['lista_fotos'] = lista_fotos
    doc.render(dados_json)
    
    nome_arquivo = f"Relatorio_{dados_json.get('cliente_local', 'Atendimento').replace(' ', '_')}.docx"
    doc.save(nome_arquivo)
    return nome_arquivo

# ==========================================
# 3. Interface Front-End
# ==========================================
st.markdown("<h1 class='titulo-app'>🚜 Agres Reports</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitulo-app'>Geração automatizada de laudos técnicos por IA</p>", unsafe_allow_html=True)

with st.container(border=True):
    st.markdown("### 🎙️ 1. Relato Técnico")
    aba1, aba2 = st.tabs(["🔴 Gravar agora", "📁 Arquivos do celular"])
    
    # --- NOVA LÓGICA DE MÚLTIPLOS ÁUDIOS ---
    with aba1:
        st.info("Grave o relato em etapas. Clique em '+' para abrir um novo bloco de gravação.")
        audios_gravados_na_hora = []
        
        for i in range(st.session_state.qtd_gravacoes):
            audio = st.audio_input(f"Trecho {i+1}", key=f"rec_{i}")
            if audio:
                audios_gravados_na_hora.append(audio)
                
        st.markdown("<div class='btn-adicionar'>", unsafe_allow_html=True)
        if st.button("➕ Adicionar novo trecho de áudio", use_container_width=True):
            st.session_state.qtd_gravacoes += 1
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
        
    with aba2:
        audios_upload = st.file_uploader("Selecione os áudios", type=['wav', 'mp3', 'm4a', 'ogg', 'aac'], accept_multiple_files=True)

with st.container(border=True):
    st.markdown("### 📸 2. Evidências Fotográficas")
    fotos = st.file_uploader("Anexe na mesma ordem do áudio", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'], label_visibility="collapsed")

# Agrupa todos os áudios gravados e enviados
audios_finais = []
if audios_gravados_na_hora: audios_finais.extend(audios_gravados_na_hora)
if audios_upload: audios_finais.extend(audios_upload)

if audios_finais and st.button("🚀 Processar Laudo Inteligente"):
    caminhos_audio_temp = []
    caminhos_foto_temp = []
    arquivo_final = None
    
    try:
        with st.status("Iniciando o processamento do atendimento...", expanded=True) as status:
            st.write("⚙️ Organizando arquivos...")
            for i, a in enumerate(audios_finais):
                extensao = a.name.split('.')[-1] if hasattr(a, 'name') and '.' in a.name else 'wav'
                path_audio = f"temp_audio_{i}.{extensao}"
                with open(path_audio, "wb") as f: f.write(a.getvalue())
                caminhos_audio_temp.append(path_audio)
            
            for i, p in enumerate(fotos):
                extensao = p.name.split('.')[-1] if hasattr(p, 'name') and '.' in p.name else 'jpg'
                path_foto = f"temp_foto_{i}.{extensao}"
                with open(path_foto, "wb") as f: f.write(p.getvalue())
                caminhos_foto_temp.append(path_foto)
                
            st.write("🤖 IA transcrevendo e unificando relato...")
            dados = processar_atendimento_completo(caminhos_audio_temp)
            
            st.write("📄 Injetando informações no Word...")
            arquivo_final = gerar_docx(dados, caminhos_foto_temp)
            
            status.update(label="Relatório finalizado com sucesso!", state="complete", expanded=False)

        st.success("✅ Tudo pronto! O documento está pronto para envio ao cliente.")
        
        if arquivo_final and os.path.exists(arquivo_final):
            with open(arquivo_final, "rb") as f:
                st.download_button(
                    label="📥 Baixar Laudo (Word)", 
                    data=f, 
                    file_name=arquivo_final,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
    except Exception as e:
        st.error(f"Erro ao gerar o relatório: {str(e)}")
        
    finally:
        for caminho in caminhos_audio_temp + caminhos_foto_temp:
            if os.path.exists(caminho):
                try: os.remove(caminho)
                except: pass