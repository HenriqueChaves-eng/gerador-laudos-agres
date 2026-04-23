import streamlit as st
import google.generativeai as genai
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import json
import os
from datetime import date

# ==========================================
# 1. Configurações Iniciais
# ==========================================
st.set_page_config(page_title="Agres | Relatórios Técnicos", page_icon="🚜", layout="centered")

# ---> O SEGREDO ESTÁ AQUI: Este bloco precisa estar exatamente assim <---
if 'lista_gravadores' not in st.session_state:
    st.session_state.lista_gravadores = [0]
if 'proximo_id' not in st.session_state:
    st.session_state.proximo_id = 1
if 'reset_audio' not in st.session_state:
    st.session_state.reset_audio = 0

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
    </style>
""", unsafe_allow_html=True)

# Blindagem 1: Verificação da Chave
try:
    GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
    genai.configure(api_key=GOOGLE_API_KEY)
    model = genai.GenerativeModel('models/gemini-2.5-flash')
except Exception as e:
    st.error("⚠️ O sistema foi interrompido. A chave do Google API não foi encontrada nos 'Secrets' do Streamlit.")
    st.stop() # Trava o app aqui para não dar erro sujo

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
    Analise os áudios e extraia dados para um relatório técnico da Agres.
    Retorne APENAS um JSON:
    {{
        "suporte": "", "instalacao": "", "treinamento": "",
        "data_visita": "{date.today().strftime('%d/%m/%Y')}",
        "tecnicos": "Henrique Chaves",
        "cliente_local": "", "equipamentos": "", "maquinas": "",
        "objetivos": "", "configuracoes": "", "calibracoes": "",
        "acompanhantes": "",
        "relato": "Texto técnico formal aqui."
    }}
    """
    
    try:
        resposta = model.generate_content([prompt] + materiais_para_ia)
        texto_json = resposta.text.strip().replace('```json', '').replace('```', '')
        return json.loads(texto_json)
    finally:
        for f in arquivos_api:
            try: genai.delete_file(f.name)
            except: pass

def gerar_docx(dados_json, dicionario_caminhos):
    doc = DocxTemplate("modelo_tags.docx")
    
    for categoria, arquivos in dicionario_caminhos.items():
        lista_fotos = []
        if arquivos:
            for i, foto_path in enumerate(arquivos):
                nome = categoria.replace('fotos_', '').replace('_', ' ').title()
                titulo = f"Figura {i+1} – Registro de {nome}."
                fonte = f"Fonte: O autor ({date.today().year})."
                imagem = InlineImage(doc, foto_path, width=Mm(145))
                lista_fotos.append({"titulo": titulo, "imagem": imagem, "fonte": fonte})
        dados_json[categoria] = lista_fotos
    
    doc.render(dados_json)
    nome = f"Relatorio_{dados_json.get('cliente_local', 'Atendimento').replace(' ', '_')}.docx"
    doc.save(nome)
    return nome

# ==========================================
# 3. Interface Front-End
# ==========================================
st.markdown("<h1 class='titulo-app'>🚜 Agres Relatórios Técnicos</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitulo-app'>Relatório Técnico de Atendimento Presencial</p>", unsafe_allow_html=True)

with st.container(border=True):
    st.markdown("### 🎙️ 1. Relato Técnico")
    aba1, aba2 = st.tabs(["🔴 Gravar agora", "📁 Arquivos do celular"])
    
  with aba1:
        st.info("Grave o relato em etapas. Use 'Remover' para excluir um trecho específico.")
        audios_rec = []
        
        for i, id_gravador in enumerate(st.session_state.lista_gravadores):
            # O vertical_alignment="bottom" resolve o problema de alinhamento no Windows
            col_gravador, col_excluir = st.columns([0.80, 0.20], vertical_alignment="bottom")
            
            with col_gravador:
                a = st.audio_input(f"Trecho {i+1}", key=f"rec_{id_gravador}_{st.session_state.reset_audio}")
                if a: audios_rec.append(a)
            
            with col_excluir:
                # Mudamos de "❌" para "Remover" para ficar elegante quando for jogado para baixo no iPhone
                if st.button("🗑️ Remover", key=f"btn_del_{id_gravador}", use_container_width=True):
                    st.session_state.lista_gravadores.remove(id_gravador)
                    st.rerun()
        
        st.markdown("<div class='btn-adicionar'>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("➕ Novo trecho", use_container_width=True):
                st.session_state.lista_gravadores.append(st.session_state.proximo_id)
                st.session_state.proximo_id += 1
                st.rerun()
        with c2:
            if st.button("🗑️ Limpar tudo", use_container_width=True):
                st.session_state.lista_gravadores = [st.session_state.proximo_id]
                st.session_state.proximo_id += 1
                st.session_state.reset_audio += 1
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    with aba2:
        audios_up = st.file_uploader("Upload de áudios", type=['wav', 'mp3', 'm4a'], accept_multiple_files=True)

with st.container(border=True):
    st.markdown("### 📸 2. Evidências Fotográficas")
    col1, col2 = st.columns(2)
    f_eq = col1.file_uploader("📋 Equipamento", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])
    f_ins = col1.file_uploader("🔧 Instalação", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])
    f_conf = col2.file_uploader("⚙️ Configurações", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])
    f_out = col2.file_uploader("📂 Outros", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])

audios_finais = audios_rec + (audios_up if audios_up else [])

if audios_finais and st.button("🚀 Processar Laudo Inteligente"):
    c_audio, c_foto = [], []
    try:
        with st.status("Processando...", expanded=True) as status:
            # Blindagem 2: Proteção de Extensões de Arquivo
            for i, a in enumerate(audios_finais):
                ext = a.name.split('.')[-1] if hasattr(a, 'name') and '.' in a.name else 'wav'
                p = f"t_aud_{i}.{ext}"
                with open(p, "wb") as f: f.write(a.getvalue())
                c_audio.append(p)
            
            fotos_map = {"fotos_equipamento": f_eq, "fotos_instalacao": f_ins, "fotos_configuracao": f_conf, "fotos_outros": f_out}
            dic_paths = {}
            for cat, files in fotos_map.items():
                paths = []
                if files:
                    for i, f in enumerate(files):
                        ext = f.name.split('.')[-1] if hasattr(f, 'name') and '.' in f.name else 'jpg'
                        p = f"t_{cat}_{i}.{ext}"
                        with open(p, "wb") as file: file.write(f.getvalue())
                        paths.append(p); c_foto.append(p)
                dic_paths[cat] = paths
                
            dados = processar_atendimento_completo(c_audio)
            arq = gerar_docx(dados, dic_paths)
            status.update(label="Concluído!", state="complete", expanded=False)

        st.success("✅ Laudo pronto!")
        with open(arq, "rb") as f:
            st.download_button("📥 Baixar Laudo", f, file_name=arq)
            
    except Exception as e:
        st.error(f"Erro: {e}")
    finally:
        for p in c_audio + c_foto:
            if os.path.exists(p): os.remove(p)
