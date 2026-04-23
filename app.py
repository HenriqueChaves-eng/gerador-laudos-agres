import streamlit as st
import google.generativeai as genai
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import json
import os
from datetime import date

# ==========================================
# 1. Configurações de Estado e Interface
# ==========================================
st.set_page_config(page_title="Agres | Relatórios Técnicos", page_icon="🚜", layout="centered")

# Inicialização da Memória do App
if 'lista_gravadores' not in st.session_state:
    st.session_state.lista_gravadores = [0]
if 'proximo_id' not in st.session_state:
    st.session_state.proximo_id = 1
if 'reset_audio' not in st.session_state:
    st.session_state.reset_audio = 0
if 'relatorio_pronto' not in st.session_state:
    st.session_state.relatorio_pronto = None
if 'nome_arquivo_pronto' not in st.session_state:
    st.session_state.nome_arquivo_pronto = None

# Estilização Customizada (Agres Style)
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

# Configuração Segura da API
try:
    GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
    genai.configure(api_key=GOOGLE_API_KEY)
    model = genai.GenerativeModel('models/gemini-2.5-flash')
except Exception:
    st.error("⚠️ Erro Crítico: Chave API não configurada nos Secrets do Streamlit.")
    st.stop()

# ==========================================
# 2. Funções de Inteligência e Documento
# ==========================================
def processar_atendimento_completo(arquivos_audio_temp):
    materiais_para_ia = []
    arquivos_api = [] 
    
    for audio in arquivos_audio_temp:
        temp_file = genai.upload_file(path=audio)
        materiais_para_ia.append(temp_file)
        arquivos_api.append(temp_file)
        
    prompt = f"""
    Analise os áudios anexados e extraia os dados para um relatório técnico da Agres.
    
    REGRA CRÍTICA DE TÓPICOS:
    - Em "objetivos", coloque apenas o propósito da visita (Ex: Instalação de kit, manutenção, etc).
    - Em "configuracoes", coloque os detalhes técnicos do que foi feito (Ex: suportes montados, chicotes passados).
    - Não misture os dois campos acima.
    - O campo "relato" deve ser um texto coeso, formal (ABNT) e profissional.

    Retorne APENAS um JSON válido:
    {{
        "suporte": "", "instalacao": "", "treinamento": "",
        "data_visita": "{date.today().strftime('%d/%m/%Y')}",
        "tecnicos": "Henrique Chaves",
        "cliente_local": "", "equipamentos": "", "maquinas": "",
        "objetivos": "", "configuracoes": "", "calibracoes": "",
        "acompanhantes": "",
        "relato": ""
    }}
    """
    
    try:
        resposta = model.generate_content([prompt] + materiais_para_ia)
        texto_bruto = resposta.text.strip()
        
        # Limpeza de JSON (ignora lixo textual da IA)
        inicio = texto_bruto.find('{')
        fim = texto_bruto.rfind('}')
        if inicio != -1 and fim != -1:
            texto_bruto = texto_bruto[inicio:fim+1]
            
        return json.loads(texto_bruto)
    except Exception as e:
        raise Exception(f"Erro na interpretação da IA: {e}")
    finally:
        for f in arquivos_api:
            try: genai.delete_file(f.name)
            except: pass

def gerar_docx(dados_json, dicionario_evidencias, caminhos_cabecalho):
    doc = DocxTemplate("modelo_tags.docx")
    
    # 2.1 Processamento de Fotos do Cabeçalho (Plaqueta 90mm / Maq 45mm)
    if caminhos_cabecalho.get('info_equip'):
        dados_json['img_info_equipamento'] = InlineImage(doc, caminhos_cabecalho['info_equip'], width=Mm(90))
    else: dados_json['img_info_equipamento'] = ""

    if caminhos_cabecalho.get('maquina'):
        dados_json['img_maquina'] = InlineImage(doc, caminhos_cabecalho['maquina'], width=Mm(45))
    else: dados_json['img_maquina'] = ""

    if caminhos_cabecalho.get('implemento'):
        dados_json['img_implemento'] = InlineImage(doc, caminhos_cabecalho['implemento'], width=Mm(45))
    else: dados_json['img_implemento'] = ""

    # 2.2 Processamento de Evidências (Corpo do Laudo - 145mm)
    for categoria, arquivos in dicionario_evidencias.items():
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
    nome_arquivo = f"Relatorio_{dados_json.get('cliente_local', 'Atendimento').replace(' ', '_')}.docx"
    doc.save(nome_arquivo)
    return nome_arquivo

# ==========================================
# 3. Interface Visual (Dashboard)
# ==========================================
st.markdown("<h1 class='titulo-app'>🚜 Agres Reports</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitulo-app'>Geração de Laudos Técnicos de Atendimento</p>", unsafe_allow_html=True)

# SEÇÃO 1: RELATO TÉCNICO
with st.container(border=True):
    st.markdown("### 🎙️ 1. Relato Técnico")
    aba1, aba2 = st.tabs(["🔴 Gravar agora", "📁 Arquivos do celular"])
    
    with aba1:
        st.info("Grave por trechos. Use 'Remover' para apagar um áudio específico.")
        audios_rec = []
        for i, id_gravador in enumerate(st.session_state.lista_gravadores):
            col_gravador, col_excluir = st.columns([0.80, 0.20], vertical_alignment="bottom")
            with col_gravador:
                a = st.audio_input(f"Trecho {i+1}", key=f"rec_{id_gravador}_{st.session_state.reset_audio}")
                if a: audios_rec.append(a)
            with col_excluir:
                if st.button("🗑️ Remover", key=f"btn_del_{id_gravador}", use_container_width=True):
                    st.session_state.lista_gravadores.remove(id_gravador)
                    st.rerun()
        
        c1, c2 = st.columns(2)
        if c1.button("➕ Novo trecho", use_container_width=True):
            st.session_state.lista_gravadores.append(st.session_state.proximo_id)
            st.session_state.proximo_id += 1
            st.rerun()
        if c2.button("🧹 Limpar tudo", use_container_width=True):
            st.session_state.lista_gravadores = [st.session_state.proximo_id]
            st.session_state.proximo_id += 1
            st.session_state.reset_audio += 1
            st.session_state.relatorio_pronto = None
            st.rerun()

    with aba2:
        audios_up = st.file_uploader("Upload de áudios", type=['wav', 'mp3', 'm4a'], accept_multiple_files=True)

# SEÇÃO 2: FOTOS DO CABEÇALHO
with st.container(border=True):
    st.markdown("### 🏷️ 2. Fotos do Cabeçalho")
    c_p1, c_p2, c_p3 = st.columns(3)
    f_plaqueta = c_p1.file_uploader("📸 Plaqueta", type=['jpg', 'jpeg', 'png'], key="up_p1")
    f_maquina = c_p2.file_uploader("🚜 Máquina", type=['jpg', 'jpeg', 'png'], key="up_p2")
    f_implemento = c_p3.file_uploader("🔧 Implemento", type=['jpg', 'jpeg', 'png'], key="up_p3")

# SEÇÃO 3: EVIDÊNCIAS
with st.container(border=True):
    st.markdown("### 📸 3. Evidências Fotográficas")
    col_e1, col_e2 = st.columns(2)
    f_eq = col_e1.file_uploader("📋 Equipamento Agres", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])
    f_ins = col_e1.file_uploader("🔨 Instalação", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])
    f_conf = col_e2.file_uploader("⚙️ Configurações", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])
    f_out = col_e2.file_uploader("📂 Outros registros", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])

# ==========================================
# 4. Lógica de Execução
# ==========================================
audios_finais = audios_rec + (audios_up if audios_up else [])

if audios_finais and st.button("🚀 Gerar Relatório Profissional"):
    temp_paths = []
    cabs_paths = {}
    try:
        with st.status("Processando dados e imagens...", expanded=True) as status:
            # Salvamento de Áudios
            for i, a in enumerate(audios_finais):
                ext = a.name.split('.')[-1] if hasattr(a, 'name') and '.' in a.name else 'wav'
                p = f"t_aud_{i}.{ext}"
                with open(p, "wb") as f: f.write(a.getvalue())
                temp_paths.append(p)
            
            # Salvamento de Fotos do Cabeçalho
            def save_img(file, name):
                if file:
                    ext = file.name.split('.')[-1]
                    p = f"{name}.{ext}"
                    with open(p, "wb") as f: f.write(file.getvalue())
                    temp_paths.append(p)
                    return p
                return None

            cabs_paths['info_equip'] = save_img(f_plaqueta, "t_plaq")
            cabs_paths['maquina'] = save_img(f_maquina, "t_maq")
            cabs_paths['implemento'] = save_img(f_implemento, "t_imp")

            # Salvamento de Evidências
            mapa_ev = {"fotos_equipamento": f_eq, "fotos_instalacao": f_ins, "fotos_configuracao": f_conf, "fotos_outros": f_out}
            dic_ev = {}
            for cat, files in mapa_ev.items():
                paths = []
                if files:
                    for i, f in enumerate(files):
                        p = save_img(f, f"t_{cat}_{i}")
                        if p: paths.append(p)
                dic_ev[cat] = paths
                
            # Chamada da IA e Word
            dados = processar_atendimento_completo(temp_paths[:len(audios_finais)])
            arq_final = gerar_docx(dados, dic_ev, cabs_paths)
            
            # Carrega para memória RAM
            with open(arq_final, "rb") as f:
                st.session_state.relatorio_pronto = f.read()
                st.session_state.nome_arquivo_pronto = arq_final
                
            os.remove(arq_final) # Limpa arquivo físico
            status.update(label="Relatório Finalizado!", state="complete", expanded=False)
            
    except Exception as e:
        st.error(f"Erro no processamento: {e}")
    finally:
        for p in temp_paths:
            if os.path.exists(p): os.remove(p)

# Exibição do Botão de Download Persistente
if st.session_state.relatorio_pronto:
    st.success("✅ O laudo está pronto para download!")
    st.download_button(
        label="📥 Baixar Relatório (Word)", 
        data=st.session_state.relatorio_pronto, 
        file_name=st.session_state.nome_arquivo_pronto,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True
    )
