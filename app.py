import streamlit as st
import google.generativeai as genai
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import json
import os
from datetime import date

# ==========================================
# 1. Configurações Iniciais e Memória
# ==========================================
st.set_page_config(page_title="Agres | Relatórios Técnicos", page_icon="🚜", layout="centered")

# Variáveis de Estado (Memória do App)
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

try:
    GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
    genai.configure(api_key=GOOGLE_API_KEY)
    model = genai.GenerativeModel('models/gemini-2.5-flash')
except Exception:
    st.error("⚠️ O sistema foi interrompido. Verifique a chave do Google API nos 'Secrets'.")
    st.stop()

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
    Retorne APENAS um JSON válido:
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
        texto_bruto = resposta.text.strip()
        
        # Filtro Inteligente: Extrai apenas o JSON, ignorando textos antes ou depois
        inicio = texto_bruto.find('{')
        fim = texto_bruto.rfind('}')
        if inicio != -1 and fim != -1:
            texto_bruto = texto_bruto[inicio:fim+1]
            
        return json.loads(texto_bruto)
    except Exception as e:
        raise Exception(f"Falha ao interpretar a resposta da IA: {e}")
    finally:
        for f in arquivos_api:
            try: genai.delete_file(f.name)
            except: pass

def gerar_docx(dados_json, dicionario_evidencias, caminhos_cabecalho):
    doc = DocxTemplate("modelo_tags.docx")
    
    # Fotos do Cabeçalho
    if caminhos_cabecalho.get('info_equip'):
        dados_json['img_info_equipamento'] = InlineImage(doc, caminhos_cabecalho['info_equip'], width=Mm(75))
    else: dados_json['img_info_equipamento'] = ""

    if caminhos_cabecalho.get('maquina'):
        dados_json['img_maquina'] = InlineImage(doc, caminhos_cabecalho['maquina'], width=Mm(35))
    else: dados_json['img_maquina'] = ""

    if caminhos_cabecalho.get('implemento'):
        dados_json['img_implemento'] = InlineImage(doc, caminhos_cabecalho['implemento'], width=Mm(35))
    else: dados_json['img_implemento'] = ""

    # Fotos de Evidências
    for categoria, arquivos in dicionario_evidencias.items():
        lista_formatada = []
        if arquivos:
            for i, foto_path in enumerate(arquivos):
                nome = categoria.replace('fotos_', '').replace('_', ' ').title()
                titulo = f"Figura {i+1} – Registro de {nome}."
                fonte = f"Fonte: O autor ({date.today().year})."
                imagem = InlineImage(doc, foto_path, width=Mm(145))
                lista_formatada.append({"titulo": titulo, "imagem": imagem, "fonte": fonte})
        dados_json[categoria] = lista_formatada
    
    doc.render(dados_json)
    nome_arquivo = f"Relatorio_{dados_json.get('cliente_local', 'Atendimento').replace(' ', '_')}.docx"
    doc.save(nome_arquivo)
    return nome_arquivo

# ==========================================
# 3. Interface Front-End
# ==========================================
st.markdown("<h1 class='titulo-app'>🚜 Agres Relatórios</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitulo-app'>Relatório Técnico de Atendimento Presencial</p>", unsafe_allow_html=True)

with st.container(border=True):
    st.markdown("### 🎙️ 1. Relato Técnico")
    aba1, aba2 = st.tabs(["🔴 Gravar agora", "📁 Arquivos do celular"])
    
    with aba1:
        st.info("Grave o relato em etapas. Use 'Remover' para excluir um trecho.")
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
            # Limpa o laudo gerado ao resetar o painel
            st.session_state.relatorio_pronto = None 
            st.rerun()

    with aba2:
        audios_up = st.file_uploader("Upload de áudios", type=['wav', 'mp3', 'm4a'], accept_multiple_files=True)

with st.container(border=True):
    st.markdown("### 🏷️ 2. Fotos do Cabeçalho")
    col_plaqueta, col_maquina, col_implemento = st.columns(3)
    with col_plaqueta: foto_plaqueta = st.file_uploader("📸 Info Equipamento", type=['jpg', 'jpeg', 'png'], key="up_plaq")
    with col_maquina: foto_maq = st.file_uploader("🚜 Máquina", type=['jpg', 'jpeg', 'png'], key="up_maq")
    with col_implemento: foto_imp = st.file_uploader("🔧 Implemento", type=['jpg', 'jpeg', 'png'], key="up_imp")

with st.container(border=True):
    st.markdown("### 📸 3. Evidências Fotográficas")
    col_ev1, col_ev2 = st.columns(2)
    f_eq = col_ev1.file_uploader("📋 Equipamento Agres", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])
    f_ins = col_ev1.file_uploader("🔨 Instalação", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])
    f_conf = col_ev2.file_uploader("⚙️ Configurações", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])
    f_out = col_ev2.file_uploader("📂 Outros", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])

audios_finais = audios_rec + (audios_up if audios_up else [])

# ==========================================
# 4. Geração e Download (Corrigido)
# ==========================================
if audios_finais and st.button("🚀 Gerar Relatório"):
    temp_files = []
    caminhos_cab = {}
    try:
        with st.status("Processando dados e imagens...", expanded=True) as status:
            # Salva áudios e resolve extensões de forma segura
            for i, a in enumerate(audios_finais):
                ext = a.name.split('.')[-1] if hasattr(a, 'name') and '.' in a.name else 'wav'
                p = f"t_aud_{i}.{ext}"
                with open(p, "wb") as f: f.write(a.getvalue())
                temp_files.append(p)
            
            # Helper function para salvar fotos
            def salvar_foto(arquivo_upload, prefixo):
                if arquivo_upload:
                    ext = arquivo_upload.name.split('.')[-1] if hasattr(arquivo_upload, 'name') and '.' in arquivo_upload.name else 'jpg'
                    path = f"{prefixo}.{ext}"
                    with open(path, "wb") as f: f.write(arquivo_upload.getvalue())
                    temp_files.append(path)
                    return path
                return None

            caminhos_cab['info_equip'] = salvar_foto(foto_plaqueta, "t_cab_plaqueta")
            caminhos_cab['maquina'] = salvar_foto(foto_maq, "t_cab_maq")
            caminhos_cab['implemento'] = salvar_foto(foto_imp, "t_cab_imp")

            mapa_evidencias = {"fotos_equipamento": f_eq, "fotos_instalacao": f_ins, "fotos_configuracao": f_conf, "fotos_outros": f_out}
            dic_evidencias = {}
            for cat, files in mapa_evidencias.items():
                paths = []
                if files:
                    for i, f in enumerate(files):
                        caminho_salvo = salvar_foto(f, f"t_{cat}_{i}")
                        if caminho_salvo: paths.append(caminho_salvo)
                dic_evidencias[cat] = paths
                
            dados = processar_atendimento_completo(temp_files[:len(audios_finais)])
            arq_final = gerar_docx(dados, dic_evidencias, caminhos_cab)
            
            # MEMÓRIA: Lê o arquivo Word para a RAM do servidor
            with open(arq_final, "rb") as f:
                st.session_state.relatorio_pronto = f.read()
                st.session_state.nome_arquivo_pronto = arq_final
                
            # HIGIENE: Deleta o arquivo Word físico gerado
            os.remove(arq_final)
            
            status.update(label="Relatório Finalizado!", state="complete", expanded=False)
            
    except Exception as e:
        st.error(f"Ocorreu um erro no processamento: {e}")
    finally:
        # Deleta todas as mídias temporárias (áudios e fotos soltas)
        for p in temp_files:
            if os.path.exists(p): os.remove(p)

# Exibe o botão de Download SE houver um relatório na memória
if st.session_state.relatorio_pronto:
    st.success("✅ Laudo gerado com sucesso! Clique abaixo para salvar.")
    st.download_button(
        label="📥 Baixar Relatório (Word)", 
        data=st.session_state.relatorio_pronto, 
        file_name=st.session_state.nome_arquivo_pronto,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary", # Deixa o botão em destaque
        use_container_width=True
    )
