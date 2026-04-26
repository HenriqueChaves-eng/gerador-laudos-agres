import json
import re
import tempfile
import unicodedata
import uuid
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import google.generativeai as genai
import streamlit as st
from docx import Document
from docx.document import Document as DocumentClass
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm, Pt
from docx.table import _Cell, Table
from docxtpl import DocxTemplate, InlineImage
from PIL import Image, ImageOps, UnidentifiedImageError


# ==========================================
# 1. Configurações gerais
# ==========================================
st.set_page_config(page_title="Agres | Relatório Técnico", page_icon="🚜", layout="centered")

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "modelo_tags.docx"

TAM_PLAQUETA = 60
TAM_MAQUINA = 32
TAM_EVIDENCIA = 120

CAMPOS_RELATORIO = (
    "suporte",
    "instalacao",
    "treinamento",
    "data_visita",
    "tecnicos",
    "cliente_local",
    "equipamentos",
    "maquinas",
    "objetivos",
    "configuracoes",
    "calibracoes",
    "acompanhantes",
    "relato",
)

CAMPOS_SERVICO = ("suporte", "instalacao", "treinamento")

ROTULOS_CAMPOS = {
    "suporte": "Suporte",
    "instalacao": "Instalação",
    "treinamento": "Treinamento",
    "configuracoes": "Configurações",
    "calibracoes": "Calibrações",
}

TERMOS_INTERVENCAO_FISICA = (
    "alimentacao",
    "cabo",
    "can h",
    "can l",
    "chicote",
    "conector",
    "confeccao",
    "defeito",
    "diagnostico",
    "falha",
    "fabricacao",
    "fixacao",
    "furacao",
    "instalacao fisica",
    "mau contato",
    "npn",
    "pinagem",
    "pino",
    "pnp",
    "rele",
    "roteamento",
    "solda",
    "substituicao",
    "suporte",
    "terminador",
    "troca",
)

TERMOS_NAO_CALIBRACAO = TERMOS_INTERVENCAO_FISICA + (
    "assistencia",
    "garantia",
    "orientacao",
    "pendencia",
    "recomendacao",
)

EXTENSOES_AUDIO = {"wav", "mp3", "m4a"}
EXTENSOES_IMAGEM = {"jpg", "jpeg", "png"}

CATEGORIAS_EVIDENCIAS = {
    "fotos_equipamento": {
        "nome": "Identificação do Equipamento",
        "titulo_padrao": "Identificação do equipamento Agres",
        "legenda_padrao": "Registro de identificação, série, versão ou componentes do equipamento Agres.",
    },
    "fotos_instalacao": {
        "nome": "Instalação e Chicotes",
        "titulo_padrao": "Instalação e chicotes do sistema",
        "legenda_padrao": "Registro da instalação física, fixação, roteamento de chicotes ou conexão elétrica.",
    },
    "fotos_configuracao": {
        "nome": "Configurações do Sistema",
        "titulo_padrao": "Configuração do sistema",
        "legenda_padrao": "Registro de tela, parâmetro, versão, calibração ou validação realizada no sistema.",
    },
    "fotos_outros": {
        "nome": "Atividades Adicionais",
        "titulo_padrao": "Registro complementar do atendimento",
        "legenda_padrao": "Registro fotográfico complementar relacionado ao atendimento técnico.",
    },
}


for chave, valor_inicial in {
    "lista_gravadores": [0],
    "proximo_id": 1,
    "reset_audio": 0,
    "relatorio_pronto": None,
    "nome_arquivo_pronto": None,
}.items():
    if chave not in st.session_state:
        st.session_state[chave] = valor_inicial.copy() if isinstance(valor_inicial, list) else valor_inicial


st.markdown(
    """
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
    """,
    unsafe_allow_html=True,
)


try:
    GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
    if not GOOGLE_API_KEY or GOOGLE_API_KEY.strip() == "cole_sua_chave_aqui":
        raise ValueError("GOOGLE_API_KEY ainda não foi preenchida no arquivo .streamlit/secrets.toml.")
    MODELO_GEMINI = st.secrets.get("GEMINI_MODEL", "models/gemini-2.5-flash")
    genai.configure(api_key=GOOGLE_API_KEY)
    model = genai.GenerativeModel(MODELO_GEMINI)
except Exception:
    st.error("⚠️ Erro crítico: chave GOOGLE_API_KEY não configurada nos Secrets do Streamlit.")
    st.stop()


# ==========================================
# 2. Tratamento de texto e validação
# ==========================================
def data_atual_brasil() -> date:
    try:
        return datetime.now(ZoneInfo("America/Sao_Paulo")).date()
    except Exception:
        return date.today()


def normalizar_busca(texto: str) -> str:
    texto = unicodedata.normalize("NFD", texto or "")
    texto = "".join(caractere for caractere in texto if unicodedata.category(caractere) != "Mn")
    return texto.lower()


def valor_para_texto(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, list):
        return "\n".join(valor_para_texto(item) for item in valor if valor_para_texto(item))
    if isinstance(valor, dict):
        linhas = []
        for chave, item in valor.items():
            texto_item = valor_para_texto(item)
            if texto_item:
                linhas.append(f"{chave}: {texto_item}")
        return "\n".join(linhas)
    return str(valor)


def limpar_texto(valor) -> str:
    texto = valor_para_texto(valor).replace("\r\n", "\n").replace("\r", "\n")
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r" *\n *", "\n", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    texto = texto.strip(" \n\t;")
    if normalizar_busca(texto) in {"null", "none", "n/a", "nao informado", "nao informada"}:
        return ""
    return texto


def texto_ou_padrao(valor, padrao="Não informado") -> str:
    texto = limpar_texto(valor)
    return texto if texto else padrao


def dividir_itens(texto: str) -> list[str]:
    itens = []
    for linha in limpar_texto(texto).split("\n"):
        for parte in re.split(r"\s*[;•]\s*", linha):
            item = parte.strip(" -–—\t")
            if item:
                itens.append(item)
    return itens


def contem_termo(texto: str, termos: tuple[str, ...]) -> bool:
    texto_normalizado = normalizar_busca(texto)
    return any(termo in texto_normalizado for termo in termos)


def normalizar_marcador_servico(valor) -> tuple[str, list[str]]:
    texto = limpar_texto(valor)
    if not texto:
        return "", []

    texto_normalizado = normalizar_busca(texto)
    marcador = ""
    if texto.strip().upper() == "X" or any(
        termo in texto_normalizado
        for termo in ("sim", "realizado", "realizada", "executado", "executada", "suporte", "instalacao", "treinamento")
    ):
        marcador = "X"

    detalhes = [] if texto.strip().upper() == "X" else [texto]
    return marcador, detalhes


def filtrar_campo_curto(texto: str, termos_bloqueados: tuple[str, ...]) -> tuple[str, list[str]]:
    itens_validos = []
    itens_para_relato = []

    for item in dividir_itens(texto):
        if contem_termo(item, termos_bloqueados):
            itens_para_relato.append(item)
        else:
            itens_validos.append(item)

    return "\n".join(itens_validos).strip(), itens_para_relato


def adicionar_ao_relato(relato: str, titulo: str, itens: list[str]) -> str:
    if not itens:
        return relato

    relato_base = limpar_texto(relato)
    relato_normalizado = normalizar_busca(relato_base)
    itens_novos = [item for item in itens if normalizar_busca(item) not in relato_normalizado]
    if not itens_novos:
        return relato_base

    bloco = titulo + "\n" + "\n".join(f"- {item}" for item in itens_novos)
    return (relato_base + "\n\n" + bloco).strip() if relato_base else bloco


def normalizar_dados_relatorio(dados: dict) -> dict:
    dados_normalizados = {campo: limpar_texto(dados.get(campo, "")) for campo in CAMPOS_RELATORIO}
    detalhes_para_relato = []

    for campo in CAMPOS_SERVICO:
        marcador, detalhes = normalizar_marcador_servico(dados_normalizados[campo])
        dados_normalizados[campo] = marcador
        detalhes_para_relato.extend(f"{ROTULOS_CAMPOS[campo]}: {detalhe}" for detalhe in detalhes)

    configuracoes, itens_config_relato = filtrar_campo_curto(
        dados_normalizados["configuracoes"],
        TERMOS_INTERVENCAO_FISICA,
    )
    calibracoes, itens_calibracao_relato = filtrar_campo_curto(
        dados_normalizados["calibracoes"],
        TERMOS_NAO_CALIBRACAO,
    )

    dados_normalizados["configuracoes"] = configuracoes or "Não informado"
    dados_normalizados["calibracoes"] = calibracoes or "Não informado"
    dados_normalizados["relato"] = adicionar_ao_relato(
        dados_normalizados["relato"],
        "Intervenções físicas e diagnósticos registrados:",
        itens_config_relato + itens_calibracao_relato + detalhes_para_relato,
    )

    if not dados_normalizados["relato"]:
        dados_normalizados["relato"] = "Não informado nos áudios ou observações encaminhadas."

    dados_normalizados["data_visita"] = dados_normalizados["data_visita"] or data_atual_brasil().strftime("%d/%m/%Y")
    dados_normalizados["tecnicos"] = dados_normalizados["tecnicos"] or "Henrique Chaves"

    for campo in ("cliente_local", "equipamentos", "maquinas", "objetivos", "acompanhantes"):
        dados_normalizados[campo] = texto_ou_padrao(dados_normalizados[campo])

    return dados_normalizados


# ==========================================
# 3. Inteligência artificial
# ==========================================
def montar_prompt(contexto_manual: str = "") -> str:
    hoje = data_atual_brasil().strftime("%d/%m/%Y")
    contexto_manual = limpar_texto(contexto_manual)
    bloco_contexto = (
        f"\nANOTAÇÕES COMPLEMENTARES INFORMADAS PELO TÉCNICO:\n{contexto_manual}\n"
        if contexto_manual
        else ""
    )

    return f"""
Você é redator técnico da Agres e deve transformar áudios/anotações de atendimento de campo em dados para um relatório formal.

Use português técnico, claro e objetivo. Reescreva falas informais em linguagem profissional, sem inventar dados, versões, medidas, peças ou conclusões que não estejam no material recebido.
{bloco_contexto}
REGRAS DE CLASSIFICAÇÃO DOS CAMPOS:
1. suporte, instalacao e treinamento: retornar somente "X" quando o serviço tiver ocorrido; caso contrário, retornar "".
2. data_visita: preserve intervalo de datas quando o atendimento ocorrer em mais de um dia, por exemplo "19 a 21/01/2026".
3. cliente_local: informar cliente, cidade/UF, revenda, fábrica, propriedade, coordenadas ou link de localização quando existirem.
4. equipamentos: organizar em linhas com modelo, série, versões de aplicação/sistema/carga, ECU, compensador, GPS e identificadores.
5. maquinas: organizar em linhas com fabricante, modelo, implemento, comando de válvulas e características relevantes.
6. objetivos: escrever somente o objetivo principal do atendimento, em uma ou duas frases.
7. configuracoes: incluir somente parâmetros de sistema, software, tela, ECU, controlador, seções, geometria, versões, módulos habilitados, ganhos ou ajustes feitos em menus. Quando houver valores, use o padrão "Parâmetro: valor".
8. calibracoes: incluir somente calibrações, aferições e validações com valores, medidas, sensores, vazão, largura, offset, angulação ou parâmetros numéricos.
9. relato: concentrar todo o detalhamento técnico e cronológico. Cabos, chicotes, conectores, soldas, conversores PNP/NPN, pinagem, relés, terminadores CAN, suportes físicos, falhas, diagnósticos, testes, correções, pendências e recomendações pertencem ao relato, não a configurações nem a calibrações.

PADRÃO DO RELATO:
- Escrever em terceira pessoa.
- Descrever contexto do atendimento, problema informado, diagnóstico inicial, procedimentos executados, configurações/calibrações relevantes, problemas encontrados, correções aplicadas, testes de funcionamento, resultado final e conclusão técnica.
- Quando houver conteúdo suficiente, usar subtítulos técnicos curtos dentro do próprio texto, como "Descrição do problema", "Diagnóstico inicial", "Ações corretivas", "Testes complementares", "Validação do sistema", "Resultado final" e "Conclusão técnica".
- Em atendimentos de vários dias, separar a sequência por data ou por etapa.
- Informar "Não informado" nos campos textuais quando o dado não for mencionado.

Retorne apenas um JSON válido, sem markdown e sem comentários, com exatamente esta estrutura:
{{
    "suporte": "",
    "instalacao": "",
    "treinamento": "",
    "data_visita": "{hoje}",
    "tecnicos": "Henrique Chaves",
    "cliente_local": "",
    "equipamentos": "",
    "maquinas": "",
    "objetivos": "",
    "configuracoes": "",
    "calibracoes": "",
    "acompanhantes": "",
    "relato": ""
}}
"""


def extrair_json_resposta(texto: str) -> dict:
    texto_bruto = (texto or "").strip()
    inicio = texto_bruto.find("{")
    fim = texto_bruto.rfind("}")
    if inicio == -1 or fim == -1 or fim <= inicio:
        raise ValueError("A IA não retornou um JSON válido.")

    texto_json = texto_bruto[inicio : fim + 1]
    try:
        return json.loads(texto_json)
    except json.JSONDecodeError as erro:
        trecho = texto_json[:500]
        raise ValueError(f"JSON inválido retornado pela IA: {erro}. Trecho recebido: {trecho}") from erro


def processar_atendimento_completo(arquivos_audio_temp: list[Path], contexto_manual: str = "") -> dict:
    materiais_para_ia = []
    arquivos_api = []

    for audio in arquivos_audio_temp:
        temp_file = genai.upload_file(path=str(audio))
        materiais_para_ia.append(temp_file)
        arquivos_api.append(temp_file)

    try:
        resposta = model.generate_content(
            [montar_prompt(contexto_manual)] + materiais_para_ia,
            generation_config=genai.GenerationConfig(
                response_mime_type="application/json",
                temperature=0.2,
            ),
        )
        dados = extrair_json_resposta(resposta.text)
        return normalizar_dados_relatorio(dados)
    except Exception as erro:
        raise Exception(f"Erro na interpretação técnica dos dados: {erro}") from erro
    finally:
        for arquivo in arquivos_api:
            try:
                genai.delete_file(arquivo.name)
            except Exception:
                pass


# ==========================================
# 4. Documento Word
# ==========================================
def imagem_docx(doc: DocxTemplate, caminho, largura_mm: int):
    if caminho:
        return InlineImage(doc, str(caminho), width=Mm(largura_mm))
    return ""


def limpar_nome_arquivo(texto: str) -> str:
    primeira_linha = next((linha for linha in limpar_texto(texto).split("\n") if linha.strip()), "Atendimento")
    nome = re.sub(r'[\\/*?:"<>|]', "", primeira_linha)
    nome = re.sub(r"\s+", "_", nome).strip("_")
    return (nome or "Atendimento")[:80]


def finalizar_frase(texto: str) -> str:
    texto = limpar_texto(texto)
    if not texto:
        return ""
    return texto if texto.endswith((".", "!", "?", ":", ";")) else f"{texto}."


def linhas_metadados(texto: str) -> list[str]:
    return [linha.strip() for linha in limpar_texto(texto).split("\n") if linha.strip()]


def separar_metadados_figura(linha: str) -> tuple[str, str, str]:
    partes = [parte.strip() for parte in linha.split("|", 2)]
    titulo = partes[0] if len(partes) >= 1 else ""
    legenda = partes[1] if len(partes) >= 2 else ""
    fonte = partes[2] if len(partes) >= 3 else ""
    return titulo, legenda, fonte


def montar_metadados_figura(categoria: str, indice_foto: int, numero_figura: int, legendas_evidencias: dict) -> dict:
    configuracao = CATEGORIAS_EVIDENCIAS[categoria]
    linhas_categoria = linhas_metadados((legendas_evidencias or {}).get(categoria, ""))
    titulo_manual, legenda_manual, fonte_manual = ("", "", "")

    if indice_foto < len(linhas_categoria):
        titulo_manual, legenda_manual, fonte_manual = separar_metadados_figura(linhas_categoria[indice_foto])

    titulo_base = finalizar_frase(titulo_manual or configuracao["titulo_padrao"])
    if normalizar_busca(titulo_base).startswith("figura"):
        titulo = titulo_base
    else:
        titulo = f"Figura {numero_figura} – {titulo_base}"

    legenda_base = finalizar_frase(legenda_manual or configuracao["legenda_padrao"])
    legenda = legenda_base if normalizar_busca(legenda_base).startswith(("legenda:", "nota:")) else f"Legenda: {legenda_base}"

    fonte_base = finalizar_frase(fonte_manual or f"O autor ({data_atual_brasil().year})")
    fonte = fonte_base if normalizar_busca(fonte_base).startswith("fonte:") else f"Fonte: {fonte_base}"

    return {"titulo": titulo, "legenda": legenda, "fonte": fonte}


def gerar_docx(
    dados_json: dict,
    dicionario_evidencias: dict,
    caminhos_cabecalho: dict,
    pasta_saida: Path,
    legendas_evidencias: dict | None = None,
) -> Path:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Modelo não encontrado: {TEMPLATE_PATH.name}")

    doc = DocxTemplate(str(TEMPLATE_PATH))
    dados_render = dict(dados_json)

    dados_render["img_info_equipamento"] = imagem_docx(doc, caminhos_cabecalho.get("info_equip"), TAM_PLAQUETA)
    dados_render["img_maquina"] = imagem_docx(doc, caminhos_cabecalho.get("maquina"), TAM_MAQUINA)
    dados_render["img_implemento"] = imagem_docx(doc, caminhos_cabecalho.get("implemento"), TAM_MAQUINA)

    contador_figura = 1
    for categoria in CATEGORIAS_EVIDENCIAS:
        lista_fotos = []
        for indice_foto, foto_path in enumerate(dicionario_evidencias.get(categoria, [])):
            metadados = montar_metadados_figura(categoria, indice_foto, contador_figura, legendas_evidencias or {})
            lista_fotos.append(
                {
                    "titulo": metadados["titulo"],
                    "imagem": InlineImage(doc, str(foto_path), width=Mm(TAM_EVIDENCIA)),
                    "fonte": metadados["fonte"],
                    "legenda": metadados["legenda"],
                }
            )
            contador_figura += 1
        dados_render[categoria] = lista_fotos

    nome_arquivo = f"Relatorio_{limpar_nome_arquivo(dados_render.get('cliente_local'))}.docx"
    caminho_saida = pasta_saida / nome_arquivo
    doc.render(dados_render)
    doc.save(str(caminho_saida))
    aplicar_paginacao_abnt_figuras(caminho_saida)
    return caminho_saida


def chave_paragrafo(paragraph) -> str:
    return paragraph._p.getroottree().getpath(paragraph._p)


def iterar_paragrafos_word(parent, vistos=None):
    if vistos is None:
        vistos = set()

    if isinstance(parent, DocumentClass):
        for paragraph in parent.paragraphs:
            chave = chave_paragrafo(paragraph)
            if chave not in vistos:
                vistos.add(chave)
                yield paragraph
        for table in parent.tables:
            yield from iterar_paragrafos_word(table, vistos)
    elif isinstance(parent, Table):
        for row in parent.rows:
            for cell in row.cells:
                yield from iterar_paragrafos_word(cell, vistos)
    elif isinstance(parent, _Cell):
        for paragraph in parent.paragraphs:
            chave = chave_paragrafo(paragraph)
            if chave not in vistos:
                vistos.add(chave)
                yield paragraph
        for table in parent.tables:
            yield from iterar_paragrafos_word(table, vistos)


def paragrafo_tem_imagem(paragraph) -> bool:
    return paragraph._p.xpath(".//*[local-name()='drawing' or local-name()='pict']")


def formatar_paragrafo_figura(paragraph, keep_with_next: bool, tamanho_fonte: int | None = 10) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    formato = paragraph.paragraph_format
    formato.keep_together = True
    formato.keep_with_next = keep_with_next
    formato.widow_control = True
    formato.line_spacing = 1
    formato.space_before = Pt(3)
    formato.space_after = Pt(3)
    if tamanho_fonte:
        for run in paragraph.runs:
            run.font.name = "Arial"
            run.font.size = Pt(tamanho_fonte)


def aplicar_paginacao_abnt_figuras(caminho_docx: Path) -> None:
    documento = Document(str(caminho_docx))
    paragrafos = list(iterar_paragrafos_word(documento))

    dentro_bloco_figura = False
    for paragraph in paragrafos:
        texto = limpar_texto(paragraph.text)
        tem_imagem = bool(paragrafo_tem_imagem(paragraph))

        if texto.startswith("Figura "):
            dentro_bloco_figura = True
            formatar_paragrafo_figura(paragraph, keep_with_next=True)
            continue

        if dentro_bloco_figura and tem_imagem:
            formatar_paragrafo_figura(paragraph, keep_with_next=True, tamanho_fonte=None)
            continue

        if dentro_bloco_figura and texto.startswith("Fonte:"):
            formatar_paragrafo_figura(paragraph, keep_with_next=True)
            continue

        if dentro_bloco_figura and (texto.startswith("Legenda:") or texto.startswith("Nota:")):
            formatar_paragrafo_figura(paragraph, keep_with_next=False)
            dentro_bloco_figura = False
            continue

        if dentro_bloco_figura and not texto:
            formatar_paragrafo_figura(paragraph, keep_with_next=True)
            continue

        if texto in {"Identificação do Equipamento", "Instalação e Chicotes", "Configurações", "Outros Registros"}:
            paragraph.paragraph_format.keep_with_next = True
            paragraph.paragraph_format.widow_control = True

    documento.save(str(caminho_docx))


def normalizar_imagem_para_docx(conteudo: bytes, caminho_saida: Path) -> Path:
    try:
        with Image.open(BytesIO(conteudo)) as imagem_original:
            imagem = ImageOps.exif_transpose(imagem_original)

            if imagem.mode in ("RGBA", "LA") or (imagem.mode == "P" and "transparency" in imagem.info):
                imagem_rgba = imagem.convert("RGBA")
                fundo = Image.new("RGB", imagem_rgba.size, "white")
                fundo.paste(imagem_rgba, mask=imagem_rgba.getchannel("A"))
                imagem = fundo
            else:
                imagem = imagem.convert("RGB")

            caminho_limpo = caminho_saida.with_suffix(".jpg")
            imagem.save(caminho_limpo, format="JPEG", quality=90, optimize=True, progressive=False)
            return caminho_limpo
    except UnidentifiedImageError as erro:
        raise ValueError("Uma das imagens enviadas não pôde ser lida. Tente reenviar em JPG ou PNG.") from erro
    except OSError as erro:
        raise ValueError("Uma das imagens está incompleta ou com metadados inválidos. Tente reenviar a foto ou tirar uma nova captura.") from erro


def salvar_upload(uploaded_file, pasta_temp: Path, prefixo: str, extensoes_permitidas: set[str], extensao_padrao: str) -> Path | None:
    if uploaded_file is None:
        return None

    nome_original = getattr(uploaded_file, "name", "")
    extensao = Path(nome_original).suffix.lower().lstrip(".")
    if extensao not in extensoes_permitidas:
        extensao = extensao_padrao

    conteudo = uploaded_file.getvalue()
    caminho = pasta_temp / f"{prefixo}_{uuid.uuid4().hex[:8]}.{extensao}"
    if extensoes_permitidas == EXTENSOES_IMAGEM:
        return normalizar_imagem_para_docx(conteudo, caminho)

    caminho.write_bytes(conteudo)
    return caminho


# ==========================================
# 5. Interface visual
# ==========================================
st.markdown("<h1 class='titulo-app'>🚜 Agres Relatórios</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitulo-app'>Geração de Relatórios Técnicos</p>", unsafe_allow_html=True)

with st.container(border=True):
    st.markdown("### 🎙️ 1. Relato Técnico")
    aba1, aba2 = st.tabs(["🔴 Gravar agora", "📁 Arquivos do celular"])

    with aba1:
        st.info("Grave por trechos. Use 'Remover' para apagar um áudio específico.")
        audios_rec = []
        for i, id_gravador in enumerate(st.session_state.lista_gravadores):
            col_gravador, col_excluir = st.columns([0.80, 0.20], vertical_alignment="bottom")
            with col_gravador:
                audio = st.audio_input(f"Trecho {i + 1}", key=f"rec_{id_gravador}_{st.session_state.reset_audio}")
                if audio:
                    audios_rec.append(audio)
            with col_excluir:
                if st.button("🗑️ Remover", key=f"btn_del_{id_gravador}", use_container_width=True):
                    st.session_state.lista_gravadores.remove(id_gravador)
                    st.rerun()

        col_novo, col_limpar = st.columns(2)
        if col_novo.button("➕ Novo trecho", use_container_width=True):
            st.session_state.lista_gravadores.append(st.session_state.proximo_id)
            st.session_state.proximo_id += 1
            st.rerun()
        if col_limpar.button("🧹 Limpar tudo", use_container_width=True):
            st.session_state.lista_gravadores = [st.session_state.proximo_id]
            st.session_state.proximo_id += 1
            st.session_state.reset_audio += 1
            st.session_state.relatorio_pronto = None
            st.rerun()

    with aba2:
        audios_up = st.file_uploader("Upload de áudios", type=list(EXTENSOES_AUDIO), accept_multiple_files=True)

    observacoes_texto = st.text_area(
        "Complemento técnico opcional",
        placeholder="Cliente, local, máquina, equipamento, séries, versões, falhas, parâmetros, calibrações ou pendências.",
        height=120,
    )

with st.container(border=True):
    st.markdown("### 🏷️ 2. Fotos do Cabeçalho")
    col_plaqueta, col_maquina, col_implemento = st.columns(3)
    f_plaqueta = col_plaqueta.file_uploader("📸 Informações do Equipamento", type=list(EXTENSOES_IMAGEM), key="up_p1")
    f_maquina = col_maquina.file_uploader("🚜 Máquina", type=list(EXTENSOES_IMAGEM), key="up_p2")
    f_implemento = col_implemento.file_uploader("🔧 Implemento", type=list(EXTENSOES_IMAGEM), key="up_p3")

with st.container(border=True):
    st.markdown("### 📸 3. Evidências Fotográficas")
    col_e1, col_e2 = st.columns(2)
    f_eq = col_e1.file_uploader("📋 Equipamento Agres", accept_multiple_files=True, type=list(EXTENSOES_IMAGEM))
    f_ins = col_e1.file_uploader("🔨 Instalação", accept_multiple_files=True, type=list(EXTENSOES_IMAGEM))
    f_conf = col_e2.file_uploader("⚙️ Configurações", accept_multiple_files=True, type=list(EXTENSOES_IMAGEM))
    f_out = col_e2.file_uploader("📂 Outros registros", accept_multiple_files=True, type=list(EXTENSOES_IMAGEM))

    with st.expander("Títulos e legendas das evidências", expanded=False):
        st.caption("Uma linha por foto, na ordem do upload. Formato: Título | legenda | fonte.")
        legendas_evidencias = {
            "fotos_equipamento": st.text_area(
                "Equipamento Agres",
                placeholder="Plaqueta de identificação da tela AgroNave 7 | Registro da série e versões do equipamento",
                height=90,
            ),
            "fotos_instalacao": st.text_area(
                "Instalação",
                placeholder="Roteamento do chicote principal | Chicote fixado no trator após instalação",
                height=90,
            ),
            "fotos_configuracao": st.text_area(
                "Configurações",
                placeholder="Tela de parâmetros do piloto | Configuração final utilizada durante os testes",
                height=90,
            ),
            "fotos_outros": st.text_area(
                "Outros registros",
                placeholder="Teste de campo após calibração | Validação operacional realizada com o cliente",
                height=90,
            ),
        }


# ==========================================
# 6. Execução
# ==========================================
audios_finais = audios_rec + (audios_up if audios_up else [])
entrada_disponivel = bool(audios_finais) or bool(limpar_texto(observacoes_texto))

if entrada_disponivel and st.button("Gerar Relatório Técnico"):
    st.session_state.relatorio_pronto = None
    st.session_state.nome_arquivo_pronto = None

    try:
        with tempfile.TemporaryDirectory() as pasta_temp_raw:
            pasta_temp = Path(pasta_temp_raw)

            with st.status("Processando dados e imagens...", expanded=True) as status:
                st.write("Salvando arquivos temporários.")
                caminhos_audio = []
                for i, audio in enumerate(audios_finais):
                    caminho_audio = salvar_upload(audio, pasta_temp, f"audio_{i}", EXTENSOES_AUDIO, "wav")
                    if caminho_audio:
                        caminhos_audio.append(caminho_audio)

                caminhos_cabecalho = {
                    "info_equip": salvar_upload(f_plaqueta, pasta_temp, "plaqueta", EXTENSOES_IMAGEM, "jpg"),
                    "maquina": salvar_upload(f_maquina, pasta_temp, "maquina", EXTENSOES_IMAGEM, "jpg"),
                    "implemento": salvar_upload(f_implemento, pasta_temp, "implemento", EXTENSOES_IMAGEM, "jpg"),
                }

                mapa_evidencias = {
                    "fotos_equipamento": f_eq,
                    "fotos_instalacao": f_ins,
                    "fotos_configuracao": f_conf,
                    "fotos_outros": f_out,
                }
                evidencias = {}
                for categoria, arquivos in mapa_evidencias.items():
                    evidencias[categoria] = []
                    for i, arquivo in enumerate(arquivos or []):
                        caminho_foto = salvar_upload(arquivo, pasta_temp, f"{categoria}_{i}", EXTENSOES_IMAGEM, "jpg")
                        if caminho_foto:
                            evidencias[categoria].append(caminho_foto)

                st.write("Extraindo e organizando informações técnicas.")
                dados = processar_atendimento_completo(caminhos_audio, observacoes_texto)

                st.write("Renderizando relatório Word.")
                arquivo_final = gerar_docx(dados, evidencias, caminhos_cabecalho, pasta_temp, legendas_evidencias)
                st.session_state.relatorio_pronto = arquivo_final.read_bytes()
                st.session_state.nome_arquivo_pronto = arquivo_final.name

                status.update(label="Relatório finalizado!", state="complete", expanded=False)

    except Exception as erro:
        st.error(f"Erro no processamento: {erro}")
        st.exception(erro)
elif not entrada_disponivel:
    st.caption("Adicione ao menos um áudio ou complemento escrito para gerar o relatório.")

if st.session_state.relatorio_pronto:
    st.success("✅ O laudo está pronto para download!")
    st.download_button(
        label="📥 Baixar Relatório (Word)",
        data=st.session_state.relatorio_pronto,
        file_name=st.session_state.nome_arquivo_pronto,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )
