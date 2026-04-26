import json
import re
import shutil
import tempfile
import unicodedata
import uuid
from datetime import date, datetime
from hashlib import sha1
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import google.generativeai as genai
import streamlit as st
from docx import Document
from docx.document import Document as DocumentClass
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
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
DRAFTS_DIR = BASE_DIR / ".rascunhos"

TAM_PLAQUETA = 60
TAM_MAQUINA = 32
TAM_EVIDENCIA = 120
FIGURA_CANVAS_PX = (1800, 1125)
FIGURAS_POR_PAGINA = 2

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
    "nome_arquivo_sugerido",
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
10. nome_arquivo_sugerido: montar no padrão "AAAAMMDD - CIDADE - UF - TIPO EQUIPAMENTO". Use a data inicial quando houver intervalo. Exemplos: "20250710 - SÃO JOSÉ DOS PINHAIS - PR - SUPORTE ISOBOX SPRAYER AGRONAVE 12" ou "20260119 - GUARANIAÇU - PR - SUPORTE AGRONAVE 7 ISOBOX SPRAYER".

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
    "nome_arquivo_sugerido": "",
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


def limpar_nome_relatorio(texto: str) -> str:
    nome = limpar_texto(texto)
    nome = nome.replace("\n", " ")
    nome = re.sub(r'[\\/*?:"<>|]', "", nome)
    nome = re.sub(r"\s*-\s*", " - ", nome)
    nome = re.sub(r"\s+", " ", nome).strip(" .-_")
    return (nome or "RELATÓRIO DE ATENDIMENTO")[:150]


def data_para_nome_arquivo(data_visita: str) -> str:
    texto = limpar_texto(data_visita)
    intervalo = re.search(r"\b(\d{1,2})\s*(?:a|até|-)\s*\d{1,2}/(\d{1,2})/(\d{4})\b", texto, flags=re.I)
    if intervalo:
        dia, mes, ano = intervalo.groups()
        return f"{int(ano):04d}{int(mes):02d}{int(dia):02d}"

    data_br = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", texto)
    if data_br:
        dia, mes, ano = data_br.groups()
        return f"{int(ano):04d}{int(mes):02d}{int(dia):02d}"

    data_iso = re.search(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b", texto)
    if data_iso:
        ano, mes, dia = data_iso.groups()
        return f"{int(ano):04d}{int(mes):02d}{int(dia):02d}"

    return data_atual_brasil().strftime("%Y%m%d")


def extrair_cidade_uf(cliente_local: str) -> tuple[str, str]:
    texto = limpar_texto(cliente_local)
    linhas = [linha.strip() for linha in texto.split("\n") if linha.strip()]

    for linha in linhas:
        match = re.search(r"(?:cidade(?:/uf)?|local(?: cliente)?|propriedade)\s*:\s*([^,\n/]+?)\s*/\s*([A-Za-z]{2})\b", linha, re.I)
        if match:
            return match.group(1).strip(), match.group(2).strip()

    for linha in linhas:
        match = re.search(r"\b([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\s.'-]{2,}?)\s*/\s*([A-Za-z]{2})\b", linha)
        if match and not linha.lower().startswith(("http", "www")):
            return match.group(1).strip(), match.group(2).strip()

    cidade = ""
    uf = ""
    for linha in linhas:
        match_cidade = re.search(r"cidade(?: revenda)?\s*:\s*(.+)", linha, re.I)
        if match_cidade and not cidade:
            cidade = match_cidade.group(1).strip()
        match_uf = re.search(r"\b(?:uf|estado)\s*:\s*([A-Za-z]{2})\b", linha, re.I)
        if match_uf:
            uf = match_uf.group(1).strip()

    if cidade and "/" in cidade:
        partes = [parte.strip() for parte in cidade.rsplit("/", 1)]
        if len(partes) == 2 and re.fullmatch(r"[A-Za-z]{2}", partes[1]):
            return partes[0], partes[1]

    if cidade and not uf:
        match_colado = re.search(r"(.+?)([A-Za-z]{2})$", cidade)
        if match_colado and len(match_colado.group(1).strip()) > 3:
            cidade = match_colado.group(1).strip(" ,-/")
            uf = match_colado.group(2)

    return cidade or "LOCAL NÃO INFORMADO", uf or "UF"


def tipo_atendimento_para_nome(dados: dict) -> str:
    tipos = []
    if dados.get("suporte") == "X":
        tipos.append("SUPORTE")
    if dados.get("instalacao") == "X":
        tipos.append("INSTALAÇÃO")
    if dados.get("treinamento") == "X":
        tipos.append("TREINAMENTO")
    return " + ".join(tipos) if tipos else "ATENDIMENTO"


def equipamento_para_nome(dados: dict) -> str:
    texto = normalizar_busca("\n".join([dados.get("objetivos", ""), dados.get("equipamentos", ""), dados.get("maquinas", "")]))
    encontrados = []
    padroes = [
        ("ISOBOX SPRAYER", r"\bisobox\s+sprayer\b"),
        ("AGRONAVE 12", r"\b(?:agronave|agronave|agn)\s*12\b"),
        ("AGRONAVE 7", r"\b(?:agronave|agronave|agn)\s*7\b"),
        ("ISOPILOT", r"\bisopilot\b"),
        ("ANP40", r"\banp40\b"),
        ("ANP21", r"\banp21\b"),
    ]
    for nome, padrao in padroes:
        if re.search(padrao, texto) and nome not in encontrados:
            encontrados.append(nome)

    if encontrados:
        return " ".join(encontrados)

    primeira_linha = next((linha for linha in limpar_texto(dados.get("equipamentos", "")).split("\n") if linha.strip()), "")
    primeira_linha = re.sub(r"^(modelo|tela|equipamento|ecu)\s*:\s*", "", primeira_linha, flags=re.I)
    return primeira_linha or "EQUIPAMENTO AGRES"


def gerar_nome_arquivo_relatorio(dados: dict) -> str:
    sugerido = limpar_nome_relatorio(dados.get("nome_arquivo_sugerido", ""))
    if re.match(r"^\d{8}\s+-\s+", sugerido):
        return sugerido.upper()

    data_nome = data_para_nome_arquivo(dados.get("data_visita", ""))
    cidade, uf = extrair_cidade_uf(dados.get("cliente_local", ""))
    partes = [
        data_nome,
        cidade.upper(),
        uf.upper(),
        f"{tipo_atendimento_para_nome(dados)} {equipamento_para_nome(dados)}".strip().upper(),
    ]
    return limpar_nome_relatorio(" - ".join(partes)).upper()


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

    nome_arquivo = f"{gerar_nome_arquivo_relatorio(dados_render)}.docx"
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
    formato.space_before = Pt(2)
    formato.space_after = Pt(2)
    if tamanho_fonte:
        for run in paragraph.runs:
            run.font.name = "Arial"
            run.font.size = Pt(tamanho_fonte)


def aplicar_paginacao_abnt_figuras(caminho_docx: Path) -> None:
    documento = Document(str(caminho_docx))
    paragrafos = list(iterar_paragrafos_word(documento))
    total_figuras = sum(1 for paragraph in paragrafos if limpar_texto(paragraph.text).startswith("Figura "))

    dentro_bloco_figura = False
    numero_figura = 0
    for paragraph in paragrafos:
        texto = limpar_texto(paragraph.text)
        tem_imagem = bool(paragrafo_tem_imagem(paragraph))

        if texto.startswith("Figura "):
            dentro_bloco_figura = True
            numero_figura += 1
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
            if numero_figura % FIGURAS_POR_PAGINA == 0 and numero_figura < total_figuras:
                paragraph.add_run().add_break(WD_BREAK.PAGE)
            dentro_bloco_figura = False
            continue

        if dentro_bloco_figura and not texto:
            formatar_paragrafo_figura(paragraph, keep_with_next=True)
            continue

        if texto in {"Identificação do Equipamento", "Instalação e Chicotes", "Configurações", "Outros Registros"}:
            paragraph.paragraph_format.keep_with_next = True
            paragraph.paragraph_format.widow_control = True

    documento.save(str(caminho_docx))


def normalizar_imagem_para_docx(conteudo: bytes, caminho_saida: Path, padronizar_figura: bool = False) -> Path:
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

            if padronizar_figura:
                quadro = Image.new("RGB", FIGURA_CANVAS_PX, "white")
                imagem.thumbnail(FIGURA_CANVAS_PX, Image.Resampling.LANCZOS)
                x = (FIGURA_CANVAS_PX[0] - imagem.width) // 2
                y = (FIGURA_CANVAS_PX[1] - imagem.height) // 2
                quadro.paste(imagem, (x, y))
                imagem = quadro

            caminho_limpo = caminho_saida.with_suffix(".jpg")
            imagem.save(caminho_limpo, format="JPEG", quality=90, optimize=True, progressive=False)
            return caminho_limpo
    except UnidentifiedImageError as erro:
        raise ValueError("Uma das imagens enviadas não pôde ser lida. Tente reenviar em JPG ou PNG.") from erro
    except OSError as erro:
        raise ValueError("Uma das imagens está incompleta ou com metadados inválidos. Tente reenviar a foto ou tirar uma nova captura.") from erro


def salvar_upload(
    uploaded_file,
    pasta_temp: Path,
    prefixo: str,
    extensoes_permitidas: set[str],
    extensao_padrao: str,
    padronizar_figura: bool = False,
) -> Path | None:
    if uploaded_file is None:
        return None

    nome_original = getattr(uploaded_file, "name", "")
    extensao = Path(nome_original).suffix.lower().lstrip(".")
    if extensao not in extensoes_permitidas:
        extensao = extensao_padrao

    conteudo = uploaded_file.getvalue()
    caminho = pasta_temp / f"{prefixo}_{uuid.uuid4().hex[:8]}.{extensao}"
    if extensoes_permitidas == EXTENSOES_IMAGEM:
        return normalizar_imagem_para_docx(conteudo, caminho, padronizar_figura)

    caminho.write_bytes(conteudo)
    return caminho


def novo_manifesto_rascunho() -> dict:
    return {
        "audios": [],
        "cabecalho": {"info_equip": None, "maquina": None, "implemento": None},
        "evidencias": {categoria: [] for categoria in CATEGORIAS_EVIDENCIAS},
        "observacoes": "",
        "legendas_evidencias": {categoria: "" for categoria in CATEGORIAS_EVIDENCIAS},
    }


def obter_id_rascunho() -> str:
    draft_url = st.query_params.get("draft", "")
    if isinstance(draft_url, list):
        draft_url = draft_url[0] if draft_url else ""

    if "draft_id" in st.session_state:
        draft_id = st.session_state.draft_id
    elif re.fullmatch(r"[0-9a-f]{12}", str(draft_url)):
        draft_id = str(draft_url)
    else:
        draft_id = uuid.uuid4().hex[:12]

    st.session_state.draft_id = draft_id
    st.query_params["draft"] = draft_id
    return draft_id


def pasta_rascunho_atual() -> Path:
    draft_dir = DRAFTS_DIR / obter_id_rascunho()
    draft_dir.mkdir(parents=True, exist_ok=True)
    return draft_dir


def caminho_manifesto(draft_dir: Path) -> Path:
    return draft_dir / "manifest.json"


def carregar_manifesto(draft_dir: Path) -> dict:
    manifesto = novo_manifesto_rascunho()
    caminho = caminho_manifesto(draft_dir)
    if caminho.exists():
        try:
            dados = json.loads(caminho.read_text(encoding="utf-8"))
            if isinstance(dados, dict):
                manifesto.update(dados)
                manifesto["cabecalho"] = {**novo_manifesto_rascunho()["cabecalho"], **manifesto.get("cabecalho", {})}
                manifesto["evidencias"] = {**novo_manifesto_rascunho()["evidencias"], **manifesto.get("evidencias", {})}
                manifesto["legendas_evidencias"] = {
                    **novo_manifesto_rascunho()["legendas_evidencias"],
                    **manifesto.get("legendas_evidencias", {}),
                }
        except (OSError, json.JSONDecodeError):
            pass
    return manifesto


def salvar_manifesto(draft_dir: Path, manifesto: dict) -> None:
    caminho_manifesto(draft_dir).write_text(
        json.dumps(manifesto, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def resolver_arquivo_rascunho(draft_dir: Path, item) -> Path | None:
    if not item:
        return None
    rel_path = item.get("path") if isinstance(item, dict) else str(item)
    caminho = (draft_dir / rel_path).resolve()
    try:
        caminho.relative_to(draft_dir.resolve())
    except ValueError:
        return None
    return caminho if caminho.exists() else None


def salvar_arquivo_rascunho(
    uploaded_file,
    draft_dir: Path,
    subpasta: str,
    prefixo: str,
    extensoes_permitidas: set[str],
    extensao_padrao: str,
    padronizar_figura: bool = False,
) -> dict | None:
    if uploaded_file is None:
        return None

    conteudo = uploaded_file.getvalue()
    if not conteudo:
        return None

    nome_original = getattr(uploaded_file, "name", f"{prefixo}.{extensao_padrao}")
    extensao = Path(nome_original).suffix.lower().lstrip(".")
    if extensao not in extensoes_permitidas:
        extensao = extensao_padrao

    pasta = draft_dir / subpasta
    pasta.mkdir(parents=True, exist_ok=True)
    digest = sha1(conteudo).hexdigest()[:12]
    caminho = pasta / f"{prefixo}_{digest}.{extensao}"

    if extensoes_permitidas == EXTENSOES_IMAGEM:
        caminho = normalizar_imagem_para_docx(conteudo, caminho, padronizar_figura)
    elif not caminho.exists():
        caminho.write_bytes(conteudo)

    return {
        "name": nome_original,
        "path": str(caminho.relative_to(draft_dir)),
        "size": len(conteudo),
    }


def atualizar_lista_rascunho(manifesto: dict, chave: str, arquivos, draft_dir: Path, subpasta: str, extensoes: set[str], extensao_padrao: str) -> None:
    if not arquivos:
        return

    itens = []
    for indice, arquivo in enumerate(arquivos):
        item = salvar_arquivo_rascunho(arquivo, draft_dir, subpasta, f"{chave}_{indice}", extensoes, extensao_padrao)
        if item:
            itens.append(item)

    if itens:
        manifesto[chave] = itens


def atualizar_rascunho_atual(
    draft_dir: Path,
    manifesto: dict,
    audios,
    cabecalho: dict,
    evidencias_upload: dict,
    observacoes: str,
    legendas: dict,
) -> dict:
    if audios:
        manifesto["audios"] = [
            item
            for indice, audio in enumerate(audios)
            if (item := salvar_arquivo_rascunho(audio, draft_dir, "audios", f"audio_{indice}", EXTENSOES_AUDIO, "wav"))
        ]

    for chave, arquivo in cabecalho.items():
        item = salvar_arquivo_rascunho(arquivo, draft_dir, "cabecalho", chave, EXTENSOES_IMAGEM, "jpg")
        if item:
            manifesto["cabecalho"][chave] = item

    for categoria, arquivos in evidencias_upload.items():
        if arquivos:
            itens = []
            for indice, arquivo in enumerate(arquivos):
                item = salvar_arquivo_rascunho(
                    arquivo,
                    draft_dir,
                    categoria,
                    f"{categoria}_{indice}",
                    EXTENSOES_IMAGEM,
                    "jpg",
                    padronizar_figura=True,
                )
                if item:
                    itens.append(item)
            if itens:
                manifesto["evidencias"][categoria] = itens

    if limpar_texto(observacoes):
        manifesto["observacoes"] = observacoes

    for categoria, texto in (legendas or {}).items():
        if limpar_texto(texto):
            manifesto["legendas_evidencias"][categoria] = texto

    salvar_manifesto(draft_dir, manifesto)
    return manifesto


def caminhos_salvos_rascunho(draft_dir: Path, manifesto: dict) -> tuple[list[Path], dict, dict]:
    audios = [caminho for item in manifesto.get("audios", []) if (caminho := resolver_arquivo_rascunho(draft_dir, item))]
    cabecalho = {
        chave: resolver_arquivo_rascunho(draft_dir, item)
        for chave, item in manifesto.get("cabecalho", {}).items()
    }
    evidencias = {
        categoria: [
            normalizar_imagem_para_docx(caminho.read_bytes(), caminho, padronizar_figura=True)
            for item in manifesto.get("evidencias", {}).get(categoria, [])
            if (caminho := resolver_arquivo_rascunho(draft_dir, item))
        ]
        for categoria in CATEGORIAS_EVIDENCIAS
    }
    return audios, cabecalho, evidencias


def contar_evidencias(manifesto: dict) -> int:
    return sum(len(itens or []) for itens in manifesto.get("evidencias", {}).values())


def limpar_rascunho_atual() -> None:
    draft_dir = DRAFTS_DIR / st.session_state.draft_id
    if draft_dir.exists():
        shutil.rmtree(draft_dir)
    st.session_state.draft_id = uuid.uuid4().hex[:12]
    st.query_params["draft"] = st.session_state.draft_id
    for chave in ("observacoes_texto", *[f"legenda_{categoria}" for categoria in CATEGORIAS_EVIDENCIAS]):
        st.session_state.pop(chave, None)
    st.session_state.relatorio_pronto = None
    st.session_state.nome_arquivo_pronto = None


# ==========================================
# 5. Interface visual
# ==========================================
draft_dir = pasta_rascunho_atual()
manifesto_rascunho = carregar_manifesto(draft_dir)

if "observacoes_texto" not in st.session_state:
    st.session_state.observacoes_texto = manifesto_rascunho.get("observacoes", "")

for categoria in CATEGORIAS_EVIDENCIAS:
    chave_legenda = f"legenda_{categoria}"
    if chave_legenda not in st.session_state:
        st.session_state[chave_legenda] = manifesto_rascunho.get("legendas_evidencias", {}).get(categoria, "")

st.markdown("<h1 class='titulo-app'>🚜 Agres Relatórios</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitulo-app'>Geração de Relatórios Técnicos</p>", unsafe_allow_html=True)

with st.container(border=True):
    col_status, col_limpar_rascunho = st.columns([0.72, 0.28], vertical_alignment="center")
    with col_status:
        st.caption(
            f"Rascunho ativo: {st.session_state.draft_id} | "
            f"Áudios salvos: {len(manifesto_rascunho.get('audios', []))} | "
            f"Fotos salvas: {contar_evidencias(manifesto_rascunho)}"
        )
    with col_limpar_rascunho:
        if st.button("Novo rascunho", use_container_width=True):
            limpar_rascunho_atual()
            st.rerun()

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
        key="observacoes_texto",
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
                key="legenda_fotos_equipamento",
            ),
            "fotos_instalacao": st.text_area(
                "Instalação",
                placeholder="Roteamento do chicote principal | Chicote fixado no trator após instalação",
                height=90,
                key="legenda_fotos_instalacao",
            ),
            "fotos_configuracao": st.text_area(
                "Configurações",
                placeholder="Tela de parâmetros do piloto | Configuração final utilizada durante os testes",
                height=90,
                key="legenda_fotos_configuracao",
            ),
            "fotos_outros": st.text_area(
                "Outros registros",
                placeholder="Teste de campo após calibração | Validação operacional realizada com o cliente",
                height=90,
                key="legenda_fotos_outros",
            ),
        }


# ==========================================
# 6. Execução
# ==========================================
audios_finais = audios_rec + (audios_up if audios_up else [])
uploads_cabecalho = {
    "info_equip": f_plaqueta,
    "maquina": f_maquina,
    "implemento": f_implemento,
}
uploads_evidencias = {
    "fotos_equipamento": f_eq,
    "fotos_instalacao": f_ins,
    "fotos_configuracao": f_conf,
    "fotos_outros": f_out,
}

manifesto_rascunho = atualizar_rascunho_atual(
    draft_dir,
    manifesto_rascunho,
    audios_finais,
    uploads_cabecalho,
    uploads_evidencias,
    observacoes_texto,
    legendas_evidencias,
)
caminhos_audio_salvos, caminhos_cabecalho_salvos, evidencias_salvas = caminhos_salvos_rascunho(
    draft_dir,
    manifesto_rascunho,
)
observacoes_salvas = limpar_texto(observacoes_texto) or manifesto_rascunho.get("observacoes", "")
legendas_salvas = {
    categoria: limpar_texto(legendas_evidencias.get(categoria, ""))
    or manifesto_rascunho.get("legendas_evidencias", {}).get(categoria, "")
    for categoria in CATEGORIAS_EVIDENCIAS
}

st.caption(
    f"Rascunho salvo agora: {len(caminhos_audio_salvos)} áudio(s), "
    f"{sum(len(lista) for lista in evidencias_salvas.values())} foto(s) de evidência."
)

entrada_disponivel = bool(caminhos_audio_salvos) or bool(limpar_texto(observacoes_salvas))

if entrada_disponivel and st.button("Gerar Relatório Técnico"):
    st.session_state.relatorio_pronto = None
    st.session_state.nome_arquivo_pronto = None

    try:
        with tempfile.TemporaryDirectory() as pasta_temp_raw:
            pasta_temp = Path(pasta_temp_raw)

            with st.status("Processando dados e imagens...", expanded=True) as status:
                st.write("Usando arquivos salvos no rascunho.")
                st.write("Extraindo e organizando informações técnicas.")
                dados = processar_atendimento_completo(caminhos_audio_salvos, observacoes_salvas)

                st.write("Renderizando relatório Word.")
                arquivo_final = gerar_docx(
                    dados,
                    evidencias_salvas,
                    caminhos_cabecalho_salvos,
                    pasta_temp,
                    legendas_salvas,
                )
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
