import streamlit as st
import pandas as pd
import os
from openai import OpenAI
import json
import io
import re
import unicodedata
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.platypus import (
    Paragraph, Preformatted, SimpleDocTemplate, Spacer, Table, TableStyle
)

st.set_page_config(page_title="🚀 Grok + Produção Excel", page_icon="🚀", layout="wide")

# ====================== CARREGAR EXCEL ======================
@st.cache_data
def carregar_excel():
    try:
        df = pd.read_excel("ProducaoIndustrial2026.xlsx")
        df['Data'] = pd.to_datetime(df['Data'])
        st.success(f"✅ Planilha carregada com {len(df)} registros!")
        return df
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo: {str(e)}")
        st.stop()

df = carregar_excel()
person_cols = [c for c in df.columns if c not in ['Data', 'Time', 'Terminal MP35', 'Terminal MP30']]

st.title("🚀 Wald Analisando sua Planilha de Produção")
st.caption("Wals entende a planilha, faz cálculos, gera insights e gráficos")

with st.sidebar:
    st.header("Configurações")
    api_key = st.text_input("xAI API Key", type="password", value=os.getenv("XAI_API_KEY", ""))
    if api_key:
        os.environ["XAI_API_KEY"] = api_key
   
    st.divider()
    st.write(f"**Período:** {df['Data'].min().date()} a {df['Data'].max().date()}")
    st.write(f"**Registros:** {len(df)}")
    st.write(f"**Times:** 5")
    st.write(f"**Colaboradores:** 20")

# ====================== FUNÇÕES / TOOLS ======================
def get_total_production(start_date=None, end_date=None, team=None):
    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['Data'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['Data'] <= pd.to_datetime(end_date)]
    if team:
        filtered = filtered[filtered['Time'] == team]
   
    total_mp35 = filtered['Terminal MP35'].sum()
    total_mp30 = filtered['Terminal MP30'].sum()
    return {
        "total_mp35": int(total_mp35),
        "total_mp30": int(total_mp30),
        "total_geral": int(total_mp35 + total_mp30),
        "periodo": f"{start_date or 'início'} até {end_date or 'fim'}",
        "time": team if team else "Todos"
    }

def get_top_team(start_date=None, end_date=None, n=5):
    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['Data'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['Data'] <= pd.to_datetime(end_date)]
   
    team_totals = filtered.groupby('Time')[['Terminal MP35', 'Terminal MP30']].sum()
    team_totals['Total'] = team_totals.sum(axis=1)
    ranking = team_totals.sort_values('Total', ascending=False).head(n)
   
    result = {}
    for idx, row in ranking.iterrows():
        result[f"Time {int(idx)}"] = {
            "total": int(row['Total']),
            "mp35": int(row['Terminal MP35']),
            "mp30": int(row['Terminal MP30'])
        }
    return result

def get_top_person(start_date=None, end_date=None, n=5):
    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['Data'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['Data'] <= pd.to_datetime(end_date)]
   
    person_totals = filtered[person_cols].sum().sort_values(ascending=False)
    ranking = person_totals.head(n)
   
    result = {}
    for idx, (name, value) in enumerate(ranking.items()):
        result[f"{idx+1}º"] = {
            "nome": name,
            "total_pecas": int(value)
        }
    return result

def get_bottom_person(start_date=None, end_date=None, n=5):
    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['Data'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['Data'] <= pd.to_datetime(end_date)]
   
    person_totals = filtered[person_cols].sum().sort_values(ascending=True)
    ranking = person_totals.head(n)
   
    result = {}
    for idx, (name, value) in enumerate(ranking.items()):
        result[f"{idx+1}º"] = {
            "nome": name,
            "total_pecas": int(value)
        }
    return result

def get_daily_production(start_date=None, end_date=None):
    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['Data'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['Data'] <= pd.to_datetime(end_date)]
   
    daily = filtered.groupby('Data')[['Terminal MP35', 'Terminal MP30']].sum()
    daily['Total'] = daily.sum(axis=1)
    return daily.reset_index().to_dict('records')[:15]

def list_colaboradores():
    team_mapping = {}
    for team in range(1, 6):
        team_people = person_cols[(team-1)*4 : team*4]
        team_mapping[f"Time {team}"] = team_people
    return team_mapping

# ====================== GROK COM TOOL CALLING ======================
tools = [
    {
        "type": "function",
        "function": {
            "name": "get_total_production",
            "description": "Calcula a produção total (MP35 + MP30). Use para totais gerais ou por período/time.",
            "parameters": {
                "type": "object",
                "properties": {
                    "start_date": {"type": "string"},
                    "end_date": {"type": "string"},
                    "team": {"type": "integer"}
                }
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_top_team",
            "description": "Retorna os N melhores times por produção total (padrão = 5).",
            "parameters": {
                "type": "object",
                "properties": {
                    "start_date": {"type": "string"},
                    "end_date": {"type": "string"},
                    "n": {"type": "integer"}
                }
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_top_person",
            "description": "Retorna os N melhores colaboradores por produção (padrão = 5).",
            "parameters": {
                "type": "object",
                "properties": {
                    "start_date": {"type": "string"},
                    "end_date": {"type": "string"},
                    "n": {"type": "integer"}
                }
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_bottom_person",
            "description": "Retorna os N colaboradores que MENOS produziram (padrão = 5).",
            "parameters": {
                "type": "object",
                "properties": {
                    "start_date": {"type": "string"},
                    "end_date": {"type": "string"},
                    "n": {"type": "integer"}
                }
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_daily_production",
            "description": "Mostra a produção diária.",
            "parameters": {
                "type": "object",
                "properties": {
                    "start_date": {"type": "string"},
                    "end_date": {"type": "string"}
                }
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "list_colaboradores",
            "description": "Lista todos os colaboradores e seus times.",
            "parameters": {"type": "object", "properties": {}}
        }
    }
]

available_functions = {
    "get_total_production": get_total_production,
    "get_top_team": get_top_team,
    "get_top_person": get_top_person,
    "get_bottom_person": get_bottom_person,
    "get_daily_production": get_daily_production,
    "list_colaboradores": list_colaboradores
}

def chat_with_grok(user_question):
    client = OpenAI(
        api_key=os.getenv("XAI_API_KEY"),
        base_url="https://api.x.ai/v1"
    )
   
    system_prompt = """Você é um analista de produção industrial muito competente.
    Use as ferramentas disponíveis para fazer cálculos precisos e gerar bons insights.
    Sempre responda em português de forma clara e profissional."""

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_question}
    ]

    response = client.chat.completions.create(
        model="grok-4",
        messages=messages,
        tools=tools,
        tool_choice="auto"
    )

    response_message = response.choices[0].message
    tool_calls = response_message.tool_calls

    if tool_calls:
        messages.append(response_message)
        
        for tool_call in tool_calls:
            function_name = tool_call.function.name
            function_args = json.loads(tool_call.function.arguments)
            
            function_to_call = available_functions[function_name]
            function_response = function_to_call(**function_args)
            
            messages.append({
                "tool_call_id": tool_call.id,
                "role": "tool",
                "name": function_name,
                "content": json.dumps(function_response, default=str)
            })
        
        second_response = client.chat.completions.create(
            model="grok-4",
            messages=messages
        )
        return second_response.choices[0].message.content
   
    return response_message.content

# ====================== EXPORTAÇÃO (TXT / PDF) ======================
AZUL = colors.HexColor("#1F3A5F")
AZUL_CLARO = colors.HexColor("#F2F5F9")
CINZA_TXT = colors.HexColor("#24292F")
CINZA_LINHA = colors.HexColor("#C9D2DD")
CINZA_FRACO = colors.HexColor("#8A8F98")

MARGEM = 18 * mm
FAIXA = 26 * mm
LARGURA_UTIL = A4[0] - 2 * MARGEM

# As fontes internas do reportlab usam WinAnsi (cp1252). O que não couber nessa
# tabela — emoji, setas, sinais matemáticos — viraria quadrado preto no PDF.
TROCAS_PDF = {
    "→": "->", "←": "<-", "↔": "<->", "⇒": "=>", "↑": "^", "↓": "v",
    "≥": ">=", "≤": "<=", "≠": "!=", "≈": "~", "×": "x", "…": "...",
    "✓": "v", "✔": "v", "✗": "x", "✘": "x", "☑": "[x]", "☐": "[ ]",
}


def _sanitizar_pdf(texto):
    """Remove ou substitui o que a fonte do PDF não consegue desenhar."""
    saida = []
    for ch in texto:
        if ch in "\n\t":
            saida.append(ch)
            continue
        try:
            ch.encode("cp1252")
        except UnicodeEncodeError:
            saida.append(TROCAS_PDF.get(ch, ""))
        else:
            saida.append(ch)
    return "".join(saida)


def _inline(texto):
    """Converte a marcação inline do markdown nas tags que o Paragraph entende."""
    texto = _sanitizar_pdf(texto)
    texto = texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    texto = re.sub(r"`([^`]+)`", r'<font face="Courier" size="9">\1</font>', texto)
    texto = re.sub(r"\*\*\*(.+?)\*\*\*", r"<b><i>\1</i></b>", texto)
    texto = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", texto)
    texto = re.sub(r"(?<!\*)\*([^*\n]+?)\*(?!\*)", r"<i>\1</i>", texto)
    texto = re.sub(r"(?<![A-Za-z0-9_])_([^_\n]+?)_(?![A-Za-z0-9_])", r"<i>\1</i>", texto)
    texto = re.sub(r"\[([^\]]+)\]\((https?://[^)]+)\)",
                   r'<link href="\2" color="#1A6DD8">\1</link>', texto)
    return re.sub(r" {2,}", " ", texto).strip()


def _estilos():
    corpo = ParagraphStyle("corpo", fontName="Helvetica", fontSize=10, leading=14.5,
                           spaceAfter=6, textColor=CINZA_TXT, alignment=TA_JUSTIFY)
    return {
        "corpo": corpo,
        "h1": ParagraphStyle("h1", parent=corpo, fontName="Helvetica-Bold", fontSize=14,
                             leading=18, spaceBefore=12, spaceAfter=6, textColor=AZUL,
                             alignment=TA_LEFT),
        "h2": ParagraphStyle("h2", parent=corpo, fontName="Helvetica-Bold", fontSize=12,
                             leading=16, spaceBefore=10, spaceAfter=5, textColor=AZUL,
                             alignment=TA_LEFT),
        "h3": ParagraphStyle("h3", parent=corpo, fontName="Helvetica-Bold", fontSize=10.5,
                             leading=14, spaceBefore=8, spaceAfter=4, textColor=CINZA_TXT,
                             alignment=TA_LEFT),
        "li": ParagraphStyle("li", parent=corpo, leftIndent=14, bulletIndent=4,
                             spaceAfter=3, alignment=TA_LEFT),
        "citacao": ParagraphStyle("citacao", parent=corpo, leftIndent=12, fontName="Helvetica-Oblique",
                                  textColor=colors.HexColor("#4A5560"), alignment=TA_LEFT),
        "rotulo": ParagraphStyle("rotulo", fontName="Helvetica-Bold", fontSize=8,
                                 leading=11, textColor=CINZA_FRACO, spaceAfter=4),
        "pergunta": ParagraphStyle("pergunta", parent=corpo, fontName="Helvetica-Oblique",
                                   fontSize=10.5, alignment=TA_LEFT, spaceAfter=0),
        "codigo": ParagraphStyle("codigo", fontName="Courier", fontSize=8.5, leading=11,
                                 textColor=CINZA_TXT),
        "th": ParagraphStyle("th", fontName="Helvetica-Bold", fontSize=8.5, leading=11,
                             textColor=colors.white, alignment=TA_LEFT),
        "td_left": ParagraphStyle("td_left", fontName="Helvetica", fontSize=8.5, leading=11,
                                  textColor=CINZA_TXT, alignment=TA_LEFT),
        "td_center": ParagraphStyle("td_center", fontName="Helvetica", fontSize=8.5, leading=11,
                                    textColor=CINZA_TXT, alignment=TA_CENTER),
        "td_right": ParagraphStyle("td_right", fontName="Helvetica", fontSize=8.5, leading=11,
                                   textColor=CINZA_TXT, alignment=TA_RIGHT),
    }


def _e_separador(linha):
    linha = linha.strip()
    return "-" in linha and bool(re.fullmatch(r"\|?[\s:|-]+\|?", linha)) and "|" in linha


def _celulas(linha):
    return [c.strip() for c in linha.strip().strip("|").split("|")]


def _montar_tabela(bloco, est):
    """Transforma uma tabela markdown em uma Table do platypus."""
    cabecalho = _celulas(bloco[0])
    n = len(cabecalho)

    alinhamentos = []
    for spec in _celulas(bloco[1])[:n]:
        if spec.startswith(":") and spec.endswith(":"):
            alinhamentos.append("td_center")
        elif spec.endswith(":"):
            alinhamentos.append("td_right")
        else:
            alinhamentos.append("td_left")
    alinhamentos += ["td_left"] * (n - len(alinhamentos))

    corpo = []
    for linha in bloco[2:]:
        celulas = _celulas(linha)[:n]
        corpo.append(celulas + [""] * (n - len(celulas)))

    # Largura de cada coluna proporcional ao conteúdo, com piso e teto para
    # nenhuma coluna espremer as outras.
    pesos = []
    for c in range(n):
        maior = max([len(cabecalho[c])] + [len(linha[c]) for linha in corpo] or [1])
        pesos.append(max(6, min(maior, 38)))
    soma = sum(pesos)
    larguras = [LARGURA_UTIL * p / soma for p in pesos]

    dados = [[Paragraph(_inline(t), est["th"]) for t in cabecalho]]
    for linha in corpo:
        dados.append([Paragraph(_inline(t), est[alinhamentos[c]]) for c, t in enumerate(linha)])

    tabela = Table(dados, colWidths=larguras, repeatRows=1, hAlign="LEFT")
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, AZUL_CLARO]),
        ("GRID", (0, 0), (-1, -1), 0.4, CINZA_LINHA),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
    ]))
    return tabela


def _caixa(flowable, fundo, borda_esq=None):
    tabela = Table([[flowable]], colWidths=[LARGURA_UTIL], hAlign="LEFT")
    estilo = [
        ("BACKGROUND", (0, 0), (-1, -1), fundo),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]
    if borda_esq:
        estilo.append(("LINEBEFORE", (0, 0), (0, -1), 2.5, borda_esq))
    tabela.setStyle(TableStyle(estilo))
    return tabela


def _markdown_para_flowables(md, est):
    """Percorre o markdown linha a linha e devolve os elementos do documento."""
    flow = []
    linhas = _sanitizar_pdf(md).replace("\r\n", "\n").split("\n")
    i = 0

    while i < len(linhas):
        linha = linhas[i]
        bruta = linha.strip()

        if not bruta:
            i += 1
            continue

        if bruta.startswith("```"):
            i += 1
            codigo = []
            while i < len(linhas) and not linhas[i].strip().startswith("```"):
                codigo.append(linhas[i])
                i += 1
            i += 1
            if codigo:
                flow.append(_caixa(Preformatted("\n".join(codigo), est["codigo"]), AZUL_CLARO))
                flow.append(Spacer(1, 6))
            continue

        if bruta.startswith("|") and i + 1 < len(linhas) and _e_separador(linhas[i + 1]):
            bloco = []
            while i < len(linhas) and linhas[i].strip().startswith("|"):
                bloco.append(linhas[i])
                i += 1
            flow.append(Spacer(1, 3))
            flow.append(_montar_tabela(bloco, est))
            flow.append(Spacer(1, 8))
            continue

        if re.fullmatch(r"(-{3,}|\*{3,}|_{3,})", bruta):
            regua = Table([[""]], colWidths=[LARGURA_UTIL], rowHeights=[1])
            regua.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), CINZA_LINHA)]))
            flow += [Spacer(1, 6), regua, Spacer(1, 8)]
            i += 1
            continue

        titulo = re.match(r"^(#{1,6})\s+(.*)", bruta)
        if titulo:
            nivel = min(len(titulo.group(1)), 3)
            flow.append(Paragraph(_inline(titulo.group(2)), est[f"h{nivel}"]))
            i += 1
            continue

        if bruta.startswith(">"):
            flow.append(_caixa(Paragraph(_inline(bruta.lstrip("> ")), est["citacao"]),
                               AZUL_CLARO, borda_esq=AZUL))
            flow.append(Spacer(1, 6))
            i += 1
            continue

        item = re.match(r"^([-*+]|\d+[.)])\s+(.+)", bruta)
        if item:
            nivel = min((len(linha) - len(linha.lstrip())) // 2, 3)
            marcador = "•" if item.group(1) in "-*+" else item.group(1)
            estilo = ParagraphStyle(
                f"li{nivel}", parent=est["li"],
                leftIndent=est["li"].leftIndent + nivel * 12,
                bulletIndent=est["li"].bulletIndent + nivel * 12,
            )
            flow.append(Paragraph(_inline(item.group(2)), estilo, bulletText=marcador))
            i += 1
            continue

        flow.append(Paragraph(_inline(bruta), est["corpo"]))
        i += 1

    return flow


class _CanvasNumerado(pdfcanvas.Canvas):
    """Guarda as páginas para poder escrever 'Página N de M' no rodapé."""
    rodape_esquerda = ""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._paginas = []

    def showPage(self):
        self._paginas.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._paginas)
        for estado in self._paginas:
            self.__dict__.update(estado)
            self._desenhar_rodape(total)
            super().showPage()
        super().save()

    def _desenhar_rodape(self, total):
        self.saveState()
        self.setStrokeColor(CINZA_LINHA)
        self.setLineWidth(0.5)
        self.line(MARGEM, 15 * mm, A4[0] - MARGEM, 15 * mm)
        self.setFont("Helvetica", 7.5)
        self.setFillColor(CINZA_FRACO)
        self.drawString(MARGEM, 11 * mm, self.rodape_esquerda)
        self.drawRightString(A4[0] - MARGEM, 11 * mm, f"Página {self._pageNumber} de {total}")
        self.restoreState()


def gerar_pdf(pergunta, resposta, gerado_em):
    """Monta o PDF de uma troca (pergunta + resposta) e devolve os bytes."""
    est = _estilos()
    periodo = f"{df['Data'].min():%d/%m/%Y} a {df['Data'].max():%d/%m/%Y}"
    subtitulo = (f"Gerado em {gerado_em:%d/%m/%Y às %H:%M}  |  "
                 f"Dados de {periodo}  |  {len(df)} registros")

    def cabecalho(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(AZUL)
        canvas.rect(0, A4[1] - FAIXA, A4[0], FAIXA, stroke=0, fill=1)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 15)
        canvas.drawString(MARGEM, A4[1] - 14 * mm, "Análise de Produção Industrial")
        canvas.setFont("Helvetica", 8)
        canvas.drawString(MARGEM, A4[1] - 20 * mm, _sanitizar_pdf(subtitulo))
        canvas.setFont("Helvetica-Bold", 9)
        canvas.drawRightString(A4[0] - MARGEM, A4[1] - 14 * mm, "Wald")
        canvas.restoreState()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=MARGEM, rightMargin=MARGEM,
        topMargin=FAIXA + 10 * mm, bottomMargin=22 * mm,
        title="Análise de Produção Industrial",
        author="Wald", subject=_sanitizar_pdf(pergunta.strip())[:120],
    )

    story = [
        Paragraph("PERGUNTA", est["rotulo"]),
        _caixa(Paragraph(_inline(pergunta), est["pergunta"]), AZUL_CLARO, borda_esq=AZUL),
        Spacer(1, 14),
        Paragraph("RESPOSTA", est["rotulo"]),
    ]
    story += _markdown_para_flowables(resposta or "(sem conteúdo)", est)

    canvas_cls = type("_CanvasDoc", (_CanvasNumerado,),
                      {"rodape_esquerda": "ProducaoIndustrial2026.xlsx  |  gerado por Wald"})
    doc.build(story, onFirstPage=cabecalho, onLaterPages=cabecalho, canvasmaker=canvas_cls)
    return buffer.getvalue()


def gerar_txt(pergunta, resposta, gerado_em):
    """Monta a versão em texto puro de uma troca."""
    larg = 74
    periodo = f"{df['Data'].min():%d/%m/%Y} a {df['Data'].max():%d/%m/%Y}"
    partes = [
        "=" * larg,
        "ANÁLISE DE PRODUÇÃO INDUSTRIAL - Wald",
        "=" * larg,
        f"Gerado em ......: {gerado_em:%d/%m/%Y às %H:%M:%S}",
        f"Período dos dados: {periodo}",
        f"Base ...........: ProducaoIndustrial2026.xlsx ({len(df)} registros, "
        f"5 times, {len(person_cols)} colaboradores)",
        "",
        "-" * larg,
        "PERGUNTA",
        "-" * larg,
        (pergunta or "").strip(),
        "",
        "-" * larg,
        "RESPOSTA",
        "-" * larg,
        (resposta or "").strip(),
        "",
        "=" * larg,
    ]
    return "\n".join(partes)


def _slug(texto, limite=40):
    ascii_only = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode()
    limpo = re.sub(r"[^A-Za-z0-9]+", "-", ascii_only).strip("-").lower()
    return limpo[:limite].strip("-") or "resposta"


def botoes_download(pergunta, resposta, gerado_em, indice):
    """Par de botões TXT/PDF de uma resposta. O conteúdo só é gerado no clique."""
    base = f"wald_{_slug(pergunta)}_{gerado_em:%Y%m%d_%H%M%S}"
    c1, c2, _ = st.columns([1, 1, 4])
    with c1:
        st.download_button(
            "📄 TXT",
            # data recebe um callable: o Streamlit só executa quando o usuário
            # clica, em vez de reconstruir todo o histórico a cada rerun.
            data=lambda p=pergunta, r=resposta, g=gerado_em: gerar_txt(p, r, g).encode("utf-8-sig"),
            file_name=f"{base}.txt",
            mime="text/plain",
            key=f"dl_txt_{indice}",
            on_click="ignore",
            use_container_width=True,
            help="Baixar esta resposta em texto puro (UTF-8)",
        )
    with c2:
        st.download_button(
            "📕 PDF",
            data=lambda p=pergunta, r=resposta, g=gerado_em: gerar_pdf(p, r, g),
            file_name=f"{base}.pdf",
            mime="application/pdf",
            key=f"dl_pdf_{indice}",
            on_click="ignore",
            use_container_width=True,
            help="Baixar esta resposta em PDF formatado",
        )


# ====================== GRÁFICOS ======================
st.subheader("📊 Gráficos de Produção")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Produção Diária Total")
    daily = df.groupby('Data')[['Terminal MP35', 'Terminal MP30']].sum()
    daily['Total'] = daily.sum(axis=1)
    st.line_chart(daily['Total'])

with col2:
    st.subheader("Produção por Time")
    team_totals = df.groupby('Time')[['Terminal MP35', 'Terminal MP30']].sum()
    team_totals['Total'] = team_totals.sum(axis=1)
    st.bar_chart(team_totals['Total'])

st.subheader("Top 10 Colaboradores")
person_totals = df[person_cols].sum().sort_values(ascending=False).head(10)
st.bar_chart(person_totals, horizontal=True)

# ====================== CHAT ======================
if "messages" not in st.session_state:
    st.session_state.messages = [{
        "role": "assistant", 
        "content": "Olá! Sou o Wald e estou conectado à sua planilha. Posso fazer cálculos, rankings, gráficos e gerar insights. O que você gostaria de saber?"
    }]

# Os botões de download precisam ser desenhados aqui, no laço do histórico: o bloco
# que responde ao chat_input é transitório e some no rerun seguinte.
ultima_pergunta = None
for indice, message in enumerate(st.session_state.messages):
    with st.chat_message(message["role"]):
        st.markdown(message["content"])
        if message["role"] == "assistant" and ultima_pergunta is not None:
            botoes_download(ultima_pergunta, message["content"],
                            message.get("ts", datetime.now()), indice)
    ultima_pergunta = message["content"] if message["role"] == "user" else None

if prompt := st.chat_input("Ex: Qual time produziu mais? Ranking dos colaboradores."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        with st.spinner("Wald está analisando..."):
            try:
                resposta = chat_with_grok(prompt)
            except Exception as e:
                resposta = f"Erro ao consultar o Wald: {str(e)}"
            st.markdown(resposta)

    st.session_state.messages.append({
        "role": "assistant", "content": resposta, "ts": datetime.now()
    })
    # Redesenha o histórico para que esta resposta já apareça com os botões.
    st.rerun()