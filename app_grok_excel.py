import streamlit as st
import pandas as pd
import numpy as np
import os
from openai import OpenAI
import json
import io
import math
import re
import unicodedata
from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.platypus import (
    Paragraph, Preformatted, SimpleDocTemplate, Spacer, Table, TableStyle
)

from sklearn.linear_model import Ridge
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler

st.set_page_config(page_title="🚀 Grok + Produção Excel", page_icon="🚀", layout="wide")

# ====================== CONFIGURAÇÃO ======================
# Tudo que era literal espalhado pelo código veio para cá herdado, e aceita variável de
# ambiente para não precisar editar o arquivo em outra máquina.
ARQUIVO_EXCEL = os.getenv("PRODUCAO_XLSX", "ProducaoIndustrial2026.xlsx")
PASTA_SESSOES = Path(os.getenv("PRODUCAO_SESSOES", "sessoes"))
MODELO = os.getenv("XAI_MODEL", "grok-4")

# Quantas mensagens da conversa vão junto com cada pergunta. Duas por troca,
# então 12 = as 6 últimas perguntas e respostas. Subir isso melhora o contexto
# e encarece a chamada.
MAX_MENSAGENS_CONTEXTO = 12

# Quantas vezes o modelo pode encadear ferramentas numa mesma pergunta. Na
# última rodada as ferramentas são retiradas para forçá-lo a concluir.
MAX_RODADAS_FERRAMENTA = 4

COLS_FIXAS = ["Data", "Time", "Terminal MP35", "Terminal MP30"]


# ====================== CARREGAR EXCEL ======================
@st.cache_data(show_spinner="Carregando planilha...")
def carregar_excel(caminho):
    """Lê e valida a planilha.

    Devolve (dataframe, erro). Nada de st.success aqui dentro: a função é
    cacheada, e a mensagem não voltaria a aparecer nas execuções seguintes —
    era o que acontecia na versão anterior.
    """
    caminho = Path(caminho)
    if not caminho.exists():
        return None, f"Arquivo não encontrado: {caminho.resolve()}"

    try:
        dados = pd.read_excel(caminho)
    except Exception as e:
        return None, f"Não foi possível ler a planilha: {e}"

    faltando = [c for c in ("Data", "Time") if c not in dados.columns]
    if faltando:
        return None, f"A planilha não tem as colunas obrigatórias: {', '.join(faltando)}"

    dados["Data"] = pd.to_datetime(dados["Data"], errors="coerce")
    if dados["Data"].isna().all():
        return None, "A coluna 'Data' não pôde ser interpretada como data."

    perdidas = int(dados["Data"].isna().sum())
    dados = dados.dropna(subset=["Data"]).reset_index(drop=True)
    aviso = f"{perdidas} linha(s) descartada(s) por data inválida." if perdidas else None
    return (dados, aviso), None


resultado, erro_carga = carregar_excel(ARQUIVO_EXCEL)
if erro_carga:
    st.error(erro_carga)
    st.info(
        "Coloque a planilha na pasta onde o app é executado, ou defina o caminho "
        "na variável de ambiente PRODUCAO_XLSX."
    )
    st.stop()

df, aviso_carga = resultado
COLS_TERMINAL = [c for c in ("Terminal MP35", "Terminal MP30") if c in df.columns]
person_cols = [c for c in df.columns if c not in COLS_FIXAS]

if not COLS_TERMINAL:
    st.error("Não encontrei nenhuma coluna de terminal (Terminal MP35 / Terminal MP30).")
    st.stop()


# ====================== COMPOSIÇÃO DOS TIMES ======================
def derivar_times(dados, colunas_pessoa):
    """Descobre quem pertence a cada time a partir dos próprios dados.

    A versão anterior fatiava a lista de colaboradores de quatro em quatro:

        person_cols[(time-1)*4 : time*4]

    Isso só acerta se as colunas estiverem exatamente na ordem dos times, e
    quando erra o app afirma com toda a confiança que alguém é de um time que
    não é o dele. Aqui a composição é deduzida: uma pessoa pertence ao time em
    cujas linhas ela tem produção registrada.

    Devolve (mapa, ambiguos, sem_time). Se a planilha não separar produção por
    time, todos aparecem em todos os times e caem em 'ambiguos' — o app avisa
    em vez de mentir.
    """
    mapa = {}
    for time in sorted(dados["Time"].dropna().unique()):
        linhas = dados[dados["Time"] == time]
        membros = [
            c for c in colunas_pessoa
            if c in linhas and linhas[c].notna().any() and (linhas[c].fillna(0) != 0).any()
        ]
        mapa[int(time)] = membros

    aparicoes = {}
    for time, membros in mapa.items():
        for pessoa in membros:
            aparicoes.setdefault(pessoa, []).append(time)

    ambiguos = {p: times for p, times in aparicoes.items() if len(times) > 1}
    sem_time = [c for c in colunas_pessoa if c not in aparicoes]
    return mapa, ambiguos, sem_time


TIMES, AMBIGUOS, SEM_TIME = derivar_times(df, person_cols)


# ====================== SESSÕES (MEMÓRIA EM DISCO) ======================
MENSAGEM_BOAS_VINDAS = {
    "role": "assistant",
    "content": (
        "Olá! Sou o Wald e estou conectado à sua planilha. Posso fazer cálculos, "
        "rankings e gráficos, e agora lembro do que já conversamos — pode perguntar "
        "\"e no mês passado?\" que eu entendo o assunto."
    ),
}


def novo_id_sessao():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def caminho_sessao(sessao_id):
    return PASTA_SESSOES / f"{sessao_id}.json"


def ts_da_mensagem(mensagem):
    """A data da mensagem volta do disco como texto; o PDF precisa de datetime."""
    bruto = mensagem.get("ts")
    if isinstance(bruto, datetime):
        return bruto
    try:
        return datetime.fromisoformat(bruto)
    except (TypeError, ValueError):
        return datetime.now()


def titulo_da_sessao(mensagens):
    for m in mensagens:
        if m["role"] == "user" and m.get("content"):
            return m["content"].strip().replace("\n", " ")[:60]
    return "Nova conversa"


def salvar_sessao(sessao_id, mensagens):
    """Grava a conversa em disco. Chamado depois de cada resposta."""
    try:
        PASTA_SESSOES.mkdir(parents=True, exist_ok=True)
        conteudo = {
            "id": sessao_id,
            "titulo": titulo_da_sessao(mensagens),
            "atualizada_em": datetime.now().isoformat(timespec="seconds"),
            "planilha": str(ARQUIVO_EXCEL),
            "mensagens": mensagens,
        }
        caminho_sessao(sessao_id).write_text(
            json.dumps(conteudo, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except (OSError, TypeError) as e:
        st.warning(f"Não foi possível salvar a sessão: {e}")


def listar_sessoes():
    """Sessões salvas, da mais recente para a mais antiga."""
    if not PASTA_SESSOES.exists():
        return []
    encontradas = []
    for arquivo in PASTA_SESSOES.glob("*.json"):
        try:
            dados = json.loads(arquivo.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        mensagens = dados.get("mensagens", [])
        encontradas.append({
            "id": dados.get("id", arquivo.stem),
            "titulo": dados.get("titulo") or arquivo.stem,
            "atualizada_em": dados.get("atualizada_em", ""),
            "trocas": sum(1 for m in mensagens if m.get("role") == "user"),
        })
    return sorted(encontradas, key=lambda s: s["atualizada_em"], reverse=True)


def carregar_sessao(sessao_id):
    try:
        dados = json.loads(caminho_sessao(sessao_id).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    return dados.get("mensagens", [])


if "sessao_id" not in st.session_state:
    st.session_state.sessao_id = novo_id_sessao()
if "messages" not in st.session_state:
    st.session_state.messages = [dict(MENSAGEM_BOAS_VINDAS)]


# ====================== INTERFACE: TÍTULO E SIDEBAR ======================
st.title("🚀 Wald Analisando sua Planilha de Produção")
st.caption("Wald entende a planilha, faz cálculos, gera insights e gráficos — e lembra da conversa")

if aviso_carga:
    st.warning(aviso_carga)

with st.sidebar:
    st.header("Configurações")
    api_key = st.text_input("xAI API Key", type="password", value=os.getenv("XAI_API_KEY", ""))
    if api_key:
        os.environ["XAI_API_KEY"] = api_key

    st.divider()
    st.subheader("🧠 Memória")
    trocas_atuais = sum(1 for m in st.session_state.messages if m["role"] == "user")
    st.caption(f"Sessão `{st.session_state.sessao_id}` · {trocas_atuais} pergunta(s)")
    st.caption(
        f"O modelo recebe as últimas {MAX_MENSAGENS_CONTEXTO // 2} trocas como contexto."
    )

    if st.button("➕ Nova conversa", use_container_width=True):
        st.session_state.sessao_id = novo_id_sessao()
        st.session_state.messages = [dict(MENSAGEM_BOAS_VINDAS)]
        st.rerun()

    sessoes = [s for s in listar_sessoes() if s["id"] != st.session_state.sessao_id]
    if sessoes:
        rotulos = {s["id"]: f"{s['titulo']} ({s['trocas']}x)" for s in sessoes}
        escolhida = st.selectbox(
            "Conversas anteriores",
            options=[""] + [s["id"] for s in sessoes],
            format_func=lambda i: "— selecione —" if i == "" else rotulos[i],
        )
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Abrir", disabled=not escolhida, use_container_width=True):
                mensagens = carregar_sessao(escolhida)
                if mensagens:
                    st.session_state.sessao_id = escolhida
                    st.session_state.messages = mensagens
                    st.rerun()
                else:
                    st.warning("Sessão vazia ou corrompida.")
        with c2:
            if st.button("Excluir", disabled=not escolhida, use_container_width=True):
                caminho_sessao(escolhida).unlink(missing_ok=True)
                st.rerun()

    st.divider()
    st.write(f"**Planilha:** {ARQUIVO_EXCEL}")
    st.write(f"**Período:** {df['Data'].min().date()} a {df['Data'].max().date()}")
    st.write(f"**Registros:** {len(df)}")
    st.write(f"**Times:** {len(TIMES)}")
    st.write(f"**Colaboradores:** {len(person_cols)}")
    st.caption("🤖 Previsão e detecção de anomalia ativas — pergunte "
               "\"qual a previsão da semana?\" ou \"algum dia fora do padrão?\"")

    if AMBIGUOS:
        st.warning(
            f"{len(AMBIGUOS)} colaborador(es) com produção em mais de um time. "
            "A composição dos times pode estar incorreta — confira a planilha."
        )
    if SEM_TIME:
        amostra = ", ".join(SEM_TIME[:4]) + ("..." if len(SEM_TIME) > 4 else "")
        st.info(f"{len(SEM_TIME)} coluna(s) sem produção registrada: {amostra}")


# ====================== FUNÇÕES / TOOLS ======================
def _janela(start_date=None, end_date=None, team=None, dados=None):
    """Recorte por período e time — base comum de todas as consultas.

    Note o 'team is not None': a versão anterior usava 'if team:', que
    silenciosamente ignorava o time 0 caso ele existisse.
    """
    base = df if dados is None else dados
    recorte = base
    if start_date:
        recorte = recorte[recorte["Data"] >= pd.to_datetime(start_date)]
    if end_date:
        recorte = recorte[recorte["Data"] <= pd.to_datetime(end_date)]
    if team is not None:
        recorte = recorte[recorte["Time"] == team]
    return recorte


def _rotulo_periodo(start_date, end_date):
    return f"{start_date or 'início'} até {end_date or 'fim'}"


def get_total_production(start_date=None, end_date=None, team=None, dados=None):
    recorte = _janela(start_date, end_date, team, dados)
    if recorte.empty:
        return {"aviso": "Nenhum registro para esse período/time.",
                "periodo": _rotulo_periodo(start_date, end_date)}

    totais = {c: int(recorte[c].sum()) for c in COLS_TERMINAL}
    return {
        "total_mp35": totais.get("Terminal MP35", 0),
        "total_mp30": totais.get("Terminal MP30", 0),
        "total_geral": int(sum(totais.values())),
        "dias_com_registro": int(recorte["Data"].nunique()),
        "periodo": _rotulo_periodo(start_date, end_date),
        "time": team if team is not None else "Todos",
    }


def get_top_team(start_date=None, end_date=None, n=5, dados=None):
    recorte = _janela(start_date, end_date, None, dados)
    if recorte.empty:
        return {"aviso": "Nenhum registro para esse período.",
                "periodo": _rotulo_periodo(start_date, end_date)}

    totais = recorte.groupby("Time")[COLS_TERMINAL].sum()
    totais["Total"] = totais.sum(axis=1)
    ranking = totais.sort_values("Total", ascending=False).head(n)

    return {
        "periodo": _rotulo_periodo(start_date, end_date),
        "times_no_periodo": int(len(totais)),
        # Lista em vez de dicionário: a posição fica explícita e o modelo não
        # depende da ordem das chaves do JSON para saber quem é o primeiro.
        "ranking": [
            {
                "posicao": i + 1,
                "time": int(indice),
                "total": int(linha["Total"]),
                "mp35": int(linha.get("Terminal MP35", 0)),
                "mp30": int(linha.get("Terminal MP30", 0)),
            }
            for i, (indice, linha) in enumerate(ranking.iterrows())
        ],
    }


def _ranking_pessoas(start_date, end_date, n, team, crescente, dados):
    recorte = _janela(start_date, end_date, team, dados)
    if recorte.empty:
        return {"aviso": "Nenhum registro para esse período/time.",
                "periodo": _rotulo_periodo(start_date, end_date)}

    if dados is None:
        composicao = TIMES
    else:
        composicao, _, _ = derivar_times(dados, person_cols)

    candidatos = composicao.get(team, []) if team is not None else person_cols
    if team is not None and not candidatos:
        return {"aviso": f"Não identifiquei colaboradores no time {team}."}

    # Sem este filtro, quem não tem nenhum lançamento na janela entra no
    # ranking com 0 peças e aparece como o "pior" colaborador do período.
    com_registro = [c for c in candidatos if c in recorte and recorte[c].notna().any()]
    ignorados = [c for c in candidatos if c not in com_registro]

    if not com_registro:
        return {"aviso": "Nenhum colaborador com produção registrada nesse recorte."}

    totais = recorte[com_registro].sum().sort_values(ascending=crescente)
    ranking = totais.head(n)

    resposta = {
        "periodo": _rotulo_periodo(start_date, end_date),
        "time": team if team is not None else "Todos",
        "ordem": "do menor para o maior" if crescente else "do maior para o menor",
        "colaboradores_considerados": len(com_registro),
        "ranking": [
            {"posicao": i + 1, "nome": nome, "total_pecas": int(valor)}
            for i, (nome, valor) in enumerate(ranking.items())
        ],
    }
    if ignorados:
        resposta["sem_registro_no_periodo"] = ignorados
    return resposta


def get_top_person(start_date=None, end_date=None, n=5, team=None, dados=None):
    return _ranking_pessoas(start_date, end_date, n, team, False, dados)


def get_bottom_person(start_date=None, end_date=None, n=5, team=None, dados=None):
    return _ranking_pessoas(start_date, end_date, n, team, True, dados)


def get_daily_production(start_date=None, end_date=None, limite=15, ordem="recentes",
                         team=None, dados=None):
    """Produção por dia.

    A versão anterior devolvia sempre [:15] — os quinze dias mais ANTIGOS —
    enquanto a descrição da ferramenta dizia ao modelo "últimos dias do
    período". Quem perguntasse pelos últimos dias recebia os primeiros.
    """
    recorte = _janela(start_date, end_date, team, dados)
    if recorte.empty:
        return {"aviso": "Nenhum registro para esse período/time.",
                "periodo": _rotulo_periodo(start_date, end_date)}

    diario = recorte.groupby("Data")[COLS_TERMINAL].sum()
    diario["Total"] = diario.sum(axis=1)
    diario = diario.sort_index()

    limite = max(1, int(limite or 15))
    dias = diario.tail(limite) if ordem == "recentes" else diario.head(limite)
    if ordem == "recentes":
        dias = dias.sort_index(ascending=False)

    return {
        "periodo": _rotulo_periodo(start_date, end_date),
        "time": team if team is not None else "Todos",
        "dias_no_periodo": int(len(diario)),
        "dias_retornados": int(len(dias)),
        "ordem": "mais recentes primeiro" if ordem == "recentes" else "mais antigos primeiro",
        "dias": [
            {
                "data": indice.strftime("%Y-%m-%d"),
                "mp35": int(linha.get("Terminal MP35", 0)),
                "mp30": int(linha.get("Terminal MP30", 0)),
                "total": int(linha["Total"]),
            }
            for indice, linha in dias.iterrows()
        ],
    }


def list_colaboradores(dados=None):
    if dados is None:
        mapa, ambiguos, sem_time = TIMES, AMBIGUOS, SEM_TIME
    else:
        mapa, ambiguos, sem_time = derivar_times(dados, person_cols)

    resposta = {
        "times": {f"Time {t}": membros for t, membros in mapa.items()},
        "total_colaboradores": sum(len(m) for m in mapa.values()),
        "origem": "deduzido dos lançamentos de produção da planilha",
    }
    if ambiguos:
        detalhe = "; ".join(
            f"{p}: times {', '.join(str(t) for t in times)}" for p, times in ambiguos.items()
        )
        resposta["aviso"] = (
            "Estes colaboradores têm produção em mais de um time, então a composição "
            f"acima pode não ser confiável: {detalhe}"
        )
    if sem_time:
        resposta["sem_producao_registrada"] = sem_time
    return resposta


# ====================== APRENDIZADO DE MÁQUINA ======================
# Aqui, e só aqui, existe aprendizado de fato: um modelo é ajustado ao
# histórico da fábrica. O resto do app são consultas determinísticas.
#
# Princípio que orienta esta seção: modelo só vale se ganhar de uma regra
# burra. Toda previsão é comparada, por validação temporal, contra três
# linhas de base ingênuas. Se não ganhar, a função DIZ isso e devolve a linha
# de base — em vez de fingir competência que não tem.

MIN_DIAS_TREINO = 30
MIN_DIAS_ANOMALIA = 21
ALPHAS = [0.01, 0.1, 1.0, 10.0, 100.0]
GANHO_MINIMO_PCT = 5.0          # abaixo disso a complexidade não se paga
DIAS_SEMANA = ["segunda", "terça", "quarta", "quinta", "sexta", "sábado", "domingo"]


def _serie_diaria(dados=None, team=None, start_date=None, end_date=None):
    """Produção total por dia — a série que o modelo aprende."""
    recorte = _janela(start_date, end_date, team, dados)
    if recorte.empty:
        return pd.Series(dtype=float)
    serie = recorte.groupby("Data")[COLS_TERMINAL].sum().sum(axis=1).sort_index()
    serie.index = pd.DatetimeIndex(serie.index)
    return serie


def _features(datas, origem):
    """Tendência linear + dia da semana.

    Deliberadamente simples. Com poucos meses de histórico, modelo mais
    elaborado decora o ruído em vez de aprender o padrão — e previsão
    decorada é pior que previsão nenhuma num painel de diretoria.
    """
    datas = pd.DatetimeIndex(datas)
    dias = (datas - origem).days.to_numpy(dtype=float)
    colunas = [dias]
    for d in range(6):                      # domingo é a categoria de referência
        colunas.append((datas.dayofweek == d).astype(float))
    return np.column_stack(colunas)


def _modelo(alpha):
    return make_pipeline(StandardScaler(), Ridge(alpha=alpha))


def _baselines(treino_y, treino_datas, teste_datas):
    """As três regras burras contra as quais o modelo precisa competir."""
    ultimo = np.full(len(teste_datas), treino_y[-1])
    media = np.full(len(teste_datas), treino_y.mean())
    por_dow, dow_treino = {}, pd.DatetimeIndex(treino_datas).dayofweek
    for d in range(7):
        v = treino_y[dow_treino == d]
        por_dow[d] = v.mean() if len(v) else treino_y.mean()
    dow = np.array([por_dow[d] for d in pd.DatetimeIndex(teste_datas).dayofweek])
    return {"ultimo valor": ultimo, "média geral": media, "média do dia da semana": dow}


def _mae(a, b):
    return float(np.mean(np.abs(np.asarray(a) - np.asarray(b))))


def _avaliar_serie(serie, n_folds=5, horizonte=7):
    """Backtest de origem móvel: treina no passado, prevê o futuro seguinte.

    Nunca embaralha. Embaralhar série temporal deixa o modelo ver o futuro
    durante o treino e produz erro artificialmente baixo — o jeito mais comum
    de alguém se enganar com previsão.
    """
    if len(serie) < MIN_DIAS_TREINO + horizonte:
        return {"suficiente": False,
                "motivo": f"São necessários ao menos {MIN_DIAS_TREINO + horizonte} dias "
                          f"com registro; a série tem {len(serie)}."}

    y = serie.to_numpy(dtype=float)
    datas, origem = serie.index, serie.index[0]
    cortes = np.linspace(MIN_DIAS_TREINO, len(serie) - horizonte, n_folds, dtype=int)
    cortes = sorted({int(c) for c in cortes if c >= MIN_DIAS_TREINO})

    erros_modelo = []
    erros_base = {k: [] for k in ["ultimo valor", "média geral", "média do dia da semana"]}
    alphas = []

    for corte in cortes:
        tr_y, tr_d = y[:corte], datas[:corte]
        te_y, te_d = y[corte:corte + horizonte], datas[corte:corte + horizonte]
        if len(te_y) == 0:
            continue

        # alpha escolhido dentro do treino, sem jamais olhar o teste
        interno = max(MIN_DIAS_TREINO, int(corte * 0.8))
        melhor_alpha, melhor_erro = ALPHAS[0], np.inf
        if interno < corte:
            for a in ALPHAS:
                m = _modelo(a).fit(_features(tr_d[:interno], origem), tr_y[:interno])
                e = _mae(tr_y[interno:], m.predict(_features(tr_d[interno:], origem)))
                if e < melhor_erro:
                    melhor_alpha, melhor_erro = a, e
        alphas.append(melhor_alpha)

        modelo = _modelo(melhor_alpha).fit(_features(tr_d, origem), tr_y)
        erros_modelo.append(_mae(te_y, modelo.predict(_features(te_d, origem))))
        for nome, pred in _baselines(tr_y, tr_d, te_d).items():
            erros_base[nome].append(_mae(te_y, pred))

    mae_modelo = float(np.mean(erros_modelo))
    mae_base = {k: float(np.mean(v)) for k, v in erros_base.items()}
    melhor = min(mae_base, key=mae_base.get)
    ganho = (mae_base[melhor] - mae_modelo) / mae_base[melhor] * 100

    return {
        "suficiente": True,
        "dias_na_serie": int(len(serie)),
        "folds": len(erros_modelo),
        "horizonte_testado_dias": horizonte,
        "erro_medio_modelo": round(mae_modelo, 1),
        "erro_medio_baselines": {k: round(v, 1) for k, v in mae_base.items()},
        "melhor_baseline": melhor,
        "ganho_sobre_baseline_pct": round(ganho, 1),
        "modelo_util": bool(ganho > GANHO_MINIMO_PCT),
        "producao_media_diaria": round(float(serie.mean()), 1),
        "erro_relativo_pct": round(mae_modelo / serie.mean() * 100, 1),
        "alpha_tipico": float(np.median(alphas)) if alphas else None,
    }


def avaliar_modelo_previsao(team=None, dados=None):
    """Quão confiável é a previsão? Compara o modelo com regras ingênuas."""
    serie = _serie_diaria(dados, team)
    if serie.empty:
        return {"aviso": "Nenhum registro para esse recorte."}
    resultado = _avaliar_serie(serie)
    if not resultado.get("suficiente"):
        return {"aviso": resultado["motivo"]}
    resultado["leitura"] = (
        f"O modelo erra em média {resultado['erro_medio_modelo']} peças por dia, "
        f"{resultado['erro_relativo_pct']}% da produção diária média. A melhor regra "
        f"simples ('{resultado['melhor_baseline']}') erra "
        f"{resultado['erro_medio_baselines'][resultado['melhor_baseline']]}. "
        + ("O modelo vale a pena." if resultado["modelo_util"]
           else "O modelo NÃO supera a regra simples; use-a no lugar dele.")
    )
    return resultado


def prever_producao(dias_a_frente=7, team=None, dados=None):
    """Previsão de produção para os próximos dias de operação."""
    serie = _serie_diaria(dados, team)
    if serie.empty:
        return {"aviso": "Nenhum registro para esse recorte."}

    aval = _avaliar_serie(serie)
    if not aval.get("suficiente"):
        return {"aviso": aval["motivo"]}

    y, datas, origem = serie.to_numpy(dtype=float), serie.index, serie.index[0]
    modelo = _modelo(aval["alpha_tipico"] or 1.0).fit(_features(datas, origem), y)
    residuos = y - modelo.predict(_features(datas, origem))
    desvio = float(np.std(residuos, ddof=1))

    # Só projeta em dias da semana que existem no histórico: prever produção
    # de domingo quando a fábrica não abre domingo não é previsão, é ruído.
    dows = set(pd.DatetimeIndex(datas).dayofweek.tolist())
    futuras, cursor = [], datas[-1]
    limite = max(1, min(int(dias_a_frente or 7), 60))
    while len(futuras) < limite:
        cursor = cursor + pd.Timedelta(days=1)
        if cursor.dayofweek in dows:
            futuras.append(cursor)
    futuras = pd.DatetimeIndex(futuras)

    if aval["modelo_util"]:
        pred = np.clip(modelo.predict(_features(futuras, origem)), 0, None)
        metodo = "regressão sobre tendência e dia da semana"
    else:
        por_dow, dow_hist = {}, pd.DatetimeIndex(datas).dayofweek
        for d in range(7):
            v = y[dow_hist == d]
            por_dow[d] = float(v.mean()) if len(v) else float(y.mean())
        pred = np.array([por_dow[d] for d in futuras.dayofweek])
        metodo = "média do dia da semana (o modelo não superou a regra simples)"

    return {
        "metodo": metodo,
        "confiavel": bool(aval["modelo_util"]),
        "time": team if team is not None else "Todos",
        "erro_medio_esperado": (aval["erro_medio_modelo"] if aval["modelo_util"]
                                else aval["erro_medio_baselines"][aval["melhor_baseline"]]),
        "erro_relativo_pct": aval["erro_relativo_pct"],
        "ganho_sobre_baseline_pct": aval["ganho_sobre_baseline_pct"],
        "treinado_com_dias": int(len(serie)),
        "previsao": [
            {"data": d.strftime("%Y-%m-%d"),
             "dia_semana": DIAS_SEMANA[d.dayofweek],
             "previsto": int(round(p)),
             "faixa_provavel": [int(max(0, round(p - 1.96 * desvio))),
                                int(round(p + 1.96 * desvio))]}
            for d, p in zip(futuras, pred)
        ],
        "observacao": ("A faixa provável cobre cerca de 95% dos casos, estimada pela "
                       "dispersão histórica dos erros. Previsão não é promessa."),
    }


def detectar_anomalias(start_date=None, end_date=None, sensibilidade=3.0,
                       team=None, dados=None):
    """Dias fora do padrão, medidos contra o que o modelo esperava para eles.

    A vantagem sobre um z-score cru está aqui: o modelo já sabe que sábado
    produz menos e que a produção vinha subindo. Um sábado baixo não é
    anomalia; um sábado baixo PARA UM SÁBADO, é.
    """
    serie = _serie_diaria(dados, team, start_date, end_date)
    if len(serie) < MIN_DIAS_ANOMALIA:
        return {"aviso": f"São necessários ao menos {MIN_DIAS_ANOMALIA} dias com "
                         f"registro; o recorte tem {len(serie)}."}

    y, datas, origem = serie.to_numpy(dtype=float), serie.index, serie.index[0]
    modelo = _modelo(1.0).fit(_features(datas, origem), y)
    esperado = modelo.predict(_features(datas, origem))
    residuo = y - esperado

    # Escala robusta: mediana e MAD não são arrastadas pelas próprias anomalias,
    # ao contrário do desvio padrão, que uma queda enorme infla e mascara.
    mad = float(np.median(np.abs(residuo - np.median(residuo))))
    escala = mad * 1.4826 if mad > 0 else float(np.std(residuo, ddof=1))
    if escala <= 0:
        return {"aviso": "A série não tem variação suficiente para avaliar anomalias."}

    sens = float(sensibilidade or 3.0)
    score = (residuo - np.median(residuo)) / escala
    marcados = np.where(np.abs(score) >= sens)[0]
    ordem = marcados[np.argsort(-np.abs(score[marcados]))][:15]

    return {
        "periodo": _rotulo_periodo(start_date, end_date),
        "time": team if team is not None else "Todos",
        "dias_analisados": int(len(serie)),
        "sensibilidade": sens,
        "anomalias_encontradas": int(len(marcados)),
        # Um limiar de N desvios deixa passar alarmes falsos por puro acaso.
        # Declarar a taxa evita que um desvio de 3.1 seja lido com o mesmo peso
        # de um de 8.0 — e o campo "desvios" de cada linha diz qual é qual.
        "alarmes_falsos_esperados_por_100_dias": round(
            math.erfc(sens / math.sqrt(2)) * 100, 1),
        "variacao_tipica_diaria": round(escala, 1),
        "anomalias": [
            {"data": datas[i].strftime("%Y-%m-%d"),
             "dia_semana": DIAS_SEMANA[datas[i].dayofweek],
             "produzido": int(round(y[i])),
             "esperado": int(round(esperado[i])),
             "diferenca": int(round(y[i] - esperado[i])),
             "desvios": round(float(score[i]), 1),
             "tipo": "queda" if score[i] < 0 else "pico"}
            for i in sorted(ordem, key=lambda k: datas[k])
        ],
        "observacao": ("'desvios' é quantas vezes a variação típica o dia se afastou do "
                       "esperado, já descontados o dia da semana e a tendência. Quanto maior, "
                       "menos chance de ser acaso: acima de 5 é praticamente certo que houve "
                       "um evento real."),
    }


# ====================== GROK COM TOOL CALLING ======================
_PERIODO = {
    "start_date": {"type": "string",
                   "description": "Data inicial no formato YYYY-MM-DD. Omita para começar do início da série."},
    "end_date": {"type": "string",
                 "description": "Data final no formato YYYY-MM-DD. Omita para ir até o fim da série."},
}

tools = [
    {
        "type": "function",
        "function": {
            "name": "get_total_production",
            "description": "Produção total somando os terminais MP35 e MP30. Use para totais gerais, por período ou por time.",
            "parameters": {
                "type": "object",
                "properties": {
                    **_PERIODO,
                    "team": {"type": "integer", "description": "Número do time. Omita para somar todos."},
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "get_top_team",
            "description": "Ranking dos times por produção total, do maior para o menor.",
            "parameters": {
                "type": "object",
                "properties": {
                    **_PERIODO,
                    "n": {"type": "integer", "description": "Quantos times no ranking. Padrão 5."},
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "get_top_person",
            "description": "Ranking dos colaboradores que produziram MAIS. Aceita filtro por time.",
            "parameters": {
                "type": "object",
                "properties": {
                    **_PERIODO,
                    "n": {"type": "integer", "description": "Quantos colaboradores. Padrão 5."},
                    "team": {"type": "integer", "description": "Restringe ao time informado. Omita para considerar todos."},
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "get_bottom_person",
            "description": "Ranking dos colaboradores que produziram MENOS. Colaboradores sem lançamento no período são excluídos, não tratados como zero. Aceita filtro por time.",
            "parameters": {
                "type": "object",
                "properties": {
                    **_PERIODO,
                    "n": {"type": "integer", "description": "Quantos colaboradores. Padrão 5."},
                    "team": {"type": "integer", "description": "Restringe ao time informado."},
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "get_daily_production",
            "description": "Produção dia a dia. Por padrão devolve os dias MAIS RECENTES do período.",
            "parameters": {
                "type": "object",
                "properties": {
                    **_PERIODO,
                    "limite": {"type": "integer", "description": "Quantos dias devolver. Padrão 15."},
                    "ordem": {"type": "string", "enum": ["recentes", "antigos"],
                              "description": "'recentes' para o fim do período, 'antigos' para o começo. Padrão 'recentes'."},
                    "team": {"type": "integer", "description": "Restringe ao time informado."},
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "list_colaboradores",
            "description": "Composição dos times, deduzida dos lançamentos da planilha. Pode vir com aviso de inconsistência.",
            "parameters": {"type": "object", "properties": {}},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "prever_producao",
            "description": "Previsão de produção para os próximos dias, com modelo treinado no histórico. Devolve faixa provável e informa se a previsão é confiável. Use quando perguntarem sobre o futuro, projeção, tendência ou meta.",
            "parameters": {
                "type": "object",
                "properties": {
                    "dias_a_frente": {"type": "integer", "description": "Quantos dias de operação projetar. Padrão 7, máximo 60."},
                    "team": {"type": "integer", "description": "Prever só um time. Omita para a fábrica inteira."},
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "detectar_anomalias",
            "description": "Encontra dias fora do padrão comparando o produzido com o que o modelo esperava para aquele dia da semana, já descontada a tendência. Use para perguntas sobre queda, pico, dia estranho ou o que saiu do normal.",
            "parameters": {
                "type": "object",
                "properties": {
                    **_PERIODO,
                    "sensibilidade": {"type": "number", "description": "Quantos desvios para marcar como anomalia. Padrão 3.0; use 2.0 para ser mais sensível."},
                    "team": {"type": "integer", "description": "Restringe a um time."},
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "avaliar_modelo_previsao",
            "description": "Diz o quanto a previsão é confiável: erro médio do modelo, erro das regras simples e se o modelo supera ou não. Use quando perguntarem se dá para confiar na projeção.",
            "parameters": {
                "type": "object",
                "properties": {
                    "team": {"type": "integer", "description": "Avaliar o modelo de um time específico."},
                },
            },
        },
    },
]

available_functions = {
    "get_total_production": get_total_production,
    "get_top_team": get_top_team,
    "get_top_person": get_top_person,
    "get_bottom_person": get_bottom_person,
    "get_daily_production": get_daily_production,
    "list_colaboradores": list_colaboradores,
    "prever_producao": prever_producao,
    "detectar_anomalias": detectar_anomalias,
    "avaliar_modelo_previsao": avaliar_modelo_previsao,
}


def descrever_dados():
    """System prompt montado a partir da planilha real.

    A versão anterior dizia '5 times, 20 colaboradores' em texto fixo, e as
    versões 4 e final tinham perdido até essa descrição. Aqui os números saem
    do dataframe, então nunca ficam desatualizados.
    """
    times = ", ".join(str(t) for t in sorted(TIMES))
    inicio = df["Data"].min().strftime("%Y-%m-%d")
    fim = df["Data"].max().strftime("%Y-%m-%d")

    return f"""Você é um analista de produção industrial competente, respondendo ao CEO da empresa.

DADOS DISPONÍVEIS (descritos a partir da planilha real; não presuma nada além disto):
- Arquivo: {ARQUIVO_EXCEL}, com {len(df)} registros.
- Período coberto: {inicio} a {fim}. Fora dessa faixa não existe dado — diga isso em vez de estimar.
- Produtos medidos: {', '.join(COLS_TERMINAL)}.
- Times: {times}.
- Colaboradores: {len(person_cols)}, um por coluna.\n- Há um modelo treinado no histórico para previsão e detecção de anomalia.

REGRAS:
- Todo número vem de ferramenta. Nunca calcule de cabeça, nunca arredonde por conta, nunca estime.
- Datas nas ferramentas sempre em YYYY-MM-DD.
- Se a ferramenta devolver "aviso" ou "erro", repasse isso ao usuário em vez de contornar.
- Você tem o histórico da conversa. Em perguntas como "e no mês passado?" ou "e o time 3?",
  descubra o assunto pelas mensagens anteriores antes de escolher a ferramenta.
- Se ainda assim a pergunta ficar ambígua, pergunte em vez de adivinhar.
- Responda em português, direto, citando os números que a ferramenta devolveu.
- Previsão NÃO é fato. Ao projetar, diga que é estimativa, informe a faixa provável e
  o erro médio esperado. Se o campo "confiavel" vier falso, avise que o modelo não superou
  uma regra simples e que a projeção deve ser lida com reserva.
- Anomalia é um sinal para investigar, não um diagnóstico. Nunca invente a causa de uma
  queda: relate o desvio e sugira apurar.
- Ao falar de desempenho individual, apresente o dado sem juízo de valor sobre a pessoa."""


def montar_contexto(historico):
    """Últimas trocas da conversa no formato que a API espera.

    Só entram texto do usuário e resposta final do assistente. As mensagens de
    ferramenta ficam de fora de propósito: elas só são válidas imediatamente
    após a mensagem que as solicitou, e reinjetá-las fora de ordem faz a API
    recusar a requisição.
    """
    uteis = [
        m for m in historico
        if m.get("role") in ("user", "assistant") and m.get("content")
    ]
    # A saudação inicial não é conversa; se sobrar como primeira mensagem, sai.
    if uteis and uteis[0]["role"] == "assistant":
        uteis = uteis[1:]
    recentes = uteis[-MAX_MENSAGENS_CONTEXTO:]
    return [{"role": m["role"], "content": m["content"]} for m in recentes]


def chat_with_grok(user_question, historico=None):
    chave = os.getenv("XAI_API_KEY")
    if not chave:
        return "Informe a xAI API Key na barra lateral para eu poder responder."

    client = OpenAI(api_key=chave, base_url="https://api.x.ai/v1")

    messages = [{"role": "system", "content": descrever_dados()}]
    messages += montar_contexto(historico or [])
    messages.append({"role": "user", "content": user_question})

    for rodada in range(MAX_RODADAS_FERRAMENTA):
        ultima = rodada == MAX_RODADAS_FERRAMENTA - 1

        parametros = {"model": MODELO, "messages": messages}
        if not ultima:
            # Na última rodada as ferramentas são retiradas: sem elas o modelo
            # é obrigado a redigir a resposta em vez de pedir mais uma consulta.
            parametros["tools"] = tools
            parametros["tool_choice"] = "auto"

        resposta = client.chat.completions.create(**parametros)
        mensagem = resposta.choices[0].message

        if not getattr(mensagem, "tool_calls", None):
            return mensagem.content

        messages.append(mensagem)

        for chamada in mensagem.tool_calls:
            nome = chamada.function.name
            try:
                argumentos = json.loads(chamada.function.arguments or "{}")
                resultado = available_functions[nome](**argumentos)
            except KeyError:
                resultado = {"erro": f"Ferramenta desconhecida: {nome}"}
            except Exception as e:
                # O erro volta para o modelo em vez de derrubar o app: ele pode
                # tentar outra ferramenta ou explicar a limitação ao usuário.
                resultado = {"erro": f"{type(e).__name__}: {e}"}

            messages.append({
                "tool_call_id": chamada.id,
                "role": "tool",
                "name": nome,
                "content": json.dumps(resultado, ensure_ascii=False, default=str),
            })

    return ("Não consegui concluir: a pergunta exigiu consultas demais em sequência. "
            "Tente dividi-la em partes menores.")


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
    diario = df.groupby("Data")[COLS_TERMINAL].sum()
    diario["Total"] = diario.sum(axis=1)
    st.line_chart(diario["Total"])

with col2:
    st.subheader("Produção por Time")
    por_time = df.groupby("Time")[COLS_TERMINAL].sum()
    por_time["Total"] = por_time.sum(axis=1)
    st.bar_chart(por_time["Total"])

st.subheader("Top 10 Colaboradores")
totais_pessoa = df[person_cols].sum().sort_values(ascending=False).head(10)
try:
    st.bar_chart(totais_pessoa, horizontal=True)
except TypeError:
    # horizontal= só existe a partir do Streamlit 1.36.
    st.bar_chart(totais_pessoa)


# ====================== CHAT ======================
# Os botões de download precisam ser desenhados aqui, no laço do histórico: o
# bloco que responde ao chat_input é transitório e some no rerun seguinte.
ultima_pergunta = None
for indice, message in enumerate(st.session_state.messages):
    with st.chat_message(message["role"]):
        st.markdown(message["content"])
        if message["role"] == "assistant" and ultima_pergunta is not None:
            botoes_download(ultima_pergunta, message["content"],
                            ts_da_mensagem(message), indice)
    ultima_pergunta = message["content"] if message["role"] == "user" else None

if prompt := st.chat_input("Ex: Qual time produziu mais? E no mês passado?"):
    st.session_state.messages.append({
        "role": "user",
        "content": prompt,
        "ts": datetime.now().isoformat(timespec="seconds"),
    })
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        with st.spinner("Wald está analisando..."):
            try:
                # Passa o histórico SEM a pergunta que acabou de entrar, que já
                # vai separada — é isso que permite entender "e no mês passado?".
                resposta = chat_with_grok(prompt, st.session_state.messages[:-1])
            except Exception as e:
                resposta = f"Erro ao consultar o Wald: {type(e).__name__}: {e}"
            st.markdown(resposta)

    st.session_state.messages.append({
        "role": "assistant",
        "content": resposta,
        "ts": datetime.now().isoformat(timespec="seconds"),
    })
    salvar_sessao(st.session_state.sessao_id, st.session_state.messages)
    # Redesenha o histórico para que esta resposta já apareça com os botões.
    st.rerun()
