import streamlit as st
import pandas as pd
import os
from openai import OpenAI
import json

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

for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

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
   
    st.session_state.messages.append({"role": "assistant", "content": resposta})