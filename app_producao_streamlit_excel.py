import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.set_page_config(page_title="🏭 Produção Industrial - Excel", page_icon="📊", layout="wide")

# ====================== CARREGAR DADOS DO EXCEL ======================
@st.cache_data
def carregar_dados_excel(caminho="producao_industrial_2026.xlsx"):
    df = pd.read_excel(caminho)
    df['Data'] = pd.to_datetime(df['Data'])
    return df

# Tentar carregar o arquivo Excel
try:
    df = carregar_dados_excel()
except FileNotFoundError:
    st.error("Arquivo 'producao_industrial_2026.xlsx' não encontrado na pasta do projeto.")
    st.stop()

# Colunas
person_cols = [col for col in df.columns if col not in ['Data', 'Time', 'Terminal MP35', 'Terminal MP30']]

st.title("🏭 Analisador de Produção Industrial")
st.caption("Leitura de planilha Excel • 5 Times • 20 Colaboradores")

with st.sidebar:
    st.header("📋 Informações")
    st.write(f"**Período:** {df['Data'].min().date()} até {df['Data'].max().date()}")
    st.write(f"**Dias úteis:** {df['Data'].nunique()}")
    st.write(f"**Total de registros:** {len(df)}")
    st.divider()
    st.subheader("Filtros")
    times_selecionados = st.multiselect("Times", options=[1,2,3,4,5], default=[1,2,3,4,5])
    data_inicio = st.date_input("Data Início", value=df['Data'].min().date())
    data_fim = st.date_input("Data Fim", value=df['Data'].max().date())

# Filtrar dados
mask = (df['Data'] >= pd.to_datetime(data_inicio)) & (df['Data'] <= pd.to_datetime(data_fim)) & (df['Time'].isin(times_selecionados))
df_filtrado = df[mask]

st.divider()

# ====================== RESUMO GERAL ======================
col1, col2, col3 = st.columns(3)

with col1:
    total_mp35 = df_filtrado['Terminal MP35'].sum()
    st.metric("Total MP35", f"{total_mp35:,}")

with col2:
    total_mp30 = df_filtrado['Terminal MP30'].sum()
    st.metric("Total MP30", f"{total_mp30:,}")

with col3:
    total_geral = total_mp35 + total_mp30
    st.metric("Produção Total", f"{total_geral:,}")

st.divider()

# ====================== TOP TIMES ======================
st.subheader("🏆 Top Times por Produção")

top_times = df_filtrado.groupby('Time')[['Terminal MP35', 'Terminal MP30']].sum()
top_times['Total'] = top_times['Terminal MP35'] + top_times['Terminal MP30']
top_times = top_times.sort_values('Total', ascending=False).reset_index()

st.dataframe(top_times, use_container_width=True)

# ====================== TOP COLABORADORES ======================
st.subheader("👥 Top 10 Colaboradores")

person_totals = df_filtrado[person_cols].sum().sort_values(ascending=False).head(10)
st.bar_chart(person_totals)

st.dataframe(person_totals.reset_index().rename(columns={'index': 'Colaborador', 0: 'Total Peças'}), use_container_width=True)

# ====================== PRODUÇÃO POR DIA ======================
st.subheader("📈 Produção Diária")

daily = df_filtrado.groupby('Data')[['Terminal MP35', 'Terminal MP30']].sum().reset_index()
daily['Total'] = daily['Terminal MP35'] + daily['Terminal MP30']

st.line_chart(daily.set_index('Data')[['Total']])

# ====================== CHAT SIMPLES ======================
st.divider()
st.subheader("💬 Pergunte sobre os dados")

if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

for msg in st.session_state.chat_history:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if prompt := st.chat_input("Ex: Qual time produziu mais? Quem produziu mais peças?"):
    st.session_state.chat_history.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)
    
    with st.chat_message("assistant"):
        resposta = ""
        p = prompt.lower()
        
        if "time" in p and ("mais" in p or "top" in p):
            top = top_times.iloc[0]
            resposta = f"O Time {int(top['Time'])} foi o que mais produziu com {int(top['Total']):,} peças no período selecionado."
        
        elif "colaborador" in p or "pessoa" in p or "quem produziu mais":
            top_person = person_totals.index[0]
            top_value = person_totals.iloc[0]
            resposta = f"O colaborador que mais produziu foi **{top_person}** com {int(top_value):,} peças."
        
        elif "total" in p or "quanto produziu":
            resposta = f"No período selecionado, a produção total foi de **{total_geral:,}** peças (MP35 + MP30)."
        
        else:
            resposta = "Posso responder sobre: Top times, Top colaboradores, totais de produção, etc. Tente reformular a pergunta."
        
        st.markdown(resposta)
    
    st.session_state.chat_history.append({"role": "assistant", "content": resposta})
