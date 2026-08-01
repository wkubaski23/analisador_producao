import streamlit as st
import pandas as pd
from analisador_producao import gerar_dados_producao, AnalisadorProducao, processar_pergunta

st.set_page_config(
    page_title="🏭 Analisador de Produção",
    page_icon="📊",
    layout="wide"
)

# Carregar dados (com cache)
@st.cache_resource
def carregar_analisador():
    with st.spinner("Carregando dados de produção..."):
        df = gerar_dados_producao(num_dias=25, data_inicio="2026-06-01")
        return AnalisadorProducao(df)

analisador = carregar_analisador()

# Sidebar
with st.sidebar:
    st.header("📋 Informações")
    st.write(f"**Registros:** {len(analisador.df):,}")
    st.write(f"**Período:** {analisador.df['data'].min().date()} até {analisador.df['data'].max().date()}")
    st.divider()
    st.subheader("Exemplos de perguntas")
    st.markdown("""
- Qual linha produziu mais em junho?
- Qual pessoa produziu mais peças?
- Quantas peças na Linha 3 no dia 2026-06-15?
- Produção total por hora na Linha 5?
- Resumo geral
    """)
    
    if st.button("🔄 Reiniciar Chat"):
        st.session_state.messages = []
        st.rerun()

# Título
st.title("🏭 Analisador de Produção")
st.caption("Converse com o sistema sobre os dados de produção")

# Chat History
if "messages" not in st.session_state:
    st.session_state.messages = []

for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Input do usuário
if prompt := st.chat_input("Digite sua pergunta sobre produção..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        with st.spinner("Analisando os dados..."):
            resposta = processar_pergunta(analisador, prompt)
            st.markdown(resposta)
    
    st.session_state.messages.append({"role": "assistant", "content": resposta})

st.caption("Versão simples e estável")