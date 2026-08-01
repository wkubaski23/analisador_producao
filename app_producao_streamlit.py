import streamlit as st
import pandas as pd
import os
from analisador_producao import gerar_dados_producao, AnalisadorProducao

# ====================== GROK (xAI) ======================
from langchain_xai import ChatXAI
from langchain_core.messages import HumanMessage, SystemMessage

st.set_page_config(page_title="🚀 Analisador Grok", page_icon="🚀", layout="wide")

# Carregar analisador
@st.cache_resource
def carregar_analisador():
    df = gerar_dados_producao(num_dias=30)
    return AnalisadorProducao(df)

analisador = carregar_analisador()

st.title("🚀 Analisador de Produção com Grok")
st.caption("Versão simplificada e estável")

with st.sidebar:
    st.header("Configurações")
    api_key = st.text_input("xAI API Key", type="password", value=os.getenv("XAI_API_KEY", ""))
    if api_key:
        os.environ["XAI_API_KEY"] = api_key
    st.divider()
    st.write(f"Registros: **{len(analisador.df):,}**")

# Chat
if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "assistant", "content": "Olá! Estou pronto para analisar seus dados de produção. Pergunte o que quiser."}]

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if prompt := st.chat_input("Faça sua pergunta..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)
    
    with st.chat_message("assistant"):
        with st.spinner("Grok está pensando..."):
            try:
                llm = ChatXAI(
                    model="grok-4",
                    temperature=0,
                    api_key=os.getenv("XAI_API_KEY")
                )
                
                # System prompt + ferramentas simples
                system_prompt = f"""Você é um analista de produção. Use os dados disponíveis.
                Dados carregados: {len(analisador.df)} registros de produção.
                Responda sempre em português."""
                
                response = llm.invoke([
                    SystemMessage(content=system_prompt),
                    HumanMessage(content=prompt)
                ])
                
                resposta_final = response.content
            except Exception as e:
                resposta_final = f"Erro ao conectar com Grok: {str(e)}\n\nVerifique se a API Key está correta."
            
            st.markdown(resposta_final)
    
    st.session_state.messages.append({"role": "assistant", "content": resposta_final})