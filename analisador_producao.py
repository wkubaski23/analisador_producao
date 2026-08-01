#!/usr/bin/env python3
"""
Analisador de Dados de Produção - Exemplo Generativo em Python
==============================================================
- 5 linhas de produção
- 10 pessoas por linha
- Dados por hora, dia e mês
- Parser de perguntas em português
"""

import pandas as pd
import numpy as np
#import re
from typing import Optional, Tuple, Dict, Any

# ============================================================
# 1. GERAÇÃO DE DADOS SINTÉTICOS (simula o banco de dados)
# ============================================================

def gerar_dados_producao(num_dias: int = 25, data_inicio: str = "2026-06-01") -> pd.DataFrame:
    np.random.seed(42)
    
    datas = pd.date_range(start=data_inicio, periods=num_dias, freq="D")
    horas_trabalho = list(range(7, 18))
    linhas = [f"Linha {i}" for i in range(1, 6)]
    
    media_por_linha = {
        "Linha 1": 82, "Linha 2": 91, "Linha 3": 87,
        "Linha 4": 79, "Linha 5": 102,
    }
    
    registros = []
    for data in datas:
        for hora in horas_trabalho:
            for linha in linhas:
                media = media_por_linha[linha]
                fator_hora = 0.88 if hora <= 8 or hora >= 16 else (1.08 if 9 <= hora <= 11 else 1.0)
                
                for pessoa_id in range(1, 11):
                    variacao = np.random.normal(0, 12)
                    pecas = max(8, int(media * fator_hora + variacao + np.random.normal(0, 15)))
                    registros.append({
                        "data": data, "hora": hora, "linha": linha,
                        "pessoa_id": pessoa_id, "pecas_produzidas": pecas
                    })
    
    df = pd.DataFrame(registros)
    df["mes"] = df["data"].dt.to_period("M").astype(str)
    df["dia"] = df["data"].dt.date
    
    print(f"✅ Dados gerados: {len(df):,} registros | Período: {df['data'].min().date()} até {df['data'].max().date()}")
    return df


# ============================================================
# 2. CLASSE DE ANÁLISE
# ============================================================

class AnalisadorProducao:
    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()
        self.df["data"] = pd.to_datetime(self.df["data"])
    
    def _filtrar(self, data_inicio=None, data_fim=None, linha=None, pessoa_id=None):
        filtrado = self.df
        if data_inicio:
            filtrado = filtrado[filtrado["data"] >= pd.to_datetime(data_inicio)]
        if data_fim:
            filtrado = filtrado[filtrado["data"] <= pd.to_datetime(data_fim)]
        if linha:
            filtrado = filtrado[filtrado["linha"] == linha]
        if pessoa_id is not None:
            filtrado = filtrado[filtrado["pessoa_id"] == pessoa_id]
        return filtrado
    
    def calcular_total(self, data_inicio=None, data_fim=None, linha=None, por="total"):
        filtrado = self._filtrar(data_inicio, data_fim, linha)
        if filtrado.empty:
            return 0 if por == "total" else pd.DataFrame()
        
        if por == "total":
            return int(filtrado["pecas_produzidas"].sum())
        elif por == "dia":
            return filtrado.groupby("data")["pecas_produzidas"].sum().reset_index().rename(columns={"pecas_produzidas": "total_pecas"})
        elif por == "mes":
            return filtrado.groupby("mes")["pecas_produzidas"].sum().reset_index().rename(columns={"pecas_produzidas": "total_pecas"})
        elif por == "hora":
            return filtrado.groupby("hora")["pecas_produzidas"].sum().reset_index().sort_values("hora").rename(columns={"pecas_produzidas": "total_pecas"})
        return 0
    
    def top_linha(self, data_inicio=None, data_fim=None):
        filtrado = self._filtrar(data_inicio, data_fim)
        if filtrado.empty:
            return None
        resumo = filtrado.groupby("linha")["pecas_produzidas"].sum().reset_index()
        idx = resumo["pecas_produzidas"].idxmax()
        return resumo.loc[idx, "linha"], int(resumo.loc[idx, "pecas_produzidas"])
    
    def top_pessoa(self, data_inicio=None, data_fim=None):
        filtrado = self._filtrar(data_inicio, data_fim)
        if filtrado.empty:
            return None
        resumo = filtrado.groupby(["linha", "pessoa_id"])["pecas_produzidas"].sum().reset_index()
        idx = resumo["pecas_produzidas"].idxmax()
        return int(resumo.loc[idx, "pessoa_id"]), resumo.loc[idx, "linha"], int(resumo.loc[idx, "pecas_produzidas"])
    
    def resumo_geral(self):
        total = self.calcular_total()
        tl = self.top_linha()
        tp = self.top_pessoa()
        texto = f"""
╔════════════════════════════════════════════════════════════╗
║           📊 RESUMO GERAL DE PRODUÇÃO                      ║
╠════════════════════════════════════════════════════════════╣
║ Total de peças produzidas: {total:>10,}                     ║
║ 🏆 Linha que mais produziu: {tl[0]} com {tl[1]:,} peças      ║
"""
        if tp:
            texto += f"║ 🏆 Pessoa que mais produziu: Pessoa {tp[0]} da {tp[1]} com {tp[2]:,} peças ║\n"
        texto += f"║ Período: {self.df['data'].min().date()} até {self.df['data'].max().date()}          ║\n╚════════════════════════════════════════════════════════════╝"
        return texto


# ============================================================
# 3. PARSER DE PERGUNTAS (Generativo)
# ============================================================

def interpretar_pergunta(pergunta: str) -> Dict[str, Any]:
    p = pergunta.lower().strip()
    params = {"tipo": None, "linha": None, "pessoa_id": None,
              "data_inicio": None, "data_fim": None, "agregacao": "total"}
    
    # Linha
    m = re.search(r"linha\s*(\d+)", p)
    if m:
        params["linha"] = f"Linha {m.group(1)}"
    
    # Data específica
    m = re.search(r"(\d{4}-\d{2}-\d{2})", pergunta)
    if m:
        params["data_inicio"] = params["data_fim"] = m.group(1)
    
    # Períodos
    if "junho" in p:
        params["data_inicio"], params["data_fim"] = "2026-06-01", "2026-06-30"
    elif "julho" in p or "este mês" in p:
        params["data_inicio"] = "2026-07-01"
    elif "mês passado" in p:
        params["data_inicio"], params["data_fim"] = "2026-06-01", "2026-06-30"
    
    # Tipo de pergunta
    if any(kw in p for kw in ["qual linha produziu mais", "linha que mais produziu", "melhor linha"]):
        params["tipo"] = "top_linha"
    elif any(kw in p for kw in ["qual pessoa produziu mais", "pessoa que mais produziu", "quem produziu mais"]):
        params["tipo"] = "top_pessoa"
    elif any(kw in p for kw in ["resumo", "visão geral", "dashboard"]):
        params["tipo"] = "resumo"
    elif any(kw in p for kw in ["quantas peças", "total de peças", "quanto produziu"]):
        params["tipo"] = "total"
        if "por hora" in p:
            params["agregacao"] = "hora"
        elif "por dia" in p:
            params["agregacao"] = "dia"
        elif "por mês" in p:
            params["agregacao"] = "mes"
    
    if params["tipo"] is None and ("produziu" in p or "peças" in p):
        params["tipo"] = "total"
    
    return params


def processar_pergunta(analisador, pergunta: str) -> str:
    params = interpretar_pergunta(pergunta)
    
    if params["tipo"] == "top_linha":
        res = analisador.top_linha(params["data_inicio"], params["data_fim"])
        if res:
            periodo = f" no período de {params['data_inicio']}" if params["data_inicio"] else ""
            return f"🏆 A linha que mais produziu{periodo} foi **{res[0]}** com **{res[1]:,} peças**."
    
    elif params["tipo"] == "top_pessoa":
        res = analisador.top_pessoa(params["data_inicio"], params["data_fim"])
        if res:
            periodo = f" no período de {params['data_inicio']}" if params["data_inicio"] else ""
            return f"🏆 A pessoa que mais produziu{periodo} foi **Pessoa {res[0]} da {res[1]}** com **{res[2]:,} peças**."
    
    elif params["tipo"] == "resumo":
        return analisador.resumo_geral()
    
    elif params["tipo"] == "total":
        total = analisador.calcular_total(
            data_inicio=params["data_inicio"], data_fim=params["data_fim"],
            linha=params["linha"], por=params["agregacao"]
        )
        linha_str = f" na {params['linha']}" if params["linha"] else ""
        if params["agregacao"] == "total":
            return f"📦 Total de peças produzidas{linha_str}: **{total:,}**"
        elif isinstance(total, pd.DataFrame) and not total.empty:
            return f"📊 Produção por {params['agregacao']}{linha_str}:\n\n{total.to_string(index=False)}"
    
    return ("Desculpe, não entendi completamente.\nAqui vai um resumo geral:\n\n" +
            analisador.resumo_geral() +
            "\n\nTente: 'Qual linha produziu mais em junho?', 'Qual pessoa produziu mais?', 'Produção por hora na Linha 3'")


# ============================================================
# 4. INTERFACE INTERATIVA
# ============================================================

def main():
    print("\n" + "="*60)
    print("🏭 ANALISADOR DE DADOS DE PRODUÇÃO - MODO INTERATIVO")
    print("="*60)
    
    df = gerar_dados_producao(num_dias=25)
    analisador = AnalisadorProducao(df)
    
    print("\n" + analisador.resumo_geral())
    print("\n💡 Digite 'resumo', 'ajuda' ou faça perguntas em português. Digite 'sair' para encerrar.\n" + "-"*60)
    
    while True:
        try:
            pergunta = input("\n❓ Sua pergunta: ").strip()
            if pergunta.lower() in ["sair", "exit", "quit"]:
                print("👋 Até logo!")
                break
            if pergunta.lower() in ["ajuda", "help", "exemplos"]:
                print("\nExemplos:\n• Qual linha produziu mais em junho?\n• Qual pessoa produziu mais peças?\n• Quantas peças na Linha 3 no dia 2026-06-15?\n• Qual foi a produção total por hora?\n• Resumo")
                continue
            print("\n" + processar_pergunta(analisador, pergunta))
        except KeyboardInterrupt:
            print("\n👋 Programa encerrado.")
            break


if __name__ == "__main__":
    main()