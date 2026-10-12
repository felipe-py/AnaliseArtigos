#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de Geração de Gráficos Analíticos de Artigos Científicos:
1. Camada (Entrada/Sensores/Dados) vs. Número de Artigos
2. Modelos (Arquiteturas IA/DL/ML) vs. Número de Artigos
3. Saída (Alvo Mineral/Mapeamento Geológico) vs. Número de Artigos
4. Acurácia (%) vs. Modelo (Desempenho comparativo)

Gera:
- 4 imagens individuais de alta resolução (300 DPI) na pasta 'graficos/'
- 1 painel consolidado integrado 2x2 ('graficos/painel_consolidado_4_graficos.png')
- Tabelas estatísticas completas no console
"""

import os
import sys
import re
from pathlib import Path

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Configuração de encoding UTF-8 para Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Estilo visual executivo e limpo
sns.set_theme(style="whitegrid", font="DejaVu Sans")
plt.rcParams["font.sans-serif"] = ["Segoe UI", "DejaVu Sans", "Arial"]
plt.rcParams["axes.edgecolor"] = "#BDC3C7"
plt.rcParams["axes.linewidth"] = 0.8


def carregar_dados(caminho_csv: str = "sintese_entrada_modelo_saida.csv") -> pd.DataFrame:
    """Carrega base consolidada em CSV com fallback para UTF-8."""
    path = Path(caminho_csv)
    if not path.exists():
        path_alt = Path("base_consolidada_classificacao_e_sota.csv")
        if path_alt.exists():
            path = path_alt
        else:
            raise FileNotFoundError(f"Arquivo não encontrado: {path.resolve()}")
            
    df = pd.read_csv(path, sep=";", encoding="utf-8-sig")
    return df


# =============================================================================
# 1. CLASSIFICAÇÃO DE CAMADAS (ENTRADA / SENSORES / DADOS)
# =============================================================================
def extrair_camadas(df: pd.DataFrame) -> pd.DataFrame:
    """Classifica as camadas de dados de entrada utilizadas em cada artigo."""
    registros = []
    
    for idx, r in df.iterrows():
        id_artigo = r.get("ID: Rank", idx + 1)
        texto = " ".join([
            str(r.get("ENTRADA: Fontes e Sensores", "")),
            str(r.get("SÍNTESE: Entrada", "")),
            str(r.get("ID: Título do Artigo", "")),
            str(r.get("Termos Sensoriamento Remoto", "")),
        ]).lower()
        
        camadas_encontradas = set()
        
        # ASTER
        if "aster" in texto or str(r.get("ENTRADA: Usa ASTER", "")).lower() in ["sim", "true", "yes"]:
            camadas_encontradas.add("Satélite ASTER (SWIR/TIR)")
            
        # Hiperespectral
        if any(k in texto for k in ["hiperespectral", "hyperspectral", "prisma", "aviris", "hyperion"]) or \
           str(r.get("ENTRADA: Hiperespectral", "")).lower() in ["sim", "true", "yes"]:
            camadas_encontradas.add("Hiperespectral (PRISMA/AVIRIS)")
            
        # Outros Satélites Multiespectrais
        if any(k in texto for k in ["landsat", "sentinel-2", "sentinel 2", "worldview", "spot", "multiespectral", "multispectral", "planetscope"]):
            camadas_encontradas.add("Satélites Ópticos (Landsat / Sentinel-2)")
            
        # Radar / SAR
        if any(k in texto for k in ["sar", "radar", "sentinel-1", "palsar", "polarimétric", "polarimetric"]):
            camadas_encontradas.add("Radar de Abertura Sintética (SAR)")
            
        # Geoquímica
        if any(k in texto for k in ["geoquím", "geochemi", "stream sediment", "soil sample", "sedimento"]):
            camadas_encontradas.add("Dados Geoquímicos (Solo/Sedimento)")
            
        # Aerogeofísica
        if any(k in texto for k in ["geofís", "geophysic", "aeromagnet", "magnetometria", "magnetic", "gama", "gamma", "radiometric", "gravimetr"]):
            camadas_encontradas.add("Aerogeofísica (Mag/Gama/Grav)")
            
        # Topografia / DEM
        if any(k in texto for k in ["dem", "srtm", "topograf", "elevation", "slope", "declividade", "geomorfolog"]):
            camadas_encontradas.add("Topografia / Modelos DEM (SRTM)")
            
        # Drone / VANT
        if any(k in texto for k in ["drone", "uav", "vant", "aerofotogrametria"]):
            camadas_encontradas.add("Imagens de Drone / VANT")
            
        if not camadas_encontradas:
            camadas_encontradas.add("Satélite Geral / Sensoriamento Óptico")
            
        for c in camadas_encontradas:
            registros.append({"ID": id_artigo, "Camada": c})
            
    df_camadas = pd.DataFrame(registros)
    return df_camadas


# =============================================================================
# 2. CLASSIFICAÇÃO DE MODELOS (ARQUITETURAS / FAMÍLIAS)
# =============================================================================
def extrair_modelos(df: pd.DataFrame) -> pd.DataFrame:
    """Classifica os modelos e famílias de algoritmos adotados."""
    registros = []
    
    for idx, r in df.iterrows():
        id_artigo = r.get("ID: Rank", idx + 1)
        texto = " ".join([
            str(r.get("MODELO: Modelos DL", "")),
            str(r.get("MODELO: Arquitetura Base", "")),
            str(r.get("SÍNTESE: Modelo", "")),
            str(r.get("ID: Título do Artigo", "")),
            str(r.get("Termos Deep Learning", "")),
        ]).lower()
        
        mods = set()
        
        # Redes Convolucionais
        if any(k in texto for k in ["cnn", "convolutional", "lenet", "resnet", "densenet", "vgg", "1d-cnn", "2d-cnn", "3d-cnn", "fcnn"]):
            mods.add("Redes Convolucionais (CNNs / ResNet)")
            
        # U-Net / Segmentação
        if any(k in texto for k in ["u-net", "unet", "fcn", "segmentation", "segmentação"]):
            mods.add("Arquiteturas U-Net / Segmentação")
            
        # Transformers / ViT
        if any(k in texto for k in ["transformer", "vit", "vision transformer", "attention", "atenção", "mamba", "roformer"]):
            mods.add("Transformers / Vision Transformer (ViT)")
            
        # Graph Neural Networks
        if any(k in texto for k in ["gcn", "gnn", "graph convolutional", "redes em grafo"]):
            mods.add("Graph Neural Networks (GCN / GNN)")
            
        # Autoencoders
        if any(k in texto for k in ["autoencoder", "dae", "vae", "auto-encoder", "deep autoencoder"]):
            mods.add("Autoencoders (DAE / VAE)")
            
        # Machine Learning Clássico / Ensembles
        if any(k in texto for k in ["random forest", "svm", "support vector", "xgboost", "catboost", "ensemble", "random trees"]):
            mods.add("Machine Learning Clássico (RF / SVM / XGBoost)")
            
        # Redes Neurais Densas (MLP / ANN)
        if any(k in texto for k in ["mlp", "multilayer perceptron", "ann", "artificial neural network", "bpnn"]):
            mods.add("Redes Neurais Densas (MLP / ANN)")
            
        # Redes Generativas
        if any(k in texto for k in ["gan", "generative adversarial", "wgan"]):
            mods.add("Redes Generativas (GANs)")
            
        if not mods:
            mods.add("Outros Modelos de IA")
            
        for m in mods:
            registros.append({"ID": id_artigo, "Modelo": m})
            
    df_modelos = pd.DataFrame(registros)
    return df_modelos


# =============================================================================
# 3. CLASSIFICAÇÃO DE SAÍDAS (ALVOS / PRODUTOS DE MAPEAMENTO)
# =============================================================================
def extrair_saidas(df: pd.DataFrame) -> pd.DataFrame:
    """Classifica os alvos minerais e produtos geológicos preditos."""
    registros = []
    
    for idx, r in df.iterrows():
        id_artigo = r.get("ID: Rank", idx + 1)
        texto = " ".join([
            str(r.get("SAÍDA: Objeto Prospecção", "")),
            str(r.get("SÍNTESE: Saída", "")),
            str(r.get("ID: Título do Artigo", "")),
            str(r.get("Termos Prospecção Mineral", "")),
        ]).lower()
        
        alvos = set()
        
        # Fosfato (prioritário por ser critério do estudo)
        if "fosfat" in texto or "phosphate" in texto or str(r.get("SAÍDA: Prospecta Fosfato", "")).lower() in ["sim", "true", "yes"]:
            alvos.add("Fosfato (Minério Fosfático)")
            
        # Cobre
        if any(k in texto for k in ["cobre", "copper", "porphyry cu", "skarn-type"]):
            alvos.add("Cobre (Pórfiro / Skarn Cu)")
            
        # Ouro
        if any(k in texto for k in ["ouro", "gold", "orogenic gold", "epithermal gold"]):
            alvos.add("Ouro (Orogênico / Epitermal)")
            
        # Alteração Hidrotermal
        if any(k in texto for k in ["alteração hidrotermal", "hydrothermal alteration", "alteration mapping", "argílica", "propilítica", "fílica"]):
            alvos.add("Zonas de Alteração Hidrotermal")
            
        # Mapeamento Litológico / Geológico
        if any(k in texto for k in ["litológ", "litholog", "geological mapping", "mapeamento geológico", "rock classification", "rochas"]):
            alvos.add("Mapeamento Litológico / Rochas")
            
        # Lítio
        if any(k in texto for k in ["lítio", "lithium", "pegmatit"]):
            alvos.add("Lítio (Pegmatitos Líticos)")
            
        # Ferro
        if any(k in texto for k in ["ferro", "iron ore", "bif", "banded iron"]):
            alvos.add("Minério de Ferro (BIF)")
            
        # Polimetálicos / Pb-Zn / Níquel / Outros
        if any(k in texto for k in ["zinco", "chumbo", "lead-zinc", "pb-zn", "nickel", "níquel", "polymetallic", "polimetál"]):
            alvos.add("Polimetálicos (Pb-Zn / Níquel)")
            
        if not alvos:
            alvos.add("Prospecção Mineral Geral / Não Explicidado")
            
        for a in alvos:
            registros.append({"ID": id_artigo, "Saida": a})
            
    df_saidas = pd.DataFrame(registros)
    return df_saidas


# =============================================================================
# 4. EXTRAÇÃO DE ACURÁCIA (%) VS. MODELO
# =============================================================================
def extrair_acuracia_modelo(df: pd.DataFrame) -> pd.DataFrame:
    """Extrai valores numéricos de acurácia (%) e associa à família de modelo."""
    def parse_acc(v):
        if not isinstance(v, str):
            return None
        v = v.strip().lower()
        if v in ["", "não mencionado", "no mencionado", "nan", "none"]:
            return None
        m = re.search(r"(\d+[\.,]?\d*)\s*%", v)
        if m:
            val = float(m.group(1).replace(",", "."))
            return val if val <= 100 else None
        m2 = re.search(r"0\.\d+", v)
        if m2:
            return float(m2.group(0)) * 100
        m3 = re.search(r"(\b\d{2}\.\d+\b)", v)
        if m3:
            val = float(m3.group(1))
            return val if val <= 100 else None
        return None

    registros = []
    for idx, r in df.iterrows():
        acc_raw = str(r.get("SAÍDA: Melhor Acurácia", ""))
        acc_val = parse_acc(acc_raw)
        
        if acc_val is not None:
            texto_mod = " ".join([
                str(r.get("MODELO: Modelos DL", "")),
                str(r.get("MODELO: Arquitetura Base", "")),
                str(r.get("SÍNTESE: Modelo", "")),
                str(r.get("ID: Título do Artigo", "")),
            ]).lower()
            
            # Identifica a família primária para o comparativo
            if "u-net" in texto_mod or "unet" in texto_mod or "fcn" in texto_mod:
                fam = "U-Net / Segmentação"
            elif any(k in texto_mod for k in ["transformer", "vit"]):
                fam = "Transformers / ViT"
            elif any(k in texto_mod for k in ["gcn", "gnn"]):
                fam = "Graph Neural Networks"
            elif any(k in texto_mod for k in ["gan", "wgan"]):
                fam = "Redes Generativas (GAN)"
            elif any(k in texto_mod for k in ["cnn", "resnet", "densenet", "conv"]):
                fam = "CNNs (1D/2D/3D / ResNet)"
            elif any(k in texto_mod for k in ["random forest", "svm", "xgboost", "mlp", "ann"]):
                fam = "ML Clássico / Ensembles"
            else:
                fam = "Outras Redes DL"
                
            titulo_curto = str(r.get("ID: Título do Artigo", ""))[:35] + "..."
            registros.append({
                "Rank": r.get("ID: Rank", idx + 1),
                "Acuracia": acc_val,
                "Familia": fam,
                "Modelo_Detalhe": str(r.get("MODELO: Modelos DL", fam))[:30],
                "Titulo": titulo_curto
            })
            
    df_acc = pd.DataFrame(registros)
    return df_acc


# =============================================================================
# FUNÇÃO PRINCIPAL DE PLOTAGEM
# =============================================================================
def gerar_todos_graficos(caminho_csv: str = "sintese_entrada_modelo_saida.csv"):
    df = carregar_dados(caminho_csv)
    total_artigos = len(df)
    
    pasta_saida = Path("graficos")
    pasta_saida.mkdir(parents=True, exist_ok=True)
    
    print("=" * 75)
    print(" PROCESSAMENTO ANALÍTICO E GERAÇÃO DE GRÁFICOS")
    print("=" * 75)
    print(f"-> Base de Artigos Carregada: {total_artigos} registros.")
    
    # 1. Processa conjuntos de dados
    df_camadas = extrair_camadas(df)
    df_modelos = extrair_modelos(df)
    df_saidas = extrair_saidas(df)
    df_acc = extrair_acuracia_modelo(df)
    
    cont_camadas = df_camadas["Camada"].value_counts()
    cont_modelos = df_modelos["Modelo"].value_counts()
    cont_saidas = df_saidas["Saida"].value_counts()
    
    # Paletas de cores elegantes
    cor_camadas = "#16A085"  # Teal / Esmeralda
    cor_modelos = "#2980B9"  # Azul Cobalto
    cor_saidas = "#D35400"   # Laranja Queimado / Âmbar
    cor_acc = "#8E44AD"      # Roxo Púrpura

    # =========================================================================
    # GRÁFICO 1: CAMADA VS. NÚMERO DE ARTIGOS
    # =========================================================================
    fig, ax = plt.subplots(figsize=(10, 6), dpi=300)
    y_pos = np.arange(len(cont_camadas))
    bars = ax.barh(y_pos, cont_camadas.values, color=cor_camadas, height=0.65, edgecolor="#0E6655", alpha=0.9)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(cont_camadas.index, fontsize=10, fontweight="bold", color="#2C3E50")
    ax.invert_yaxis()
    ax.set_xlabel("Número de Artigos", fontsize=11, fontweight="bold", labelpad=10)
    ax.set_title(f"Camadas de Entrada e Sensores vs. Número de Artigos (Total N={total_artigos})", 
                 fontsize=13, fontweight="bold", pad=15, color="#1A252F")
    
    for bar in bars:
        w = bar.get_width()
        pct = (w / total_artigos) * 100
        ax.text(w + 0.8, bar.get_y() + bar.get_height() / 2, f"{int(w)} ({pct:.1f}%)",
                va="center", ha="left", fontsize=9.5, fontweight="bold", color="#1A252F")
    
    ax.set_xlim(0, max(cont_camadas.values) * 1.18)
    plt.tight_layout()
    p1 = pasta_saida / "1_camadas_vs_artigos.png"
    plt.savefig(p1, bbox_inches="tight")
    plt.close()
    print(f"   ✓ Gráfico 1 salvo: {p1.resolve()}")

    # =========================================================================
    # GRÁFICO 2: MODELOS VS. NÚMERO DE ARTIGOS
    # =========================================================================
    fig, ax = plt.subplots(figsize=(10, 6), dpi=300)
    y_pos = np.arange(len(cont_modelos))
    bars = ax.barh(y_pos, cont_modelos.values, color=cor_modelos, height=0.65, edgecolor="#1B4F72", alpha=0.9)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(cont_modelos.index, fontsize=10, fontweight="bold", color="#2C3E50")
    ax.invert_yaxis()
    ax.set_xlabel("Número de Artigos", fontsize=11, fontweight="bold", labelpad=10)
    ax.set_title(f"Modelos e Arquiteturas Computacionais vs. Número de Artigos", 
                 fontsize=13, fontweight="bold", pad=15, color="#1A252F")
    
    for bar in bars:
        w = bar.get_width()
        pct = (w / total_artigos) * 100
        ax.text(w + 0.8, bar.get_y() + bar.get_height() / 2, f"{int(w)} ({pct:.1f}%)",
                va="center", ha="left", fontsize=9.5, fontweight="bold", color="#1A252F")
                
    ax.set_xlim(0, max(cont_modelos.values) * 1.18)
    plt.tight_layout()
    p2 = pasta_saida / "2_modelos_vs_artigos.png"
    plt.savefig(p2, bbox_inches="tight")
    plt.close()
    print(f"   ✓ Gráfico 2 salvo: {p2.resolve()}")

    # =========================================================================
    # GRÁFICO 3: SAÍDA VS. NÚMERO DE ARTIGOS
    # =========================================================================
    fig, ax = plt.subplots(figsize=(10, 6), dpi=300)
    y_pos = np.arange(len(cont_saidas))
    bars = ax.barh(y_pos, cont_saidas.values, color=cor_saidas, height=0.65, edgecolor="#935116", alpha=0.9)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(cont_saidas.index, fontsize=10, fontweight="bold", color="#2C3E50")
    ax.invert_yaxis()
    ax.set_xlabel("Número de Artigos", fontsize=11, fontweight="bold", labelpad=10)
    ax.set_title(f"Saídas e Alvos de Prospecção / Mapeamento vs. Número de Artigos", 
                 fontsize=13, fontweight="bold", pad=15, color="#1A252F")
    
    for bar in bars:
        w = bar.get_width()
        pct = (w / total_artigos) * 100
        ax.text(w + 0.8, bar.get_y() + bar.get_height() / 2, f"{int(w)} ({pct:.1f}%)",
                va="center", ha="left", fontsize=9.5, fontweight="bold", color="#1A252F")
                
    ax.set_xlim(0, max(cont_saidas.values) * 1.18)
    plt.tight_layout()
    p3 = pasta_saida / "3_saidas_vs_artigos.png"
    plt.savefig(p3, bbox_inches="tight")
    plt.close()
    print(f"   ✓ Gráfico 3 salvo: {p3.resolve()}")

    # =========================================================================
    # GRÁFICO 4: ACURÁCIA (%) VS. MODELO
    # =========================================================================
    fig, ax = plt.subplots(figsize=(10, 6), dpi=300)
    # Médias de acurácia por família
    acc_stats = df_acc.groupby("Familia")["Acuracia"].agg(["mean", "count", "min", "max"]).sort_values(by="mean", ascending=False)
    
    y_pos = np.arange(len(acc_stats))
    bars = ax.barh(y_pos, acc_stats["mean"].values, color=cor_acc, height=0.6, edgecolor="#5B2C6F", alpha=0.88)
    ax.set_yticks(y_pos)
    labels = [f"{fam} (n={count})" for fam, count in zip(acc_stats.index, acc_stats["count"])]
    ax.set_yticklabels(labels, fontsize=10, fontweight="bold", color="#2C3E50")
    ax.invert_yaxis()
    ax.set_xlabel("Acurácia Média Reportada (%)", fontsize=11, fontweight="bold", labelpad=10)
    ax.set_xlim(80, 103)  # Zoom nos desempenhos de 80% a 100%
    ax.set_title(f"Acurácia Média por Família de Modelos (Artigos c/ Métrica Numérica)", 
                 fontsize=13, fontweight="bold", pad=15, color="#1A252F")
    
    for idx_b, bar in enumerate(bars):
        w = bar.get_width()
        mn = acc_stats["min"].iloc[idx_b]
        mx = acc_stats["max"].iloc[idx_b]
        texto_rotulo = f"{w:.2f}% (min: {mn:.1f}% | max: {mx:.1f}%)" if mn != mx else f"{w:.2f}%"
        ax.text(w + 0.3, bar.get_y() + bar.get_height() / 2, texto_rotulo,
                va="center", ha="left", fontsize=9.2, fontweight="bold", color="#1A252F")

    plt.tight_layout()
    p4 = pasta_saida / "4_acuracia_vs_modelo.png"
    plt.savefig(p4, bbox_inches="tight")
    plt.close()
    print(f"   ✓ Gráfico 4 salvo: {p4.resolve()}")

    # =========================================================================
    # PAINEL CONSOLIDADO 2x2 (TODOS EM UMA ÚNICA IMAGEM INTEGRADA)
    # =========================================================================
    fig, axes = plt.subplots(2, 2, figsize=(20, 14), dpi=300)
    fig.suptitle("PANORAMA GERAL DO ESTADO DA ARTE: ENTRADA, MODELOS, SAÍDAS E DESEMPENHO", 
                 fontsize=17, fontweight="bold", y=0.98, color="#1A252F")

    # Subplot 1: Camadas
    ax1 = axes[0, 0]
    y1 = np.arange(len(cont_camadas))
    b1 = ax1.barh(y1, cont_camadas.values, color=cor_camadas, height=0.62, edgecolor="#0E6655")
    ax1.set_yticks(y1)
    ax1.set_yticklabels(cont_camadas.index, fontsize=9.5, fontweight="bold")
    ax1.invert_yaxis()
    ax1.set_title("A) Camadas de Entrada e Sensores (Nº Artigos)", fontsize=12, fontweight="bold", pad=10)
    ax1.set_xlabel("Número de Artigos", fontsize=9.5, fontweight="bold")
    for bar in b1:
        w = bar.get_width()
        ax1.text(w + 0.6, bar.get_y() + bar.get_height()/2, f"{int(w)} ({w/total_artigos*100:.0f}%)", va="center", fontsize=9, fontweight="bold")
    ax1.set_xlim(0, max(cont_camadas.values) * 1.18)

    # Subplot 2: Modelos
    ax2 = axes[0, 1]
    y2 = np.arange(len(cont_modelos))
    b2 = ax2.barh(y2, cont_modelos.values, color=cor_modelos, height=0.62, edgecolor="#1B4F72")
    ax2.set_yticks(y2)
    ax2.set_yticklabels(cont_modelos.index, fontsize=9.5, fontweight="bold")
    ax2.invert_yaxis()
    ax2.set_title("B) Modelos Computacionais / Redes Neurais (Nº Artigos)", fontsize=12, fontweight="bold", pad=10)
    ax2.set_xlabel("Número de Artigos", fontsize=9.5, fontweight="bold")
    for bar in b2:
        w = bar.get_width()
        ax2.text(w + 0.6, bar.get_y() + bar.get_height()/2, f"{int(w)} ({w/total_artigos*100:.0f}%)", va="center", fontsize=9, fontweight="bold")
    ax2.set_xlim(0, max(cont_modelos.values) * 1.18)

    # Subplot 3: Saídas
    ax3 = axes[1, 0]
    y3 = np.arange(len(cont_saidas))
    b3 = ax3.barh(y3, cont_saidas.values, color=cor_saidas, height=0.62, edgecolor="#935116")
    ax3.set_yticks(y3)
    ax3.set_yticklabels(cont_saidas.index, fontsize=9.5, fontweight="bold")
    ax3.invert_yaxis()
    ax3.set_title("C) Saídas e Alvos Minerais Prospectados (Nº Artigos)", fontsize=12, fontweight="bold", pad=10)
    ax3.set_xlabel("Número de Artigos", fontsize=9.5, fontweight="bold")
    for bar in b3:
        w = bar.get_width()
        ax3.text(w + 0.6, bar.get_y() + bar.get_height()/2, f"{int(w)} ({w/total_artigos*100:.0f}%)", va="center", fontsize=9, fontweight="bold")
    ax3.set_xlim(0, max(cont_saidas.values) * 1.18)

    # Subplot 4: Acurácia vs Modelo
    ax4 = axes[1, 1]
    y4 = np.arange(len(acc_stats))
    b4 = ax4.barh(y4, acc_stats["mean"].values, color=cor_acc, height=0.62, edgecolor="#5B2C6F")
    ax4.set_yticks(y4)
    ax4.set_yticklabels([f"{fam} (n={count})" for fam, count in zip(acc_stats.index, acc_stats["count"])], fontsize=9.5, fontweight="bold")
    ax4.invert_yaxis()
    ax4.set_title("D) Acurácia Média Reportada (%) vs. Família de Modelo", fontsize=12, fontweight="bold", pad=10)
    ax4.set_xlabel("Acurácia (%)", fontsize=9.5, fontweight="bold")
    ax4.set_xlim(80, 103)
    for bar in b4:
        w = bar.get_width()
        ax4.text(w + 0.3, bar.get_y() + bar.get_height()/2, f"{w:.1f}%", va="center", fontsize=9, fontweight="bold")

    plt.tight_layout(rect=[0, 0, 1, 0.96])
    p_painel = pasta_saida / "painel_consolidado_4_graficos.png"
    plt.savefig(p_painel, bbox_inches="tight")
    plt.close()
    print(f"   ✓ Painel 2x2 integrado salvo: {p_painel.resolve()}")

    print("=" * 75)
    print(" RESUMO ESTATÍSTICO DAS FREQUÊNCIAS")
    print("=" * 75)
    print("\n1. Camadas / Sensores:")
    for k, v in cont_camadas.items():
        print(f"   • {k:<45}: {v:>3} artigos ({v/total_artigos*100:5.1f}%)")

    print("\n2. Modelos / Arquiteturas:")
    for k, v in cont_modelos.items():
        print(f"   • {k:<45}: {v:>3} artigos ({v/total_artigos*100:5.1f}%)")

    print("\n3. Saídas / Alvos Minerais:")
    for k, v in cont_saidas.items():
        print(f"   • {k:<45}: {v:>3} artigos ({v/total_artigos*100:5.1f}%)")

    print("\n4. Acurácia Média por Modelo:")
    for fam, r in acc_stats.iterrows():
        print(f"   • {fam:<35}: Média={r['mean']:6.2f}% (Min: {r['min']:5.1f}% | Max: {r['max']:5.1f}% | n={int(r['count'])})")
    print("=" * 75)


if __name__ == "__main__":
    gerar_todos_graficos()
