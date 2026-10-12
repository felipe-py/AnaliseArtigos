#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para gerar Diagrama de Sankey de Fluxo Multidimensional:
Camadas de Entrada (Sensores/Dados) -> Modelos Computacionais (IA/DL/ML) -> Saídas (Alvos Minerais/Mapeamento)

Gera:
1. 'graficos/sankey_entrada_modelo_saida.html': Diagrama interativo de alta performance (HTML/Plotly)
2. 'graficos/sankey_entrada_modelo_saida.png': Versão estática em alta resolução (300 DPI) para relatórios
"""

import sys
from pathlib import Path
from collections import defaultdict

import pandas as pd
import numpy as np
import plotly.graph_objects as go

# Configuração de encoding UTF-8 para Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def carregar_dados(caminho_csv: str = "sintese_entrada_modelo_saida.csv") -> pd.DataFrame:
    """Carrega dados da base consolidada."""
    path = Path(caminho_csv)
    if not path.exists():
        path_alt = Path("base_consolidada_classificacao_e_sota.csv")
        if path_alt.exists():
            path = path_alt
        else:
            raise FileNotFoundError(f"Arquivo de dados não encontrado: {path.resolve()}")
    return pd.read_csv(path, sep=";", encoding="utf-8-sig")


def extrair_dimensoes_artigo(r: pd.Series):
    """
    Identifica as dimensões de Entrada, Modelo e Saída para um artigo específico.
    Retorna três listas: (camadas, modelos, saidas).
    """
    # 1. ENTRADA (Camadas de dados e sensores)
    texto_in = " ".join([
        str(r.get("ENTRADA: Fontes e Sensores", "")),
        str(r.get("SÍNTESE: Entrada", "")),
        str(r.get("ID: Título do Artigo", "")),
        str(r.get("Termos Sensoriamento Remoto", "")),
    ]).lower()
    
    camadas = []
    if "aster" in texto_in or str(r.get("ENTRADA: Usa ASTER", "")).lower() in ["sim", "true", "yes"]:
        camadas.append("ASTER (SWIR/TIR)")
    if any(k in texto_in for k in ["hiperespectral", "hyperspectral", "prisma", "aviris", "hyperion"]) or \
       str(r.get("ENTRADA: Hiperespectral", "")).lower() in ["sim", "true", "yes"]:
        camadas.append("Hiperespectral (PRISMA/AVIRIS)")
    if any(k in texto_in for k in ["landsat", "sentinel-2", "sentinel 2", "worldview", "spot", "multiespectral", "multispectral"]):
        camadas.append("Óptico (Landsat/Sentinel)")
    if any(k in texto_in for k in ["sar", "radar", "sentinel-1", "palsar"]):
        camadas.append("Radar SAR")
    if any(k in texto_in for k in ["geoquím", "geochemi", "sedimento", "stream sediment", "soil"]):
        camadas.append("Geoquímica")
    if any(k in texto_in for k in ["geofís", "geophysic", "aeromagnet", "magnetometria", "gamma", "gama", "gravimetr"]):
        camadas.append("Aerogeofísica")
    if any(k in texto_in for k in ["drone", "uav", "vant"]):
        camadas.append("Drone / VANT")
        
    if not camadas:
        camadas.append("Satélite Convencional")

    # 2. MODELO (Arquiteturas computacionais)
    texto_mod = " ".join([
        str(r.get("MODELO: Modelos DL", "")),
        str(r.get("MODELO: Arquitetura Base", "")),
        str(r.get("SÍNTESE: Modelo", "")),
        str(r.get("ID: Título do Artigo", "")),
        str(r.get("Termos Deep Learning", "")),
    ]).lower()
    
    modelos = []
    if any(k in texto_mod for k in ["cnn", "convolutional", "lenet", "resnet", "densenet", "vgg", "1d-cnn", "2d-cnn", "3d-cnn", "fcnn"]):
        modelos.append("Redes Convolucionais (CNNs)")
    if any(k in texto_mod for k in ["u-net", "unet", "fcn", "segmentation", "segmentação"]):
        modelos.append("U-Net / Segmentação")
    if any(k in texto_mod for k in ["transformer", "vit", "vision transformer", "attention", "atenção", "mamba"]):
        modelos.append("Transformers / ViT")
    if any(k in texto_mod for k in ["gcn", "gnn", "graph convolutional", "redes em grafo"]):
        modelos.append("Graph Networks (GCN)")
    if any(k in texto_mod for k in ["autoencoder", "dae", "vae", "auto-encoder", "deep autoencoder"]):
        modelos.append("Autoencoders (DAE/VAE)")
    if any(k in texto_mod for k in ["random forest", "svm", "support vector", "xgboost", "catboost", "ensemble"]):
        modelos.append("Machine Learning Clássico")
    if any(k in texto_mod for k in ["mlp", "multilayer perceptron", "ann", "artificial neural network"]):
        modelos.append("Redes Neurais Densas (MLP)")
    if any(k in texto_mod for k in ["gan", "generative adversarial", "wgan"]):
        modelos.append("Redes Generativas (GAN)")
        
    if not modelos:
        modelos.append("Outros Modelos IA")

    # 3. SAÍDA (Alvos minerais e produtos)
    texto_out = " ".join([
        str(r.get("SAÍDA: Objeto Prospecção", "")),
        str(r.get("SÍNTESE: Saída", "")),
        str(r.get("ID: Título do Artigo", "")),
        str(r.get("Termos Prospecção Mineral", "")),
    ]).lower()
    
    saidas = []
    if "fosfat" in texto_out or "phosphate" in texto_out or str(r.get("SAÍDA: Prospecta Fosfato", "")).lower() in ["sim", "true", "yes"]:
        saidas.append("Fosfato")
    if any(k in texto_out for k in ["cobre", "copper", "porphyry cu", "skarn"]):
        saidas.append("Cobre (Pórfiro/Skarn)")
    if any(k in texto_out for k in ["ouro", "gold", "orogenic", "epithermal"]):
        saidas.append("Ouro (Orogênico/Epitermal)")
    if any(k in texto_out for k in ["alteração hidrotermal", "hydrothermal alteration", "alteration"]):
        saidas.append("Alteração Hidrotermal")
    if any(k in texto_out for k in ["litológ", "litholog", "geological mapping", "mapeamento geológico", "rochas"]):
        saidas.append("Mapeamento Litológico")
    if any(k in texto_out for k in ["lítio", "lithium", "pegmatit"]):
        saidas.append("Lítio (Pegmatitos)")
    if any(k in texto_out for k in ["ferro", "iron ore", "bif"]):
        saidas.append("Minério de Ferro")
    if any(k in texto_out for k in ["zinco", "chumbo", "pb-zn", "nickel", "níquel", "polymetallic", "polimetál"]):
        saidas.append("Polimetálicos (Pb-Zn/Ni)")
        
    if not saidas:
        saidas.append("Prospecção Mineral Geral")
        
    return camadas, modelos, saidas


def gerar_diagrama_sankey(caminho_csv: str = "sintese_entrada_modelo_saida.csv"):
    df = carregar_dados(caminho_csv)
    total_artigos = len(df)
    
    pasta_saida = Path("graficos")
    pasta_saida.mkdir(parents=True, exist_ok=True)
    
    print("=" * 75)
    print(" CONSTRUÇÃO DO DIAGRAMA DE SANKEY (ENTRADA -> MODELO -> SAÍDA)")
    print("=" * 75)
    print(f"-> Processando {total_artigos} artigos científicos...")

    # Mapeamentos de fluxos com conservação balanceada
    fluxos_entrada_modelo = defaultdict(float)
    fluxos_modelo_saida = defaultdict(float)
    
    rotas_completas = defaultdict(float)

    for _, r in df.iterrows():
        camadas, modelos, saidas = extrair_dimensoes_artigo(r)
        
        # Ponderação para que cada artigo contribua com peso unitário (1.0)
        n_c = len(camadas)
        n_m = len(modelos)
        n_s = len(saidas)
        
        peso_in_mod = 1.0 / (n_c * n_m)
        peso_mod_out = 1.0 / (n_m * n_s)
        
        for c in camadas:
            for m in modelos:
                fluxos_entrada_modelo[(c, m)] += peso_in_mod
                
        for m in modelos:
            for s in saidas:
                fluxos_modelo_saida[(m, s)] += peso_mod_out

        # Rota tripla para ranking textual
        for c in camadas:
            for m in modelos:
                for s in saidas:
                    rotas_completas[(c, m, s)] += 1.0 / (n_c * n_m * n_s)

    # Identifica nós únicos em cada estágio
    nos_entrada = sorted(list({c for c, m in fluxos_entrada_modelo.keys()}))
    nos_modelo = sorted(list({m for c, m in fluxos_entrada_modelo.keys()} | {m for m, s in fluxos_modelo_saida.keys()}))
    nos_saida = sorted(list({s for m, s in fluxos_modelo_saida.keys()}))

    # Atribuição de índices globais
    todos_nos = []
    custom_data = []
    mapa_indices = {}
    
    for n in nos_entrada:
        mapa_indices[f"IN_{n}"] = len(todos_nos)
        todos_nos.append(n)
        custom_data.append(f"[Entrada] {n}")
        
    for n in nos_modelo:
        mapa_indices[f"MOD_{n}"] = len(todos_nos)
        todos_nos.append(n)
        custom_data.append(f"[Modelo] {n}")
        
    for n in nos_saida:
        mapa_indices[f"OUT_{n}"] = len(todos_nos)
        todos_nos.append(n)
        custom_data.append(f"[Saída] {n}")

    # Paletas de cores para nós
    cores_entrada = ["#16A085", "#1ABC9C", "#2ECC71", "#27AE60", "#48C9B0", "#76D7C4", "#117864", "#0E6251"]
    cores_modelo = ["#2980B9", "#3498DB", "#5DADE2", "#1F618D", "#2874A6", "#154360", "#2471A3", "#5499C7", "#4A90E2"]
    cores_saida = ["#D35400", "#E67E22", "#F39C12", "#CA6F1E", "#BA4A00", "#873600", "#E59866", "#DC7633", "#B9770E"]

    cores_nos = []
    for i, n in enumerate(nos_entrada):
        cores_nos.append(cores_entrada[i % len(cores_entrada)])
    for i, n in enumerate(nos_modelo):
        cores_nos.append(cores_modelo[i % len(cores_modelo)])
    for i, n in enumerate(nos_saida):
        cores_nos.append(cores_saida[i % len(cores_saida)])

    # Construção dos links (Source, Target, Value, Link Color)
    sources = []
    targets = []
    values = []
    cores_links = []

    # Links 1: Entrada -> Modelo (RGBA translúcido baseado no nó de origem)
    for (c, m), peso in fluxos_entrada_modelo.items():
        if peso > 0.05:
            src_idx = mapa_indices[f"IN_{c}"]
            tgt_idx = mapa_indices[f"MOD_{m}"]
            sources.append(src_idx)
            targets.append(tgt_idx)
            values.append(round(peso, 2))
            
            # Converte cor hex para rgba com 40% opacidade
            hex_cor = cores_nos[src_idx].lstrip("#")
            r_c, g_c, b_c = tuple(int(hex_cor[j:j+2], 16) for j in (0, 2, 4))
            cores_links.append(f"rgba({r_c}, {g_c}, {b_c}, 0.40)")

    # Links 2: Modelo -> Saída (RGBA translúcido baseado no modelo)
    for (m, s), peso in fluxos_modelo_saida.items():
        if peso > 0.05:
            src_idx = mapa_indices[f"MOD_{m}"]
            tgt_idx = mapa_indices[f"OUT_{s}"]
            sources.append(src_idx)
            targets.append(tgt_idx)
            values.append(round(peso, 2))
            
            hex_cor = cores_nos[src_idx].lstrip("#")
            r_c, g_c, b_c = tuple(int(hex_cor[j:j+2], 16) for j in (0, 2, 4))
            cores_links.append(f"rgba({r_c}, {g_c}, {b_c}, 0.40)")

    # Script JavaScript para reposicionar os rótulos das colunas dos extremos para FORA do gráfico
    js_posicionar_rotulos_extremos = r"""
    function reposicionarRotulosExtremos() {
        const gd = document.getElementById('{plot_id}');
        if (!gd) return;
        const nodes = gd.querySelectorAll('.sankey-node');
        if (!nodes || nodes.length === 0) return;

        let minX = Infinity, maxX = -Infinity;
        const xList = [];
        nodes.forEach(n => {
            const tr = n.getAttribute('transform') || '';
            const m = /translate\(([-\d\.]+)/.exec(tr);
            if (m) {
                const x = parseFloat(m[1]);
                if (x < minX) minX = x;
                if (x > maxX) maxX = x;
                xList.push({node: n, x: x});
            }
        });

        xList.forEach(({node, x}) => {
            const label = node.querySelector('.node-label');
            const rect = node.querySelector('rect');
            if (!label || !rect) return;

            const rectW = parseFloat(rect.getAttribute('width')) || 22;
            const tr = label.getAttribute('transform') || '';
            const m = /translate\(([-\d\.]+),\s*([-\d\.]+)\)/.exec(tr);
            const currY = m ? m[2] : '0';

            // Coluna da Esquerda (Entrada): Posiciona o rótulo FORA, à esquerda do nó
            if (Math.abs(x - minX) < 15) {
                label.setAttribute('text-anchor', 'end');
                label.setAttribute('transform', 'translate(-8, ' + currY + ')');
            }
            // Coluna da Direita (Saída): Posiciona o rótulo FORA, à direita do nó
            else if (Math.abs(x - maxX) < 15) {
                label.setAttribute('text-anchor', 'start');
                label.setAttribute('transform', 'translate(' + (rectW + 8) + ', ' + currY + ')');
            }
        });
    }

    // Execuções temporizadas para garantir renderização imediata e estável
    setTimeout(reposicionarRotulosExtremos, 40);
    setTimeout(reposicionarRotulosExtremos, 180);
    setTimeout(reposicionarRotulosExtremos, 450);

    const graphDiv = document.getElementById('{plot_id}');
    if (graphDiv) {
        if (graphDiv.on) {
            graphDiv.on('plotly_afterplot', reposicionarRotulosExtremos);
        }
        const observer = new MutationObserver(reposicionarRotulosExtremos);
        observer.observe(graphDiv, { childList: true, subtree: true });
    }
    """

    def construir_sankey_obj(com_rotulos: bool = True, proximidade_colunas: float = 0.70, largura: int = 1080, altura: int = 780):
        rotulos = todos_nos if com_rotulos else ["" for _ in todos_nos]
        sub = "Rótulos nos Extremos Externos" if com_rotulos else "Visão Minimalista / Sem Rótulos (Hover Interativo)"
        
        prox = max(0.15, min(1.0, float(proximidade_colunas)))
        x_in = round(0.5 - prox / 2.0, 3)
        x_mod = 0.50
        x_out = round(0.5 + prox / 2.0, 3)

        fig = go.Figure(data=[go.Sankey(
            arrangement="snap",
            domain=dict(x=[x_in, x_out], y=[0.02, 0.92]),  # APROXIMAÇÃO DAS COLUNAS
            node=dict(
                pad=18,
                thickness=22 if com_rotulos else 16,
                line=dict(color="#2C3E50", width=0.8),
                label=rotulos,
                customdata=custom_data,
                color=cores_nos,
                hovertemplate="<b>%{customdata}</b><br>Fluxo Ponderado: %{value:.1f} artigos<extra></extra>"
            ),
            link=dict(
                source=sources,
                target=targets,
                value=values,
                color=cores_links,
                hovertemplate="<b>%{source.customdata}</b>  →  <b>%{target.customdata}</b><br>Volume: %{value:.2f} artigos<extra></extra>"
            )
        )])

        fig.update_layout(
            title=dict(
                text="<b>DIAGRAMA DE SANKEY: FLUXO METODOLÓGICO DA LITERATURA</b><br>"
                     f"<span style='font-size: 13px; color: #566573;'>Rastreamento de Camadas de Entrada → Modelos de Inteligência Artificial → Saídas | {sub} | Proximidade: {int(prox*100)}% (N = {total_artigos} artigos)</span>",
                font=dict(size=18, family="Segoe UI, Arial, sans-serif", color="#1A252F"),
                x=0.03,
                y=0.97
            ),
            annotations=[
                dict(x=x_in, y=1.06, xref="paper", yref="paper", text="<b>🟩 1. ENTRADA</b>", showarrow=False, font=dict(size=12, color="#16A085"), xanchor="center"),
                dict(x=x_mod, y=1.06, xref="paper", yref="paper", text="<b>🟦 2. MODELOS</b>", showarrow=False, font=dict(size=12, color="#2980B9"), xanchor="center"),
                dict(x=x_out, y=1.06, xref="paper", yref="paper", text="<b>🟧 3. SAÍDAS</b>", showarrow=False, font=dict(size=12, color="#D35400"), xanchor="center"),
            ],
            font=dict(size=11, family="Segoe UI, Arial, sans-serif", color="#2C3E50"),
            height=altura,
            width=largura,
            margin=dict(l=40, r=40, t=95, b=35),
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF"
        )
        return fig

    def exportar_sankey_png(caminho_html: Path, caminho_png: Path, fig_obj: go.Figure = None, scale: float = 2.5, largura: int = 1080, altura: int = 780) -> bool:
        """
        Exporta a imagem PNG em alta resolução (300 DPI) a partir do HTML,
        garantindo que os rótulos posicionados fora das colunas fiquem perfeitamente renderizados.
        """
        caminho_png.parent.mkdir(parents=True, exist_ok=True)
        sucesso = False
        try:
            from playwright.sync_api import sync_playwright
            with sync_playwright() as p:
                browser = p.chromium.launch()
                page = browser.new_page(viewport={'width': largura, 'height': altura}, device_scale_factor=scale)
                page.goto(caminho_html.resolve().as_uri(), wait_until='networkidle')
                page.wait_for_timeout(400)
                page.screenshot(path=str(caminho_png))
                browser.close()
                sucesso = True
        except Exception:
            pass

        if not sucesso and fig_obj is not None:
            try:
                fig_obj.write_image(str(caminho_png), scale=scale)
                sucesso = True
            except Exception as err:
                print(f"   [Aviso exportação PNG]: {err}")

        return sucesso

    # 1. Versão COM RÓTULOS (Rótulos externos para colunas dos extremos)
    fig_com = construir_sankey_obj(com_rotulos=True)
    html_com_1 = pasta_saida / "sankey_entrada_modelo_saida.html"
    html_com_2 = pasta_saida / "sankey_com_rotulos.html"
    fig_com.write_html(str(html_com_1), include_plotlyjs="cdn", post_script=js_posicionar_rotulos_extremos)
    fig_com.write_html(str(html_com_2), include_plotlyjs="cdn", post_script=js_posicionar_rotulos_extremos)

    png_com_1 = pasta_saida / "sankey_entrada_modelo_saida.png"
    png_com_2 = pasta_saida / "sankey_com_rotulos.png"
    ok1 = exportar_sankey_png(html_com_1, png_com_1, fig_com, scale=2.5)
    ok2 = exportar_sankey_png(html_com_2, png_com_2, fig_com, scale=2.5)
    if ok1 or ok2:
        print("   ✓ Arquivos COM rótulos nos extremos externos salvos (HTML e PNG).")

    # 2. Versão SEM RÓTULOS (Minimalista)
    fig_sem = construir_sankey_obj(com_rotulos=False)
    html_sem = pasta_saida / "sankey_sem_rotulos.html"
    fig_sem.write_html(str(html_sem), include_plotlyjs="cdn")

    png_sem = pasta_saida / "sankey_sem_rotulos.png"
    ok_sem = exportar_sankey_png(html_sem, png_sem, fig_sem, scale=2.5)
    if ok_sem:
        print("   ✓ Arquivos SEM rótulos salvos (HTML e PNG).")

    # Estatísticas de Top Rotas
    print("\n" + "=" * 75)
    print(" PRINCIPAIS ROTAS METODOLÓGICAS (TOP 10 CAMINHOS MAIS FREQUENTES)")
    print("=" * 75)
    top_rotas = sorted(rotas_completas.items(), key=lambda x: x[1], reverse=True)[:10]
    for idx_r, ((c, m, s), v) in enumerate(top_rotas, 1):
        print(f"{idx_r:>2}. {c}  ──►  {m}  ──►  {s} ({v:.2f} eq. artigos)")
    print("=" * 75)


if __name__ == "__main__":
    gerar_diagrama_sankey()
