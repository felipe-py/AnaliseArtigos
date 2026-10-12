#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para consolidar e sintetizar as dimensões de ENTRADA, MODELO e SAÍDA
a partir da base consolidada de classificação e SOTA.

Gera:
1. 'sintese_entrada_modelo_saida.xlsx': Planilha multi-aba com design executivo:
   - Aba 1: 'Síntese Executiva (E-M-S)': Visão compacta, limpa e direta para apresentações e leitura rápida.
   - Aba 2: 'Matriz Completa E-M-S': Todas as variáveis categorizadas por blocos temáticos coloridos.
   - Aba 3: 'Painel Estatístico': Rankings de modelos mais usados, sensores e alvos minerais.
2. 'sintese_entrada_modelo_saida.csv': Versão tabular unificada em UTF-8 com BOM (separador ';').
"""

import sys
import re
import unicodedata
from pathlib import Path
from typing import Any

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Encoding UTF-8 para Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def limpar_texto(valor: Any) -> str:
    """Normaliza texto removendo quebras extras e valores nulos."""
    if valor is None or pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.lower() in ["nan", "none", "null", "não mencionado", "no mencionado", "n/a", "-"]:
        return ""
    return texto


def gerar_resumo_entrada(row: pd.Series) -> str:
    """Gera síntese concisa dos dados de entrada (sensores, dados, bandas)."""
    partes = []
    
    # 1. Fonte/Sensores
    fontes = limpar_texto(row.get("SOTA: Fonte Imagem", ""))
    usa_aster = limpar_texto(row.get("SOTA: Usa ASTER", ""))
    hiper = limpar_texto(row.get("SOTA: Hiperespectral", ""))
    
    sensores_destaque = []
    if usa_aster.lower() in ["sim", "true", "yes"]:
        sensores_destaque.append("ASTER")
    if hiper.lower() in ["sim", "true", "yes"]:
        sensores_destaque.append("Hiperespectral")
    
    if fontes:
        # Se tiver sensores específicos destacados, agrega
        if sensores_destaque:
            partes.append(f"{fontes} ({', '.join(sensores_destaque)})")
        else:
            partes.append(fontes)
    elif sensores_destaque:
        partes.append(", ".join(sensores_destaque))
    
    # 2. Bandas
    bandas = limpar_texto(row.get("SOTA: Nº Bandas", ""))
    if bandas:
        partes.append(f"{bandas} bandas")
        
    # 3. Tamanho Dataset
    tam = limpar_texto(row.get("SOTA: Tamanho Dataset", ""))
    if tam:
        partes.append(f"Tam: {tam}")
        
    # 4. Augmentation
    aug = limpar_texto(row.get("SOTA: Data Augmentation", ""))
    if aug.lower() in ["sim", "true", "yes"]:
        partes.append("c/ Augmentation")
        
    if partes:
        return " | ".join(partes)
    
    # Fallback inteligente para termos extraídos no ranking inicial
    termos_rs = limpar_texto(row.get("Termos Sensoriamento Remoto", ""))
    if termos_rs:
        return f"[Termos RS]: {termos_rs}"
    return "Dados não especificados no SOTA"


def gerar_resumo_modelo(row: pd.Series) -> str:
    """Gera síntese concisa do modelo, arquitetura e treinamento."""
    partes = []
    
    # 1. Modelos principais
    modelos = limpar_texto(row.get("SOTA: Modelos DL", ""))
    if modelos:
        partes.append(modelos)
        
    # 2. Arquitetura Base
    arq = limpar_texto(row.get("SOTA: Arquitetura Base", ""))
    if arq and arq.lower() not in modelos.lower():
        partes.append(f"Base: {arq}")
        
    # 3. Transfer learning
    tl = limpar_texto(row.get("SOTA: Transfer Learning", ""))
    if tl.lower() in ["sim", "true", "yes"]:
        partes.append("Transfer Learning: Sim")
        
    # 4. Otimizador / Loss
    opt = limpar_texto(row.get("SOTA: Otimizador", ""))
    loss = limpar_texto(row.get("SOTA: Loss Functions", ""))
    detalhes_treino = []
    if opt:
        detalhes_treino.append(f"Opt: {opt}")
    if loss:
        detalhes_treino.append(f"Loss: {loss}")
    if detalhes_treino:
        partes.append(", ".join(detalhes_treino))
        
    if partes:
        return " | ".join(partes)
        
    # Fallback inteligente para termos de IA/DL
    termos_dl = limpar_texto(row.get("Termos Deep Learning", ""))
    if termos_dl:
        return f"[Termos DL]: {termos_dl}"
    return "Modelo não especificado no SOTA"


def gerar_resumo_saida(row: pd.Series) -> str:
    """Gera síntese concisa do alvo de prospecção e métricas atingidas."""
    partes = []
    
    # 1. Alvo
    alvo = limpar_texto(row.get("SOTA: Objeto Prospecção", ""))
    fosfato = limpar_texto(row.get("SOTA: Prospecta Fosfato", ""))
    if fosfato.lower() in ["sim", "true", "yes"]:
        if alvo:
            partes.append(f"Alvo: {alvo} (Fosfato)")
        else:
            partes.append("Alvo: Fosfato")
    elif alvo:
        partes.append(f"Alvo: {alvo}")
        
    # 2. Melhores Métricas
    acc = limpar_texto(row.get("SOTA: Melhor Acurácia", ""))
    iou = limpar_texto(row.get("SOTA: Melhor IoU", ""))
    f1 = limpar_texto(row.get("SOTA: Melhor F1", ""))
    
    metricas = []
    if acc:
        metricas.append(f"Acc: {acc}")
    if f1:
        metricas.append(f"F1: {f1}")
    if iou:
        metricas.append(f"IoU: {iou}")
        
    if metricas:
        partes.append(" | ".join(metricas))
    else:
        outras_met = limpar_texto(row.get("SOTA: Métricas", ""))
        if outras_met:
            partes.append(f"Métricas: {outras_met}")
            
    if partes:
        return " | ".join(partes)
        
    # Fallback inteligente para termos de prospecção mineral
    termos_mp = limpar_texto(row.get("Termos Prospecção Mineral", ""))
    if termos_mp:
        return f"[Termos Prospecção]: {termos_mp}"
    return "Resultados não especificados no SOTA"


def construir_planilha_sintetizada(
    arquivo_base_csv: str = "base_consolidada_classificacao_e_sota.csv",
    arquivo_saida_xlsx: str = "sintese_entrada_modelo_saida.xlsx",
    arquivo_saida_csv: str = "sintese_entrada_modelo_saida.csv",
):
    print("=" * 75)
    print(" CONSTRUÇÃO DA SÍNTESE EXECUTIVA: ENTRADA | MODELO | SAÍDA")
    print("=" * 75)
    
    path_csv = Path(arquivo_base_csv)
    if not path_csv.exists():
        raise FileNotFoundError(f"Arquivo base não encontrado: {path_csv.resolve()}")
        
    # Carrega dados consolidados
    df_base = pd.read_csv(path_csv, sep=";", encoding="utf-8-sig")
    print(f"-> Base carregada com sucesso: {len(df_base)} artigos.")

    # Normaliza colunas com quebras de linha (ex: Score\n(0–100))
    df_base.columns = [c.replace("\n", " ").strip() for c in df_base.columns]
    
    # 1. GERAÇÃO DAS COLUNAS SINTÉTICAS
    print("-> Computando sínteses de Entrada, Modelo e Saída por artigo...")
    df_base["SÍNTESE: Entrada"] = df_base.apply(gerar_resumo_entrada, axis=1)
    df_base["SÍNTESE: Modelo"] = df_base.apply(gerar_resumo_modelo, axis=1)
    df_base["SÍNTESE: Saída"] = df_base.apply(gerar_resumo_saida, axis=1)
    
    # Extrai o score total para ordenação e filtro
    col_score = [c for c in df_base.columns if "score" in c.lower() and "100" in c][0]
    col_class = [c for c in df_base.columns if "classific" in c.lower()][0]
    col_titulo = [c for c in df_base.columns if "título" in c.lower() or "titulo" in c.lower()][0]
    
    # =========================================================================
    # ABA 1: SÍNTESE EXECUTIVA (VISÃO RÁPIDA, DIRETA E LIMPA)
    # =========================================================================
    cols_executiva = [
        "Rank",
        col_class,
        col_score,
        col_titulo,
        "Ano",
        "Status SOTA",
        "SÍNTESE: Entrada",
        "SÍNTESE: Modelo",
        "SÍNTESE: Saída",
        "SOTA: Objeto Prospecção",
        "SOTA: Melhor Acurácia",
        "SOTA: Melhor F1",
    ]
    cols_exec_existentes = [c for c in cols_executiva if c in df_base.columns]
    df_executiva = df_base[cols_exec_existentes].copy()
    
    # Renomeia para exibição elegante
    mapa_renomear_exec = {
        col_class: "Classificação",
        col_score: "Score (0–100)",
        col_titulo: "Título do Artigo",
        "SÍNTESE: Entrada": "ENTRADA (Sensores, Dados e Bandas)",
        "SÍNTESE: Modelo": "MODELO (Arquitetura e Treinamento)",
        "SÍNTESE: Saída": "SAÍDA (Alvo e Desempenho)",
        "SOTA: Objeto Prospecção": "Minério / Alvo",
        "SOTA: Melhor Acurácia": "Acurácia",
        "SOTA: Melhor F1": "F1-Score",
    }
    df_executiva = df_executiva.rename(columns=mapa_renomear_exec)

    # =========================================================================
    # ABA 2: MATRIZ COMPLETA CATEGORIZADA POR DIMENSÕES
    # =========================================================================
    colunas_ordenadas = [
        # Identificação
        ("Rank", "ID: Rank"),
        (col_class, "ID: Classificação"),
        (col_score, "ID: Score (0–100)"),
        (col_titulo, "ID: Título do Artigo"),
        ("Ano", "ID: Ano"),
        ("Periódico/Fonte", "ID: Periódico"),
        ("Status SOTA", "ID: Status Pareamento"),
        
        # Sínteses Centrais
        ("SÍNTESE: Entrada", "SÍNTESE: Entrada"),
        ("SÍNTESE: Modelo", "SÍNTESE: Modelo"),
        ("SÍNTESE: Saída", "SÍNTESE: Saída"),
        
        # Bloco ENTRADA Detalhado
        ("SOTA: Fonte Imagem", "ENTRADA: Fontes e Sensores"),
        ("SOTA: Usa ASTER", "ENTRADA: Usa ASTER"),
        ("SOTA: Hiperespectral", "ENTRADA: Hiperespectral"),
        ("SOTA: Nº Bandas", "ENTRADA: Nº Bandas"),
        ("SOTA: Tamanho Dataset", "ENTRADA: Tamanho Dataset"),
        ("SOTA: Split Dataset", "ENTRADA: Divisão Treino/Teste"),
        ("SOTA: Índices Espectrais", "ENTRADA: Índices Espectrais"),
        ("SOTA: Data Augmentation", "ENTRADA: Data Augmentation"),
        
        # Bloco MODELO Detalhado
        ("SOTA: Modelos DL", "MODELO: Modelos DL"),
        ("SOTA: Arquitetura Base", "MODELO: Arquitetura Base"),
        ("SOTA: Transfer Learning", "MODELO: Transfer Learning"),
        ("SOTA: Loss Functions", "MODELO: Função de Perda (Loss)"),
        ("SOTA: Otimizador", "MODELO: Otimizador"),
        ("SOTA: LR Scheduler", "MODELO: LR Scheduler"),
        ("SOTA: GPU/VRAM", "MODELO: GPU/VRAM"),
        ("SOTA: RAM Sistema", "MODELO: RAM Sistema"),
        
        # Bloco SAÍDA Detalhado
        ("SOTA: Objeto Prospecção", "SAÍDA: Objeto Prospecção"),
        ("SOTA: Prospecta Fosfato", "SAÍDA: Prospecta Fosfato"),
        ("SOTA: Melhor Acurácia", "SAÍDA: Melhor Acurácia"),
        ("SOTA: Melhor IoU", "SAÍDA: Melhor IoU"),
        ("SOTA: Melhor F1", "SAÍDA: Melhor F1"),
        ("SOTA: Métricas", "SAÍDA: Conjunto de Métricas"),
        ("SOTA: Desafios e Limitações", "SAÍDA: Desafios e Limitações"),
        ("SOTA: País Área Estudo", "SAÍDA: País Área Estudo"),
    ]
    
    cols_matriz = []
    mapa_renomear_matriz = {}
    for col_orig, col_novo in colunas_ordenadas:
        if col_orig in df_base.columns:
            cols_matriz.append(col_orig)
            mapa_renomear_matriz[col_orig] = col_novo
            
    df_matriz = df_base[cols_matriz].copy().rename(columns=mapa_renomear_matriz)

    # =========================================================================
    # ABA 3: PAINEL ESTATÍSTICO DE DISTRIBUIÇÃO
    # =========================================================================
    # Estatísticas de Sensores
    total_artigos = len(df_base)
    cruzados = (df_base["Status SOTA"] == "Cruzado com Sucesso").sum()
    com_aster = df_base["SOTA: Usa ASTER"].astype(str).str.lower().isin(["sim", "true", "yes"]).sum()
    com_hiper = df_base["SOTA: Hiperespectral"].astype(str).str.lower().isin(["sim", "true", "yes"]).sum()
    com_fosfato = df_base["SOTA: Prospecta Fosfato"].astype(str).str.lower().isin(["sim", "true", "yes"]).sum()
    
    df_kpis = pd.DataFrame([
        {"Indicador": "Total de Artigos Selecionados", "Valor": total_artigos, "% do Total": "100%"},
        {"Indicador": "Artigos com SOTA Completo Mapeado", "Valor": cruzados, "% do Total": f"{cruzados/total_artigos*100:.1f}%"},
        {"Indicador": "Artigos que usam Sensor ASTER", "Valor": com_aster, "% do Total": f"{com_aster/total_artigos*100:.1f}%"},
        {"Indicador": "Artigos que usam Dados Hiperespectrais", "Valor": com_hiper, "% do Total": f"{com_hiper/total_artigos*100:.1f}%"},
        {"Indicador": "Artigos com Foco Específico em Fosfato", "Valor": com_fosfato, "% do Total": f"{com_fosfato/total_artigos*100:.1f}%"},
    ])

    # Frequência de Objetos de Prospecção / Minérios
    alvos_series = df_base["SOTA: Objeto Prospecção"].dropna().astype(str)
    alvos_series = alvos_series[~alvos_series.str.lower().isin(["não mencionado", "no mencionado", ""])]
    alvos_contagem = alvos_series.value_counts().reset_index()
    alvos_contagem.columns = ["Objeto / Minério Prospectado", "Quantidade de Artigos"]

    # Frequência de Modelos DL
    modelos_series = df_base["SOTA: Modelos DL"].dropna().astype(str)
    modelos_series = modelos_series[~modelos_series.str.lower().isin(["não mencionado", "no mencionado", ""])]
    
    # Separa por ponto-e-vírgula se houver múltiplos modelos
    todos_modelos = []
    for m_str in modelos_series:
        for m in m_str.split(";"):
            m_clean = m.strip()
            if m_clean and m_clean.lower() not in ["não mencionado", "no mencionado"]:
                todos_modelos.append(m_clean)
    df_freq_modelos = pd.Series(todos_modelos).value_counts().head(15).reset_index()
    df_freq_modelos.columns = ["Modelo Computacional / Rede Neural", "Frequência Citada"]

    # =========================================================================
    # EXPORTAÇÃO PARA EXCEL COM FORMATAÇÃO PREMIUM
    # =========================================================================
    path_out_xlsx = Path(arquivo_saida_xlsx)
    print(f"-> Gerando Excel estilizado: {path_out_xlsx.resolve()}")
    
    with pd.ExcelWriter(path_out_xlsx, engine="openpyxl") as writer:
        df_executiva.to_excel(writer, sheet_name="Síntese Executiva (E-M-S)", index=False)
        df_matriz.to_excel(writer, sheet_name="Matriz Completa E-M-S", index=False)
        
        # Painel estatístico na mesma aba com tabelas empilhadas
        df_kpis.to_excel(writer, sheet_name="Painel Estatístico", index=False, startrow=0)
        df_freq_modelos.to_excel(writer, sheet_name="Painel Estatístico", index=False, startrow=len(df_kpis) + 3)
        alvos_contagem.head(10).to_excel(writer, sheet_name="Painel Estatístico", index=False, startrow=len(df_kpis) + len(df_freq_modelos) + 6)

    # Estilização visual no OpenPyXL
    formatar_excel_sintetizado(path_out_xlsx)
    print("   ✓ Arquivo Excel gerado e estilizado com sucesso.")

    # =========================================================================
    # EXPORTAÇÃO PARA CSV (UTF-8 com BOM e sep=';')
    # =========================================================================
    path_out_csv = Path(arquivo_saida_csv)
    print(f"-> Gerando CSV unificado  : {path_out_csv.resolve()}")
    df_matriz.to_csv(path_out_csv, sep=";", index=False, encoding="utf-8-sig")
    print("   ✓ Arquivo CSV gerado com sucesso (UTF-8 com BOM).")

    print("=" * 75)
    print(" CONCLUÍDO COM SUCESSO!")
    print("=" * 75)


def formatar_excel_sintetizado(caminho_xlsx: Path):
    """Aplica design profissional com paleta executiva por dimensão."""
    try:
        wb = openpyxl.load_workbook(caminho_xlsx)

        # Paleta de Cores
        fill_id = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")       # Azul Petróleo Escuro (Identificação)
        fill_entrada = PatternFill(start_color="16A085", end_color="16A085", fill_type="solid")  # Verde Esmeralda (Entrada)
        fill_modelo = PatternFill(start_color="2980B9", end_color="2980B9", fill_type="solid")   # Azul Cobalto (Modelo)
        fill_saida = PatternFill(start_color="D35400", end_color="D35400", fill_type="solid")    # Laranja Queimado (Saída)
        fill_sintese = PatternFill(start_color="8E44AD", end_color="8E44AD", fill_type="solid")  # Roxo Executivo (Sínteses Centrais)
        fill_zebra = PatternFill(start_color="F8F9F9", end_color="F8F9F9", fill_type="solid")    # Fundo sutil alternado

        font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        font_data = Font(name="Segoe UI", size=9)
        border_light = Border(
            left=Side(style="thin", color="E2E4E6"),
            right=Side(style="thin", color="E2E4E6"),
            top=Side(style="thin", color="E2E4E6"),
            bottom=Side(style="thin", color="E2E4E6"),
        )

        for ws in wb.worksheets:
            ws.views.sheetView[0].showGridLines = True
            
            # Freeze panes
            if ws.title.startswith("Síntese"):
                ws.freeze_panes = "E2"  # Congela até Título
            elif ws.title.startswith("Matriz"):
                ws.freeze_panes = "E2"
            else:
                ws.freeze_panes = "A2"

            # Cabeçalho da Linha 1
            ws.row_dimensions[1].height = 34
            for cell in ws[1]:
                val = str(cell.value or "").upper()
                
                # Seleciona cor temático do cabeçalho
                if any(k in val for k in ["ENTRADA"]):
                    cell.fill = fill_entrada
                elif any(k in val for k in ["MODELO"]):
                    cell.fill = fill_modelo
                elif any(k in val for k in ["SAÍDA", "SAIDA", "ALVO", "DESEMPENHO"]):
                    cell.fill = fill_saida
                elif any(k in val for k in ["SÍNTESE", "SINTESE"]):
                    cell.fill = fill_sintese
                else:
                    cell.fill = fill_id
                    
                cell.font = font_header
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # Células de Dados
            for row_idx in range(2, ws.max_row + 1):
                ws.row_dimensions[row_idx].height = 24
                is_zebra = (row_idx % 2 == 0)

                for col_idx in range(1, ws.max_column + 1):
                    c = ws.cell(row=row_idx, column=col_idx)
                    c.font = font_data
                    c.border = border_light
                    
                    if is_zebra and not ws.title.startswith("Painel"):
                        c.fill = fill_zebra

                    col_title = str(ws.cell(1, col_idx).value or "").lower()
                    
                    # Alinhamento inteligente
                    if any(term in col_title for term in ["rank", "ano", "score", "acurácia", "f1", "iou", "aster", "hiper"]):
                        c.alignment = Alignment(horizontal="center", vertical="center")
                    elif "classificação" in col_title or "status" in col_title:
                        c.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            # Largura das Colunas
            for col in ws.columns:
                col_letter = get_column_letter(col[0].column)
                header_text = str(col[0].value or "")
                
                # Largura baseada no conteúdo
                max_content_len = max(len(str(c.value or "")) for c in col[:40])
                largura = max(max_content_len + 3, len(header_text) + 4, 12)
                
                # Limites inteligentes
                if "título" in header_text.lower():
                    largura = 50
                elif "síntese" in header_text.lower() or "entrada" in header_text.lower() or "modelo" in header_text.lower() or "saída" in header_text.lower():
                    largura = min(max(largura, 35), 65)
                elif largura > 70:
                    largura = 70
                    
                ws.column_dimensions[col_letter].width = largura

        wb.save(caminho_xlsx)
    except Exception as e:
        print(f"[Aviso] Formatação visual adicional ignorada: {e}")


if __name__ == "__main__":
    construir_planilha_sintetizada()
