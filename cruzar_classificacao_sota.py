#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para cruzar a planilha de classificação de artigos (Média-Alta, Alta, Excelente)
com os dados consolidados do Estado da Arte (state_of_the_art_final.csv e fontes).

Funcionalidades:
- Mapeamento difuso inteligente entre títulos de artigos e paper_id do SOTA
- Fusão completa das notas de relevância (Scores 0–100, DL, RS, MP) com as 26 variáveis de SOTA
- Integração de modelos de Deep Learning, sensores (ASTER, Hiperespectral), métricas e limitações
- Geração de planilha Excel multi-aba formatada profissionalmente:
    * Aba 1: 'Consolidado Classificação + SOTA'
    * Aba 2: 'Resumo por Classificação' (estatísticas de modelos e sensores por estrela)
    * Aba 3: 'Rastreabilidade de Fontes' (páginas e seções do PDF de onde os dados vieram)
- Exportação em CSV UTF-8 com BOM (compatível com Excel e Power BI)
"""

import os
import sys
import re
import unicodedata
import argparse
from pathlib import Path
from typing import Dict, List, Tuple, Any

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuração de encoding de saída para Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def normalizar_texto(texto: str) -> str:
    """Remove pontuações, acentos e separa CamelCase para matching robusto."""
    if not isinstance(texto, str):
        return ""
    texto = re.sub(r"([a-z])([A-Z])", r"\1 \2", texto)
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("utf-8").lower()
    texto = re.sub(r"[^a-z0-9]", " ", texto)
    return " ".join(texto.split())


def construir_mapa_cruzamento(
    titulos_classificacao: List[str], papers_sota: List[str]
) -> Dict[str, str]:
    """
    Cria mapa de correspondência de alta precisão entre títulos da classificação
    e paper_id do State of the Art.
    """
    sota_preparados = []
    for p in papers_sota:
        if not p or pd.isna(p):
            continue
        # Limpa sufixos de duplicatas como (1)
        p_str = str(p).strip()
        p_sem_dup = re.sub(r"\(\d+\)$", "", p_str).strip()
        p_norm = normalizar_texto(p_sem_dup)
        sota_preparados.append((p_norm, p_str))

    mapa = {}
    for titulo in titulos_classificacao:
        if not titulo or pd.isna(titulo):
            continue
        t_norm = normalizar_texto(str(titulo))
        if not t_norm:
            continue

        melhor_sota = None
        melhor_score = 0

        for s_norm, orig_p in sota_preparados:
            if t_norm == s_norm:
                melhor_sota = orig_p
                melhor_score = 100
                break

            # Prefixo truncado
            if len(s_norm) >= 20 and t_norm.startswith(s_norm):
                score = 98
            elif len(t_norm) >= 20 and s_norm.startswith(t_norm):
                score = 98
            elif s_norm in t_norm or t_norm in s_norm:
                score = min(len(s_norm), len(t_norm)) / max(len(s_norm), len(t_norm)) * 95
            else:
                p_t = set(t_norm.split())
                p_s = set(s_norm.split())
                if p_t and p_s:
                    inter = len(p_t & p_s)
                    union = len(p_t | p_s)
                    score = (inter / union) * 85
                else:
                    score = 0

            if score > melhor_score and score >= 45:
                melhor_score = score
                melhor_sota = orig_p

        if melhor_sota:
            mapa[str(titulo)] = melhor_sota

    return mapa


def formatar_planilha_cruzada(caminho_xlsx: Path):
    """Aplica design executivo na planilha Excel gerada."""
    try:
        wb = openpyxl.load_workbook(caminho_xlsx)

        fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        fill_resumo = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        font_data = Font(name="Segoe UI", size=9)
        border_light = Border(
            left=Side(style="thin", color="E0E0E0"),
            right=Side(style="thin", color="E0E0E0"),
            top=Side(style="thin", color="E0E0E0"),
            bottom=Side(style="thin", color="E0E0E0"),
        )

        for ws in wb.worksheets:
            ws.views.sheetView[0].showGridLines = True
            ws.freeze_panes = "B2" if ws.title.startswith("Consolidado") else "A2"

            header_fill = fill_resumo if "Resumo" in ws.title else fill_header

            # Cabeçalho
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = font_header
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 32

            # Células de dados
            for row_idx in range(2, ws.max_row + 1):
                ws.row_dimensions[row_idx].height = 20
                for col_idx in range(1, ws.max_column + 1):
                    c = ws.cell(row=row_idx, column=col_idx)
                    c.font = font_data
                    c.border = border_light
                    c.alignment = Alignment(vertical="center")

                    col_name = str(ws.cell(1, col_idx).value or "").lower()
                    if any(term in col_name for term in ["score", "acurácia", "iou", "f1", "rank", "ano"]):
                        if isinstance(c.value, (int, float)):
                            c.alignment = Alignment(horizontal="center", vertical="center")
                    elif "classificação" in col_name or "status" in col_name:
                        c.alignment = Alignment(horizontal="center", vertical="center")

            # Largura inteligente das colunas
            for col in ws.columns:
                col_letter = get_column_letter(col[0].column)
                header_val = str(col[0].value or "")
                max_len = max(len(str(c.value or "")) for c in col[:50])
                largura = max(max_len + 3, len(header_val) + 4, 11)
                if largura > 65:
                    largura = 65
                ws.column_dimensions[col_letter].width = largura

        wb.save(caminho_xlsx)
    except Exception as e:
        print(f"[Aviso] Não foi possível formatar esteticamente a planilha: {e}")


def executar_cruzamento(
    arquivo_classificacao: str = "artigos_media_alta_alta_excelente.xlsx",
    arquivo_sota_final: str = "state_of_the_art_final.csv",
    arquivo_sota_fontes: str = "state_of_the_art_v2_fontes.csv",
    arquivo_saida_xlsx: str = "base_consolidada_classificacao_e_sota.xlsx",
    arquivo_saida_csv: str = "base_consolidada_classificacao_e_sota.csv",
):
    path_class = Path(arquivo_classificacao)
    path_sota = Path(arquivo_sota_final)
    path_fontes = Path(arquivo_sota_fontes)
    path_out_xlsx = Path(arquivo_saida_xlsx)
    path_out_csv = Path(arquivo_saida_csv)

    if not path_class.exists():
        raise FileNotFoundError(f"Arquivo de classificação não encontrado: {path_class.resolve()}")
    if not path_sota.exists():
        raise FileNotFoundError(f"Arquivo SOTA final não encontrado: {path_sota.resolve()}")

    print("=" * 75)
    print(" CRUZAMENTO: CLASSIFICAÇÃO DE ARTIGOS x ESTADO DA ARTE (SOTA)")
    print("=" * 75)
    print(f"-> Base de Classificação: {path_class.resolve()}")
    print(f"-> Base SOTA Final       : {path_sota.resolve()}")
    print(f"-> Base SOTA Fontes      : {path_fontes.resolve() if path_fontes.exists() else 'Não fornecida'}")

    # 1. Carrega a base de classificação
    df_class = pd.read_excel(path_class, sheet_name=0)
    col_titulo_class = None
    for c in df_class.columns:
        if "titulo" in normalizar_texto(c) or "title" in normalizar_texto(c):
            col_titulo_class = c
            break

    if not col_titulo_class:
        raise ValueError(f"Coluna de título não encontrada em {path_class.name}")

    total_artigos_class = len(df_class)
    print(f"-> Total de artigos na classificação: {total_artigos_class}")

    # 2. Carrega a base do Estado da Arte Final
    df_sota = pd.read_csv(path_sota, encoding="utf-8-sig")
    if "paper_id" not in df_sota.columns:
        raise ValueError("Coluna 'paper_id' não encontrada no arquivo SOTA final.")
    print(f"-> Total de registros no SOTA Final: {len(df_sota)}")

    # 3. Constrói o mapeamento difuso
    titulos_class = df_class[col_titulo_class].dropna().tolist()
    papers_sota = df_sota["paper_id"].dropna().tolist()
    mapa_titulos_sota = construir_mapa_cruzamento(titulos_class, papers_sota)
    total_cruzados = len(mapa_titulos_sota)

    pct_cruzados = (total_cruzados / total_artigos_class) * 100
    print("\n" + "-" * 75)
    print(f" RESULTADO DO PAREAMENTO: {total_cruzados} de {total_artigos_class} artigos cruzados ({pct_cruzados:.1f}%)")
    print("-" * 75)

    # Cria coluna temporária de junção
    df_class["_paper_id_sota"] = df_class[col_titulo_class].map(mapa_titulos_sota)

    # Remove duplicatas em df_sota mantendo a primeira ocorrência (ex: sem sufixo (1))
    df_sota_dedup = df_sota.drop_duplicates(subset=["paper_id"]).copy()

    # Renomeia colunas do SOTA para nomes amigáveis em português
    mapa_colunas_sota = {
        "modelos_dl": "SOTA: Modelos DL",
        "arquitetura_base": "SOTA: Arquitetura Base",
        "usa_transfer_learning": "SOTA: Transfer Learning",
        "fonte_imagem": "SOTA: Fonte Imagem",
        "usa_aster": "SOTA: Usa ASTER",
        "hiperespectral": "SOTA: Hiperespectral",
        "n_bandas": "SOTA: Nº Bandas",
        "tamanho_dataset": "SOTA: Tamanho Dataset",
        "split_dataset": "SOTA: Split Dataset",
        "indices_espectrais": "SOTA: Índices Espectrais",
        "rotulagem_assistida": "SOTA: Rotulagem Assistida",
        "data_augmentation": "SOTA: Data Augmentation",
        "metricas": "SOTA: Métricas",
        "melhor_acuracia": "SOTA: Melhor Acurácia",
        "melhor_iou": "SOTA: Melhor IoU",
        "melhor_f1": "SOTA: Melhor F1",
        "objeto_prospeccao": "SOTA: Objeto Prospecção",
        "prospecta_fosfato": "SOTA: Prospecta Fosfato",
        "desafios_limitacoes": "SOTA: Desafios e Limitações",
        "paises_area_estudo": "SOTA: País Área Estudo",
        "paises_afiliacao": "SOTA: País Afiliação",
        "loss_functions": "SOTA: Loss Functions",
        "otimizador": "SOTA: Otimizador",
        "learning_rate_scheduler": "SOTA: LR Scheduler",
        "gpu_vram": "SOTA: GPU/VRAM",
        "ram_sistema": "SOTA: RAM Sistema",
    }
    colunas_sota_uteis = [c for c in mapa_colunas_sota.keys() if c in df_sota_dedup.columns]
    df_sota_subset = df_sota_dedup[["paper_id"] + colunas_sota_uteis].copy()
    df_sota_subset = df_sota_subset.rename(columns=mapa_colunas_sota)

    # Realiza o Merge (Left Join mantendo todos os artigos da classificação)
    df_consolidado = pd.merge(
        df_class,
        df_sota_subset,
        left_on="_paper_id_sota",
        right_on="paper_id",
        how="left",
    )

    # Status de Cruzamento
    df_consolidado["Status SOTA"] = df_consolidado["paper_id"].apply(
        lambda x: "Cruzado com Sucesso" if pd.notna(x) else "Não Encontrado no SOTA"
    )

    # Organização das colunas: coloca Status SOTA logo após a Classificação
    df_consolidado = df_consolidado.drop(columns=["_paper_id_sota", "paper_id"])
    col_class_name = [c for c in df_consolidado.columns if "classific" in normalizar_texto(c)][0]
    cols = list(df_consolidado.columns)
    cols.remove("Status SOTA")
    idx_class = cols.index(col_class_name) + 1
    cols.insert(idx_class, "Status SOTA")
    df_consolidado = df_consolidado[cols]

    # 4. Gera Resumo Estatístico por Categoria de Classificação
    resumo_linhas = []
    grupos = df_consolidado.groupby(col_class_name)
    for cat, grp in grupos:
        tot = len(grp)
        cruz = len(grp[grp["Status SOTA"] == "Cruzado com Sucesso"])
        com_dl = len(grp[grp["SOTA: Modelos DL"].notna() & ~grp["SOTA: Modelos DL"].isin(["Não mencionado", "No mencionado"])])
        com_hiper = len(grp[grp["SOTA: Hiperespectral"].notna() & ~grp["SOTA: Hiperespectral"].isin(["Não mencionado", "No mencionado"])])
        com_aster = len(grp[grp["SOTA: Usa ASTER"].notna() & ~grp["SOTA: Usa ASTER"].isin(["Não mencionado", "No mencionado"])])
        resumo_linhas.append({
            "Classificação": cat,
            "Total de Artigos": tot,
            "Artigos Cruzados SOTA": cruz,
            "% Cruzamento": f"{cruz/tot*100:.1f}%",
            "Com Modelo DL Identificado": com_dl,
            "Usa Dados Hiperespectrais": com_hiper,
            "Usa Dados ASTER": com_aster,
        })
    df_resumo = pd.DataFrame(resumo_linhas)

    print("\n Resumo por Categoria de Relevância:")
    for _, r in df_resumo.iterrows():
        print(f"   • {r['Classificação']:<18}: {r['Artigos Cruzados SOTA']}/{r['Total de Artigos']} cruzados ({r['% Cruzamento']}) | DL: {r['Com Modelo DL Identificado']} | Hiperespectral: {r['Usa Dados Hiperespectrais']}")

    # 5. Prepara aba de Fontes/Rastreabilidade se existir
    df_fontes_cruzado = None
    if path_fontes.exists():
        df_fontes_raw = pd.read_csv(path_fontes, encoding="utf-8-sig")
        df_fontes_dedup = df_fontes_raw.drop_duplicates(subset=["paper_id"]).copy()
        df_class_temp = df_class[[col_titulo_class, "_paper_id_sota"]].copy()
        df_fontes_cruzado = pd.merge(
            df_class_temp,
            df_fontes_dedup,
            left_on="_paper_id_sota",
            right_on="paper_id",
            how="left",
        ).drop(columns=["_paper_id_sota"])

    # 6. Salva em Excel com múltiplas abas formatadas
    print("\n" + "-" * 75)
    print(f"-> Gerando planilha Excel consolidada: {path_out_xlsx.resolve()}")
    with pd.ExcelWriter(path_out_xlsx, engine="openpyxl") as writer:
        df_consolidado.to_excel(writer, sheet_name="Classificação e SOTA", index=False)
        df_resumo.to_excel(writer, sheet_name="Resumo por Classificação", index=False)
        if df_fontes_cruzado is not None:
            df_fontes_cruzado.to_excel(writer, sheet_name="Rastreabilidade Fontes", index=False)

    formatar_planilha_cruzada(path_out_xlsx)
    print("   ✓ Arquivo Excel gerado com formatação profissional e abas dinâmicas.")

    # 7. Salva em CSV com separador ';' e UTF-8 com BOM
    print(f"-> Gerando arquivo CSV consolidado   : {path_out_csv.resolve()}")
    df_consolidado.to_csv(path_out_csv, index=False, sep=";", encoding="utf-8-sig")
    print("   ✓ Arquivo CSV exportado com sucesso (UTF-8 com BOM).")

    print("=" * 75)
    print(" CRUZAMENTO CONCLUÍDO COM SUCESSO!")
    print("=" * 75)
    return df_consolidado


def main():
    parser = argparse.ArgumentParser(
        description="Cruza a planilha de artigos classificados com os dados consolidados do Estado da Arte (SOTA)."
    )
    parser.add_argument(
        "--classificacao", "-c",
        default="artigos_media_alta_alta_excelente.xlsx",
        help="Planilha de artigos classificados (padrão: artigos_media_alta_alta_excelente.xlsx)",
    )
    parser.add_argument(
        "--sota-final", "-s",
        default="state_of_the_art_final.csv",
        help="Arquivo CSV do SOTA final (padrão: state_of_the_art_final.csv)",
    )
    parser.add_argument(
        "--sota-fontes", "-f",
        default="state_of_the_art_v2_fontes.csv",
        help="Arquivo CSV do SOTA com fontes/páginas (padrão: state_of_the_art_v2_fontes.csv)",
    )
    parser.add_argument(
        "--output-xlsx", "-o",
        default="base_consolidada_classificacao_e_sota.xlsx",
        help="Arquivo Excel de saída (padrão: base_consolidada_classificacao_e_sota.xlsx)",
    )
    parser.add_argument(
        "--output-csv",
        default="base_consolidada_classificacao_e_sota.csv",
        help="Arquivo CSV de saída (padrão: base_consolidada_classificacao_e_sota.csv)",
    )

    args = parser.parse_args()

    executar_cruzamento(
        arquivo_classificacao=args.classificacao,
        arquivo_sota_final=args.sota_final,
        arquivo_sota_fontes=args.sota_fontes,
        arquivo_saida_xlsx=args.output_xlsx,
        arquivo_saida_csv=args.output_csv,
    )


if __name__ == "__main__":
    main()
