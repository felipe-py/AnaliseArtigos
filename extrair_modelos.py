#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para extrair automaticamente o modelo de Machine Learning / Deep Learning
utilizado em cada artigo científico em PDF, utilizando a API do Google Gemini.

Recursos:
- Upload seguro via cópia temporária com nome ASCII (elimina erros de encoding no Windows)
- Suporte a caminhos longos no Windows (> 260 caracteres com prefixo \\\\?\\)
- Tratamento adaptativo de Rate Limits (429) e sobrecargas (503) com backoff de janela móvel (65s+)
- Sistema de logs completo em arquivo (extracao_modelos.log) e console (UTF-8)
- Checkpoint contínuo em Excel e reprocessamento inteligente de falhas anteriores
- Extração estruturada (Pydantic / Structured Outputs)
"""

import os
import sys
import time
import re
import unicodedata
import argparse
import logging
import tempfile
import shutil
from pathlib import Path
from typing import Optional, List, Dict, Any

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field

# Configuração de encoding no Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def configurar_logger(arquivo_log: str = "extracao_modelos.log") -> logging.Logger:
    """Configura logger para arquivo (UTF-8 detalhado) e terminal."""
    logger = logging.getLogger("ExtracaoModelos")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    # Handler de Arquivo com UTF-8
    file_handler = logging.FileHandler(arquivo_log, mode="a", encoding="utf-8")
    file_formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)

    # Handler de Console
    console_handler = logging.StreamHandler(sys.stdout)
    console_formatter = logging.Formatter("%(message)s")
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)

    return logger


logger = configurar_logger()


class ExtracaoModeloArtigo(BaseModel):
    modelo_principal: str = Field(
        description="Nome específico do modelo ou arquitetura de IA/ML/DL principal proposto ou utilizado (ex: 'U-Net com Mecanismo de Atenção', '1D-CNN com Covariance Pooling', 'VAE-BIRCH Híbrido', 'Random Forest')."
    )
    categoria_modelo: str = Field(
        description="Categoria: 'Deep Learning', 'Machine Learning Clássico', 'Híbrido (DL + ML)', 'Estatístico / Processamento Digital' ou 'Revisão / Teórico'."
    )
    modelos_comparados: List[str] = Field(
        default_factory=list,
        description="Lista de modelos utilizados como baselines ou comparações no estudo (ex: ['SVM', 'Random Forest', 'MLP']). Vazio se nenhum."
    )
    sensores_dados: str = Field(
        description="Fontes de dados ou sensores alimentados ao modelo (ex: 'PRISMA Hiperespectral', 'ASTER + Geoquímica', 'Sentinel-2', 'Dados Aerogeofísicos')."
    )
    alvo_ou_aplicacao: str = Field(
        description="Objeto do estudo (ex: 'Cobre Porfirítico', 'Ouro Orogênico', 'Mapeamento Litológico', 'Alteração Hidrotermal')."
    )
    trecho_evidencia: str = Field(
        description="Frase curta extraída da metodologia ou resumo comprovando a escolha/uso do modelo principal."
    )


PROMPT_SISTEMA = """
Você é um cientista pesquisador sênior especialista em Sensoriamento Remoto, Inteligência Artificial e Geociências.
Sua missão é analisar criteriosamente o artigo acadêmico em PDF fornecido e identificar exatamente qual MODELO COMPUTACIONAL (Machine Learning, Deep Learning ou método estatístico) os autores implementaram ou propuseram como metodologia principal da pesquisa.

DIRETRIZES FUNDAMENTAIS:
1. DISTINGA O MODELO PRINCIPAL DOS BASELINES: Muitas vezes os autores comparam o modelo proposto com outros (ex: comparam uma CNN com SVM e Random Forest). Identifique a contribuição central como modelo_principal e os demais em modelos_comparados.
2. SEJA ESPECÍFICO NA ARQUITETURA: Não responda apenas 'Deep Learning' ou 'CNN'. Se o artigo usa '3D-CNN com Attention', 'ResNet-18', 'ViT', 'U-Net', 'CatBoost + GMM', informe a variante exata.
3. SE FOR ARTIGO DE REVISÃO (Review): Indique 'Artigo de Revisão (Sem modelo proprietário)' e resuma as principais abordagens discutidas.
4. EVIDÊNCIA: Forneça um trecho literal conciso da metodologia ou resumo que ateste a identificação.
"""


def normalizar_texto(texto: str) -> str:
    """Remove pontuações, acentos e separa palavras coladas em CamelCase."""
    if not isinstance(texto, str):
        return ""
    texto = re.sub(r"([a-z])([A-Z])", r"\1 \2", texto)
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("utf-8").lower()
    texto = re.sub(r"[^a-z0-9]", " ", texto)
    return " ".join(texto.split())


def normalizar_caminho_windows(caminho: Path) -> str:
    """Garante suporte a caminhos longos (> 260 caracteres) no Windows usando o prefixo \\\\?\\."""
    caminho_abs = str(caminho.resolve())
    if sys.platform == "win32" and not caminho_abs.startswith("\\\\?\\"):
        return f"\\\\?\\{caminho_abs}"
    return caminho_abs


def mapear_pdfs_para_titulos(pasta_pdfs: Path, titulos_excel: List[str]) -> Dict[str, Path]:
    """
    Cria um mapa de correspondência entre o título da planilha e o arquivo PDF real na pasta.
    Suporta pequenas variações de caracteres especiais, sublinhados, truncamentos e CamelCase.
    """
    arquivos_pdf = list(pasta_pdfs.glob("*.pdf"))
    mapa_resultado = {}

    pdfs_norm = []
    for pdf_path in arquivos_pdf:
        stem_limpo = re.sub(r"\(\d+\)$", "", pdf_path.stem).strip()
        stem_norm = normalizar_texto(stem_limpo)
        pdfs_norm.append((stem_norm, pdf_path))

    for titulo in titulos_excel:
        if not titulo or pd.isna(titulo):
            continue
        titulo_norm = normalizar_texto(str(titulo))
        if not titulo_norm:
            continue

        melhor_pdf = None
        melhor_score = 0

        for stem_norm, pdf_path in pdfs_norm:
            if titulo_norm == stem_norm:
                melhor_pdf = pdf_path
                melhor_score = 100
                break

            if len(stem_norm) >= 20 and titulo_norm.startswith(stem_norm):
                score = 98
            elif len(titulo_norm) >= 20 and stem_norm.startswith(titulo_norm):
                score = 98
            elif stem_norm in titulo_norm or titulo_norm in stem_norm:
                score = min(len(stem_norm), len(titulo_norm)) / max(len(stem_norm), len(titulo_norm)) * 95
            else:
                palavras_titulo = set(titulo_norm.split())
                palavras_stem = set(stem_norm.split())
                if palavras_titulo and palavras_stem:
                    intersecao = len(palavras_titulo & palavras_stem)
                    uniao = len(palavras_titulo | palavras_stem)
                    score = (intersecao / uniao) * 85
                else:
                    score = 0

            if score > melhor_score and score >= 45:
                melhor_score = score
                melhor_pdf = pdf_path

        if melhor_pdf:
            mapa_resultado[str(titulo)] = melhor_pdf

    return mapa_resultado


def extrair_informacoes_pdf(client, modelo_gemini: str, caminho_pdf: Path) -> ExtracaoModeloArtigo:
    """
    Envia o PDF para a API do Gemini de forma segura:
    1. Cria uma cópia temporária com nome ASCII puro (resolve erros de encod 'ascii' no Windows)
    2. Suporta caminhos longos (> 260 chars)
    3. Faz fallback transparente entre versões estáveis do Gemini
    4. Garante deleção do arquivo temporário local e do arquivo remoto
    """
    from google.genai import types

    caminho_seguro_origem = normalizar_caminho_windows(caminho_pdf)

    # Cria arquivo temporário com nome ASCII puro (ex: tmp_doc_12345.pdf)
    # Isso evita que o httpx tente codificar caracteres não-ASCII em headers HTTP
    with tempfile.NamedTemporaryFile(suffix=".pdf", prefix="gemini_doc_", delete=False) as tmp_file:
        caminho_tmp = Path(tmp_file.name)

    try:
        shutil.copy2(caminho_seguro_origem, caminho_tmp)
        caminho_upload = normalizar_caminho_windows(caminho_tmp)

        arquivo_remoto = client.files.upload(file=caminho_upload)
        try:
            modelos_candidatos = [
                modelo_gemini,
                "gemini-flash-latest",
                "gemini-3.8-flash",
                "gemini-3.5-flash",
            ]
            modelos_unicos = []
            for m in modelos_candidatos:
                if m and m not in modelos_unicos:
                    modelos_unicos.append(m)

            ultimo_erro = None
            for mod in modelos_unicos:
                try:
                    resposta = client.models.generate_content(
                        model=mod,
                        contents=[
                            arquivo_remoto,
                            "Extraia as informações do modelo computacional e metodologia deste artigo conforme o schema fornecido.",
                        ],
                        config=types.GenerateContentConfig(
                            system_instruction=PROMPT_SISTEMA,
                            response_mime_type="application/json",
                            response_schema=ExtracaoModeloArtigo,
                            temperature=0.1,
                        ),
                    )
                    return resposta.parsed
                except Exception as e:
                    err_text = str(e)
                    ultimo_erro = e
                    if "404" in err_text or "NOT_FOUND" in err_text or "503" in err_text or "UNAVAILABLE" in err_text:
                        logger.warning(f"   [Aviso Modelo] {mod} indisponível temporariamente ({err_text[:80]}). Tentando fallback...")
                        continue
                    raise e

            if ultimo_erro:
                raise ultimo_erro
        finally:
            try:
                client.files.delete(name=arquivo_remoto.name)
            except Exception:
                pass
    finally:
        # Garante a remoção da cópia temporária do disco
        if caminho_tmp.exists():
            try:
                caminho_tmp.unlink()
            except Exception:
                pass


def formatar_excel(caminho_xlsx: Path):
    """Aplica estilo visual executivo no arquivo final."""
    try:
        wb = openpyxl.load_workbook(caminho_xlsx)
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        regular_font = Font(name="Segoe UI", size=10)
        border_thin = Border(
            left=Side(style="thin", color="E0E0E0"),
            right=Side(style="thin", color="E0E0E0"),
            top=Side(style="thin", color="E0E0E0"),
            bottom=Side(style="thin", color="E0E0E0"),
        )

        for ws in wb.worksheets:
            ws.views.sheetView[0].showGridLines = True
            ws.freeze_panes = "A2"

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 30

            for row_idx in range(2, ws.max_row + 1):
                ws.row_dimensions[row_idx].height = 24
                for col_idx in range(1, ws.max_column + 1):
                    c = ws.cell(row=row_idx, column=col_idx)
                    c.font = regular_font
                    c.border = border_thin
                    c.alignment = Alignment(vertical="center")

            for col in ws.columns:
                col_letter = get_column_letter(col[0].column)
                header_val = str(col[0].value or "")
                max_len = max(len(str(c.value or "")) for c in col[:50])
                largura = max(max_len + 3, len(header_val) + 4, 12)
                if largura > 65:
                    largura = 65
                ws.column_dimensions[col_letter].width = largura

        wb.save(caminho_xlsx)
    except Exception as e:
        logger.warning(f"[Aviso] Não foi possível formatar esteticamente a planilha: {e}")


def executar_extracao(
    api_key: str,
    pasta_pdfs: str = "artigos",
    planilha_entrada: str = "artigos_media_alta_alta_excelente.xlsx",
    planilha_saida: str = "artigos_com_modelos_extraidos.xlsx",
    modelo_gemini: str = "gemini-flash-latest",
    delay_segundos: float = 6.0,
    limite: Optional[int] = None,
    reprocessar_erros: bool = True,
):
    from google import genai

    path_pdfs = Path(pasta_pdfs)
    path_entrada = Path(planilha_entrada)
    path_saida = Path(planilha_saida)

    if not path_pdfs.exists():
        raise FileNotFoundError(f"Pasta de PDFs não encontrada: {path_pdfs.resolve()}")
    if not path_entrada.exists():
        raise FileNotFoundError(f"Planilha de entrada não encontrada: {path_entrada.resolve()}")

    logger.info("=" * 75)
    logger.info(" EXTRAÇÃO DE MODELOS CIENTÍFICOS COM GEMINI API")
    logger.info("=" * 75)
    logger.info(f"-> Pasta com PDFs   : {path_pdfs.resolve()}")
    logger.info(f"-> Base de Artigos   : {path_entrada.resolve()}")
    logger.info(f"-> Planilha Saída    : {path_saida.resolve()}")
    logger.info(f"-> Modelo Gemini     : {modelo_gemini}")
    logger.info(f"-> Delay entre req.  : {delay_segundos:.1f}s")
    logger.info(f"-> Reprocessar erros : {reprocessar_erros}")
    logger.info(f"-> Arquivo de log    : extracao_modelos.log")

    client = genai.Client(api_key=api_key)

    # Carrega a planilha de saída existente para checkpoint
    if path_saida.exists():
        logger.info(f"\n[Checkpoint] Planilha existente encontrada. Lendo progresso...")
        df = pd.read_excel(path_saida)
    else:
        df = pd.read_excel(path_entrada, sheet_name=0)

    # Identifica coluna de título
    col_titulo = None
    for c in df.columns:
        norm = normalizar_texto(c)
        if "titulo" in norm or "title" in norm:
            col_titulo = c
            break

    if not col_titulo:
        raise ValueError(f"Coluna de título não encontrada na planilha. Colunas: {list(df.columns)}")

    novas_colunas = [
        "Modelo Principal",
        "Categoria Modelo",
        "Modelos Comparados",
        "Sensores e Dados",
        "Alvo / Aplicação",
        "Trecho Evidência",
        "Arquivo PDF",
    ]
    for col in novas_colunas:
        if col not in df.columns:
            df[col] = None

    titulos = df[col_titulo].dropna().tolist()
    mapa_pdfs = mapear_pdfs_para_titulos(path_pdfs, titulos)
    logger.info(f"-> PDFs localizados e pareados: {len(mapa_pdfs)} de {len(titulos)} artigos")

    linhas_processadas = 0
    linhas_totais = len(df) if limite is None else min(len(df), limite)

    for idx, row in df.iterrows():
        if limite is not None and linhas_processadas >= limite:
            logger.info(f"\n[Limite atingido] Finalizando após {limite} artigos solicitados.")
            break

        titulo = row[col_titulo]
        val_modelo = str(row.get("Modelo Principal") or "").strip()

        # Verifica se já está preenchido com sucesso
        ja_concluido = pd.notna(row.get("Modelo Principal")) and val_modelo != "" and not val_modelo.startswith("[ERRO")
        if ja_concluido:
            continue

        # Se já tem erro e o usuário NÃO quer reprocessar erros, pula
        if val_modelo.startswith("[ERRO") and not reprocessar_erros:
            continue

        caminho_pdf = mapa_pdfs.get(str(titulo))
        if not caminho_pdf:
            logger.warning(f"\n[Aviso] PDF não encontrado na pasta para o artigo [{idx+1}/{linhas_totais}]: '{str(titulo)[:50]}...'")
            continue

        logger.info(f"\n[{idx+1}/{linhas_totais}] Processando: {caminho_pdf.name[:65]}...")

        max_tentativas = 5
        sucesso = False

        for tentativa in range(1, max_tentativas + 1):
            try:
                dados_extraidos = extrair_informacoes_pdf(client, modelo_gemini, caminho_pdf)

                df.at[idx, "Modelo Principal"] = dados_extraidos.modelo_principal
                df.at[idx, "Categoria Modelo"] = dados_extraidos.categoria_modelo
                df.at[idx, "Modelos Comparados"] = (
                    ", ".join(dados_extraidos.modelos_comparados) if dados_extraidos.modelos_comparados else "Nenhum"
                )
                df.at[idx, "Sensores e Dados"] = dados_extraidos.sensores_dados
                df.at[idx, "Alvo / Aplicação"] = dados_extraidos.alvo_ou_aplicacao
                df.at[idx, "Trecho Evidência"] = dados_extraidos.trecho_evidencia
                df.at[idx, "Arquivo PDF"] = caminho_pdf.name

                logger.info(f"   ✓ Modelo: {dados_extraidos.modelo_principal} ({dados_extraidos.categoria_modelo})")
                logger.info(f"   ✓ Dados : {dados_extraidos.sensores_dados}")
                sucesso = True
                linhas_processadas += 1
                break

            except Exception as e:
                err_str = str(e)
                # 429 = Rate Limit / Quota Exceeded (TPM ou RPM)
                if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                    # Janela móvel do Gemini é de 60 segundos. Espera 65s para resetar 100% da cota por minuto!
                    espera = 65 + (tentativa - 1) * 30
                    logger.warning(
                        f"   [Rate Limit 429] Limite por minuto atingido. Aguardando reset da janela ({espera}s) (tentativa {tentativa}/{max_tentativas})..."
                    )
                    time.sleep(espera)
                elif "503" in err_str or "UNAVAILABLE" in err_str:
                    espera = 25 * tentativa
                    logger.warning(
                        f"   [Serviço Ocupado 503] Servidor Google sob alta demanda. Aguardando {espera}s (tentativa {tentativa}/{max_tentativas})..."
                    )
                    time.sleep(espera)
                else:
                    logger.error(f"   [Erro] Falha ao extrair do PDF: {e}")
                    if tentativa < max_tentativas:
                        time.sleep(5)
                    else:
                        df.at[idx, "Modelo Principal"] = f"[ERRO NA EXTRAÇÃO]: {e}"
                        df.at[idx, "Arquivo PDF"] = caminho_pdf.name

        # Salva o checkpoint continuamente
        df.to_excel(path_saida, index=False)

        # Intervalo respeitando o rate limit da API
        if sucesso:
            time.sleep(delay_segundos)

    # Formatação final
    formatar_excel(path_saida)
    logger.info("\n" + "=" * 75)
    logger.info(" PROCESSO CONCLUÍDO COM SUCESSO!")
    logger.info(f"-> Arquivo salvo e formatado em: {path_saida.resolve()}")
    logger.info(f"-> Log de operações salvo em   : Path('extracao_modelos.log').resolve()")
    logger.info("=" * 75)


def main():
    parser = argparse.ArgumentParser(
        description="Extrai os modelos de ML/DL de artigos científicos em PDF usando a API do Gemini."
    )
    parser.add_argument(
        "--api-key",
        default=os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY"),
        help="Chave da API do Google Gemini (se não informada, lê de GEMINI_API_KEY)",
    )
    parser.add_argument(
        "--pdf-dir", "-p",
        default="artigos",
        help="Diretório onde estão os PDFs (padrão: artigos)",
    )
    parser.add_argument(
        "--input", "-i",
        default="artigos_media_alta_alta_excelente.xlsx",
        help="Planilha de entrada com os artigos (padrão: artigos_media_alta_alta_excelente.xlsx)",
    )
    parser.add_argument(
        "--output", "-o",
        default="artigos_com_modelos_extraidos.xlsx",
        help="Planilha de saída com os dados extraídos (padrão: artigos_com_modelos_extraidos.xlsx)",
    )
    parser.add_argument(
        "--model", "-m",
        default="gemini-flash-latest",
        help="Modelo do Gemini a utilizar (padrão: gemini-flash-latest)",
    )
    parser.add_argument(
        "--delay", "-d",
        type=float,
        default=6.0,
        help="Segundos de pausa entre artigos para respeitar o limite gratuito (padrão: 6.0)",
    )
    parser.add_argument(
        "--limit", "-l",
        type=int,
        default=None,
        help="Número máximo de artigos a processar (útil para teste piloto)",
    )
    parser.add_argument(
        "--no-retry-errors",
        action="store_true",
        help="Não tentar reprocessar linhas que falharam com [ERRO] em execuções anteriores",
    )

    args = parser.parse_args()

    api_key = args.api_key
    if not api_key:
        print("\n" + "!" * 70)
        print(" AVISO: Chave da API do Gemini não detectada nas variáveis de ambiente.")
        print("!" * 70)
        try:
            api_key = input("Por favor, cole sua GEMINI_API_KEY aqui: ").strip()
        except EOFError:
            pass

    if not api_key:
        print("[Erro Fatal] É necessária uma GEMINI_API_KEY para executar a extração.")
        print("Obtenha gratuitamente em: https://aistudio.google.com/")
        sys.exit(1)

    executar_extracao(
        api_key=api_key,
        pasta_pdfs=args.pdf_dir,
        planilha_entrada=args.input,
        planilha_saida=args.output,
        modelo_gemini=args.model,
        delay_segundos=args.delay,
        limite=args.limit,
        reprocessar_erros=not args.no_retry_errors,
    )


if __name__ == "__main__":
    main()
