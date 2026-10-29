"""
Módulo de download automatizado de extratos do SGF (Silvicultura / Fechamento).
Executa o fluxo via Playwright em modo headless (segundo plano).
"""

import sys
import calendar
import os
import zipfile
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# Carrega variáveis do arquivo .env
load_dotenv(Path(__file__).parent / ".env")

URL_LOGIN = "https://sgfp.celulosers.com.br/sgf/Default.aspx?ReturnUrl=%2fsgf%2fDefaultModulos.aspx"
URL_RELATORIO = "https://sgfp.celulosers.com.br/sgf/Modulos/Relatorios/Silvicultura/VwSilConsultaCompletaBoletimFrm.aspx?cdprocesso=2&cdfuncao=3730"

USUARIO_PADRAO = os.getenv("SGF_USER", "fapereira")
SENHA_PADRAO = os.getenv("SGF_PASSWORD", "Celulose2026")


def calcular_periodo_mes(mes: int, ano: int):
    """Retorna as datas no formato DD/MM/AAAA para o 1º e último dia do mês."""
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    data_ini = f"01/{mes:02d}/{ano}"
    data_fim = f"{ultimo_dia:02d}/{mes:02d}/{ano}"
    return data_ini, data_fim


def baixar_extrato_sgf(data_ini: str, data_fim: str, pasta_destino: Path, status_callback=None):
    """
    Realiza o login no SGF, preenche o período e baixa o extrato Excel.
    Extrai o .xls do arquivo .zip baixado para dentro de pasta_destino.
    """
    def log(msg):
        print(msg)
        if status_callback:
            status_callback(msg)

    pasta_destino.mkdir(parents=True, exist_ok=True)

    log(f"➜ Iniciando conexão com o SGF para o período {data_ini} a {data_fim}...")

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # 1. Login
        log("➜ Efetuando login no SGF...")
        page.goto(URL_LOGIN, wait_until="networkidle")
        page.fill("#txtLogin", USUARIO_PADRAO)
        page.fill("#txtSenha", SENHA_PADRAO)
        page.click("#btnOk")
        page.wait_for_load_state("networkidle")

        # 2. Tela de Consulta Completa
        log("➜ Acessando tela de Consulta Completa de Boletins...")
        page.goto(URL_RELATORIO, wait_until="networkidle")

        # 3. Preenchimento de Datas
        log(f"➜ Aplicando filtros de data ({data_ini} a {data_fim})...")
        page.fill("#txtDataIni", data_ini)
        page.fill("#txtDataFim", data_fim)
        page.click("#btnGerar")
        page.wait_for_load_state("networkidle")

        # 4. Exportação
        log("➜ Solicitando exportação para o Excel (isso pode levar alguns segundos)...")
        with page.expect_download(timeout=120000) as download_info:
            page.click("input[title='Exportar dados para Excel']")

        download = download_info.value
        nome_arquivo = download.suggested_filename
        caminho_baixado = pasta_destino / nome_arquivo

        download.save_as(caminho_baixado)
        browser.close()

    log(f"✓ Download concluído: {nome_arquivo}")

    # Se o arquivo for um .zip gerado pelo SGF, descompacta o .xls
    arquivo_final = caminho_baixado
    if caminho_baixado.suffix.lower() == ".zip":
        log("➜ Descompactando arquivo .zip do SGF...")
        with zipfile.ZipFile(caminho_baixado, "r") as z:
            nomes = z.namelist()
            z.extractall(pasta_destino)
            if nomes:
                arquivo_final = pasta_destino / nomes[0]
        # Remove o zip temporário
        try:
            caminho_baixado.unlink()
        except Exception:
            pass

    log(f"✅ Extrato pronto na pasta: {arquivo_final.name}")
    return arquivo_final


if __name__ == "__main__":
    import sys
    destino = Path(__file__).parent / "exports_sgf"
    d_ini, d_fim = calcular_periodo_mes(9, 2026)
    baixar_extrato_sgf(d_ini, d_fim, destino)
