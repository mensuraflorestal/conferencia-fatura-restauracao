import sys
from datetime import datetime
from pathlib import Path
import re
import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from download_sgf import baixar_extrato_sgf, calcular_periodo_mes

PASTA_EXPORTS = Path(__file__).parent / "exports_sgf"

MESES_NOME = [
    "01 - Janeiro",
    "02 - Fevereiro",
    "03 - Março",
    "04 - Abril",
    "05 - Maio",
    "06 - Junho",
    "07 - Julho",
    "08 - Agosto",
    "09 - Setembro",
    "10 - Outubro",
    "11 - Novembro",
    "12 - Dezembro",
]


# --- FUNÇÕES AUXILIARES ---
def converter_valor(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace("R$", "").replace(" ", "")
    if not s or s.lower() == "nan":
        return 0.0
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def extrair_mes_fatura(texto):
    texto = str(texto)
    m = re.search(r"(20\d{2})[-_](0[1-9]|1[0-2])", texto)
    if m:
        return f"{m.group(2)}/{m.group(1)}"
    m = re.search(r"(0[1-9]|1[0-2])[-_](20\d{2})", texto)
    if m:
        return f"{m.group(1)}/{m.group(2)}"
    m = re.search(r"(\d{2})[-_](0[1-9]|1[0-2])", texto)
    if m:
        return f"{m.group(2)}/20{m.group(1)}"
    m = re.search(r"(0[1-9]|1[0-2])[-_](\d{2})", texto)
    if m:
        return f"{m.group(1)}/20{m.group(2)}"
    return "Indefinido"


def adicionar_total(df, cols_texto):
    if df.empty:
        return df
    dict_total = {col: "" for col in cols_texto}
    dict_total[cols_texto[-1]] = "TOTAL"
    dict_total["Custo Total (R$)"] = df["Custo Total (R$)"].sum()
    dict_total["Qtd Produção/Horas"] = df["Qtd Produção/Horas"].sum()
    linha = pd.DataFrame([dict_total])
    return pd.concat([df, linha], ignore_index=True)


def descompactar_zips_pendentes(pasta: Path):
    """
    Localiza qualquer arquivo .zip na pasta, descompacta apenas as planilhas
    (.xls ou .xlsx) diretamente na pasta e remove o .zip após a extração.
    """
    import zipfile
    for zip_path in pasta.glob("*.zip"):
        try:
            print(f"📦 Descompactando '{zip_path.name}'...")
            with zipfile.ZipFile(zip_path, "r") as z:
                for member in z.namelist():
                    if member.startswith("__MACOSX") or member.endswith("/"):
                        continue
                    if member.lower().endswith((".xls", ".xlsx")):
                        nome_arquivo = Path(member).name
                        destino_arquivo = pasta / nome_arquivo
                        with z.open(member) as fonte, open(destino_arquivo, "wb") as destino:
                            destino.write(fonte.read())
                        print(f"   -> Planilha extraída: {nome_arquivo}")
            # Remove o .zip original após extrair com sucesso
            try:
                zip_path.unlink()
            except Exception:
                pass
        except Exception as e:
            print(f"⚠️ Erro ao descompactar '{zip_path.name}': {e}")



# --- INTERFACE DE SELEÇÃO DO MÊS / DOWNLOAD ---
def selecionar_mes_download():
    """Janela inicial para o usuário escolher o mês do extrato ou usar o arquivo local existente."""
    janela = tk.Tk()
    janela.title("SGF - Extração de Extrato")
    janela.geometry("460x300")
    janela.resizable(False, False)

    hoje = datetime.now()
    # Padrão: mês anterior se for início do mês, ou mês atual
    mes_padrao_idx = hoje.month - 1  # 0-indexed

    resultado = {"acao": None, "mes": hoje.month, "ano": hoje.year}

    tk.Label(
        janela,
        text="Conferência de Fechamento (SGF)",
        font=("Segoe UI", 12, "bold"),
        fg="#1b5e20",
        pady=10,
    ).pack()

    frame_corpo = tk.Frame(janela, padx=20, pady=10)
    frame_corpo.pack(fill="both", expand=True)

    tk.Label(
        frame_corpo,
        text="Selecione o mês de extração do SGF:",
        font=("Segoe UI", 10),
    ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

    tk.Label(frame_corpo, text="Mês:", font=("Segoe UI", 9, "bold")).grid(
        row=1, column=0, sticky="w", pady=5
    )
    cb_mes = ttk.Combobox(
        frame_corpo, values=MESES_NOME, state="readonly", width=18
    )
    cb_mes.current(mes_padrao_idx)
    cb_mes.grid(row=1, column=1, sticky="w", pady=5, padx=5)

    tk.Label(frame_corpo, text="Ano:", font=("Segoe UI", 9, "bold")).grid(
        row=2, column=0, sticky="w", pady=5
    )
    spin_ano = ttk.Spinbox(
        frame_corpo, from_=2020, to=2035, width=8, justify="center"
    )
    spin_ano.set(hoje.year)
    spin_ano.grid(row=2, column=1, sticky="w", pady=5, padx=5)

    lbl_info = tk.Label(
        frame_corpo,
        text="O robô fará o login e baixará do dia 1º até o último dia do mês.",
        font=("Segoe UI", 8),
        fg="#555",
    )
    lbl_info.grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 5))

    frame_botoes = tk.Frame(janela, padx=15, pady=15)
    frame_botoes.pack(fill="x", side="bottom")

    def acao_baixar():
        resultado["acao"] = "baixar"
        resultado["mes"] = cb_mes.current() + 1
        try:
            resultado["ano"] = int(spin_ano.get())
        except ValueError:
            resultado["ano"] = hoje.year
        janela.destroy()

    def acao_local():
        resultado["acao"] = "local"
        janela.destroy()

    btn_local = tk.Button(
        frame_botoes,
        text="📁 Usar Arquivo Já Baixado",
        command=acao_local,
        font=("Segoe UI", 9),
        padx=8,
        pady=4,
    )
    btn_local.pack(side="left")

    btn_baixar = tk.Button(
        frame_botoes,
        text="📥 Baixar do SGF e Avançar ➔",
        command=acao_baixar,
        bg="#2e7d32",
        fg="white",
        font=("Segoe UI", 9, "bold"),
        padx=10,
        pady=4,
    )
    btn_baixar.pack(side="right")

    janela.mainloop()
    return resultado


# --- INTERFACE GRÁFICA DE SELEÇÃO MANUAL DE OS ---
def selecionar_os_manual(lista_os):
    """Abre uma janela estilo filtro do Excel para seleção visual das OSs."""
    janela = tk.Tk()
    janela.title("Filtro de Ordens de Serviço (SGF)")
    janela.geometry("640x600")
    janela.minsize(520, 420)

    selecionados = []
    variaveis = {}

    # Cabeçalho
    lbl_instrucao = tk.Label(
        janela,
        text="Selecione as Ordens de Serviço que deseja processar:",
        font=("Segoe UI", 10, "bold"),
        pady=8,
    )
    lbl_instrucao.pack(fill="x")

    # Campo de Busca Rápida
    frame_busca = tk.Frame(janela, padx=10, pady=4)
    frame_busca.pack(fill="x")

    tk.Label(
        frame_busca, text="🔍 Buscar OS:", font=("Segoe UI", 9, "bold")
    ).pack(side="left", padx=4)
    ent_busca = tk.Entry(frame_busca, font=("Segoe UI", 9))
    ent_busca.pack(side="left", fill="x", expand=True, padx=4)

    # Frame com Scrollbar e Checkboxes
    frame_lista = tk.Frame(janela, padx=10, pady=5)
    frame_lista.pack(fill="both", expand=True)

    canvas = tk.Canvas(frame_lista, borderwidth=1, relief="solid")
    scrollbar = ttk.Scrollbar(
        frame_lista, orient="vertical", command=canvas.yview
    )
    scroll_frame = tk.Frame(canvas)

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
    )
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Cria os Checkboxes
    widgets_chk = {}
    for os_nome in sorted(lista_os):
        var = tk.BooleanVar(value=False)
        chk = tk.Checkbutton(
            scroll_frame,
            text=os_nome,
            variable=var,
            font=("Segoe UI", 9),
            anchor="w",
        )
        chk.pack(fill="x", padx=4, pady=2)
        variaveis[os_nome] = var
        widgets_chk[os_nome] = chk

    # Função de Filtro Dinâmico na Tela
    def filtrar_lista(*args):
        termo = ent_busca.get().lower()
        for os_nome, chk in widgets_chk.items():
            if termo in os_nome.lower():
                chk.pack(fill="x", padx=4, pady=2)
            else:
                chk.pack_forget()

    ent_busca.bind("<KeyRelease>", filtrar_lista)

    # Botões de Ação Rápida
    frame_acoes = tk.Frame(janela, padx=10, pady=6)
    frame_acoes.pack(fill="x")

    def marcar_todos():
        for os_nome, chk in widgets_chk.items():
            if ent_busca.get().lower() in os_nome.lower():
                variaveis[os_nome].set(True)

    def desmarcar_todos():
        for var in variaveis.values():
            var.set(False)

    def confirmar():
        nonlocal selecionados
        selecionados = [os for os, var in variaveis.items() if var.get()]
        if not selecionados:
            messagebox.showwarning(
                "Atenção", "Selecione ao menos uma Ordem de Serviço!"
            )
            return
        janela.destroy()

    btn_marcar = tk.Button(
        frame_acoes, text="Selecionar Visíveis", command=marcar_todos, width=16
    )
    btn_marcar.pack(side="left", padx=4)

    btn_desmarcar = tk.Button(
        frame_acoes, text="Limpar Seleção", command=desmarcar_todos, width=14
    )
    btn_desmarcar.pack(side="left", padx=4)

    btn_confirmar = tk.Button(
        frame_acoes,
        text="Processar Seleção ➔",
        command=confirmar,
        bg="#2e7d32",
        fg="white",
        font=("Segoe UI", 9, "bold"),
        padx=10,
    )
    btn_confirmar.pack(side="right", padx=4)

    janela.mainloop()
    return selecionados


# --- EXPORTAÇÃO EXCEL COM FORMATAÇÃO ---
def exportar_relatorio(df_filtrado, col_os, nome_arquivo):
    if df_filtrado.empty:
        return

    df = df_filtrado.copy()
    df["Mês Fatura"] = df[col_os].apply(extrair_mes_fatura)
    df["Custo_Num"] = df["Custo Total Operação"].apply(converter_valor)
    df["Prod_Num"] = df["Valor Produção"].apply(converter_valor)

    # REGRA 9004: Zera a quantidade de horas/produção se for OS de taxa 9004
    mask_9004 = df[col_os].astype(str).str.contains("9004", case=False, na=False)
    df.loc[mask_9004, "Prod_Num"] = 0.0

    colunas = list(df.columns)
    col_talhao = next(
        (c for c in colunas if "talh" in str(c).lower()), "Talhão"
    )
    col_municipio = next(
        (
            c
            for c in colunas
            if "munic" in str(c).lower() or "cidade" in str(c).lower()
        ),
        "Município",
    )

    for col in [
        col_os,
        "ID Projeto",
        "Nome Projeto",
        col_talhao,
        col_municipio,
    ]:
        if col not in df.columns:
            df[col] = "Não Informado"

    for col in [
        col_os,
        "Mês Fatura",
        "ID Projeto",
        "Nome Projeto",
        col_talhao,
        col_municipio,
    ]:
        df[col] = df[col].fillna("Não Informado")

    # 1. Por Horto
    df_horto = (
        df.groupby(
            [col_os, "Mês Fatura", "ID Projeto", "Nome Projeto"],
            dropna=False,
            as_index=False,
        )
        .agg({"Custo_Num": "sum", "Prod_Num": "sum"})
        .rename(
            columns={
                "Custo_Num": "Custo Total (R$)",
                "Prod_Num": "Qtd Produção/Horas",
                col_os: "Ordem de Serviço",
            }
        )
        .sort_values(
            by=["Mês Fatura", "Ordem de Serviço", "Custo Total (R$)"],
            ascending=[True, True, False],
        )
        .reset_index(drop=True)
    )
    df_horto_total = adicionar_total(
        df_horto,
        ["Ordem de Serviço", "Mês Fatura", "ID Projeto", "Nome Projeto"],
    )

    # 2. Por Horto + Talhão
    df_horto_talhao = (
        df.groupby(
            [
                col_os,
                "Mês Fatura",
                "ID Projeto",
                "Nome Projeto",
                col_talhao,
            ],
            dropna=False,
            as_index=False,
        )
        .agg({"Custo_Num": "sum", "Prod_Num": "sum"})
        .rename(
            columns={
                "Custo_Num": "Custo Total (R$)",
                "Prod_Num": "Qtd Produção/Horas",
                col_os: "Ordem de Serviço",
                col_talhao: "Talhão",
            }
        )
        .sort_values(
            by=["Mês Fatura", "Ordem de Serviço", "Custo Total (R$)"],
            ascending=[True, True, False],
        )
        .reset_index(drop=True)
    )
    df_horto_talhao_total = adicionar_total(
        df_horto_talhao,
        [
            "Ordem de Serviço",
            "Mês Fatura",
            "ID Projeto",
            "Nome Projeto",
            "Talhão",
        ],
    )

    # 3. Por Município
    df_municipio = (
        df.groupby(
            [col_os, "Mês Fatura", col_municipio],
            dropna=False,
            as_index=False,
        )
        .agg({"Custo_Num": "sum", "Prod_Num": "sum"})
        .rename(
            columns={
                "Custo_Num": "Custo Total (R$)",
                "Prod_Num": "Qtd Produção/Horas",
                col_os: "Ordem de Serviço",
                col_municipio: "Município",
            }
        )
        .sort_values(
            by=["Mês Fatura", "Ordem de Serviço", "Custo Total (R$)"],
            ascending=[True, True, False],
        )
        .reset_index(drop=True)
    )
    df_municipio_total = adicionar_total(
        df_municipio, ["Ordem de Serviço", "Mês Fatura", "Município"]
    )

    caminho = Path(__file__).parent / nome_arquivo

    try:
        with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
            df_horto_total.to_excel(writer, sheet_name="Por Horto", index=False)
            df_horto_talhao_total.to_excel(
                writer, sheet_name="Horto + Talhão", index=False
            )
            df_municipio_total.to_excel(
                writer, sheet_name="Por Município", index=False
            )

            # Autoajuste de largura de colunas
            for ws in writer.sheets.values():
                for col in ws.columns:
                    max_len = max(len(str(cell.value or "")) for cell in col)
                    col_letter = col[0].column_letter
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        print(f"💾 Relatório gerado com sucesso: {caminho.resolve()}")
    except PermissionError:
        print(f"\n❌ ERRO: O arquivo '{nome_arquivo}' está aberto no Excel!")
        print("Por favor, feche a planilha e tente novamente.")


# --- EXECUÇÃO PRINCIPAL ---
def main():
    # 1. Pergunta inicial: Baixar ou usar arquivo existente
    escolha = selecionar_mes_download()

    if not escolha["acao"]:
        print("Operação cancelada pelo usuário.")
        return

    if escolha["acao"] == "baixar":
        d_ini, d_fim = calcular_periodo_mes(escolha["mes"], escolha["ano"])
        print(f"\n=======================================================")
        print(f"Iniciando download do SGF ({d_ini} até {d_fim})...")
        print(f"=======================================================")
        try:
            baixar_extrato_sgf(d_ini, d_fim, PASTA_EXPORTS)
        except Exception as e:
            print(f"\n❌ Falha durante o download no SGF: {e}")
            print("Tentando prosseguir com os arquivos já existentes...")

    # 2. Descompacta automaticamente qualquer .zip presente na pasta
    descompactar_zips_pendentes(PASTA_EXPORTS)

    # 3. Carrega arquivos de planilha da pasta exports_sgf
    arquivos = [
        f
        for f in PASTA_EXPORTS.glob("*.xls*")
        if f.is_file() and not f.name.startswith("~$")
    ]

    if not arquivos:
        print(f"⚠️ Nenhum arquivo encontrado em '{PASTA_EXPORTS.resolve()}'!")
        return

    print(f"\n➜ Lendo {len(arquivos)} arquivo(s)...")
    lista_dfs = []
    for caminho in arquivos:
        try:
            df_temp = pd.read_html(
                caminho, encoding="utf-8", decimal=",", thousands="."
            )[0]
            df_temp.columns = df_temp.iloc[0]
            df_temp = df_temp.iloc[1:].reset_index(drop=True)
        except Exception:
            df_temp = pd.read_excel(caminho)

        df_temp.columns = df_temp.columns.astype(str).str.strip()  # type: ignore
        lista_dfs.append(df_temp)

    df_sgf = pd.concat(lista_dfs, ignore_index=True)
    print(f"✓ Total de linhas carregadas: {len(df_sgf)}")

    # Detecta a coluna de Ordem de Serviço de forma resiliente
    col_os = next(
        (c for c in df_sgf.columns if "ordem" in c.lower() and "servi" in c.lower()),
        None,
    )
    if not col_os:
        print("❌ Coluna de Ordem de Serviço não encontrada na planilha!")
        return

    # Puxa todas as OSs únicas
    lista_os_unicas = df_sgf[col_os].dropna().astype(str).unique().tolist()

    print(
        f"➜ Abrindo janela com {len(lista_os_unicas)} OS(s) encontradas na base..."
    )
    os_selecionadas = selecionar_os_manual(lista_os_unicas)

    if not os_selecionadas:
        print("❌ Nenhuma OS selecionada. Processamento cancelado.")
        return

    print(
        f"\n✓ {len(os_selecionadas)} OS(s) selecionada(s) manualmente pelo usuário."
    )

    # Filtra o DataFrame somente com as OSs marcadas
    df_selecionado = df_sgf[df_sgf[col_os].astype(str).isin(os_selecionadas)].copy()

    # Gera o relatório consolidado final
    exportar_relatorio(
        df_selecionado, col_os, "Consolidado_Manual_Selecionado.xlsx"
    )
    print("✅ Processamento finalizado com sucesso!")


if __name__ == "__main__":
    main()
