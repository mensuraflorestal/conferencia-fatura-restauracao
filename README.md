# Conferência de Fechamento - Silvicultura (SGF)

Automação para extração e conferência de apontamentos de boletins de silvicultura a partir do sistema SGF (INFLOR).

---

## 🚀 Funcionalidades

1. **Download Automático (RPA / Playwright):**
   * Realiza login no SGF em segundo plano (headless).
   * Acessa a Consulta Completa de Boletins de Silvicultura.
   * Filtra automaticamente o mês e ano selecionados (do 1º ao último dia).
   * Baixa a exportação Excel e descompacta arquivos `.zip` automaticamente para a pasta `exports_sgf/`.

2. **Seleção Manual de OSs:**
   * Interface gráfica com busca em tempo real para marcar as Ordens de Serviço a serem consolidadas.

3. **Consolidação e Relatório:**
   * Agrupa e totaliza por:
     - **Por Horto**
     - **Horto + Talhão**
     - **Por Município**
   * Aplica regras de negócio (ex: zera produção/horas em OSs de taxa `9004`).
   * Gera a planilha `Consolidado_Manual_Selecionado.xlsx` com colunas autoajustadas.

---

## 🛠️ Configuração e Instalação

### 1. Pré-requisitos
* **Python 3.10+** instalado.
* **Microsoft Edge** ou **Google Chrome** instalado na máquina.

### 2. Configurar o Ambiente Virtual
No terminal, na pasta do projeto:
```cmd
python -m venv .venv
call .venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Configurar Credenciais
Copie o arquivo `.env.example` para `.env` e ajuste suas credenciais do SGF se necessário:
```env
SGF_USER=seu_usuario
SGF_PASSWORD=sua_senha
```

---

## ▶️ Como Executar

Basta dar dois cliques no arquivo:
```cmd
run.bat
```
Ou rodar diretamente no terminal:
```cmd
python extract.py
```
