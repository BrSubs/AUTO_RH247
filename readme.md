# 🤖 Automatização de Justificativas de Ponto (RH247)

Script em Python desenvolvido para automatizar o lançamento de **abonos de ponto** no sistema RH247 a partir de um arquivo CSV de justificativas. O objetivo é eliminar o trabalho manual, padronizar o processo e garantir rastreabilidade total via logs.

---

## 📋 Visão Geral

O sistema utiliza o **Playwright** para interagir com a interface web do RH247 e o **Pandas** para a manipulação dos dados. Para cada linha do CSV, o script localiza o funcionário, abre o cartão de ponto, realiza o abono e registra o status da operação.

### Pilares do Projeto
* **Playwright (Síncrono):** Automação de navegador Chromium de fácil manutenção.
* **Pandas:** Carregamento, normalização de datas e atualização de status em lote.
* **Persistência:** O arquivo CSV é reescrito a cada lote, permitindo retomar o processo em caso de falhas.

---

## ⚙️ Configuração e Instalação

### Pré-requisitos
* **Python 3.8** ou superior.
* Bibliotecas: `playwright`, `pandas`.

### Instalação Rápida
1.  **Instalar dependências:**
    ```bash
    pip install playwright pandas
    playwright install
    ```
2.  **Configurar variáveis:** No topo do script `just.py`, ajuste a `URL_PONTO` e os caminhos dos arquivos CSV conforme sua necessidade.

---

## 📊 Estrutura de Dados (CSV)

### Entrada: `justificativas.csv`
O arquivo deve seguir a ordem de colunas abaixo para o processamento correto:

| Coluna | Descrição | Exemplos |
| :--- | :--- | :--- |
| **Status** | Controle (Vazio = Processar; `OK` ou `CONFLITO` = Ignorar) | ``, `OK`, `ERRO` |
| **Nome_Completo** | Nome exatamente como consta no RH247 | `Fulano da Silva Sicrano` |
| **Dia_Inicio** | Dia inicial do período | `01`, `1` |
| **Dia_Fim** | Dia final (se vazio, assume o mesmo do início) | `03`, `` |
| **Mes/Ano** | Competência no formato MM/AA ou MM/AAAA | `03/26`, `03/2026` |
| **Motivo** | Descrição que será gravada no sistema | `Folga`, `Atestado` |

---

## 🚀 Fluxo de Operação

1.  **Preparação:** Preencha o `justificativas.csv` com os dados dos funcionários e deixe a coluna `Status` em branco.
2.  **Execução:** Rode o comando `python just.py`.
3.  **Login e Navegação:** O navegador abrirá. Realize o login manualmente e navegue até a **Lista de Funcionários**.
4.  **Início:** Quando estiver na tela de busca, volte ao terminal e pressione **ENTER** para iniciar a automação.
5.  **Logs:** Acompanhe o arquivo `logs_justificativas.csv` para verificar sucessos, conflitos de data ou erros de sistema.

---

## 🛠️ Regras de Negócio e Segurança

* **Tratamento de Erros:** O script detecta popups de "período conflitante", marca a linha como `CONFLITO` e segue para o próximo item sem travar.
* **Validação de Identidade:** Antes de salvar, o script valida se o nome exibido no espelho de ponto coincide com o nome no CSV.
* **Intervenção Manual:** Caso o script perca o contexto da página (ex: login expirado), ele solicita intervenção manual via terminal antes de continuar.

---

## 🔮 Melhorias Futuras
- [ ] Implementar modo `headless` total para rodar em segundo plano.
- [ ] Migrar logs de `print` para o módulo nativo `logging`.
- [ ] Criar testes unitários para as funções de conversão de data.
