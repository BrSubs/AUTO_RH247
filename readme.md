```markdown
# Automatização de Justificativas de Ponto (RH247)

Script em Python para automatizar o lançamento de **abonos de ponto** no RH247 a partir de um arquivo CSV de justificativas.  
O objetivo é reduzir trabalho manual, padronizar o processo e manter rastreabilidade completa via logs.

---

## Sumário

- [Visão Geral](#visão-geral)
- [Arquitetura e Decisões de Projeto](#arquitetura-e-decisões-de-projeto)
- [Formato dos Arquivos CSV](#formato-dos-arquivos-csv)
  - [justificativas.csv](#justificativascsv)
  - [logs_justificativas.csv](#logs_justificativascsv)
- [Pré-requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Como Usar](#como-usar)
  - [Passo a passo de execução](#passo-a-passo-de-execução)
  - [Fluxo de processamento](#fluxo-de-processamento)
- [Regras de Negócio Implementadas](#regras-de-negócio-implementadas)
- [Detalhes Técnicos Importantes](#detalhes-técnicos-importantes)
- [Boas Práticas e Possíveis Melhorias Futuras](#boas-práticas-e-possíveis-melhorias-futuras)
- [Licença](#licença)

---

## Visão Geral

Este projeto automatiza o cadastro de abonos de ponto de funcionários no sistema RH247, lendo um arquivo `justificativas.csv` e interagindo com a interface web via Playwright.  
Para cada linha do CSV, o script:

- Localiza o funcionário na lista do RH247.  
- Abre o cartão de ponto (espelho) daquele funcionário.  
- Cria um abono de ponto no período informado.  
- Atualiza o status da linha no CSV.  
- Registra logs de negócio e de sistema em `logs_justificativas.csv`.

---

## Arquitetura e Decisões de Projeto

### Principais componentes

- **Playwright (modo síncrono)**  
  Usado para automação de navegador (Chromium) e interação com o RH247.  
  Optou-se pelo modo síncrono (`playwright.sync_api`) por ser mais simples de ler e manter em scripts de automação operacional.

- **Pandas**  
  Responsável por:
  - Carregar e validar o CSV de justificativas.  
  - Ordenar e normalizar os dados.  
  - Atualizar o status e reescrever o arquivo no novo formato.

- **CSV + datetime nativos**  
  Utilizados para:
  - Logs em `logs_justificativas.csv`.  
  - Conversão e validação de datas (funções `parse_data_br` e `montar_intervalo_datas`).

### Decisões de projeto

1. **Formato de datas no CSV**  
   - Input do usuário: dias e mês/ano separados (`Dia_Inicio`, `Dia_Fim`, `Mes/Ano`).  
   - Internamente: o script converte para `Data_Inicio` e `Data_Fim` no formato `dd/mm/aaaa` para simplificar ordenação e envio para o RH247.  
   - Isso permite que o usuário trabalhe só com dia e mês/ano, o que é mais natural na rotina.

2. **Persistência de estado via CSV**  
   - O mesmo arquivo `justificativas.csv` é lido e reescrito a cada “lote” por funcionário, com a coluna `Status` atualizada.  
   - Isso protege contra falhas no meio do processo: se o script for interrompido, ele consegue retomar ignorando linhas já processadas.

3. **Logs detalhados em CSV**  
   - Todos os eventos relevantes são registrados em `logs_justificativas.csv`:  
     - Erros de sistema (popups, falha de navegação).  
     - Conflitos de período.  
     - Sucesso na criação de abonos.  
   - Logs em CSV facilitam auditoria posterior, filtro por funcionário e análise em Excel.

4. **Intervenção manual mínima, mas possível**  
   - Em situações em que a tela esperada não é encontrada (mudança de layout, login expirado etc.), o script pede intervenção manual e aguarda `ENTER`.  
   - Isso evita travar o processo de forma silenciosa e permite “resgatar” a execução sem reiniciar tudo.

---

## Formato dos Arquivos CSV

### justificativas.csv

É o arquivo **principal** de entrada e saída do processo.  
Formato atual (ordem das colunas):

```text
Status,Nome_Completo,Dia_Inicio,Dia_Fim,Mes/Ano,Motivo
```

- `Status` (string, pode ser vazio)
  - `""` (vazio): ainda não processado (será processado).  
  - `"OK"`: abono criado com sucesso (não será reprocessado).  
  - `"CONFLITO"`: houve conflito de período (não será reprocessado automaticamente).  
  - `"ERRO"` ou outro texto: erro genérico, o script tentará processar novamente.

- `Nome_Completo` (string)  
  - Nome do funcionário, conforme aparece na lista do RH247.

- `Dia_Inicio` (inteiro ou texto conversível, ex.: `1`, `01`, `1.0`)  
  - Dia inicial do período de abono dentro do mês (`Mes/Ano`).

- `Dia_Fim` (inteiro, vazio ou `NaN`, ex.: `3`, `3.0`, `""`)  
  - Dia final do período de abono dentro do mês (`Mes/Ano`).  
  - Se vazio, será assumido igual a `Dia_Inicio`.

- `Mes/Ano` (string)  
  - Formatos aceitos: `MM/AA` ou `MM/AAAA`.  
  - Ex.: `03/26` ou `03/2026`.

- `Motivo` (string)  
  - Descrição do abono que será lançada no RH247.

#### Exemplos

```text
Status,Nome_Completo,Dia_Inicio,Dia_Fim,Mes/Ano,Motivo
,yan brasil,1,3,03/26,folga
,yan brasil,1,,03/2026,folga
OK,maria silva,10,12,02/2026,atestado médico
```

> Observação: o próprio script reescreve esse arquivo no mesmo formato após o processamento, apenas atualizando o `Status`.

---

### logs_justificativas.csv

Arquivo de saída com os logs do processo.  

Formato típico:

```text
timestamp,Nome_Completo,Motivo,Data_Inicio,Data_Fim,acao,status,detalhe
2026-03-23 09:00:00,yan brasil,folga,01/03/2026,03/03/2026,criar_abono,OK,Abono criado sem erro de sistema.
2026-03-23 09:02:10,yan brasil,folga,01/03/2026,03/03/2026,popup_erro,ERRO,Período conflitante - Já existe abono neste intervalo.
2026-03-23 09:05:30,,,,"esperar_tela_lista",ERRO,Nao encontrou campo de busca ou tabela da lista. ...
```

- Cada linha descreve uma ação ou evento de sistema:
  - `acao`: exemplo `preparar_abono`, `criar_abono`, `popup_erro`, `erro_geral`, `voltar_lista`, `esperar_tela_lista`.  
  - `status`: `INFO`, `OK`, `ERRO`, `CONFLITO`.  
  - `detalhe`: mensagem descritiva (incluindo textos de popups do RH247).

---

## Pré-requisitos

- **Sistema operacional**  
  - Windows ou outro sistema compatível com Playwright (Chromium).

- **Python**  
  - Python 3.8 ou superior instalado e disponível no `PATH`. [web:60]

- **Bibliotecas Python**  
  - `playwright`
  - `pandas`

- **Dependências do Playwright**  
  - Navegadores instalados via `playwright install`.

---

## Instalação

1. **Clonar o projeto ou copiar os arquivos**

   Coloque o script principal (por exemplo `just.py`) em uma pasta de trabalho, junto com os arquivos CSV.

2. **Criar e ativar um ambiente virtual (opcional, recomendado)**

   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   ```

3. **Instalar dependências**

   ```bash
   pip install playwright pandas
   ```

4. **Instalar os navegadores do Playwright** [web:57][web:60]

   ```bash
   playwright install
   ```

5. **Verificar a instalação**

   ```bash
   python -m playwright --version
   python -c "import pandas; print(pandas.__version__)"
   ```

---

## Configuração

As principais configurações ficam no topo do script:

```python
CAMINHO_CSV_JUSTIFICATIVAS = "justificativas.csv"
CAMINHO_CSV_LOGS = "logs_justificativas.csv"
URL_PONTO = "https://rh247.com.br/230540701/ponto"
PERFIL_BROWSER = "perfil_rh247"
```

- **CAMINHO_CSV_JUSTIFICATIVAS**  
  - Caminho do arquivo de justificativas. Pode ser relativo ou absoluto.

- **CAMINHO_CSV_LOGS**  
  - Onde será criado/atualizado o arquivo de logs.

- **URL_PONTO**  
  - URL inicial do módulo de ponto do RH247.

- **PERFIL_BROWSER**  
  - Diretório do perfil persistente do Chromium (usado pelo Playwright).  
  - Permite reaproveitar sessão/logins entre execuções.

Se necessário, ajuste esses valores de acordo com o seu ambiente (por exemplo, outro CNPJ ou outra URL do RH247).

---

## Como Usar

### Passo a passo de execução

1. **Preparar o `justificativas.csv`**  
   - Preencha as linhas no formato descrito em [justificativas.csv](#justificativascsv).  
   - Deixe `Status` vazio para as linhas que devem ser processadas.

2. **Executar o script**

   No terminal/prompt de comando, na pasta do projeto:

   ```bash
   python just.py
   ```

3. **Primeira execução (login)**  
   - O navegador Chromium será aberto com o perfil configurado.  
   - Caso seja a primeira vez, faça o login no RH247 manualmente.  
   - Navegue até a tela de **LISTA de funcionários**, onde exista o campo “Pesquisar...” e a tabela de funcionários.

4. **Confirmar a tela de lista**  
   - No terminal, o script exibirá instruções.  
   - Quando você estiver na tela correta, pressione `ENTER` no terminal.

5. **Acompanhar o processamento**  
   - O script irá:
     - Buscar cada funcionário.  
     - Abrir o cartão de ponto.  
     - Lançar os abonos.  
     - Voltar para a lista quando mudar de funcionário.  
   - O arquivo `justificativas.csv` será atualizado ao longo do processo, e os logs serão gravados em `logs_justificativas.csv`.

6. **Finalização**  
   - Ao fim, o script informará que o processamento foi concluído.  
   - Pressione `ENTER` para encerrar (você pode fechar o navegador manualmente em seguida).

---

### Fluxo de processamento

Para cada linha do `justificativas.csv`:

1. Converte `Dia_Inicio`, `Dia_Fim` e `Mes/Ano` para `Data_Inicio` e `Data_Fim` (formato `dd/mm/aaaa`).  
2. Ordena as linhas por `Nome_Completo` e `Data_Inicio`.  
3. Para cada funcionário (bloco de linhas com o mesmo nome):
   - Se necessário, volta para a lista e abre o cartão de ponto daquele funcionário.  
   - Confere se o nome exibido no espelho corresponde ao CSV.  
   - Para cada justificativa com `Status` processável:
     - Tenta criar o abono.
     - Trata popups de erro ou conflito.  
     - Atualiza `Status` no DataFrame (`OK`, `CONFLITO`, `ERRO`).  
   - Ao mudar de funcionário (ou no fim do arquivo):
     - Volta para a lista.  
     - Reescreve o `justificativas.csv` no **formato novo** com o `Status` atualizado.

---

## Regras de Negócio Implementadas

### Decisão de processamento por `Status`

Função `deve_processar`:

- `Status == ""` (vazio) → **processa** (primeira vez).  
- `Status == "OK"` → **não processa** (já abonado).  
- `Status == "CONFLITO"` → **não processa** (conflito já detectado).  
- Qualquer outro valor (ex.: `"ERRO"`) → **processa novamente**.

### Interpretação das datas

- Entrada do CSV:  
  - `Dia_Inicio`, `Dia_Fim`, `Mes/Ano`.  
- Regra de intervalo:
  - Se `Dia_Fim` estiver vazio ou `NaN`, é assumido igual a `Dia_Inicio`.  
  - `Mes/Ano` aceita `MM/AA` ou `MM/AAAA`; anos com 2 dígitos são convertidos para `2000 + AA`.  
  - Validação garante que `Data_Fim >= Data_Inicio`.

### Tratamento de popups de erro

- Se o RH247 exibir um popup de erro (SweetAlert2):  
  - A mensagem é lida e registrada em `logs_justificativas.csv`.  
  - O botão “OK” do popup é clicado.  
  - Se a mensagem contiver “conflitante” ou “conflito”:
    - `Status` da justificativa = `"CONFLITO"`.  
  - Caso contrário:
    - `Status` da justificativa = `"ERRO"`.

---

## Detalhes Técnicos Importantes

- **Automação de lista de funcionários**
  - Localiza o campo de busca “Pesquisar...” com lógica robusta (vários fallbacks).  
  - Dispara a busca via botão de lupa ou tecla ENTER.  
  - Aguarda mudanças na tabela e depois garante que haja ao menos uma linha visível.

- **Seleção de funcionário e cartão de ponto**
  - Procura pela linha cujo nome (coluna 2) corresponde ao `Nome_Completo` normalizado.  
  - Clica no botão de edição (`.btn-grid-edit` ou equivalentes) daquela linha.

- **Validação no espelho de ponto**
  - Busca um elemento `<strong>` com o nome normalizado do funcionário.  
  - Se o nome exibido for diferente do CSV, lança erro e registra log.

- **Criação do abono**
  - Localiza o botão “Abono de Ponto” (via função de acessibilidade ou texto).  
  - Preenche descrição, data inicial e final.  
  - Clica em “Salvar” e trata eventuais popups.

- **Retorno à lista**
  - Tenta clicar em botões conhecidos de “Buscar” ou “Voltar” na tela de espelho.  
  - Se falhar, pede intervenção manual e registra log.

---

## Boas Práticas e Possíveis Melhorias Futuras

- **Desempenho**  
  - Reduzir ou remover `slow_mo` quando o script estiver estável. [web:41][web:43]  
  - Substituir `wait_for_timeout` por esperas específicas de elementos (`locator.wait_for`).

- **Melhoria de logging**  
  - Migrar os `print` para o módulo `logging`, com níveis (`INFO`, `DEBUG`, `ERROR`) e possibilidade de ativar/desativar debug sem alterar código. [web:50]

- **Configuração externa**  
  - Mover parâmetros (URLs, caminhos de arquivos, perfis) para um `.env` ou arquivo de configuração (`.ini`, `yaml`).

- **Testes automatizados**  
  - Criar pequenos testes unitários para as funções de data (`parse_data_br`, `montar_intervalo_datas`) e para a transformação de CSV.

- **Modo totalmente headless**  
  - Após estabilidade, rodar com `headless=True` e tudo automatizado (login via script), para execução em servidor/agenda.

```