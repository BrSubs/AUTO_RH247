# Automatizador de Abonos RH247

Script em Python + Playwright para automatizar a criação de **abonos de ponto** no sistema RH247, a partir de uma planilha CSV de justificativas.

---

## Visão geral

Fluxo resumido:

1. Você faz login no RH247 e abre a **lista de funcionários** (tela com campo “Pesquisar...” e a tabela).
2. O script lê o arquivo `justificativas.csv`.
3. Para cada funcionário:
   - Busca o nome na lista.
   - Abre o cartão de ponto.
   - Confere se o nome da tela bate com o CSV.
   - Cria os abonos (um ou mais períodos) conforme a planilha.
   - Volta para a lista para pegar o próximo funcionário.
4. O CSV é atualizado com o **Status** de cada linha e um arquivo de **logs** (`logs_justificativas.csv`) é alimentado com todo o histórico.

---

## Requisitos

- Python 3.9+
- Playwright (modo síncrono)
- pandas

Instalação recomendada:

```bash
pip install playwright pandas
playwright install


Estrutura dos arquivos
1. Planilha de justificativas (justificativas.csv)
A planilha deve estar na mesma pasta do script, com exatamente estas colunas e nesta ordem:

Nome_Completo

Motivo

Data_Inicio

Data_Fim

Status

Regras importantes:

Nome_Completo: nome que aparece na lista do RH247.

Motivo: texto que será preenchido na descrição do abono.

Data_Inicio: data no formato dd/mm/aa ou dd/mm/aaaa.

Data_Fim:

Se estiver vazia, é tratada como igual a Data_Inicio (abono de 1 dia).

Se preenchida, é o fim do intervalo de abono.

Status:

Campo de controle do próprio script, atualizado a cada execução.

Você pode deixar em branco para novas justificativas.

Exemplo (CSV simplificado):

text
Nome_Completo,Motivo,Data_Inicio,Data_Fim,Status
JOAO,FOLGA,13/02/26,, 
MARIA,ATESTADO,27/02/26,27/02/26,
Observação: Data_Fim vazia na primeira linha será assumida como 13/02/26.

2. Arquivo de logs (logs_justificativas.csv)
Gerado automaticamente pelo script, com eventos de:

Preparação de linha (início de processamento).

Busca de funcionário.

Conferência de nome.

Criação de abono (OK / erro / conflito).

Popups de erro do sistema.

Problemas para voltar para a lista, etc.

Estrutura de colunas:

text
timestamp,Nome_Completo,Motivo,Data_Inicio,Data_Fim,acao,status,detalhe
Regras de negócio
Status da justificativa
O campo Status controla o que deve ou não ser reprocessado:

"" (vazio)

Linha ainda não processada.

O script vai tentar criar o abono normalmente.

OK

Abono já foi criado com sucesso em uma execução anterior.

A linha não é reprocessada.

CONFLITO

O RH247 retornou mensagem de conflito de período (ex.: já existe outro atestado no intervalo).

A linha não é reprocessada (fica para análise manual).

Qualquer outro valor (ex.: ERRO)

Linha será reprocessada em execuções futuras.

Datas
Data_Inicio e Data_Fim aceitam dd/mm/aa ou dd/mm/aaaa.

Se Data_Fim for vazia/nula, o script assume Data_Fim = Data_Inicio.

Se Data_Fim < Data_Inicio, a linha é marcada com ERRO.

Erros do sistema (popups)
Quando o RH247 exibe um popup de erro (SweetAlert):

A mensagem é lida e gravada em logs_justificativas.csv.

O modal de abono é fechado.

O Status da linha é atualizado conforme o texto:

Se contiver “conflito” / “conflitante” → CONFLITO.

Caso contrário → ERRO.

Comportamento na tela do sistema
Tela de lista (pesquisa)
Na tela de lista de funcionários o script faz:

Confere se existe:

Campo de busca com placeholder Pesquisar... (ou equivalente).

Tabela principal de funcionários com tbody > tr.

Para cada funcionário novo:

Lê a quantidade de linhas da tabela antes da busca.

Digita o nome no campo “Pesquisar...”.

Clica na lupa (ou envia ENTER).

Fica em “polling” até a quantidade de linhas da tabela mudar ou até estourar um timeout.

Garante que a tabela resultante está visível.

Localiza a linha do funcionário pela coluna de nome.

Clica no botão Editar da linha para abrir o cartão de ponto.

Esse loop evita avançar antes da tabela ser de fato atualizada, reduzindo erros de “funcionário não encontrado”.

Tela de cartão de ponto (espelho)
No cartão de ponto o script:

Confere o nome do funcionário no cabeçalho (<strong>).

Clica no botão “Abono de Ponto”.

Preenche:

Descrição (Motivo).

Data inicial.

Data final.

Clica em Salvar.

Trata possíveis popups de erro (conflito de período, etc.).

Ao terminar todas as justificativas daquele funcionário:

Tenta voltar automaticamente para a lista (botão Voltar, Buscar ou #btn-buscar-crud).

Se não conseguir, registra log de erro e pede que você volte manualmente, confirmando com ENTER.

Fluxo de execução
Garanta que justificativas.csv está com as colunas e formatos esperados.

Execute o script:

bash
python nome_do_script.py
O Playwright abre o navegador com um perfil persistente (perfil_rh247):

Se for a primeira vez, faça login no RH247.

Navegue até a tela de lista de funcionários (pesquisa + tabela).

No terminal, pressione ENTER quando estiver na lista.

O script começa a processar linha a linha:

Logs aparecem no terminal e em logs_justificativas.csv.

O arquivo justificativas.csv é atualizado em disco com o Status de cada linha.

Em alguns cenários (ex.: layout inesperado, botão de voltar diferente), o script pode:

Não conseguir encontrar o campo de busca ou a tabela ao tentar voltar para a lista.

Nesses casos, ele registra o erro e pede no terminal para você arrumar a tela e apertar ENTER.

Ao final, o terminal exibirá uma mensagem de conclusão e você pode encerrar o navegador.

Boas práticas de uso
Sempre faça backup de justificativas.csv antes de rodar um lote grande.

Use um ambiente de teste do RH247 (se houver) para validar comportamentos novos.

Depois de uma grande execução, revise:

justificativas.csv → Status de cada linha.

logs_justificativas.csv → eventuais ERRO ou CONFLITO para tratativa manual.

Evite mexer manualmente no campo Status; deixe o script manter esse controle.

Se a interface do RH247 mudar (botões, IDs, estrutura da tabela), os seletores de Playwright podem precisar ser ajustados.

Limitações atuais
Alguns retornos à tela de lista ainda podem depender de intervenção manual (quando o layout não oferece os botões esperados ou quando há modais inesperados abertos).

A detecção de conflito de período é baseada em texto (palavras “conflito” / “conflitante” na mensagem do popup); se o sistema mudar essa mensagem, será necessário ajustar o código.

O script assume que o nome na planilha bate exatamente com o nome apresentado na coluna da tabela de funcionários.

Próximos passos (ideias de melhoria)
Refino dos seletores de “Voltar para lista” com base no HTML real do cartão de ponto.

Mecanismo mais sofisticado de “detectar mudança na tabela”, não só por quantidade de linhas, mas também por conteúdo (primeira célula da primeira linha, por exemplo).

Suporte a filtros adicionais (matrícula, CPF etc.) se a tela de lista oferecer.