import csv
import datetime
import os
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd
from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
    Page,
)

# ============================================================
# CONFIGURAÇÕES GERAIS
# ============================================================

CAMINHO_CSV_JUSTIFICATIVAS = "justificativas.csv"
CAMINHO_CSV_LOGS = "logs_justificativas.csv"
URL_PONTO = "https://rh247.com.br/230540701/ponto"
PERFIL_BROWSER = "perfil_rh247"

TIMEOUT_CURTO_MS = 5_000
TIMEOUT_PADRAO_MS = 15_000


# ============================================================
# MODELOS DE DADOS
# ============================================================

@dataclass
class Justificativa:
    """
    Modelo de uma justificativa carregada da planilha.
    Representa uma linha lógica do CSV.
    """
    nome_completo: str
    motivo: str
    data_inicio: str
    data_fim: str
    status: str
    indice_df: int  # índice da linha no DataFrame para atualizar Status


# ============================================================
# FUNÇÕES UTILITÁRIAS
# ============================================================

def normalizar_nome(nome: str) -> str:
    """
    Normaliza nomes para comparação:
    - Remove espaços extras.
    - Converte para maiúsculas.
    """
    if not isinstance(nome, str):
        return ""
    return " ".join(nome.strip().upper().split())


def parse_data_br(data_str: str) -> Optional[datetime.date]:
    """
    Converte datas nos formatos:
    - dd/mm/aa
    - dd/mm/aaaa

    Retorna None se não conseguir converter.
    """
    if not isinstance(data_str, str):
        return None
    data_str = data_str.strip()
    if not data_str:
        return None

    for fmt in ("%d/%m/%y", "%d/%m/%Y"):
        try:
            return datetime.datetime.strptime(data_str, fmt).date()
        except ValueError:
            continue

    return None


def registrar_log_csv(just: Justificativa, acao: str, status: str, detalhe: str) -> None:
    """
    Registra logs de negócio (criação de abono, conflitos, erros) no arquivo CSV.
    """
    Path(CAMINHO_CSV_LOGS).touch(exist_ok=True)
    arquivo_existe = os.path.getsize(CAMINHO_CSV_LOGS) > 0

    with open(CAMINHO_CSV_LOGS, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        if not arquivo_existe:
            writer.writerow([
                "timestamp",
                "Nome_Completo",
                "Motivo",
                "Data_Inicio",
                "Data_Fim",
                "acao",
                "status",
                "detalhe",
            ])

        agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        writer.writerow([
            agora,
            just.nome_completo,
            just.motivo,
            just.data_inicio,
            just.data_fim,
            acao,
            status,
            detalhe,
        ])


def registrar_log_sistema(etapa: str, status: str, mensagem: str) -> None:
    """
    Log leve para eventos de sistema (navegação, validações, etc.).
    Campos de justificativa ficam vazios.
    """
    Path(CAMINHO_CSV_LOGS).touch(exist_ok=True)

    agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(CAMINHO_CSV_LOGS, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([agora, "", "", "", "", etapa, status, mensagem])


# ============================================================
# CARREGAMENTO E PRÉ-PROCESSAMENTO DO CSV
# ============================================================

def carregar_justificativas(caminho_csv: str) -> pd.DataFrame:
    """
    Carrega o CSV de justificativas, garante colunas esperadas e
    aplica os ajustes de formato solicitados:

    Colunas (ordem nova):
    - Nome_Completo
    - Motivo
    - Data_Inicio
    - Data_Fim
    - Status

    Regras:
    - Se Data_Fim estiver vazia, considera igual a Data_Inicio.
    - Status nulo é tratado como string vazia.
    """
    df = pd.read_csv(caminho_csv, encoding="latin1")
    df.columns = [c.strip() for c in df.columns]

    # Ajuste de nomes de colunas (se vierem com pequenas variações)
    rename_map = {}
    if "Nome Completo" in df.columns and "Nome_Completo" not in df.columns:
        rename_map["Nome Completo"] = "Nome_Completo"
    if "Data Inicio" in df.columns and "Data_Inicio" not in df.columns:
        rename_map["Data Inicio"] = "Data_Inicio"
    if "Data Fim" in df.columns and "Data_Fim" not in df.columns:
        rename_map["Data Fim"] = "Data_Fim"
    if rename_map:
        df = df.rename(columns=rename_map)

    colunas_esperadas = ["Nome_Completo", "Motivo", "Data_Inicio", "Data_Fim", "Status"]
    for col in colunas_esperadas:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatória ausente no CSV: {col}")

    # Normaliza Status nulo para string vazia
    df["Status"] = df["Status"].fillna("")

    # Regra: se Data_Fim estiver vazia, assume igual à Data_Inicio
    df["Data_Fim"] = df.apply(
        lambda row: row["Data_Inicio"] if (str(row["Data_Fim"]).strip() == "" or str(row["Data_Fim"]).lower() == "nan")
        else row["Data_Fim"],
        axis=1
    )

    return df


def ordenar_justificativas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ordena justificativas por Nome_Completo e Data_Inicio (convertida),
    para processar registros do mesmo funcionário agrupados por data.
    """
    df["Data_Ordenacao"] = df["Data_Inicio"].apply(parse_data_br)
    df_ordenado = df.sort_values(
        by=["Nome_Completo", "Data_Ordenacao"],
        kind="stable"
    ).drop(columns=["Data_Ordenacao"]).reset_index(drop=True)
    return df_ordenado


def linha_para_justificativa(df: pd.DataFrame, idx: int) -> Justificativa:
    """
    Converte uma linha do DataFrame em um objeto Justificativa.
    """
    row = df.iloc[idx]
    return Justificativa(
        nome_completo=str(row["Nome_Completo"]),
        motivo=str(row["Motivo"]),
        data_inicio=str(row["Data_Inicio"]),
        data_fim=str(row["Data_Fim"]),
        status=str(row["Status"]),
        indice_df=idx,
    )


# ============================================================
# LOCALIZAÇÃO DE ELEMENTOS NA LISTA
# ============================================================

def localizar_campo_busca_lista(page: Page):
    """
    Localiza o input de busca da lista (placeholder 'Pesquisar...'),
    sempre retornando um elemento visível.
    """
    campos = page.locator("input.form-control[placeholder='Pesquisar...']")
    total = campos.count()
    print(f"[DEBUG] Inputs com placeholder 'Pesquisar...': {total}")

    for i in range(total):
        try:
            campo = campos.nth(i)
            if campo.is_visible():
                print(f"[DEBUG] Usando input visível de índice {i} para busca.")
                return campo
        except Exception:
            continue

    grupos = page.locator("div.input-group.input-group-sm")
    for i in range(grupos.count()):
        try:
            grupo = grupos.nth(i)
            if not grupo.is_visible():
                continue
            campo = grupo.locator("input.form-control").first
            if campo.is_visible():
                print(f"[DEBUG] Usando input em grupo visível de índice {i} para busca.")
                return campo
        except Exception:
            continue

    texto_visiveis = page.locator("input[type='text']").filter(has=page.locator(":visible"))
    if texto_visiveis.count() > 0:
        print("[DEBUG] Usando fallback: primeiro input[type='text'] visível.")
        return texto_visiveis.first

    print("[DEBUG] Nenhum input 'Pesquisar...' visível encontrado; usando primeiro input.texto como fallback duro.")
    return page.locator("input[type='text']").first


def localizar_tbody_lista(page: Page):
    """
    Localiza o tbody da tabela principal de funcionários.

    Critérios:
    - table que tenha tbody>tr
    - e tenha ao menos uma linha com primeira coluna .min-width (matrícula)
    - e algum botão .btn-grid-edit (editar cartão) na tabela.
    """
    todas_tabelas = page.locator("table").filter(has=page.locator("tbody tr"))
    total_tabelas = todas_tabelas.count()
    print(f"[DEBUG] Tabelas com tbody e linhas encontradas: {total_tabelas}")

    if total_tabelas == 0:
        raise RuntimeError("Nenhuma tabela com tbody e linhas foi encontrada na tela de lista.")

    for i in range(total_tabelas):
        tabela = todas_tabelas.nth(i)
        tbody = tabela.locator("tbody")
        linhas = tbody.locator("tr")
        if linhas.count() == 0:
            continue

        primeira_coluna = linhas.first.locator("td").first
        classes = (primeira_coluna.get_attribute("class") or "").split()
        tem_min_width = "min-width" in classes

        botoes_editar = tabela.locator("button.btn-grid-edit")

        if tem_min_width and botoes_editar.count() > 0:
            print(f"[DEBUG] Selecionando tabela índice {i} como tabela de funcionários.")
            return tbody

    print("[DEBUG] Nenhuma tabela bateu nos critérios fortes; usando a primeira tabela como fallback.")
    return todas_tabelas.first.locator("tbody")


def esperar_tabela_resultados(page: Page, timeout_ms: int = TIMEOUT_PADRAO_MS):
    """
    Após a busca, espera o tbody da tabela principal ter ao menos uma linha visível.
    Não exige mudança no total de linhas.
    """
    deadline = time.time() + timeout_ms / 1000
    while time.time() < deadline:
        tbody = localizar_tbody_lista(page)
        linhas = tbody.locator("tr")
        total = linhas.count()
        print(f"[DEBUG] Tabela de resultados: total de linhas = {total}")

        if total > 0:
            try:
                linhas.first.wait_for(state="visible", timeout=1_000)
                return tbody
            except PlaywrightTimeoutError:
                pass

        page.wait_for_timeout(200)

    raise RuntimeError("Timeout ao aguardar a tabela de resultados apos a busca.")


def esperar_tabela_mudar(page: Page, total_antes: int, timeout_ms: int = 15000):
    """
    Espera até que a quantidade de linhas da tabela principal
    seja diferente de total_antes, ou até estourar timeout.

    Se continuar igual até o fim do timeout, apenas retorna.
    """
    deadline = time.time() + timeout_ms / 1000
    while time.time() < deadline:
        tbody = localizar_tbody_lista(page)
        linhas = tbody.locator("tr")
        total_atual = linhas.count()
        print(f"[DEBUG] Aguardando mudança na tabela: antes={total_antes}, atual={total_atual}")

        if total_atual != total_antes:
            print("[DEBUG] Mudança na quantidade de linhas da tabela detectada.")
            return

        page.wait_for_timeout(300)


def garantir_tela_lista(page: Page):
    """
    Garante que a tela de lista está carregada, validando:
    - Campo de busca visível.
    - Tabela principal de funcionários com tbody.

    Se não conseguir, registra erro e pede intervenção manual
    (você volta para a tela de lista e aperta ENTER).
    """
    print("[INFO] Verificando se a tela de lista está carregada...")
    registrar_log_sistema(
        "esperar_tela_lista",
        "INFO",
        f"Verificando tela de lista. URL='{page.url}', titulo='{page.title()}'",
    )

    try:
        # tenta localizar o campo de busca da tela de lista
        campo_busca = localizar_campo_busca_lista(page)
        print(f"[DEBUG] Locator do campo de busca: {campo_busca}")
        campo_busca.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)

        tbody = localizar_tbody_lista(page)
        qtd_linhas = tbody.locator("tr").count()

        if qtd_linhas == 0:
            print("[ALERTA] Tela de lista carregou, mas a tabela está vazia (sem linhas).")
        else:
            print("[INFO] Tela de lista carregada, tabela com linhas detectada.")

        registrar_log_sistema(
            "esperar_tela_lista",
            "OK",
            "Campo de busca e tabela encontrados na tela de lista.",
        )
        return campo_busca

    except Exception as e:
        print("[ALERTA] Não consegui confirmar a tela de lista (campo ou tabela).")
        print(f"[ALERTA] URL atual: {page.url}")
        print(f"[ALERTA] Erro: {e}")
        print("Provavelmente você ainda está na tela de espelho, em um modal, ou em outra tela.")
        print("Volte manualmente para a LISTA de funcionários (onde tem o campo 'Pesquisar...' e a tabela).")
        print("Depois, pressione ENTER aqui no terminal para continuar...")

        registrar_log_sistema(
            "esperar_tela_lista",
            "ERRO",
            f"Nao encontrou campo de busca ou tabela da lista. URL='{page.url}', erro='{e}'",
        )
        input()
        return None


# ============================================================
# AÇÕES NA LISTA
# ============================================================

def buscar_funcionario_na_lista(page: Page, nome_funcionario: str) -> None:
    """
    Executa a busca pelo funcionário na tela de lista:
    - Lê o total de linhas da tabela ANTES da busca.
    - Preenche o campo 'Pesquisar...' com o nome normalizado.
    - Clica no botão de lupa (ou ENTER).
    - Espera a tabela mudar de quantidade de linhas (ou até timeout).
    - Depois disso chama esperar_tabela_resultados.
    """
    nome_norm = normalizar_nome(nome_funcionario)
    print(f"[INFO] Buscando funcionário pelo nome: '{nome_norm}'")

    registrar_log_sistema(
        "buscar_funcionario",
        "INFO",
        f"Iniciando busca do funcionario. URL='{page.url}', titulo='{page.title()}'",
    )

    # estado pré-busca
    tbody_antes = localizar_tbody_lista(page)
    total_antes = tbody_antes.locator("tr").count()
    print(f"[DEBUG] Total de linhas ANTES da busca: {total_antes}")

    campo_busca = localizar_campo_busca_lista(page)
    campo_busca.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
    campo_busca.fill("")
    campo_busca.fill(nome_norm)

    # clicar na lupa
    try:
        grupo = campo_busca.locator("xpath=ancestor::div[contains(@class,'input-group')]")
        botao_lupa = grupo.locator("button.btn.btn-info.btn-flat").first
        if botao_lupa.count() == 0:
            botao_lupa = grupo.locator("button:has(i.flaticon-search-1)").first
        botao_lupa.click()
    except Exception as e:
        print(f"[ALERTA] Não consegui clicar na lupa da busca, tentando com ENTER. Erro: {e}")
        campo_busca.press("Enter")

    # espera “inteligente”: aguarda a tabela ter quantidade de linhas diferente
    esperar_tabela_mudar(page, total_antes, timeout_ms=15000)

    # e depois garante que a tabela final está visível/estável
    tbody = esperar_tabela_resultados(page)
    total = tbody.locator("tr").count()
    print(f"[DEBUG] Após a busca, total de linhas na tabela: {total}")


def abrir_cartao_ponto_funcionario(page: Page, nome_funcionario: str) -> None:
    """
    Localiza a linha do funcionário na tabela (coluna de nome) e
    clica no botão 'Editar' dessa linha para abrir o cartão de ponto.
    """
    nome_norm = normalizar_nome(nome_funcionario)
    print(f"[INFO] Procurando na tabela o funcionário '{nome_norm}' para abrir o ponto...")

    tbody = localizar_tbody_lista(page)

    linha = tbody.locator(
        "tr:has(td:nth-child(2) a span:has-text('%s'))" % nome_norm
    )

    total_linhas = linha.count()
    print(f"[DEBUG] Linhas encontradas com nome '{nome_norm}': {total_linhas}")
    if total_linhas == 0:
        raise Exception(f"Nenhuma linha da tabela corresponde exatamente a '{nome_norm}'.")

    linha_target = linha.first
    linha_target.scroll_into_view_if_needed()
    page.wait_for_timeout(200)

    print("[INFO] Linha do funcionário localizada, tentando localizar botão Editar...")

    botao_editar = linha_target.locator("button.btn-grid-edit")
    if botao_editar.count() == 0:
        print("[DEBUG] Nenhum .btn-grid-edit na linha, tentando por data-original-title='Edit Task'...")
        botao_editar = linha_target.locator("[data-original-title='Edit Task']")

    if botao_editar.count() == 0:
        raise Exception("Não encontrei botão de Editar na linha correspondente ao funcionário.")

    botao_editar.first.scroll_into_view_if_needed()
    page.wait_for_timeout(200)
    botao_editar.first.click()
    print("[INFO] Clique no botão Editar realizado.")


def conferir_nome_no_espelho(page: Page, nome_esperado: str) -> None:
    """
    Valida se o nome exibido no cabeçalho do cartão de ponto
    corresponde ao nome da justificativa.
    """
    nome_norm = normalizar_nome(nome_esperado)
    print(f"[INFO] Conferindo nome na tela de espelho (esperado: '{nome_norm}')")

    try:
        locator_nome = page.locator("strong", has_text=nome_norm)
        locator_nome.first.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
        texto_lido = normalizar_nome(locator_nome.first.inner_text())
        print(f"[DEBUG] Nome lido na tela de espelho: '{texto_lido}'")
    except Exception as e:
        raise Exception(
            f"Nao foi possivel ler o nome do funcionario na tela de espelho "
            f"(strong com texto '{nome_norm}' nao encontrado ou invisivel): {e}"
        )

    if texto_lido != nome_norm:
        raise Exception(
            f"Nome na tela de espelho diferente do CSV. Esperado='{nome_norm}', lido='{texto_lido}'"
        )

    print("[INFO] Nome na tela de espelho confere com o CSV.")


# ============================================================
# AÇÕES NO CARTÃO DE PONTO
# ============================================================

def fechar_popup_erro(page: Page, justificativa: Justificativa) -> Tuple[bool, Optional[str]]:
    """
    Verifica se há um popup de erro (SweetAlert2).
    Se houver:
      - Lê a mensagem.
      - Clica em OK.
      - Registra log de negócio.
      - Retorna (True, mensagem).

    Se não houver popup, retorna (False, None).
    """
    try:
        popup = page.locator("div.swal2-popup.swal2-modal.swal2-icon-error.swal2-show")
        if not popup.is_visible(timeout=1_000):
            return False, None
    except PlaywrightTimeoutError:
        return False, None
    except Exception:
        return False, None

    try:
        titulo = page.locator("#swal2-title").inner_text(timeout=TIMEOUT_CURTO_MS)
    except Exception:
        titulo = ""

    try:
        conteudo = page.locator("#swal2-content").inner_text(timeout=TIMEOUT_CURTO_MS)
    except Exception:
        conteudo = ""

    mensagem = f"{titulo} - {conteudo}".strip() or "Popup de erro sem mensagem legível."
    print(f"[ALERTA] Popup de erro detectado: {mensagem}")

    try:
        botao_ok = page.locator("button.swal2-confirm")
        botao_ok.scroll_into_view_if_needed()
        page.wait_for_timeout(500)
        botao_ok.click()
        page.wait_for_timeout(1_000)
    except Exception:
        print("[ALERTA] Não consegui clicar no botão OK do popup.")

    registrar_log_csv(justificativa, "popup_erro", "ERRO", mensagem)
    return True, mensagem


def criar_abono_no_cartao(page: Page, justificativa: Justificativa) -> str:
    """
    Cria um abono no cartão de ponto para a justificativa.

    Retornos de Status:
    - "OK": abono criado sem erro do sistema.
    - "CONFLITO": houve popup com mensagem de conflito, não cria novo abono.
    - "ERRO": qualquer outra falha (ex.: erro genérico do sistema).
    """
    nome = justificativa.nome_completo
    data_ini_str_original = justificativa.data_inicio
    data_fim_str_original = justificativa.data_fim
    motivo = justificativa.motivo

    print(
        f"[INFO] Criando abono para {nome} | {data_ini_str_original} -> "
        f"{data_fim_str_original} | Motivo: {motivo}"
    )

    data_ini = parse_data_br(data_ini_str_original)
    data_fim = parse_data_br(data_fim_str_original)

    if data_ini is None or data_fim is None:
        raise RuntimeError(
            f"Datas inválidas na justificativa: "
            f"Data_Inicio='{data_ini_str_original}', Data_Fim='{data_fim_str_original}'."
        )

    if data_fim < data_ini:
        raise RuntimeError(
            f"Data_Fim menor que Data_Inicio: {data_ini_str_original} > {data_fim_str_original}."
        )

    data_ini_str = data_ini.strftime("%d/%m/%Y")
    data_fim_str = data_fim.strftime("%d/%m/%Y")

    try:
        print("[INFO] Procurando botão 'Abono de Ponto'...")

        try:
            botao_abono = page.get_by_role("button", name="Abono de Ponto")
            botao_abono.wait_for(state="visible", timeout=TIMEOUT_CURTO_MS)
        except PlaywrightTimeoutError:
            botao_abono = page.get_by_text("Abono", exact=False).filter(
                has=page.locator("button")
            ).first
            botao_abono.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)

        botao_abono.scroll_into_view_if_needed()
        page.wait_for_timeout(250)
        botao_abono.click()
        page.wait_for_timeout(500)
    except PlaywrightTimeoutError:
        raise RuntimeError("Botão 'Abono de Ponto' não encontrado ou não visível.")

    try:
        campo_descricao = page.locator("#descricao")
        campo_data_ini = page.locator("#data_ini")
        campo_data_fim = page.locator("#data_fim")

        campo_descricao.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
        campo_data_ini.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)
        campo_data_fim.wait_for(state="visible", timeout=TIMEOUT_PADRAO_MS)

        campo_descricao.fill(motivo)
        campo_data_ini.fill(data_ini_str)
        campo_data_fim.fill(data_fim_str)

        botao_salvar = page.get_by_role("button", name="Salvar")
        botao_salvar.click()
    except PlaywrightTimeoutError:
        raise RuntimeError("Erro ao preencher campos de abono ou clicar em 'Salvar'.")

    page.wait_for_timeout(2_000)

    houve_popup, mensagem = fechar_popup_erro(page, justificativa)
    if not houve_popup:
        registrar_log_csv(justificativa, "criar_abono", "OK", "Abono criado sem erro de sistema.")
        return "OK"

    mensagem_lower = (mensagem or "").lower()
    if "conflitante" in mensagem_lower or "conflito" in mensagem_lower:
        print("[INFO] Abono não criado devido a conflito de período (status CONFLITO). Fechando modal...")
        status = "CONFLITO"
    else:
        print("[INFO] Abono não criado devido a erro do sistema (status ERRO). Fechando modal...")
        status = "ERRO"

    try:
        botao_fechar = page.locator(
            "body > div.fade.modal.show > div > div > div.modal-body "
            "> form > div.row > div > div > button:nth-child(2)"
        )
        botao_fechar.click()
        page.wait_for_timeout(200)
    except Exception:
        try:
            botao_x = page.locator(
                "body > div.fade.modal.show > div > div > div.modal-header > button > span:nth-child(1)"
            )
            botao_x.click()
            page.wait_for_timeout(200)
        except Exception:
            print("[ALERTA] Não consegui fechar o modal de abono após erro.")

    registrar_log_csv(justificativa, "criar_abono", status, f"Popup de erro: {mensagem}")
    return status


def voltar_da_tela_espelho_para_lista(page: Page) -> None:
    """
    Volta do cartão de ponto para a tela de lista.

    Tenta, nesta ordem:
    - #btn-buscar-crud
    - botão 'Voltar' visível
    - botão 'Buscar' visível dentro do conteúdo (mas não links de menu)

    Se nada funcionar, registra erro e deixa para você voltar manualmente.
    """
    print("[INFO] Voltando para tela de lista (tela de espelho)...")

    try:
        candidatos = []

        # id legado
        candidatos.append(page.locator("#btn-buscar-crud"))

        area_conteudo = page.locator("div.content")

        # botões reais
        candidatos.append(area_conteudo.get_by_role("button", name="Voltar"))
        candidatos.append(area_conteudo.get_by_role("button", name="Buscar"))

        botao_validado = None
        for loc in candidatos:
            try:
                if loc.count() == 0:
                    continue
                candidato = loc.first
                if not candidato.is_visible():
                    continue
                botao_validado = candidato
                break
            except Exception:
                continue

        if not botao_validado:
            raise RuntimeError("Nenhum botão conhecido de voltar/buscar foi encontrado no cartão de ponto.")

        botao_validado.scroll_into_view_if_needed()
        page.wait_for_timeout(200)
        botao_validado.click(force=True)
        time.sleep(1)
    except Exception as e:
        print("[ALERTA] Nao consegui clicar em nenhum botão de voltar/buscar na tela de espelho.")
        print(f"[ALERTA] Erro: {e}")
        print("Se necessário, volte manualmente para a LISTA pelo sistema.")
        registrar_log_sistema(
            "voltar_lista",
            "ERRO",
            f"Falha ao voltar para lista automaticamente: {e}",
        )

    # depois do clique (ou da intervenção manual), valida/ajusta a lista
    garantir_tela_lista(page)



# ============================================================
# REGRA DE NEGÓCIO / SERVIÇO
# ============================================================

def deve_processar(j: Justificativa) -> bool:
    """
    Decide se uma justificativa deve ser processada.

    Regras:
    - Status ""        -> processa (primeira vez).
    - Status "OK"      -> não processa (já abonado).
    - Status "CONFLITO"-> não processa (já detectado conflito).
    - Qualquer outro   -> processa novamente.
    """
    status_norm = (j.status or "").strip().upper()
    if status_norm in ("OK", "CONFLITO"):
        return False
    return True


def processar_uma_justificativa(
    page: Page,
    justificativa: Justificativa,
    cartao_aberto_ok: bool,
) -> Tuple[bool, str]:
    """
    Aplica uma justificativa no cartão (se ele estiver aberto) e devolve:

    (cartao_aberto_ainda_valido, novo_status_da_justificativa)
    """
    if not cartao_aberto_ok:
        msg = "Não foi possível abrir o cartão de ponto para este funcionário; pulando criação de abono."
        print("[ALERTA]", msg)
        registrar_log_csv(justificativa, "criar_abono", "ERRO", msg)
        return cartao_aberto_ok, "ERRO"

    status = criar_abono_no_cartao(page, justificativa)
    return cartao_aberto_ok, status


# ============================================================
# ORQUESTRAÇÃO PRINCIPAL
# ============================================================

def executar_processamento():
    """
    Orquestra o fluxo completo.
    """
    df = carregar_justificativas(CAMINHO_CSV_JUSTIFICATIVAS)
    print(f"[INFO] Total de linhas no CSV (original): {len(df)}")

    df = ordenar_justificativas(df)
    print(f"[INFO] Total de linhas no CSV (ordenado): {len(df)}")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            PERFIL_BROWSER,
            headless=False,
            slow_mo=350,
            viewport={"width": 1366, "height": 768},
        )

        page = context.new_page()
        page.goto(URL_PONTO)

        print("=" * 60)
        print("Se for a primeira vez neste perfil, faça login agora no RH247.")
        print("Depois de logado, vá até a LISTA de funcionários, com campo 'Pesquisar...'.")
        print("Nas próximas execuções, o login deve ser reaproveitado automaticamente.")
        print("=" * 60)
        input("Quando estiver na tela de LISTA de funcionários, pressione ENTER aqui no terminal...")

        garantir_tela_lista(page)

        nome_anterior_norm = None
        cartao_aberto_ok = False

        for idx in range(len(df)):
            justificativa = linha_para_justificativa(df, idx)
            nome_atual_norm = normalizar_nome(justificativa.nome_completo)

            print("\n" + "=" * 40)
            print(f"[LOOP] Linha {idx} - Funcionário: '{justificativa.nome_completo}'")

            if not deve_processar(justificativa):
                print(f"[INFO] Linha {idx} já processada anteriormente com status '{justificativa.status}'. Pulando.")
                continue

            registrar_log_csv(
                justificativa,
                "preparar_abono",
                "INFO",
                f"Iniciando processamento da linha {idx} para '{justificativa.nome_completo}'.",
            )

            try:
                # se mudou de funcionário, voltar para lista e abrir novo cartão
                if nome_atual_norm != nome_anterior_norm:
                    cartao_aberto_ok = False
                    buscar_funcionario_na_lista(page, justificativa.nome_completo)
                    abrir_cartao_ponto_funcionario(page, justificativa.nome_completo)
                    conferir_nome_no_espelho(page, justificativa.nome_completo)
                    registrar_log_csv(justificativa, "conferir_nome", "OK", "Nome na tela confere com o CSV.")
                    cartao_aberto_ok = True

                cartao_aberto_ok, novo_status = processar_uma_justificativa(
                    page,
                    justificativa,
                    cartao_aberto_ok,
                )
                df.at[justificativa.indice_df, "Status"] = novo_status

            except Exception as e:
                print(f"[ERRO] Problema ao processar linha {idx}: {e}")
                registrar_log_csv(justificativa, "erro_geral", "ERRO", str(e))
                df.at[justificativa.indice_df, "Status"] = "ERRO"

            # Verifica o próximo funcionário
            try:
                proxima_just = linha_para_justificativa(df, idx + 1)
                nome_proximo_norm = normalizar_nome(proxima_just.nome_completo)
            except IndexError:
                proxima_just = None
                nome_proximo_norm = None

            # mudança de funcionário ou fim de arquivo -> voltar para lista e salvar CSV parcial
            if nome_proximo_norm is None or nome_proximo_norm != nome_atual_norm:
                try:
                    voltar_da_tela_espelho_para_lista(page)
                    registrar_log_csv(
                        justificativa,
                        "voltar_lista",
                        "INFO",
                        "Voltando para lista para próximo funcionário.",
                    )
                except Exception as e:
                    registrar_log_csv(
                        justificativa,
                        "voltar_lista",
                        "ERRO",
                        f"Falha ao voltar para lista: {e}",
                    )

                df.to_csv(
                    CAMINHO_CSV_JUSTIFICATIVAS,
                    index=False,
                    encoding="latin1",
                )

            nome_anterior_norm = nome_atual_norm

        print("\n[INFO] Processamento concluído.")
        print("Verifique o arquivo de logs:", CAMINHO_CSV_LOGS)
        print("Verifique também o arquivo de justificativas atualizado:", CAMINHO_CSV_JUSTIFICATIVAS)
        input("Pressione ENTER para encerrar o script (o navegador pode ser fechado manualmente depois)...")
        context.close()


if __name__ == "__main__":
    executar_processamento()
