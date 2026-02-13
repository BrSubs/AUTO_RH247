import csv
import datetime
import os
from pathlib import Path

import pandas as pd
from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)

# ==============================
# CONFIGURAÇÕES BÁSICAS
# ==============================

CAMINHO_CSV_JUSTIFICATIVAS = "justificativas.csv"
CAMINHO_CSV_LOGS = "logs_justificativas.csv"

TIMEOUT_CURTO = 5000      # 5s
TIMEOUT_PADRAO = 15000    # 15s


# ==============================
# FUNÇÕES DE APOIO
# ==============================

def normalizar_nome(nome: str) -> str:
    if not isinstance(nome, str):
        return ""
    return " ".join(nome.strip().upper().split())


def parse_data_br(data_str: str):
    """
    Converte datas do formato dd/mm/aa ou dd/mm/aaaa em datetime.date.
    Se não conseguir converter, retorna None.
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


def registrar_log(linha, acao, status, detalhe):
    """
    Escreve uma linha no arquivo de log CSV.
    linha: pandas Series com campos Nome_Completo, Data_Inicio, Data_Fim, Motivo
    """
    Path(CAMINHO_CSV_LOGS).touch(exist_ok=True)
    arquivo_existe = os.path.getsize(CAMINHO_CSV_LOGS) > 0

    with open(CAMINHO_CSV_LOGS, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not arquivo_existe:
            writer.writerow([
                "timestamp",
                "Nome_Completo",
                "Data_Inicio",
                "Data_Fim",
                "Motivo",
                "acao",
                "status",
                "detalhe",
            ])

        agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        nome = linha["Nome_Completo"]
        data_ini = linha["Data_Inicio"]
        data_fim = linha["Data_Fim"]
        motivo = linha["Motivo"]

        writer.writerow([
            agora,
            nome,
            data_ini,
            data_fim,
            motivo,
            acao,
            status,
            detalhe,
        ])


def ler_justificativas(caminho_csv):
    # encoding latin1 para CSV do Windows
    df = pd.read_csv(caminho_csv, encoding="latin1")
    df.columns = [c.strip() for c in df.columns]

    # Renomeia se vier como "Nome Completo"
    if "Nome Completo" in df.columns and "Nome_Completo" not in df.columns:
        df = df.rename(columns={"Nome Completo": "Nome_Completo"})

    for col in ["Nome_Completo", "Data_Inicio", "Data_Fim", "Motivo"]:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatória ausente no CSV: {col}")
    return df


# ==============================
# FUNÇÕES DE NAVEGAÇÃO
# ==============================

def esperar_tela_lista(page):
    """
    Verifica/espera que a tela de lista de funcionários esteja carregada.
    """
    print("[INFO] Verificando se a tela de lista está carregada...")
    try:
        campo_busca = page.locator('xpath=//*[@id="Dr"]/div[3]/div/div[2]/div[2]/div/input')
        campo_busca.wait_for(state="visible", timeout=TIMEOUT_PADRAO * 2)
        print("[OK] Tela de lista detectada (campo de busca visível).")
    except PlaywrightTimeoutError:
        print("[ALERTA] Não encontrei o campo de busca da tela de lista.")
        print("Confira se você está na tela correta do RH247 nesta aba (lista de funcionários).")
        input("Ajuste a tela no navegador e pressione ENTER para tentar continuar...")


def buscar_funcionario_por_nome(page, nome):
    """
    Preenche o campo de busca com o nome e clica em pesquisar.
    """
    print(f"[INFO] Buscando funcionário pelo nome: '{nome}'")
    try:
        campo_busca = page.locator('xpath=//*[@id="Dr"]/div[3]/div/div[2]/div[2]/div/input')
        campo_busca.wait_for(state="visible", timeout=TIMEOUT_PADRAO * 2)
        campo_busca.click()
        campo_busca.fill("")
        campo_busca.fill(nome)

        # Pausa para você ver o nome sendo preenchido
        page.wait_for_timeout(1000)

        botao_pesquisar = page.locator('xpath=//*[@id="Dr"]/div[3]/div/div[2]/div[2]/div/div/button/i')
        botao_pesquisar.wait_for(state="visible", timeout=TIMEOUT_PADRAO * 2)
        botao_pesquisar.click()

        # Pausa para a tabela atualizar
        page.wait_for_timeout(2000)
    except PlaywrightTimeoutError:
        raise RuntimeError("Timeout ao tentar usar campo de busca ou botão pesquisar.")


def debug_listar_nomes_tabela(page):
    """
    Lista no terminal os nomes que estão na tabela de resultados após a busca.
    """
    print("[DEBUG] Lendo linhas da tabela de resultados...")
    try:
        linhas = page.locator('xpath=//*[@id="Dr"]/div[3]/div/table/tbody/tr')
        linhas.first.wait_for(state="visible", timeout=TIMEOUT_PADRAO)
    except PlaywrightTimeoutError:
        print("[DEBUG] Nenhuma linha visível na tabela (timeout).")
        return

    qtd_linhas = linhas.count()
    print(f"[DEBUG] Quantidade de linhas na tabela: {qtd_linhas}")

    for i in range(qtd_linhas):
        idx = i + 1
        xpath_nome = f'//*[@id="Dr"]/div[3]/div/table/tbody/tr[{idx}]/td[2]//span'
        celula_nome = page.locator(f'xpath={xpath_nome}')
        try:
            texto_nome = celula_nome.inner_text(timeout=TIMEOUT_CURTO)
            print(f"[DEBUG] Linha {idx} - Nome: '{texto_nome}'")
        except PlaywrightTimeoutError:
            print(f"[DEBUG] Linha {idx} - Nome: (timeout)")


def abrir_ponto_funcionario_por_nome(page, nome_busca):
    """
    Na tela de lista, encontra a linha cujo nome da coluna 2 corresponde a nome_busca
    e clica no botão Editar (coluna 12) dessa mesma linha.
    """
    nome_normalizado = normalizar_nome(nome_busca)

    try:
        linhas = page.locator('xpath=//*[@id="Dr"]/div[3]/div/table/tbody/tr')
        linhas.first.wait_for(state="visible", timeout=TIMEOUT_PADRAO)
    except PlaywrightTimeoutError:
        raise RuntimeError("Timeout ao aguardar a tabela de resultados.")

    qtd_linhas = linhas.count()
    if qtd_linhas == 0:
        raise RuntimeError("Nenhuma linha na tabela de resultados após a busca.")

    print(f"[INFO] Procurando '{nome_busca}' na tabela ({qtd_linhas} linhas)...")

    for i in range(qtd_linhas):
        idx = i + 1
        xpath_nome = f'//*[@id="Dr"]/div[3]/div/table/tbody/tr[{idx}]/td[2]//span'
        celula_nome = page.locator(f'xpath={xpath_nome}')
        try:
            texto_nome = celula_nome.inner_text(timeout=TIMEOUT_CURTO)
        except PlaywrightTimeoutError:
            print(f"[DEBUG] Linha {idx} - não consegui ler o nome (timeout).")
            continue

        print(f"[DEBUG] Linha {idx} - nome lido: '{texto_nome}'")

        if normalizar_nome(texto_nome) == nome_normalizado:
            print(f"[OK] Encontrado '{texto_nome}' na linha {idx}, clicando em Editar...")
            xpath_editar = f'//*[@id="Dr"]/div[3]/div/table/tbody/tr[{idx}]/td[12]/button/i'
            botao_editar = page.locator(f'xpath={xpath_editar}')
            botao_editar.click()
            page.wait_for_timeout(2000)
            return

    raise RuntimeError(f"Funcionário com nome '{nome_busca}' não encontrado na tabela.")


def conferir_nome_na_tela(page, nome_csv):
    """
    Lê o nome em <strong> na tela do espelho e compara com o nome do CSV.
    """
    nome_normalizado_csv = normalizar_nome(nome_csv)
    x_nome_strong = 'xpath=//*[@id="Dr"]/div/div/div[3]/div/div/div/div/div/div/div[1]/div/strong'

    try:
        strong_nome = page.locator(x_nome_strong)
        strong_nome.wait_for(state="visible", timeout=TIMEOUT_PADRAO)
        texto_tela = strong_nome.inner_text(timeout=TIMEOUT_CURTO)
    except PlaywrightTimeoutError:
        raise RuntimeError("Não foi possível ler o nome do funcionário na tela de espelho.")

    print(f"[INFO] Nome na tela: '{texto_tela}' | Nome do CSV: '{nome_csv}'")

    if normalizar_nome(texto_tela) != nome_normalizado_csv:
        raise RuntimeError(f"Nome na tela '{texto_tela}' difere do CSV '{nome_csv}'.")


def criar_abono(page, linha):
    """
    Cria um abono na tela do espelho para a linha de justificativa informada.
    Garante que as datas sejam preenchidas no formato dd/mm/aaaa.
    """
    nome = linha["Nome_Completo"]
    data_ini_str_original = str(linha["Data_Inicio"])
    data_fim_str_original = str(linha["Data_Fim"])
    motivo = str(linha["Motivo"])

    print(f"[INFO] Criando abono para {nome} | {data_ini_str_original} -> {data_fim_str_original} | Motivo: {motivo}")

    # Validar e normalizar datas
    data_ini = parse_data_br(data_ini_str_original)
    data_fim = parse_data_br(data_fim_str_original)

    if data_ini is None or data_fim is None:
        raise RuntimeError(
            f"Datas inválidas: Data_Inicio='{data_ini_str_original}', Data_Fim='{data_fim_str_original}'."
        )

    if data_fim < data_ini:
        raise RuntimeError(
            f"Data_Fim menor que Data_Inicio: {data_ini_str_original} > {data_fim_str_original}."
        )

    # Formata sempre como dd/mm/aaaa
    data_ini_str = data_ini.strftime("%d/%m/%Y")
    data_fim_str = data_fim.strftime("%d/%m/%Y")

    # Clicar no botão "Abono de Ponto"
    try:
        print("[INFO] Procurando botão 'Abono de Ponto'...")
        try:
            botao_abono = page.get_by_role("button", name="Abono de Ponto")
            botao_abono.wait_for(state="visible", timeout=TIMEOUT_CURTO)
        except PlaywrightTimeoutError:
            botao_abono = page.locator("button:has-text('Abono de Ponto')")
            botao_abono.wait_for(state="visible", timeout=TIMEOUT_PADRAO)

        botao_abono.click()
        page.wait_for_timeout(1000)
    except PlaywrightTimeoutError:
        raise RuntimeError("Botão 'Abono de Ponto' não encontrado ou não visível.")

    # Preencher campos de abono
    try:
        campo_descricao = page.locator("#descricao")
        campo_data_ini = page.locator("#data_ini")
        campo_data_fim = page.locator("#data_fim")

        campo_descricao.wait_for(state="visible", timeout=TIMEOUT_PADRAO)
        campo_data_ini.wait_for(state="visible", timeout=TIMEOUT_PADRAO)
        campo_data_fim.wait_for(state="visible", timeout=TIMEOUT_PADRAO)

        campo_descricao.fill("")
        campo_descricao.fill(motivo)

        campo_data_ini.fill("")
        campo_data_ini.fill(data_ini_str)

        campo_data_fim.fill("")
        campo_data_fim.fill(data_fim_str)

        botao_salvar = page.get_by_role("button", name="Salvar")
        botao_salvar.click()

    except PlaywrightTimeoutError:
        raise RuntimeError("Erro ao preencher campos de abono ou clicar em 'Salvar'.")

    # Espera o sistema processar
    page.wait_for_timeout(2000)

    # Verifica se apareceu popup de erro (ex.: período conflitante)
    if tratar_popup_erro(page, linha):
        print("[INFO] Abono não criado devido ao erro informado pelo sistema.")
        return




def tratar_popup_erro(page, linha):
    """
    Verifica se há um popup de erro (SweetAlert2) aberto.
    Se houver, lê a mensagem, clica em OK e registra no log como ERRO.
    Retorna True se tratou um erro, False se não havia popup.
    """
    try:
        # Localiza o container principal do popup de erro (swal2-popup com ícone de erro)
        popup = page.locator("div.swal2-popup.swal2-modal.swal2-icon-error.swal2-show")
        if not popup.is_visible(timeout=1000):
            return False
    except PlaywrightTimeoutError:
        return False
    except Exception:
        return False

    # Tenta ler o título e o conteúdo
    try:
        titulo = page.locator("#swal2-title").inner_text(timeout=TIMEOUT_CURTO)
    except Exception:
        titulo = ""

    try:
        conteudo = page.locator("#swal2-content").inner_text(timeout=TIMEOUT_CURTO)
    except Exception:
        conteudo = ""

    mensagem = f"{titulo} - {conteudo}".strip()
    if not mensagem:
        mensagem = "Popup de erro sem mensagem legível."

    print(f"[ALERTA] Popup de erro detectado: {mensagem}")

    # Clica no botão OK do popup
    try:
        botao_ok = page.locator("button.swal2-confirm")
        botao_ok.click()
        page.wait_for_timeout(1000)
    except Exception:
        print("[ALERTA] Não consegui clicar no botão OK do popup.")

    # Registra no log que houve erro para essa linha
    registrar_log(linha, "popup_erro", "ERRO", mensagem)

    return True



def voltar_para_lista(page):
    """
    Clica no botão Buscar (#btn-buscar-crud) para voltar/recarregar a lista.
    """
    print("[INFO] Voltando para tela de lista (botão Buscar)...")
    try:
        botao_buscar = page.locator("#btn-buscar-crud")
        botao_buscar.wait_for(state="visible", timeout=TIMEOUT_PADRAO)
        botao_buscar.click()
        page.wait_for_timeout(2000)
    except PlaywrightTimeoutError:
        print("[ALERTA] Não encontrei o botão Buscar; tentando go_back()...")
        try:
            page.go_back()
            page.wait_for_timeout(2000)
        except Exception:
            print("[ERRO] Falha ao voltar com go_back().")


# ==============================
# FUNÇÃO PRINCIPAL
# ==============================

def main():
    justificativas = ler_justificativas(CAMINHO_CSV_JUSTIFICATIVAS)
    print(f"[INFO] Total de linhas no CSV: {len(justificativas)}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.set_viewport_size({"width": 1366, "height": 768})

        print("=" * 60)
        print("Abra o RH247 NESTA ABA, faça login e deixe na LISTA de funcionários.")
        print("Certifique-se de:")
        print("- Campo 'Buscar por' esteja em NOME.")
        print("- Tabela de funcionários apareça com coluna de Nome e botão Editar.")
        print("=" * 60)
        input("Quando estiver pronto nessa tela, pressione ENTER aqui no terminal...")

        esperar_tela_lista(page)

        nome_anterior = None

        for idx, linha in justificativas.iterrows():
            nome_atual = str(linha["Nome_Completo"])
            print("\n" + "=" * 40)
            print(f"[LOOP] Linha {idx} - Funcionário: '{nome_atual}'")
            registrar_log(linha, "preparar_abono", "INFO",
                          f"Iniciando processamento da linha {idx} para '{nome_atual}'.")

            try:
                # Se for um novo funcionário (nome diferente do anterior)
                if normalizar_nome(nome_atual) != normalizar_nome(nome_anterior):
                    # PRIMEIRA VEZ desse funcionário:
                    # Pressupomos que já estamos na tela de lista correta.
                    # Só vamos voltar_para_lista nas próximas vezes (quando trocar de funcionário).

                    if nome_anterior is not None:
                        # Só tenta voltar se NÃO for a primeira linha de todo o CSV
                        voltar_para_lista(page)

                    # Buscar funcionário
                    buscar_funcionario_por_nome(page, nome_atual)


                    # Debug opcional (pode comentar depois de confiar)
                    debug_listar_nomes_tabela(page)

                    # Abrir espelho
                    abrir_ponto_funcionario_por_nome(page, nome_atual)

                    # Conferir nome na tela
                    conferir_nome_na_tela(page, nome_atual)
                    registrar_log(linha, "conferir_nome", "OK", "Nome na tela confere com o CSV.")

                # Criar abono para esta linha
                criar_abono(page, linha)
                registrar_log(linha, "criar_abono", "OK", "Abono criado com sucesso.")

            except Exception as e:
                print(f"[ERRO] Problema ao processar linha {idx}: {e}")
                registrar_log(linha, "erro_geral", "ERRO", str(e))

            # Ver próxima linha para decidir se muda de funcionário
            try:
                proxima_linha = justificativas.iloc[idx + 1]
                nome_proximo = str(proxima_linha["Nome_Completo"])
            except IndexError:
                proxima_linha = None
                nome_proximo = None

            if nome_proximo is None or normalizar_nome(nome_proximo) != normalizar_nome(nome_atual):
                # Vai mudar de funcionário -> volta para lista
                try:
                    voltar_para_lista(page)
                    registrar_log(linha, "voltar_lista", "INFO",
                                  "Voltando para lista para próximo funcionário.")
                except Exception as e:
                    registrar_log(linha, "voltar_lista", "ERRO",
                                  f"Falha ao voltar para lista: {e}")

            nome_anterior = nome_atual

        print("\n[INFO] Processamento concluído.")
        print("Verifique o arquivo de logs:", CAMINHO_CSV_LOGS)
        input("Pressione ENTER para fechar o navegador...")
        browser.close()


if __name__ == "__main__":
    main()
