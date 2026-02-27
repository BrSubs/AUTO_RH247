import csv
import datetime
import os
import time
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

PERFIL_BROWSER = "perfil_rh247"

TIMEOUT_CURTO = 5000
TIMEOUT_PADRAO = 15000


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


def registrar_log_basico(nome, etapa, status, mensagem):
    agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(CAMINHO_CSV_LOGS, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([agora, nome, "", "", "", etapa, status, mensagem])


def ler_justificativas(caminho_csv):
    df = pd.read_csv(caminho_csv, encoding="latin1")
    df.columns = [c.strip() for c in df.columns]

    if "Nome Completo" in df.columns and "Nome_Completo" not in df.columns:
        df = df.rename(columns={"Nome Completo": "Nome_Completo"})

    for col in ["Nome_Completo", "Data_Inicio", "Data_Fim", "Motivo", "Status"]:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatória ausente no CSV: {col}")

    df["Status"] = df["Status"].fillna("")

    return df


# ==============================
# POPUP DE ERRO (SWEETALERT2)
# ==============================

def tratar_popup_erro(page, linha):
    """
    Verifica se há um popup de erro (SweetAlert2) aberto.
    Se houver, lê a mensagem, clica em OK e registra no log como ERRO.
    Retorna True se tratou um erro, False se não havia popup.
    """
    try:
        popup = page.locator("div.swal2-popup.swal2-modal.swal2-icon-error.swal2-show")
        if not popup.is_visible(timeout=1000):
            return False
    except PlaywrightTimeoutError:
        return False
    except Exception:
        return False

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

    try:
        botao_ok = page.locator("button.swal2-confirm")
        botao_ok.scroll_into_view_if_needed()
        page.wait_for_timeout(500)
        botao_ok.click()
        page.wait_for_timeout(1000)
    except Exception:
        print("[ALERTA] Não consegui clicar no botão OK do popup.")

    registrar_log(linha, "popup_erro", "ERRO", mensagem)
    return True


# ==============================
# FUNÇÕES DE NAVEGAÇÃO
# ==============================

def esperar_tela_lista(page):
    print("[INFO] Verificando se a tela de lista está carregada...")
    registrar_log_basico(
        "sistema",
        "esperar_tela_lista",
        "INFO",
        f"Verificando tela de lista. URL='{page.url}', titulo='{page.title()}'"
    )

    try:
        campo_busca = page.wait_for_selector(
            'xpath=//*[@id="Fr"]/div[3]/div/div[2]/div[2]/div/input',
            timeout=TIMEOUT_PADRAO
        )
        print("[INFO] Campo de busca da lista encontrado.")
        registrar_log_basico(
            "sistema",
            "esperar_tela_lista",
            "OK",
            "Campo de busca encontrado na tela de lista."
        )
        return campo_busca
    except Exception as e:
        print("[ALERTA] Não encontrei o campo de busca da tela de lista.")
        print("Confira se você está na tela correta do RH247 nesta aba (lista de funcionários).")
        print("Ajuste a tela no navegador e pressione ENTER para tentar continuar...")
        registrar_log_basico(
            "sistema",
            "esperar_tela_lista",
            "ERRO",
            f"Nao encontrou campo de busca da lista. URL='{page.url}', titulo='{page.title()}', erro='{e}'"
        )
        input()
        return None


def buscar_funcionario_por_nome(page, nome):
    nome_normalizado = normalizar_nome(nome)
    print(f"[INFO] Buscando funcionário pelo nome: '{nome_normalizado}'")
    registrar_log_basico(
        nome_normalizado,
        "buscar_funcionario",
        "INFO",
        f"Iniciando busca do funcionario. URL='{page.url}', titulo='{page.title()}'"
    )

    try:
        campo_busca = page.wait_for_selector(
            'css=div.container-fluid:nth-child(1) > div:nth-child(3) > '
            'div:nth-child(2) > div:nth-child(1) > input:nth-child(1)',
            timeout=TIMEOUT_PADRAO
        )
    except Exception as e:
        registrar_log_basico(
            nome_normalizado,
            "buscar_funcionario",
            "ERRO",
            f"Nao encontrou barra de pesquisa. URL='{page.url}', titulo='{page.title()}', erro='{e}'"
        )
        raise

    campo_busca.fill("")
    campo_busca.fill(nome_normalizado)
    campo_busca.press("Enter")

    linhas = page.locator(
        "#Fr > div.content.bg-white.card.crud-content.aos-init.aos-animate > "
        "div > table > tbody > tr"
    )

    try:
        linhas.first.wait_for(state="visible", timeout=TIMEOUT_PADRAO)
    except PlaywrightTimeoutError:
        raise RuntimeError("Timeout ao aguardar a tabela de resultados apos a busca.")

    deadline = time.time() + 5
    while time.time() < deadline:
        total = linhas.count()
        for i in range(total):
            try:
                celula_nome = linhas.nth(i).locator("td:nth-child(2) span")
                texto = normalizar_nome(celula_nome.inner_text())
                if texto == nome_normalizado:
                    print(f"[INFO] Tabela atualizada, funcionário '{nome_normalizado}' encontrado na linha {i}.")
                    return
            except Exception:
                pass
        page.wait_for_timeout(200)

    print(f"[ALERTA] Tabela não exibiu '{nome_normalizado}' após a busca; seguindo assim mesmo.")


def debug_listar_nomes_tabela(page):
    print("[DEBUG] Lendo linhas da tabela de resultados...")

    linhas = page.locator(
        "#Fr > div.content.bg-white.card.crud-content.aos-init.aos-animate > "
        "div > table > tbody > tr"
    )
    try:
        total = linhas.count()
    except Exception as e:
        print(f"[DEBUG] Erro ao contar linhas da tabela: {e}")
        return

    if total == 0:
        print("[DEBUG] Nenhuma linha na tabela (count == 0).")
        return

    print(f"[DEBUG] Total de linhas na tabela: {total}")
    for i in range(total):
        try:
            celula_nome = linhas.nth(i).locator("td:nth-child(2) span")
            texto = celula_nome.inner_text().strip()
            print(f"[DEBUG] Linha {i}: '{texto}'")
        except Exception as e:
            print(f"[DEBUG] Falha ao ler nome da linha {i}: {e}")


def abrir_ponto_funcionario_por_nome(page, nome_alvo):
    nome_alvo_normalizado = normalizar_nome(nome_alvo)
    print(f"[INFO] Procurando na tabela o funcionário '{nome_alvo_normalizado}' para abrir o ponto...")

    linhas = page.locator(
        "#Fr > div.content.bg-white.card.crud-content.aos-init.aos-animate > "
        "div > table > tbody > tr"
    )

    total = linhas.count()
    if total == 0:
        raise Exception("Tabela de resultados vazia ao tentar abrir ponto do funcionario.")

    for i in range(total):
        linha = linhas.nth(i)
        celula_nome = linha.locator("td:nth-child(2) span")
        texto = normalizar_nome(celula_nome.inner_text())
        print(f"[DEBUG] Comparando linha {i}: '{texto}'")

        if texto == nome_alvo_normalizado:
            print(f"[INFO] Encontrado na linha {i}, clicando no botão Editar...")
            botao_editar = linha.locator("td:nth-child(12) button.btn-grid-edit")
            botao_editar.click()
            return

    raise Exception(f"Nenhuma linha da tabela corresponde exatamente a '{nome_alvo_normalizado}'.")


def conferir_nome_na_tela(page, nome_esperado):
    nome_esperado_norm = normalizar_nome(nome_esperado)
    print(f"[INFO] Conferindo nome na tela de espelho (esperado: '{nome_esperado_norm}')")

    try:
        elem_nome = page.wait_for_selector(
            "#Fr > div > div > div:nth-child(4) > div > div > div > "
            "div > div > div > div:nth-child(1) > div > strong",
            timeout=TIMEOUT_PADRAO
        )
        texto_lido = normalizar_nome(elem_nome.inner_text())
        print(f"[DEBUG] Nome lido na tela de espelho: '{texto_lido}'")
    except Exception as e:
        raise Exception(
            f"Nao foi possivel ler o nome do funcionario na tela de espelho "
            f"(elemento nao encontrado): {e}"
        )

    if texto_lido != nome_esperado_norm:
        raise Exception(
            f"Nome na tela de espelho diferente do CSV. "
            f"Esperado='{nome_esperado_norm}', lido='{texto_lido}'"
        )


def criar_abono(page, linha):
    """
    Cria um abono na tela do espelho para a linha de justificativa informada.
    Retorna True se considerarmos que o abono foi criado com sucesso,
    False se houve popup de erro do sistema.
    """
    nome = linha["Nome_Completo"]
    data_ini_str_original = str(linha["Data_Inicio"])
    data_fim_str_original = str(linha["Data_Fim"])
    motivo = str(linha["Motivo"])

    print(f"[INFO] Criando abono para {nome} | {data_ini_str_original} -> {data_fim_str_original} | Motivo: {motivo}")

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

    data_ini_str = data_ini.strftime("%d/%m/%Y")
    data_fim_str = data_fim.strftime("%d/%m/%Y")

    try:
        print("[INFO] Procurando botão 'Abono de Ponto'...")
        try:
            botao_abono = page.get_by_role("button", name="Abono de Ponto")
            botao_abono.wait_for(state="visible", timeout=TIMEOUT_CURTO)
        except PlaywrightTimeoutError:
            botao_abono = page.locator("button:has-text('Abono de Ponto')")
            botao_abono.wait_for(state="visible", timeout=TIMEOUT_PADRAO)

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

        campo_descricao.wait_for(state="visible", timeout=TIMEOUT_PADRAO)
        campo_data_ini.wait_for(state="visible", timeout=TIMEOUT_PADRAO)
        campo_data_fim.wait_for(state="visible", timeout=TIMEOUT_PADRAO)

        campo_descricao.fill(motivo)
        campo_data_ini.fill(data_ini_str)
        campo_data_fim.fill(data_fim_str)

        botao_salvar = page.get_by_role("button", name="Salvar")
        botao_salvar.click()
    except PlaywrightTimeoutError:
        raise RuntimeError("Erro ao preencher campos de abono ou clicar em 'Salvar'.")

    page.wait_for_timeout(2000)

    if tratar_popup_erro(page, linha):
        print("[INFO] Abono não criado devido ao erro informado pelo sistema. Fechando modal de abono...")

        try:
            botao_fechar = page.locator(
                "body > div.fade.modal.show > div > div > div.modal-body > "
                "form > div.row > div > div > button:nth-child(2)"
            )
            botao_fechar.click()
            page.wait_for_timeout(200)
            return False
        except Exception:
            pass

        try:
            botao_x = page.locator(
                "body > div.fade.modal.show > div > div > div.modal-header > button > span:nth-child(1)"
            )
            botao_x.click()
            page.wait_for_timeout(200)
        except Exception:
            print("[ALERTA] Não consegui fechar o modal de abono após erro.")
        return False

    return True


def voltar_para_lista(page):
    print("[INFO] Voltando para tela de lista (tela de espelho)...")

    try:
        botao_buscar = page.wait_for_selector(
            "#btn-buscar-crud",
            timeout=TIMEOUT_CURTO
        )
        print("[INFO] Botão Buscar (btn-buscar-crud) encontrado no cartão de ponto, clicando...")
        botao_buscar.click()
        time.sleep(1)
    except Exception as e:
        raise Exception(
            f"Nao consegui voltar para a tela de lista a partir do espelho: "
            f"botao '#btn-buscar-crud' nao encontrado. Erro: {e}"
        )

    esperar_tela_lista(page)


# ==============================
# FUNÇÃO PRINCIPAL
# ==============================

def main():
    justificativas = ler_justificativas(CAMINHO_CSV_JUSTIFICATIVAS)
    print(f"[INFO] Total de linhas no CSV (original): {len(justificativas)}")

    justificativas["Data_Ordenacao"] = justificativas["Data_Inicio"].apply(parse_data_br)
    justificativas = justificativas.sort_values(
        by=["Nome_Completo", "Data_Ordenacao"],
        kind="stable"
    ).drop(columns=["Data_Ordenacao"]).reset_index(drop=True)

    print(f"[INFO] Total de linhas no CSV (ordenado): {len(justificativas)}")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            PERFIL_BROWSER,
            headless=False,
            slow_mo=500,
            viewport={"width": 1366, "height": 768},
        )

        page = context.new_page()
        page.goto("https://rh247.com.br/230540701/ponto")

        print("=" * 60)
        print("Se for a primeira vez neste perfil, faça login agora no RH247.")
        print("Depois de logado, vá até a LISTA de funcionários, com 'Buscar por' = NOME.")
        print("Nas próximas execuções, o login deve ser reaproveitado automaticamente.")
        print("=" * 60)
        input("Quando estiver na tela de LISTA de funcionários, pressione ENTER aqui no terminal...")

        esperar_tela_lista(page)

        nome_anterior = None

        for idx, linha in justificativas.iterrows():
            nome_atual = str(linha["Nome_Completo"])
            print("\n" + "=" * 40)
            print(f"[LOOP] Linha {idx} - Funcionário: '{nome_atual}'")
            registrar_log(
                linha,
                "preparar_abono",
                "INFO",
                f"Iniciando processamento da linha {idx} para '{nome_atual}'."
            )

            try:
                if normalizar_nome(nome_atual) != normalizar_nome(nome_anterior):
                    buscar_funcionario_por_nome(page, nome_atual)
                    debug_listar_nomes_tabela(page)
                    abrir_ponto_funcionario_por_nome(page, nome_atual)
                    conferir_nome_na_tela(page, nome_atual)
                    registrar_log(linha, "conferir_nome", "OK", "Nome na tela confere com o CSV.")

                sucesso = criar_abono(page, linha)
                if sucesso:
                    registrar_log(
                        linha,
                        "criar_abono",
                        "OK",
                        "Abono criado sem erro de sistema."
                    )
                    justificativas.loc[idx, "Status"] = "OK"
                else:
                    registrar_log(
                        linha,
                        "criar_abono",
                        "ERRO",
                        "Sistema exibiu popup de erro ao tentar criar abono."
                    )
                    justificativas.loc[idx, "Status"] = "ERRO"

            except Exception as e:
                print(f"[ERRO] Problema ao processar linha {idx}: {e}")
                registrar_log(linha, "erro_geral", "ERRO", str(e))
                justificativas.loc[idx, "Status"] = "ERRO"

            try:
                proxima_linha = justificativas.iloc[idx + 1]
                nome_proximo = str(proxima_linha["Nome_Completo"])
            except IndexError:
                proxima_linha = None
                nome_proximo = None

            if nome_proximo is None or normalizar_nome(nome_proximo) != normalizar_nome(nome_atual):
                try:
                    voltar_para_lista(page)
                    registrar_log(
                        linha,
                        "voltar_lista",
                        "INFO",
                        "Voltando para lista para próximo funcionário."
                    )
                except Exception as e:
                    registrar_log(
                        linha,
                        "voltar_lista",
                        "ERRO",
                        f"Falha ao voltar para lista: {e}"
                    )

                justificativas.to_csv(
                    CAMINHO_CSV_JUSTIFICATIVAS,
                    index=False,
                    encoding="latin1"
                )

            nome_anterior = nome_atual

        print("\n[INFO] Processamento concluído.")
        print("Verifique o arquivo de logs:", CAMINHO_CSV_LOGS)
        print("Verifique também o arquivo de justificativas atualizado:", CAMINHO_CSV_JUSTIFICATIVAS)
        input("Pressione ENTER para encerrar o script (o navegador pode ser fechado manualmente depois)...")
        context.close()


if __name__ == "__main__":
    main()
