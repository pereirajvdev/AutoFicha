import argparse
import time

from pywinauto import Desktop
from pywinauto import mouse


TITULO_JANELA = (
    "Gerador de Relatórios "
    "( Versão: 4.0 ) - Recursos Humanos e Folha de Pagamento"
)

TITULO_ERRO = "Erro Não Previsto"


def clicar_controle(janela, texto):
    controles = janela.descendants()

    for controle in controles:
        try:
            if controle.window_text() == texto:
                controle.click()
                return True
        except Exception:
            pass

    return False


def encontrar_janela_carregamento():
    desktop = Desktop(backend="win32")

    for janela in desktop.windows():
        try:
            botao_cancelar = janela.child_window(
                title="Cancelar",
                class_name="TButton"
            )

            if botao_cancelar.exists():
                return janela

        except Exception:
            pass

    return None


def encontrar_preview():
    desktop = Desktop(backend="win32")

    for janela in desktop.windows():
        try:
            controle = janela.child_window(
                title="frpPreview",
                class_name="TfrPreview"
            )

            if controle.exists():
                return janela

        except Exception:
            pass

    return None


def tratar_erro():
    desktop = Desktop(backend="win32")

    janela_erro = desktop.window(title=TITULO_ERRO)

    if not janela_erro.exists(timeout=0.2):
        return False

    print("Janela 'Erro Não Previsto' detectada.")

    if not clicar_controle(janela_erro, "&Não"):
        print("Botão 'Não' não encontrado.")
        return False

    print("Clicando em 'Não'.")

    return True


def processar_resultado():
    print("Aguardando resultado...")

    janela_carregamento = None
    tela_anterior = None

    while True:

        tela_atual = get_screen_name()

        if tela_atual != tela_anterior:
            print(tela_atual)
            tela_anterior = tela_atual

        # Guarda a janela de carregamento assim que ela aparecer
        if janela_carregamento is None:
            janela_carregamento = encontrar_janela_carregamento()

            if janela_carregamento is not None:
                print("Janela de carregamento encontrada.")

        # Verifica se apareceu erro
        if tratar_erro():
            print("Erro tratado.")

            if janela_carregamento is not None:
                if clicar_controle(janela_carregamento, "Cancelar"):
                    print("Clicando em 'Cancelar'.")
                    return "erro"

                print("Botão 'Cancelar' não encontrado.")
                return None

            print("Janela de carregamento não foi localizada.")
            return None

        # Verifica se abriu o preview
        if tela_atual == "Preview":
            print("Relatório aberto no preview.")

            clicar_botao_preview()

            imprimir_como_pdf()
            
            return "preview"

        time.sleep(0.2)


def get_screen_name():
    desktop = Desktop(backend="win32")

    for janela in desktop.windows():
        try:
            titulo = janela.window_text()
            classe = janela.class_name()

            if titulo == TITULO_JANELA:
                return "Gerador"

            if (
                titulo == "Ficha_Financeira_Resumo_Geral"
                and classe == "TfrmPreview"
            ):
                return "Preview"

        except Exception:
            pass

    return "Desconhecida"


def alterar_ano(ano):
    desktop = Desktop(backend="win32")

    janela = desktop.window(title=TITULO_JANELA)

    if not janela.exists(timeout=5):
        print("Janela do Gerador de Relatórios não encontrada.")
        return

    print("Janela encontrada.")
    
    controles = janela.descendants()

    campo_ano = None

    for controle in controles:
        try:
            if controle.class_name() != "TcxCustomInnerTextEdit":
                continue

            rect = controle.rectangle()

            if rect.left < 500 and rect.top < 450:
                campo_ano = controle
                break

        except Exception:
            pass

    if campo_ano is None:
        print("Campo do ano não encontrado.")
        return

    print("Campo do ano encontrado.")
    print("Valor atual:", campo_ano.window_text())
    print("Alterando para:", ano)

    campo_ano.set_focus()

    campo_ano.type_keys("^a")
    campo_ano.type_keys(str(ano))

    print("Valor depois da alteração:", campo_ano.window_text())

    botao_ok = None

    for controle in controles:
        try:
            if (
                controle.class_name() == "TBitBtn"
                and controle.window_text() == "&OK"
            ):
                botao_ok = controle
                break
        except Exception:
            pass

    if botao_ok is None:
        print("Botão OK não encontrado.")
        return

    print("Botão OK encontrado.")
    print("Clicando em OK...")

    botao_ok.click()

    print("OK clicado.")

    # Trata o resultado do processamento.
    resultado = processar_resultado()

    if resultado == "erro":
        print("Fluxo de erro concluído.")

    elif resultado == "preview":
        print("Processamento concluído.")


def clicar_botao_preview():
    x = 306
    y = 39

    print(f"Clicando no botão em ({x}, {y})...")

    mouse.click(coords=(x, y))

    print("Botão clicado.")


def imprimir_como_pdf():
    desktop = Desktop(backend="win32")

    print("Aguardando janela de impressão...")

    janela = None

    while janela is None:
        janelas = desktop.windows(
            class_name="TfrPrintForm"
        )

        for item in janelas:
            try:
                if item.is_visible():
                    janela = item
                    break
            except Exception:
                pass

        if janela is None:
            time.sleep(0.2)

    print("Janela de impressão encontrada.")
    print("Handle:", janela.handle)

    listar_combos_impressao(janela)

    # ---------------------------------------------------------
    # Localiza a ComboBox da impressora
    # ---------------------------------------------------------

    combo_impressora = None

    for controle in janela.descendants():
        try:
            if controle.class_name() != "TComboBox":
                continue

            itens = controle.texts()

            print(
                "Analisando TComboBox:",
                repr(controle.window_text()),
                "| Handle:",
                controle.handle
            )

            if "Microsoft Print to PDF" in itens:
                combo_impressora = controle
                break

        except Exception:
            pass

    if combo_impressora is None:
        print("ComboBox da impressora não encontrada.")
        return False

    # ---------------------------------------------------------
    # Seleciona Microsoft Print to PDF
    # ---------------------------------------------------------

    print(
        "ComboBox da impressora encontrada:",
        combo_impressora.handle
    )

    print(
        "Impressora atual:",
        repr(combo_impressora.window_text())
    )

    print("Selecionando Microsoft Print to PDF...")

    try:
        combo_impressora.select("Microsoft Print to PDF")
    except Exception as e:
        print("Erro ao selecionar impressora:", e)
        return False

    time.sleep(0.5)

    print(
        "Impressora depois da seleção:",
        repr(combo_impressora.window_text())
    )

    # ---------------------------------------------------------
    # Localiza o botão OK
    # ---------------------------------------------------------

    botao_ok = None

    for controle in janela.descendants():
        try:
            if (
                controle.class_name() == "TButton"
                and controle.window_text() == "OK"
            ):
                botao_ok = controle
                break

        except Exception:
            pass

    if botao_ok is None:
        print("Botão OK da impressão não encontrado.")
        return False

    print("Botão OK da impressão encontrado.")
    print("Handle:", botao_ok.handle)
    print("Clicando em OK...")

    botao_ok.click()

    print("OK da impressão clicado.")

    return True


def listar_combos_impressao(janela):
    print("\nCOMBOBOXES ENCONTRADAS")
    print("=" * 70)

    combos = janela.descendants(class_name="TComboBox")

    print(f"Quantidade: {len(combos)}\n")

    for i, combo in enumerate(combos):
        try:
            print("-" * 70)
            print(f"COMBO [{i}]")
            print("Handle :", combo.handle)
            print("Texto  :", repr(combo.window_text()))
            print("Classe :", combo.class_name())

            try:
                itens = combo.texts()

                print("Quantidade de itens:", len(itens))

                for j, item in enumerate(itens):
                    print(f"  [{j}] {item!r}")

            except Exception as e:
                print("Não foi possível obter os itens:", e)

        except Exception as e:
            print(f"Erro na ComboBox [{i}]: {e}")

def selecionar_impressora(janela, nome_impressora):
    combos = janela.descendants(
        class_name="TComboBox"
    )

    for combo in combos:
        try:
            itens = combo.texts()

            if nome_impressora in itens:
                print("ComboBox da impressora encontrada.")
                print("Handle:", combo.handle)
                print("Impressora atual:", repr(combo.window_text()))

                print("Selecionando:", nome_impressora)

                combo.select(nome_impressora)

                time.sleep(0.5)

                print(
                    "Impressora depois da seleção:",
                    repr(combo.window_text())
                )

                return True

        except Exception as e:
            print(
                f"Erro ao analisar ComboBox "
                f"{combo.handle}: {e}"
            )

    print(
        f"Impressora {nome_impressora!r} "
        "não encontrada."
    )

    return False

def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "intervalo",
        help="Intervalo de anos no formato ANO_INICIAL-ANO_FINAL"
    )

    args = parser.parse_args()

    ano_inicial, ano_final = args.intervalo.split("-")

    ano_inicial = int(ano_inicial)
    ano_final = int(ano_final)

    if ano_inicial > ano_final:
        raise ValueError("O ano inicial não pode ser maior que o ano final.")

    print(f"Intervalo recebido: {ano_inicial}-{ano_final}")

    alterar_ano(ano_inicial)


if __name__ == "__main__":
    main()