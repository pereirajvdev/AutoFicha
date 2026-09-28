import argparse

from pywinauto import Desktop


TITULO_JANELA = (
    "Gerador de Relatórios "
    "( Versão: 4.0 ) - Recursos Humanos e Folha de Pagamento"
)


def alterar_ano(ano):
    desktop = Desktop(backend="win32")

    janela = desktop.window(title=TITULO_JANELA)

    if not janela.exists(timeout=5):
        print("Janela do Gerador de Relatórios não encontrada.")
        return

    print("Janela encontrada.")

    controles = janela.descendants()

    campo_ano = None
    botao_ok = None

    for controle in controles:
        try:
            classe = controle.class_name()
            texto = controle.window_text()

            if (
                classe == "TcxCustomInnerTextEdit"
                and texto == "2026"
            ):
                campo_ano = controle

            if (
                classe == "TBitBtn"
                and texto == "&OK"
            ):
                botao_ok = controle

        except Exception:
            pass

    if campo_ano is None:
        print("Campo do ano não encontrado.")
        return

    if botao_ok is None:
        print("Botão OK não encontrado.")
        return

    # Altera o ano
    print("Campo do ano encontrado.")
    print("Valor atual:", campo_ano.window_text())
    print("Alterando para:", ano)

    campo_ano.set_focus()
    campo_ano.type_keys("^a")
    campo_ano.type_keys(str(ano))

    print("Ano alterado para:", campo_ano.window_text())

    # Clica no OK
    print("Clicando em OK...")
    botao_ok.click()

    print("OK clicado.")
    print("Fim da automação.")


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