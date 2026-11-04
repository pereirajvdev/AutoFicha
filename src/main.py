"""
Automação do Gerador de Relatórios (RH e Folha) para exportar a
Ficha Financeira em PDF.

Uso:
    python main.py vinculos.xlsx          # colunas: matrícula | início | fim
    python main.py --mat 22057 2017-2018  # matrícula + intervalo (ou um ano só)
    python main.py --padrao               # matrícula 3198, ano 2026

Opções: --perf (relatório de tempos), --debug (diagnóstico detalhado)
"""

import argparse
import ctypes
import time
from contextlib import contextmanager
from ctypes import wintypes
from datetime import date, datetime
from functools import wraps
from pathlib import Path

from openpyxl import load_workbook
from pywinauto import Desktop, keyboard, mouse


# ======================================================================
# CONFIGURAÇÃO
# ======================================================================

TITULO_GERADOR = (
    "Gerador de Relatórios "
    "( Versão: 4.0 ) - Recursos Humanos e Folha de Pagamento"
)
TITULO_ERRO = "Erro Não Previsto"
TITULO_NAO_ENCONTRADO = "Arquivo Não Encontrado"
TITULO_SAIDA_IMPRESSAO = "Salvar saída de impressão como"
TITULO_CONFIRMAR_SOBRESCRITA = "Confirmar Salvar como"
TITULO_FICHA = "Ficha_Financeira_Resumo_Geral"
TITULO_IMPRIMIR = "Imprimir"
TITULO_PREPARANDO_IMPRESSAO = "Preparando - Ficha_Financeira_Resumo_Geral"

DIRETORIO_SERVIDORES = Path(
    r"C:\Users\Joao Castro\Desktop\JOAO\Servidores"
)

IMPRESSORA_PDF = "Microsoft Print to PDF"

# Coordenadas (tela) do botão de imprimir no preview.
BOTAO_PREVIEW_XY = (306, 39)

TEMPO_LIMITE = 60
TEMPO_RESULTADO = 180
INTERVALO_VERIFICACAO = 0.3
MAX_TENTATIVAS_FICHA = 3

PERF_ATIVO = False
DEBUG = False


# ======================================================================
# WIN32 (ctypes): enumeração rápida de janelas
# ======================================================================

_user32 = ctypes.windll.user32

_WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

_user32.EnumWindows.argtypes = [_WNDENUMPROC, wintypes.LPARAM]
_user32.EnumWindows.restype = wintypes.BOOL
_user32.IsWindow.argtypes = [wintypes.HWND]
_user32.IsWindow.restype = wintypes.BOOL
_user32.IsWindowVisible.argtypes = [wintypes.HWND]
_user32.IsWindowVisible.restype = wintypes.BOOL
_user32.IsHungAppWindow.argtypes = [wintypes.HWND]
_user32.IsHungAppWindow.restype = wintypes.BOOL
_user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
_user32.GetWindowTextLengthW.restype = ctypes.c_int
_user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
_user32.GetWindowTextW.restype = ctypes.c_int

_DESKTOP = None
_AVISOS_TRAVADA = set()


def _desktop():
    """Um único objeto Desktop durante toda a execução."""
    global _DESKTOP

    if _DESKTOP is None:
        _DESKTOP = Desktop(backend="win32")

    return _DESKTOP


# ======================================================================
# MONITORAMENTO DE DESEMPENHO
# ======================================================================

_TEMPOS = {}


def _registrar_tempo(nome, duracao):
    registro = _TEMPOS.setdefault(nome, [0, 0.0])
    registro[0] += 1
    registro[1] += duracao


@contextmanager
def medir(nome):
    """Mede o tempo de um bloco. Sem efeito se --perf estiver desligado."""
    if not PERF_ATIVO:
        yield
        return

    inicio = time.perf_counter()
    try:
        yield
    finally:
        _registrar_tempo(nome, time.perf_counter() - inicio)


def cronometrar(nome=None):
    """Decorator que mede o tempo total de uma função."""
    def decorador(func):
        rotulo = nome or func.__name__

        @wraps(func)
        def wrapper(*args, **kwargs):
            if not PERF_ATIVO:
                return func(*args, **kwargs)

            inicio = time.perf_counter()
            try:
                return func(*args, **kwargs)
            finally:
                _registrar_tempo(rotulo, time.perf_counter() - inicio)

        return wrapper

    return decorador


def imprimir_relatorio_tempos():
    if not PERF_ATIVO or not _TEMPOS:
        return

    print("\n" + "=" * 66)
    print(" RELATÓRIO DE DESEMPENHO (--perf)")
    print(" Obs.: tempos aninhados se sobrepõem; não some as linhas.")
    print("=" * 66)
    print(
        f"{'Etapa':<42}{'Chamadas':>9}{'Total (s)':>11}{'Média (s)':>11}"
    )

    for nome, (chamadas, total) in sorted(
        _TEMPOS.items(), key=lambda item: item[1][1], reverse=True
    ):
        media = total / chamadas if chamadas else 0.0
        print(f"{nome:<42}{chamadas:>9}{total:>11.2f}{media:>11.3f}")

    print("=" * 66)


def log_debug(*args):
    """Mensagens de diagnóstico; só aparecem com --debug."""
    if DEBUG:
        print(*args)


# ======================================================================
# INFRAESTRUTURA
# ======================================================================

def _janela_valida(handle, padrao=False):
    """Verifica (barato) se um handle ainda existe."""
    try:
        return bool(_user32.IsWindow(handle))
    except Exception:
        return padrao


def _janela_visivel(hwnd):
    return bool(_user32.IsWindow(hwnd)) and bool(_user32.IsWindowVisible(hwnd))


def _mesmo_titulo(a, b):
    return a.strip().casefold() == b.strip().casefold()


class Controles(list):
    """
    Lista de controles da janela principal, indexada por classe
    (construída uma única vez) e com cache de controles já localizados.

    A classe de um controle nunca muda; o texto (window_text) é lido
    na hora da busca, pois pode mudar.
    """

    def __init__(self, iteravel=()):
        super().__init__(iteravel)
        self._por_classe = {}
        self._cache = {}

        for controle in self:
            try:
                nome = controle.class_name()
            except Exception:
                continue

            self._por_classe.setdefault(nome, []).append(controle)

    def da_classe(self, nome):
        return self._por_classe.get(nome, [])

    def em_cache(self, chave, buscar):
        """Devolve o controle em cache se o handle for válido; senão busca."""
        controle = self._cache.get(chave)

        if controle is not None and _janela_valida(controle.handle):
            return controle

        controle = buscar()

        if controle is not None:
            self._cache[chave] = controle

        return controle


# ======================================================================
# LOCALIZAÇÃO DE JANELAS
# ======================================================================

def _listar_janelas_visiveis():
    """Enumera janelas de nível superior visíveis: [(hwnd, titulo), ...]."""
    resultado = []

    def callback(hwnd, _lparam):
        if _user32.IsWindowVisible(hwnd):
            tamanho = _user32.GetWindowTextLengthW(hwnd)
            if tamanho:
                buffer = ctypes.create_unicode_buffer(tamanho + 1)
                _user32.GetWindowTextW(hwnd, buffer, tamanho + 1)
                resultado.append((hwnd, buffer.value))
        return True

    proc = _WNDENUMPROC(callback)  # mantém referência viva durante a chamada
    _user32.EnumWindows(proc, 0)
    return resultado


def _buscar_janela(titulos):
    """
    UMA enumeração; retorna (hwnd, titulo_real) da primeira janela válida,
    respeitando a ordem dos títulos. Título idêntico (sem diferenciar
    maiúsculas/minúsculas). Janelas travadas são ignoradas.
    """
    candidatas = {}

    for hwnd, texto in _listar_janelas_visiveis():
        texto = texto.strip()
        candidatas.setdefault(texto.casefold(), []).append((hwnd, texto))

    for titulo in titulos:
        for hwnd, texto in candidatas.get(titulo.strip().casefold(), []):
            if _user32.IsHungAppWindow(hwnd):
                if hwnd not in _AVISOS_TRAVADA:
                    _AVISOS_TRAVADA.add(hwnd)
                    print(f"[AGUARDANDO] Janela sem resposta: {texto}")
                continue

            return hwnd, texto

    return None


def _criar_wrapper(hwnd):
    return _desktop().window(handle=hwnd).wrapper_object()


@cronometrar()
def localizar_janela(titulo_esperado):
    """Verificação única (sem espera). Retorna wrapper ou None."""
    achado = _buscar_janela([titulo_esperado])

    if achado is None:
        return None

    hwnd, titulo = achado
    log_debug(f"Janela identificada: {titulo}")
    return _criar_wrapper(hwnd)


@cronometrar()
def aguardar_janela(*titulos_esperados, time_limit=TEMPO_LIMITE):
    """
    Aguarda até que uma janela com um dos títulos apareça.
    Retorna o wrapper (seguindo a ordem dos títulos) ou None no timeout.
    """
    print(f"Aguardando uma das janelas: {', '.join(titulos_esperados)}")

    limite = time.monotonic() + time_limit

    while True:
        with medir("enumerar janelas"):
            achado = _buscar_janela(titulos_esperados)

        if achado is not None:
            hwnd, titulo = achado
            print(f"Janela identificada: {titulo}")
            return _criar_wrapper(hwnd)

        if time.monotonic() >= limite:
            break

        time.sleep(INTERVALO_VERIFICACAO)

    print(
        f"Nenhuma janela válida apareceu em {time_limit} segundos: "
        f"{', '.join(titulos_esperados)}"
    )
    return None


def exigir_janela(*titulos_esperados, time_limit=TEMPO_LIMITE):
    janela = aguardar_janela(*titulos_esperados, time_limit=time_limit)

    if janela is None:
        raise RuntimeError(
            "Janela não encontrada: " + ", ".join(titulos_esperados)
        )

    return janela


def aguardar_fechamento(hwnd, time_limit=10):
    """Espera até a janela deixar de existir/ficar visível."""
    limite = time.monotonic() + time_limit

    while _janela_visivel(hwnd) and time.monotonic() < limite:
        time.sleep(INTERVALO_VERIFICACAO)

    return not _janela_visivel(hwnd)


# ======================================================================
# CONTROLES E AÇÕES
# ======================================================================

@cronometrar()
def clicar_controle(janela, texto):
    for controle in janela.descendants():
        try:
            if controle.window_text() == texto:
                controle.click()
                return True
        except Exception:
            pass

    return False


def tratar_erro(janela_erro):
    """Clica em 'Não' na janela 'Erro Não Previsto' já localizada."""
    print("Janela 'Erro Não Previsto' detectada.")

    if not clicar_controle(janela_erro, "&Não"):
        print("Botão 'Não' não encontrado.")
        return False

    print("Clicado em 'Não'.")
    return True


@cronometrar()
def alterar_ano(ano, controles):
    """Preenche o campo do ano. Retorna True se o valor ficou correto."""

    def buscar():
        # Só percorre os TcxCustomInnerTextEdit (índice por classe).
        for controle in controles.da_classe("TcxCustomInnerTextEdit"):
            try:
                texto = controle.window_text().strip()

                if texto.isdigit() and 1900 <= int(texto) <= 2100:
                    return controle
            except Exception:
                pass

        return None

    # O campo é achado uma vez e reaproveitado: depois da primeira
    # alteração, outro campo numérico poderia casar com o filtro.
    campo_ano = controles.em_cache("campo_ano", buscar)

    if campo_ano is None:
        print("Campo do ano não encontrado.")
        return False

    print("Alterando ano:", campo_ano.window_text().strip(), "->", ano)

    campo_ano.set_focus()
    campo_ano.type_keys("^a")
    campo_ano.type_keys(str(ano))

    valor = campo_ano.window_text().strip()

    if valor != str(ano):
        print(f"ATENÇÃO: campo do ano ficou com {valor!r}, esperado {ano}.")
        return False

    return True


@cronometrar()
def clicar_ok(controles):
    def buscar():
        for controle in controles.da_classe("TBitBtn"):
            try:
                if controle.window_text() == "&OK":
                    return controle
            except Exception:
                pass

        return None

    botao_ok = controles.em_cache("botao_ok", buscar)

    if botao_ok is None:
        print("Botão OK não encontrado.")
        return False

    print("Clicando em OK...")
    botao_ok.click()
    return True


def clicar_botao_preview():
    x, y = BOTAO_PREVIEW_XY
    print(f"Clicando no botão do preview em ({x}, {y})...")
    mouse.click(coords=(x, y))


# ======================================================================
# RESULTADO / PREVIEW
# ======================================================================

@cronometrar()
def fechar_preview(aguardar=True):
    """
    aguardar=True : espera a janela aparecer (até TEMPO_LIMITE).
    aguardar=False: fecha só se ela já estiver aberta, sem esperar.
    """
    try:
        if aguardar:
            janela = aguardar_janela(TITULO_FICHA)
        else:
            janela = localizar_janela(TITULO_FICHA)

        if janela is None:
            print("Preview já estava fechado.")
            return True

        janela.set_focus()
        print("Fechando preview...")
        janela.close()
        print("Preview fechado.")
        return True

    except Exception as e:
        print("Erro ao fechar o preview:", e)
        return False


@cronometrar()
def processar_resultado(nome_funcionario, ano):
    """
    Após clicar em OK pode aparecer: erro, arquivo não encontrado ou preview.

    Retorna: "preview" | "erro" | "nao_encontrado" | "falhou" | "sem_janela"
    """
    print("Aguardando resultado...")

    janela = aguardar_janela(
        TITULO_ERRO,
        TITULO_NAO_ENCONTRADO,
        TITULO_FICHA,
        time_limit=TEMPO_RESULTADO,
    )

    if janela is None:
        return "sem_janela"

    janela.set_focus()
    titulo = janela.window_text()

    if _mesmo_titulo(titulo, TITULO_NAO_ENCONTRADO):
        input("Sistema possivelmente caiu, aguarde estabilidade e tecle ENTER. ")

        try:
            janela.close()
        except Exception:
            pass

        return "nao_encontrado"

    if _mesmo_titulo(titulo, TITULO_ERRO):
        tratar_erro(janela)
        return "erro"

    # Preview da ficha
    print("Relatório aberto no preview.")

    try:
        clicar_botao_preview()
        salvo = imprimir_como_pdf(nome_funcionario, ano)
    finally:
        # Fecha o preview em qualquer caso, sem esperar por ele.
        fechar_preview(aguardar=False)

    return "preview" if salvo else "falhou"


@cronometrar()
def executar_ficha(nome_funcionario, ano, controles):
    """
    Clica em OK e trata o resultado. Se nenhuma janela aparecer, aperta
    ENTER (provável aviso bloqueando a tela) e tenta de novo, no máximo
    MAX_TENTATIVAS_FICHA vezes.
    """
    resultado = "falhou"

    for tentativa in range(1, MAX_TENTATIVAS_FICHA + 1):
        if not clicar_ok(controles):
            return "falhou"

        resultado = processar_resultado(nome_funcionario, ano)

        if resultado == "preview":
            print("[OK] Processamento concluído.")
            return resultado

        if resultado == "erro":
            print("[ERRO] Fluxo de erro concluído.")
            return resultado

        if resultado == "sem_janela":
            print(
                f"[ERRO] Nenhuma janela apareceu (tentativa {tentativa}/"
                f"{MAX_TENTATIVAS_FICHA}). Enviando ENTER."
            )
            janela = aguardar_janela(TITULO_PREPARANDO_IMPRESSAO)
            janela.set_focus()
            janela.close()

        else:
            print(
                f"Resultado {resultado!r} (tentativa {tentativa}/"
                f"{MAX_TENTATIVAS_FICHA})."
            )

    return resultado


# ======================================================================
# IMPRESSÃO / PDF
# ======================================================================

def listar_combos_impressao(janela):
    """Diagnóstico (--debug): lista as ComboBoxes da janela de impressão."""
    print("\nCOMBOBOXES ENCONTRADAS")
    print("=" * 70)

    combos = janela.descendants(class_name="TComboBox")
    print(f"Quantidade: {len(combos)}\n")

    for i, combo in enumerate(combos):
        try:
            print("-" * 70)
            print(f"COMBO [{i}] handle={combo.handle} "
                  f"texto={combo.window_text()!r}")

            for j, item in enumerate(combo.texts()):
                print(f"  [{j}] {item!r}")

        except Exception as e:
            print(f"Erro na ComboBox [{i}]: {e}")


@cronometrar()
def imprimir_como_pdf(nome_funcionario, ano):
    print("Aguardando janela de impressão...")

    janela = aguardar_janela(TITULO_IMPRIMIR)

    if janela is None:
        print("Janela de impressão não apareceu.")
        return False

    janela.set_focus()
    log_debug("Handle:", janela.handle)

    if DEBUG:
        listar_combos_impressao(janela)

    # Localiza a ComboBox da impressora (só TComboBox são envolvidos)
    combo_impressora = None

    for controle in janela.descendants(class_name="TComboBox"):
        try:
            if IMPRESSORA_PDF in controle.texts():
                combo_impressora = controle
                break
        except Exception:
            pass

    if combo_impressora is None:
        print("ComboBox da impressora não encontrada.")
        return False

    print(f"Selecionando {IMPRESSORA_PDF}...")

    try:
        combo_impressora.select(IMPRESSORA_PDF)
    except Exception as e:
        print("Erro ao selecionar impressora:", e)
        return False

    time.sleep(0.05)

    # Botão OK
    botao_ok = None

    for controle in janela.descendants(class_name="TButton"):
        try:
            if controle.window_text() == "OK":
                botao_ok = controle
                break
        except Exception:
            pass

    if botao_ok is None:
        print("Botão OK da impressão não encontrado.")
        return False

    print("Clicando em OK da impressão...")
    botao_ok.click()

    janela_salvar = aguardar_janela(TITULO_SAIDA_IMPRESSAO)

    if janela_salvar is None:
        print("Janela de salvamento não apareceu.")
        return False

    janela_salvar.set_focus()

    return salvar_pdf(nome_funcionario, ano, janela_salvar)


@cronometrar()
def salvar_pdf(nome_funcionario, ano, janela):
    pasta_servidor = DIRETORIO_SERVIDORES / nome_funcionario
    pasta_servidor.mkdir(parents=True, exist_ok=True)

    caminho_arquivo = pasta_servidor / f"FICHA FINANCEIRA {ano}.pdf"

    print("Arquivo:", caminho_arquivo)

    # Campo grande de nome do arquivo
    campo_nome = None

    for controle in janela.descendants(class_name="Edit"):
        try:
            if controle.rectangle().width() > 500:
                campo_nome = controle
                break
        except Exception:
            pass

    if campo_nome is None:
        print("Campo do nome do arquivo não encontrado.")
        return False

    campo_nome.set_focus()
    campo_nome.set_edit_text(str(caminho_arquivo))

    # Botão Salvar
    botao_salvar = None

    for controle in janela.descendants(class_name="Button"):
        try:
            if controle.window_text() == "Sa&lvar":
                botao_salvar = controle
                break
        except Exception:
            pass

    if botao_salvar is None:
        print("Botão Salvar não encontrado.")
        return False

    hwnd_salvar = janela.handle

    print("Clicando em Salvar...")
    botao_salvar.click()

    confirmar_sobrescrita_pdf(hwnd_salvar)

    # Espera o diálogo de salvamento sumir antes de seguir.
    if not aguardar_fechamento(hwnd_salvar):
        print("ATENÇÃO: janela de salvamento continua aberta.")
        return False

    print("PDF salvo.")
    return True


@cronometrar()
def confirmar_sobrescrita_pdf(hwnd_salvar=None, time_limit=5):
    """
    Clica em "&Sim" se aparecer "Confirmar Salvar como".

    Se a janela de salvamento fechar sem a confirmação ter aparecido,
    o arquivo era novo: encerra na hora em vez de esperar o limite todo.
    """
    print("Verificando confirmação de sobrescrita...")

    limite = time.monotonic() + time_limit

    while time.monotonic() < limite:
        achado = _buscar_janela([TITULO_CONFIRMAR_SOBRESCRITA])

        if achado is not None:
            janela = _criar_wrapper(achado[0])

            for botao in janela.descendants(class_name="Button"):
                try:
                    if botao.window_text().strip() == "&Sim":
                        botao.click()
                        print("Sobrescrita confirmada.")
                        return True
                except Exception:
                    pass

            print('Janela de confirmação sem botão "&Sim".')

        if hwnd_salvar is not None and not _janela_valida(hwnd_salvar, True):
            print("Janela de salvamento fechada sem pedir confirmação.")
            return False

        time.sleep(0.1)

    print("Confirmação de sobrescrita não apareceu.")
    return False


# ======================================================================
# FUNCIONÁRIO / MATRÍCULA
# ======================================================================

_CACHE_LISTBOX = {}


def _obter_listboxes(janela):
    """TListBox da janela principal, com cache revalidado por handle."""
    cache = _CACHE_LISTBOX.get(janela.handle)

    if cache and all(_janela_valida(lb.handle) for lb in cache):
        return cache

    listboxes = janela.descendants(class_name="TListBox")
    _CACHE_LISTBOX[janela.handle] = listboxes
    return listboxes


@cronometrar()
def obter_nome_funcionario(janela):
    for listbox in _obter_listboxes(janela):
        try:
            for item in listbox.texts():
                item = item.strip()

                if " - " not in item:
                    continue

                codigo, nome = item.split(" - ", 1)

                if codigo.strip().isdigit() and nome.strip():
                    return nome.strip()

        except Exception:
            pass

    print("Nome do funcionário não encontrado.")
    return input("Digite o nome a ser considerado: ").strip()


def _obter_campo_matricula(controles):
    """Localiza (uma vez) o TelLookupCombo da aba 'Funcionário'."""

    def buscar():
        aba = None

        for controle in controles.da_classe("TcxTabSheet"):
            try:
                if controle.window_text().strip() == "Funcionário":
                    aba = controle
                    break
            except Exception:
                pass

        if aba is None:
            print("Aba 'Funcionário' não encontrada.")
            return None

        combos = aba.descendants(class_name="TelLookupCombo")
        print(f"TelLookupCombo encontrados na aba: {len(combos)}")

        if not combos:
            print("Nenhum TelLookupCombo na aba 'Funcionário'.")
            return None

        if len(combos) > 1:
            print(
                "ATENÇÃO: vários TelLookupCombo na aba 'Funcionário'; "
                "usando o primeiro (matrícula)."
            )

        if DEBUG:
            for combo in combos:
                try:
                    rect = combo.rectangle()
                    print(
                        f"TelLookupCombo | texto={combo.window_text()!r} | "
                        f"pos=({rect.left},{rect.top}) | "
                        f"tam=({rect.width()}x{rect.height()})"
                    )
                except Exception:
                    pass

        return combos[0]

    return controles.em_cache("campo_matricula", buscar)


@cronometrar()
def preencher_matricula(matricula, controles):
    janela = aguardar_janela(TITULO_GERADOR)

    if janela is None:
        return False

    janela.set_focus()

    campo = _obter_campo_matricula(controles)

    if campo is None:
        return False

    campo.set_focus()
    campo.type_keys("^a")
    campo.type_keys(str(matricula))
    campo.type_keys("{ENTER}")

    print("Matrícula preenchida e confirmada:", matricula)
    return True


# ======================================================================
# EXCEL
# ======================================================================

def _converter_data(valor, numero_linha):
    if isinstance(valor, datetime):
        return valor.date()

    if isinstance(valor, date):
        return valor

    if isinstance(valor, str):
        return datetime.strptime(valor.strip(), "%d/%m/%Y").date()

    raise ValueError(f"Data inválida na linha {numero_linha}: {valor!r}")


@cronometrar()
def ler_vinculos(caminho_excel):
    caminho_excel = Path(caminho_excel)

    if not caminho_excel.exists():
        raise FileNotFoundError(
            f"Arquivo de vínculos não encontrado: {caminho_excel}"
        )

    wb = load_workbook(caminho_excel, read_only=True, data_only=True)

    try:
        ws = wb.active
        vinculos = []

        for numero_linha, linha in enumerate(
            ws.iter_rows(min_col=1, max_col=3, values_only=True), start=1
        ):
            # Em modo read_only a linha pode vir mais curta que 3 colunas.
            matricula, inicio, fim = (tuple(linha) + (None,) * 3)[:3]

            if matricula is None and inicio is None and fim is None:
                continue

            try:
                matricula_num = int(matricula)
            except (TypeError, ValueError, OverflowError):
                continue  # cabeçalho ou linha inválida

            data_inicio = _converter_data(inicio, numero_linha)
            data_fim = _converter_data(fim, numero_linha)

            if data_inicio > data_fim:
                raise ValueError(
                    f"Data inicial posterior à final na linha {numero_linha}."
                )

            vinculos.append({
                "matricula": matricula_num,
                "inicio": data_inicio,
                "fim": data_fim,
            })

        if not vinculos:
            raise ValueError(
                "Nenhum vínculo válido encontrado nas colunas A, B e C."
            )

        return vinculos

    finally:
        wb.close()


# ======================================================================
# PLANO DE EXECUÇÃO
# ======================================================================

def _parse_intervalo(texto, parser):
    """'2017-2024' ou '2017' -> lista de anos."""
    try:
        partes = [int(p) for p in texto.split("-")]
    except ValueError:
        parser.error("Intervalo inválido. Use ANO_INICIAL-ANO_FINAL.")

    if len(partes) == 1:
        partes = partes * 2

    if len(partes) != 2:
        parser.error("Intervalo inválido. Use ANO_INICIAL-ANO_FINAL.")

    ano_inicial, ano_final = partes

    if ano_inicial > ano_final:
        parser.error("O ano inicial não pode ser maior que o ano final.")

    return list(range(ano_inicial, ano_final + 1))


def montar_plano(args, parser):
    """
    Valida a entrada ANTES de mexer na interface.
    Retorna (matriculas, anos).
    """
    if args.padrao:
        print("MODO PADRÃO")
        return [3198], [2026]

    if args.entrada and args.entrada.lower().endswith(".xlsx"):
        vinculos = ler_vinculos(args.entrada)

        # Sem duplicatas, mantendo a ordem do arquivo.
        matriculas = list(dict.fromkeys(v["matricula"] for v in vinculos))

        # Todas as matrículas ficam na seleção ao mesmo tempo, então a
        # ficha de cada ano é gerada UMA vez, cobrindo a união dos anos
        # de todos os vínculos.
        anos = sorted({
            ano
            for v in vinculos
            for ano in range(v["inicio"].year, v["fim"].year + 1)
        })

        return matriculas, anos

    if args.mat is None:
        parser.error(
            "Informe o caminho de um vinculos.xlsx, "
            "--mat com o intervalo, ou --padrao."
        )

    if not args.entrada:
        parser.error("Informe o intervalo de anos. Exemplo: 2017-2024")

    return [args.mat], _parse_intervalo(args.entrada, parser)


# ======================================================================
# EXECUÇÃO
# ======================================================================

@cronometrar("main")
def executar(args, parser):
    matriculas, anos = montar_plano(args, parser)

    print(f"Matrículas: {matriculas}")
    print(f"Anos: {anos}\n")

    print(f"Esperando janela: {TITULO_GERADOR}\n")
    janela = exigir_janela(TITULO_GERADOR)
    janela.set_focus()

    # Controles coletados uma única vez e indexados por classe.
    controles = Controles(janela.descendants())

    print("=" * 70)
    print("ETAPA 1: PREENCHER AS MATRÍCULAS")
    print("=" * 70)

    for indice, matricula in enumerate(matriculas, start=1):
        print(f"[{indice}/{len(matriculas)}] Matrícula: {matricula}")

        if not preencher_matricula(matricula, controles):
            print(
                f"ERRO ao preencher a matrícula {matricula}. "
                "Processamento interrompido."
            )
            return

    # O nome não muda entre os anos: lido uma vez.
    nome_funcionario = obter_nome_funcionario(janela)
    print("Funcionário:", nome_funcionario)

    print()
    print("=" * 70)
    print("ETAPA 2: PROCESSAR OS ANOS")
    print("=" * 70)

    falhas = []
    tempos_ano = {}

    for indice, ano in enumerate(anos, start=1):
        print()
        print("#" * 70)
        print(f"ANO {ano} ({indice}/{len(anos)})")
        print("#" * 70)

        inicio_ano = time.perf_counter()

        if not alterar_ano(ano, controles):
            falhas.append((ano, "ano não alterado"))
            tempos_ano[ano] = time.perf_counter() - inicio_ano
            continue

        resultado = executar_ficha(nome_funcionario, ano, controles)
        tempos_ano[ano] = time.perf_counter() - inicio_ano

        print(f"Tempo do ano {ano}: {tempos_ano[ano]:.2f} s")

        if resultado != "preview":
            falhas.append((ano, resultado))

    print()
    print("=" * 70)
    print("PROCESSAMENTO FINALIZADO.")

    if tempos_ano:
        media = sum(tempos_ano.values()) / len(tempos_ano)
        print(f"Tempo médio por ano: {media:.2f} s")

    if falhas:
        print("Anos com problema:")
        for ano, motivo in falhas:
            print(f"  - {ano}: {motivo}")
    else:
        print("Todos os anos foram processados com sucesso.")

    print("=" * 70)


def main():
    global PERF_ATIVO, DEBUG

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "entrada",
        nargs="?",
        help="Caminho do vinculos.xlsx ou intervalo de anos",
    )
    parser.add_argument(
        "--mat", type=int, help="Matrícula para processamento manual"
    )
    parser.add_argument(
        "--padrao",
        action="store_true",
        help="Executa matrícula 3198 no ano 2026",
    )
    parser.add_argument(
        "--perf",
        action="store_true",
        help="Exibe relatório de desempenho da execução",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Exibe mensagens detalhadas de diagnóstico",
    )

    args = parser.parse_args()

    PERF_ATIVO = args.perf
    DEBUG = args.debug

    inicio_execucao = time.perf_counter()

    try:
        executar(args, parser)
    finally:
        tempo_total = time.perf_counter() - inicio_execucao

        print("\n" + "=" * 60)
        print(f" TEMPO TOTAL DE EXECUÇÃO: {tempo_total:.2f} segundos")
        print("=" * 60)

        # Detalhamento por etapa (só com --perf)
        imprimir_relatorio_tempos()


if __name__ == "__main__":
    main()
