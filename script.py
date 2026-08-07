import os
import re
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from pywinauto import Application, Desktop


TITLE_RE = r"^IWOS$"
ARQUIVO_ENV = Path(__file__).with_name(".env")
ATALHO_IWOS = Path(
    r"C:\Users\anali\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\..appref-ms"
)
TEMPO_LIMITE_ABERTURA = 60
TEMPO_LIMITE_BUSCA = 120
EMPRESAS = (
    ("MG Setel", "MGS"),
    ("Delta", "DELTA"),
)
PAINEIS = (
    {
        "menu": "Faltas",
        "aba": "Painel de Faltas",
        "controle": "FaltaControl",
        "grid": "grid",
        "coluna_data": "Data",
        "sheet": "Faltas",
    },
    {
        "menu": "Atestados",
        "aba": "Painel de Atestados",
        "controle": "AtestadoControl",
        "grid": "gridAtestado",
        "coluna_data": "Início",
        "sheet": "Atestados",
    },
)
COLUNAS_EXPORTACAO = (
    "ID",
    "Nome",
    "Centro",
    "Função",
    "Admissão",
    "Disponibilidade",
    "Demissão",
)
PADRAO_CELULA = re.compile(
    r"^((?:ID|Nome|Centro|Função|Admissão|Disponibilidade|Demissão)) row (\d+)$"
)


def carregar_variaveis_env(caminho):
    """Carrega configurações simples do .env sem expor os valores."""
    if not caminho.is_file():
        return

    for linha_bruta in caminho.read_text(encoding="utf-8").splitlines():
        linha = linha_bruta.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue

        chave, valor = linha.split("=", 1)
        chave = chave.strip()
        valor = valor.strip()
        if len(valor) >= 2 and valor[0] == valor[-1] and valor[0] in "\"'":
            valor = valor[1:-1]
        if chave:
            os.environ.setdefault(chave, valor)


def localizar_janela_iwos():
    return next(
        (
            janela
            for janela in Desktop(backend="uia").windows(
                title_re=TITLE_RE,
                visible_only=True,
            )
            if "Visual Studio Code" not in janela.window_text()
        ),
        None,
    )


def localizar_janela_login():
    for janela in Desktop(backend="uia").windows(
        title="Login",
        visible_only=True,
    ):
        auto_ids = {
            controle.element_info.automation_id
            for controle in janela.descendants()
        }
        if {"editUsuario", "editSenha"}.issubset(auto_ids):
            return janela
    return None


def aguardar_inicio_iwos():
    limite = time.monotonic() + TEMPO_LIMITE_ABERTURA
    while time.monotonic() < limite:
        janela_iwos = localizar_janela_iwos()
        if janela_iwos is not None:
            return janela_iwos, None

        janela_login = localizar_janela_login()
        if janela_login is not None:
            return None, janela_login
        time.sleep(1)

    return None, None


def obter_janela_iwos():
    carregar_variaveis_env(ARQUIVO_ENV)
    janela_candidata = localizar_janela_iwos()
    janela_login = localizar_janela_login() if janela_candidata is None else None

    if janela_candidata is None and janela_login is None:
        if not ATALHO_IWOS.is_file():
            raise FileNotFoundError(f"Atalho do IWOS não encontrado: {ATALHO_IWOS}")

        print("IWOS não está aberto. Iniciando o aplicativo...")
        os.startfile(ATALHO_IWOS)
        janela_candidata, janela_login = aguardar_inicio_iwos()

    if janela_candidata is None and janela_login is not None:
        usuario = os.getenv("IWOS_USUARIO", "").strip()
        senha = os.getenv("IWOS_SENHA", "")
        if not usuario or not senha:
            raise RuntimeError(
                "Preencha IWOS_USUARIO e IWOS_SENHA no arquivo .env para entrar."
            )

        app_login = Application(backend="uia").connect(handle=janela_login.handle)
        login = app_login.window(handle=janela_login.handle)
        login.set_focus()
        login.child_window(
            auto_id="editUsuario",
            control_type="Edit",
        ).set_edit_text(usuario)
        login.child_window(
            auto_id="editSenha",
            control_type="Edit",
        ).set_edit_text(senha)
        print("Credenciais preenchidas. Entrando no IWOS...")
        login.child_window(
            auto_id="buttonEntrar",
            control_type="Button",
        ).click_input()

        limite = time.monotonic() + TEMPO_LIMITE_ABERTURA
        while time.monotonic() < limite:
            time.sleep(1)
            janela_candidata = localizar_janela_iwos()
            if janela_candidata is not None:
                break

    if janela_candidata is None:
        raise RuntimeError(
            f"A janela principal do IWOS não apareceu em "
            f"{TEMPO_LIMITE_ABERTURA} segundos."
        )

    app = Application(backend="uia").connect(handle=janela_candidata.handle)
    return app, app.window(handle=janela_candidata.handle)


def abrir_quadro_funcionarios(janela):
    """Abre Recursos Humanos > Quadro de Funcionários."""
    janela.set_focus()
    recursos_humanos = next(
        (
            item
            for item in janela.descendants(control_type="MenuItem")
            if item.window_text() == "Recursos Humanos"
            and item.parent().element_info.control_type == "MenuBar"
        ),
        None,
    )
    if recursos_humanos is None:
        raise RuntimeError("O menu Recursos Humanos não foi encontrado no IWOS.")

    recursos_humanos.click_input()
    time.sleep(0.5)
    quadro_funcionarios = next(
        (
            item
            for item in recursos_humanos.descendants(control_type="MenuItem")
            if item.window_text() == "Quadro de Funcionários"
        ),
        None,
    )
    if quadro_funcionarios is None:
        raise RuntimeError(
            "O item Quadro de Funcionários não foi encontrado em Recursos Humanos."
        )

    quadro_funcionarios.click_input()
    janela.child_window(
        auto_id="gridFuncionario",
        control_type="Table",
    ).wait("exists visible ready", timeout=30)


def adicionar_coluna_demissao(janela):
    """Adiciona Demissão pelo Column Chooser, caso ainda esteja oculta."""
    grid = janela.child_window(
        auto_id="gridFuncionario",
        control_type="Table",
    )
    if grid.descendants(title="Demissão", control_type="Header"):
        print("A coluna Demissão já está visível.")
        return

    cabecalhos_id = grid.descendants(title="ID", control_type="Header")
    if not cabecalhos_id:
        raise RuntimeError("O cabeçalho ID não foi encontrado na grade.")

    cabecalhos_id[0].click_input(button="right")
    menu_contexto = Desktop(backend="uia").window(title="Contexto")
    opcao_column_chooser = menu_contexto.child_window(
        title="Column Chooser",
        control_type="MenuItem",
    )
    opcao_column_chooser.wait("exists visible enabled", timeout=10)
    opcao_column_chooser.click_input()

    customization = janela.child_window(
        title="Customization",
        control_type="Window",
    )
    customization.wait("exists visible ready", timeout=10)
    janela_customization = customization.wrapper_object()
    coluna_demissao = next(
        (
            item
            for item in janela_customization.descendants(control_type="ListItem")
            if item.window_text() == "Demissão"
        ),
        None,
    )
    if coluna_demissao is None:
        raise RuntimeError("A coluna Demissão não foi encontrada no Column Chooser.")

    coluna_demissao.double_click_input()
    limite = time.monotonic() + 10
    while time.monotonic() < limite:
        if grid.descendants(title="Demissão", control_type="Header"):
            break
        time.sleep(0.5)
    else:
        raise RuntimeError("A coluna Demissão não foi adicionada à grade.")

    botoes_fechar = janela_customization.descendants(
        title="Fechar",
        control_type="Button",
    )
    if not botoes_fechar:
        raise RuntimeError("O botão Fechar do Column Chooser não foi encontrado.")
    botoes_fechar[0].click_input()
    print("Coluna Demissão adicionada à grade.")


def localizar_controle_visivel(janela, auto_id, control_type):
    controles = [
        controle
        for controle in janela.descendants(control_type=control_type)
        if controle.element_info.automation_id == auto_id
        and controle.is_visible()
        and controle.rectangle().width() > 0
        and controle.rectangle().height() > 0
    ]
    return controles[0] if controles else None


def selecionar_aba_modulo(janela, titulo_aba, controle_conteudo):
    """Seleciona uma aba já aberta no gerenciador de módulos do IWOS."""
    gerenciador = janela.child_window(
        auto_id="moduleManager",
        control_type="Tab",
    )
    candidatos = [
        item
        for item in gerenciador.descendants(
            title=titulo_aba,
            control_type="TabItem",
        )
        if item.rectangle().width() > 0 and item.rectangle().height() > 0
    ]
    if not candidatos:
        return False

    # A aba de navegação tem cerca de 20 px de altura; a página selecionada
    # também é exposta como TabItem, mas ocupa quase toda a janela.
    aba = min(candidatos, key=lambda item: item.rectangle().height())
    aba.click_input()
    limite = time.monotonic() + 15
    while time.monotonic() < limite:
        if localizar_controle_visivel(janela, controle_conteudo, "Pane"):
            return True
        time.sleep(0.5)
    raise RuntimeError(f"O conteúdo da aba {titulo_aba} não ficou visível.")


def abrir_painel_rh(janela, painel):
    """Abre Recursos Humanos > Painéis > Faltas/Atestados."""
    if selecionar_aba_modulo(janela, painel["aba"], painel["controle"]):
        return

    janela.set_focus()
    recursos_humanos = next(
        (
            item
            for item in janela.descendants(control_type="MenuItem")
            if item.window_text() == "Recursos Humanos"
            and item.parent().element_info.control_type == "MenuBar"
        ),
        None,
    )
    if recursos_humanos is None:
        raise RuntimeError("O menu Recursos Humanos não foi encontrado no IWOS.")

    recursos_humanos.click_input()
    time.sleep(0.3)
    menu_paineis = next(
        (
            item
            for item in recursos_humanos.descendants(control_type="MenuItem")
            if item.window_text() == "Painéis"
        ),
        None,
    )
    if menu_paineis is None:
        raise RuntimeError("O submenu Painéis não foi encontrado em Recursos Humanos.")

    menu_paineis.click_input()
    time.sleep(0.3)
    item_painel = next(
        (
            item
            for item in menu_paineis.descendants(control_type="MenuItem")
            if item.window_text() == painel["menu"]
        ),
        None,
    )
    if item_painel is None:
        raise RuntimeError(
            f"O item {painel['menu']} não foi encontrado em Recursos Humanos > Painéis."
        )

    item_painel.click_input()
    janela.child_window(
        auto_id=painel["controle"],
        control_type="Pane",
    ).wait("exists visible ready", timeout=30)


def localizar_checkbox_periodo(janela):
    for controle in janela.descendants(control_type="CheckBox"):
        if (
            controle.window_text() == "Earlier this year"
            and controle.rectangle().width() > 0
            and controle.rectangle().height() > 0
        ):
            return controle
    return None


def aplicar_filtro_ano(janela, painel):
    grid = localizar_controle_visivel(janela, painel["grid"], "Table")
    if grid is None:
        raise RuntimeError(f"A grade de {painel['aba']} não está visível.")
    cabecalhos = grid.descendants(
        title=painel["coluna_data"],
        control_type="Header",
    )
    if not cabecalhos:
        raise RuntimeError(
            f"A coluna {painel['coluna_data']} não foi encontrada em {painel['aba']}."
        )

    cabecalho = cabecalhos[0]
    retangulo = cabecalho.rectangle()
    coordenadas_filtro = (retangulo.width() - 5, retangulo.height() // 2)
    cabecalho.move_mouse_input(coords=coordenadas_filtro)
    time.sleep(0.3)
    cabecalho.click_input(coords=coordenadas_filtro)

    limite = time.monotonic() + 10
    checkbox = None
    while time.monotonic() < limite:
        checkbox = localizar_checkbox_periodo(janela)
        if checkbox is not None:
            break
        time.sleep(0.5)
    if checkbox is None:
        raise RuntimeError(
            f"O filtro Earlier this year não foi encontrado em {painel['aba']}."
        )

    if checkbox.get_toggle_state() == 0:
        checkbox.click_input()
    else:
        checkbox.type_keys("{ESC}")


def localizar_item_empresa(nome_empresa, retangulo_combo):
    """Localiza o item realmente aberto no popup do ComboBox DevExpress."""
    for janela_desktop in Desktop(backend="uia").windows(visible_only=True):
        try:
            itens = janela_desktop.descendants(control_type="ListItem")
        except Exception:
            continue

        for item in itens:
            retangulo = item.rectangle()
            if (
                item.window_text().casefold() == nome_empresa.casefold()
                and retangulo.width() > 0
                and retangulo.height() > 0
                and retangulo.top >= retangulo_combo.bottom
                and abs(retangulo.left - retangulo_combo.left) <= 10
            ):
                return item
    return None


def aguardar_resultado_busca(nome_empresa):
    """Aguarda o IWOS fechar a tela durante a consulta e reconstruí-la."""
    inicio = time.monotonic()
    limite = inicio + TEMPO_LIMITE_BUSCA
    janela_sumiu = False
    assinatura_anterior = None
    verificacoes_estaveis = 0

    while time.monotonic() < limite:
        candidata = localizar_janela_iwos()
        if candidata is None:
            janela_sumiu = True
            time.sleep(0.5)
            continue

        # Em algumas versões a janela não desaparece; cinco segundos evitam
        # reutilizar a grade anterior antes que a consulta seja aplicada.
        if not janela_sumiu and time.monotonic() - inicio < 5:
            time.sleep(0.5)
            continue

        try:
            app = Application(backend="uia").connect(handle=candidata.handle)
            janela = app.window(handle=candidata.handle)
            grade = janela.child_window(
                auto_id="gridFuncionario",
                control_type="Table",
            )
            grade.wait("exists visible ready", timeout=2)
            combo = janela.child_window(
                auto_id="comboEmpresa",
                control_type="ComboBox",
            )
            ids_visiveis = tuple(
                controle.window_text()
                for controle in grade.descendants(control_type="DataItem")
                if controle.window_text().startswith("ID row ")
            )
            if (
                combo.selected_text().casefold() == nome_empresa.casefold()
                and ids_visiveis
            ):
                if ids_visiveis == assinatura_anterior:
                    verificacoes_estaveis += 1
                else:
                    assinatura_anterior = ids_visiveis
                    verificacoes_estaveis = 1

            # Só continua após três leituras iguais da primeira página.
            if verificacoes_estaveis >= 3:
                return app, janela
        except Exception:
            assinatura_anterior = None
            verificacoes_estaveis = 0
        time.sleep(1)

    raise RuntimeError(
        f"A busca da empresa {nome_empresa} não terminou em "
        f"{TEMPO_LIMITE_BUSCA} segundos."
    )


def selecionar_e_buscar_empresa(janela, nome_empresa):
    combo = janela.child_window(
        auto_id="comboEmpresa",
        control_type="ComboBox",
    )
    combo.child_window(title="Open", control_type="Button").click_input()
    time.sleep(0.5)

    item_empresa = localizar_item_empresa(nome_empresa, combo.rectangle())
    if item_empresa is None:
        raise RuntimeError(f"A empresa {nome_empresa} não foi encontrada na lista.")
    item_empresa.click_input()

    if combo.selected_text().casefold() != nome_empresa.casefold():
        raise RuntimeError(f"Não foi possível selecionar a empresa {nome_empresa}.")

    print(f"Empresa selecionada: {nome_empresa}. Executando a busca...")
    janela.child_window(
        auto_id="buttonGerar",
        control_type="Button",
    ).click_input()
    return aguardar_resultado_busca(nome_empresa)


def aguardar_resultado_painel(painel):
    inicio = time.monotonic()
    limite = inicio + TEMPO_LIMITE_BUSCA
    janela_sumiu = False
    assinatura_anterior = None
    verificacoes_estaveis = 0

    while time.monotonic() < limite:
        candidata = localizar_janela_iwos()
        if candidata is None:
            janela_sumiu = True
            time.sleep(0.5)
            continue

        if not janela_sumiu and time.monotonic() - inicio < 5:
            time.sleep(0.5)
            continue

        try:
            app = Application(backend="uia").connect(handle=candidata.handle)
            janela = app.window(handle=candidata.handle)
            if not selecionar_aba_modulo(
                janela,
                painel["aba"],
                painel["controle"],
            ):
                raise RuntimeError(f"A aba {painel['aba']} não está aberta.")

            grade = localizar_controle_visivel(
                janela,
                painel["grid"],
                "Table",
            )
            if grade is None:
                raise RuntimeError(f"A grade de {painel['aba']} não está visível.")
            assinatura = tuple(
                controle.window_text()
                for controle in grade.descendants(control_type="DataItem")
            )
            if assinatura == assinatura_anterior:
                verificacoes_estaveis += 1
            else:
                assinatura_anterior = assinatura
                verificacoes_estaveis = 1

            if verificacoes_estaveis >= 3:
                return app, janela
        except Exception:
            assinatura_anterior = None
            verificacoes_estaveis = 0
        time.sleep(1)

    raise RuntimeError(
        f"A busca de {painel['aba']} não terminou em "
        f"{TEMPO_LIMITE_BUSCA} segundos."
    )


def filtrar_e_buscar_painel(janela, painel):
    if not selecionar_aba_modulo(
        janela,
        painel["aba"],
        painel["controle"],
    ):
        raise RuntimeError(f"A aba {painel['aba']} não está aberta.")

    aplicar_filtro_ano(janela, painel)
    print(
        f"{painel['aba']}: filtro Earlier this year aplicado. "
        "Executando a busca..."
    )
    controle = localizar_controle_visivel(janela, painel["controle"], "Pane")
    botoes_busca = [
        botao
        for botao in controle.descendants(control_type="Button")
        if botao.element_info.automation_id == "buttonGerar"
        and botao.is_visible()
    ]
    if not botoes_busca:
        raise RuntimeError(f"O botão de busca não foi encontrado em {painel['aba']}.")
    botoes_busca[0].click_input()
    return aguardar_resultado_painel(painel)


def extrair_grade_painel(janela, painel):
    """Extrai todas as colunas da grade principal, exceto Ações."""
    grid = localizar_controle_visivel(janela, painel["grid"], "Table")
    if grid is None:
        raise RuntimeError(f"A grade de {painel['aba']} não está visível.")
    cabecalhos = sorted(
        (
            cabecalho
            for cabecalho in grid.descendants(control_type="Header")
            if cabecalho.window_text()
            and cabecalho.window_text().casefold() != "ações"
        ),
        key=lambda cabecalho: cabecalho.rectangle().left,
    )
    colunas = tuple(cabecalho.window_text() for cabecalho in cabecalhos)
    if not colunas:
        raise RuntimeError(f"Nenhuma coluna foi encontrada em {painel['aba']}.")

    coluna_referencia = "ID" if "ID" in colunas else colunas[0]

    def obter_celulas():
        return grid.descendants(control_type="DataItem")

    barras_verticais = [
        barra
        for barra in grid.descendants(control_type="ScrollBar")
        if barra.rectangle().height() > barra.rectangle().width()
    ]
    barra_vertical = barras_verticais[0] if barras_verticais else None

    def obter_botao_rolagem(titulo):
        if barra_vertical is None:
            return None
        for botao in barra_vertical.descendants(
            title=titulo,
            control_type="Button",
        ):
            retangulo = botao.rectangle()
            if retangulo.width() > 0 and retangulo.height() > 0:
                return botao
        return None

    def obter_linhas_visiveis(celulas):
        linhas = {}
        for celula in celulas:
            titulo = celula.window_text()
            coluna, separador, numero_linha = titulo.rpartition(" row ")
            if not separador or coluna not in colunas or not numero_linha.isdigit():
                continue
            try:
                valor = celula.iface_value.CurrentValue
            except Exception:
                valor = ""
            linhas.setdefault(int(numero_linha), {})[coluna] = valor
        return linhas

    def assinatura_linhas(celulas):
        return tuple(
            celula.window_text()
            for celula in celulas
            if celula.window_text().startswith(f"{coluna_referencia} row ")
        )

    def clicar_e_aguardar_pagina(botao, assinatura_anterior):
        botao.click_input()
        limite_atualizacao = time.monotonic() + 10
        while time.monotonic() < limite_atualizacao:
            time.sleep(0.5)
            nova_assinatura = assinatura_linhas(obter_celulas())
            if nova_assinatura and nova_assinatura != assinatura_anterior:
                return True
        return False

    celulas_iniciais = obter_celulas()
    if not obter_linhas_visiveis(celulas_iniciais):
        return colunas, []

    if barra_vertical is not None:
        thumbs = barra_vertical.descendants(control_type="Thumb")
        if thumbs:
            retangulo_barra = barra_vertical.rectangle()
            thumbs[0].drag_mouse_input(
                dst=(
                    retangulo_barra.left + retangulo_barra.width() // 2,
                    retangulo_barra.top + 20,
                )
            )
            time.sleep(0.8)

    # Usa os botões reais da barra vertical. O Page Up/Page Down enviado ao
    # teclado pode saltar ao fim quando o DevExpress preserva a linha focada.
    for _ in range(300):
        celulas_visiveis = obter_celulas()
        titulos = {celula.window_text() for celula in celulas_visiveis}
        if f"{coluna_referencia} row 0" in titulos:
            break
        botao_subir = obter_botao_rolagem("Page Up")
        if botao_subir is None:
            break
        clicar_e_aguardar_pagina(
            botao_subir,
            assinatura_linhas(celulas_visiveis),
        )

    registros = {}
    tentativas_sem_novidade = 0
    for _ in range(300):
        celulas_visiveis = obter_celulas()
        linhas_visiveis = obter_linhas_visiveis(celulas_visiveis)
        novas_linhas = 0
        for numero_linha, dados in linhas_visiveis.items():
            if numero_linha not in registros:
                novas_linhas += 1
            registros[numero_linha] = dados

        tentativas_sem_novidade = (
            tentativas_sem_novidade + 1 if novas_linhas == 0 else 0
        )
        botao_descer = obter_botao_rolagem("Page Down")
        if botao_descer is None or tentativas_sem_novidade >= 3:
            break
        pagina_mudou = clicar_e_aguardar_pagina(
            botao_descer,
            assinatura_linhas(celulas_visiveis),
        )
        if not pagina_mudou:
            tentativas_sem_novidade += 1

    linhas = [
        {coluna: registros[numero].get(coluna, "") for coluna in colunas}
        for numero in sorted(registros)
    ]
    return colunas, linhas


def extrair_registros(janela):
    grid = janela.child_window(
        auto_id="gridFuncionario",
        control_type="Table",
    )

    def rolar_pagina(tecla):
        grid.set_focus()
        grid.type_keys(tecla)
        time.sleep(0.8)

    # Volta ao topo para garantir que nenhuma linha fique de fora.
    for _ in range(200):
        ids_visiveis = [
            celula.window_text()
            for celula in grid.descendants(control_type="DataItem")
            if celula.window_text().startswith("ID row ")
        ]
        if "ID row 0" in ids_visiveis:
            break
        rolar_pagina("{PGUP}")

    registros = {}
    tentativas_sem_novidade = 0

    for _ in range(200):
        linhas_visiveis = {}
        for celula in grid.descendants(control_type="DataItem"):
            correspondencia = PADRAO_CELULA.match(celula.window_text())
            if not correspondencia:
                continue

            coluna, numero_linha = correspondencia.groups()
            try:
                valor = celula.iface_value.CurrentValue
            except Exception:
                valor = ""
            linhas_visiveis.setdefault(int(numero_linha), {})[coluna] = valor

        novas_linhas = 0
        for numero_linha, dados in linhas_visiveis.items():
            if numero_linha not in registros:
                novas_linhas += 1
            registros[numero_linha] = dados

        tentativas_sem_novidade = (
            tentativas_sem_novidade + 1 if novas_linhas == 0 else 0
        )
        if tentativas_sem_novidade >= 3:
            break
        rolar_pagina("{PGDN}")

    if not registros:
        raise RuntimeError("Nenhum funcionário foi encontrado na grade do IWOS.")

    return registros


def preparar_lista_exportacao(registros):
    lista = []
    for numero_linha in sorted(registros):
        dados = registros[numero_linha]
        registro = {
            coluna: dados.get(coluna, "")
            for coluna in COLUNAS_EXPORTACAO
        }
        demissao = str(registro["Demissão"] or "").strip()
        if not demissao or demissao.casefold() == "demissão":
            registro["Demissão"] = "N/A"
        lista.append(registro)
    return lista


def salvar_excel(arquivo_excel, planilhas):
    with pd.ExcelWriter(arquivo_excel, engine="openpyxl") as writer:
        for nome_sheet, (colunas, linhas) in planilhas.items():
            pd.DataFrame(linhas, columns=colunas).to_excel(
                writer,
                index=False,
                sheet_name=nome_sheet,
            )

    workbook = load_workbook(arquivo_excel)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    nomes_tabelas = {
        "Funcionários": "TabelaFuncionarios",
        "Faltas": "TabelaFaltas",
        "Atestados": "TabelaAtestados",
    }
    larguras_por_titulo = {
        "ID": 14,
        "Nome": 32,
        "Centro": 22,
        "Funcionário": 32,
        "Função": 26,
        "Admissão": 16,
        "Disponibilidade": 24,
        "Demissão": 16,
        "Data": 16,
        "Motivo": 32,
        "Status": 18,
        "Patologia": 38,
        "Início": 16,
        "Término": 16,
        "Dias": 12,
    }

    for indice, worksheet in enumerate(workbook.worksheets):
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.sheet_view.showGridLines = False
        worksheet.sheet_properties.tabColor = ("1F4E78", "548235", "A64D79")[
            indice % 3
        ]
        worksheet.print_title_rows = "1:1"
        worksheet.page_setup.orientation = "landscape"
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_margins.left = 0.25
        worksheet.page_margins.right = 0.25
        worksheet.page_margins.top = 0.5
        worksheet.page_margins.bottom = 0.5
        worksheet.row_dimensions[1].height = 24

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell_width = larguras_por_titulo.get(str(cell.value), 22)
            worksheet.column_dimensions[cell.column_letter].width = cell_width

        for linha in worksheet.iter_rows(min_row=2):
            for cell in linha:
                cell.alignment = Alignment(vertical="center")

        if worksheet.max_row > 1:
            tabela = Table(
                displayName=nomes_tabelas[worksheet.title],
                ref=worksheet.dimensions,
            )
            tabela.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
            worksheet.add_table(tabela)

    workbook.save(arquivo_excel)
    print(f"Arquivo Excel gerado: {arquivo_excel}")


def criar_pasta_exportacao(data_execucao):
    pasta = Path(__file__).with_name("exports") / data_execucao.strftime(
        "%d %m %Y - %H %M"
    )
    pasta.mkdir(parents=True, exist_ok=True)
    return pasta


def exportar_excel(registros, codigo_empresa, pasta_exportacao, data_execucao):
    data_relatorio = data_execucao.strftime("%Y-%m-%d")
    arquivo_excel = pasta_exportacao / (
        f"IWOS-{codigo_empresa}-{data_relatorio}.xlsx"
    )
    salvar_excel(
        arquivo_excel,
        {
            "Funcionários": (
                COLUNAS_EXPORTACAO,
                preparar_lista_exportacao(registros),
            )
        },
    )


def exportar_excel_painel(
    colunas,
    linhas,
    painel,
    pasta_exportacao,
    data_execucao,
):
    data_relatorio = data_execucao.strftime("%Y-%m-%d")
    arquivo_excel = pasta_exportacao / (
        f"IWOS-{painel['sheet'].upper()}-{data_relatorio}.xlsx"
    )
    salvar_excel(
        arquivo_excel,
        {painel["sheet"]: (colunas, linhas)},
    )


def fechar_iwos():
    janela_candidata = localizar_janela_iwos()
    if janela_candidata is None:
        return

    app = Application(backend="uia").connect(handle=janela_candidata.handle)
    app.window(handle=janela_candidata.handle).close()

    limite = time.monotonic() + 10
    while time.monotonic() < limite:
        if localizar_janela_iwos() is None:
            print("IWOS fechado automaticamente.")
            return
        time.sleep(0.5)
    raise RuntimeError("O processo terminou, mas não foi possível fechar o IWOS.")


def main():
    data_execucao = datetime.now()
    pasta_exportacao = criar_pasta_exportacao(data_execucao)
    _app, janela = obter_janela_iwos()
    abrir_quadro_funcionarios(janela)
    adicionar_coluna_demissao(janela)

    # Os painéis são globais: ficam abertos antes das consultas por empresa,
    # mas serão pesquisados e exportados apenas uma vez.
    for painel in PAINEIS:
        abrir_painel_rh(janela, painel)

    for nome_empresa, codigo_empresa in EMPRESAS:
        if not selecionar_aba_modulo(
            janela,
            "RH - Quadro de Funcionários",
            "QuadroFuncionarioControl",
        ):
            raise RuntimeError("A aba RH - Quadro de Funcionários não está aberta.")
        _app, janela = selecionar_e_buscar_empresa(janela, nome_empresa)
        registros = extrair_registros(janela)
        print(f"{nome_empresa}: {len(registros)} registros identificados.")
        exportar_excel(
            registros,
            codigo_empresa,
            pasta_exportacao,
            data_execucao,
        )

    for painel in PAINEIS:
        _app, janela = filtrar_e_buscar_painel(janela, painel)
        colunas, linhas = extrair_grade_painel(janela, painel)
        print(f"{painel['aba']}: {len(linhas)} registros identificados.")
        exportar_excel_painel(
            colunas,
            linhas,
            painel,
            pasta_exportacao,
            data_execucao,
        )

    fechar_iwos()


if __name__ == "__main__":
    main()
