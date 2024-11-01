# coding: utf-8
from __future__ import unicode_literals

import uno
from com.sun.star.script.provider import XScript  # type: ignore
# import unicodedata
from calendar import monthrange

# import apso_utils  # type: ignore

TIPO_TEC_ADM = "Técnico Administrativo"
TIPO_DOC = "Docente"

PLA_TIT = "Titulares"
PLA_SUP = "Suplentes"
PLA_OCO = "Ocorrências"
PLA_MODL = "Modelo"

titulares = {}
valores_grs = {}


class Servidor:
    def __init__(self):
        super().__init__()
        self.nome_formatado = None
        self.nome = None
        self.matricula = None
        self.funcao_titular = None
        self.lotacao = None
        self.deduzir_insalubridade = False
        self.funcao_confianca = None
        self.valor_275 = None
        self.valor_402 = None
        self.valor_404 = None
        self.categoria = None


class Substituto(Servidor):
    def __init__(self):
        super().__init__()
        self.ordem_substituicao = None


class Titular(Servidor):
    def __init__(self):
        super().__init__()
        self.substitutos = []


class Ocorrencia:
    def __init__(self, titular):
        self.titular = titular
        self.motivo_impedimento = None
        self.dias_ocorrencia = 0
        self.mes_ocorrencia = None
        self.ano_ocorrencia = None
        self.periodo = None


def get_doc():
    return XSCRIPTCONTEXT.getDocument()  # type: ignore


def get_planilhas():
    return get_doc().getSheets()


def get_nome_mes(numero_mes):
    mes_int = int(numero_mes)
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril",
        "Maio", "Junho", "Julho", "Agosto",
        "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    return meses[mes_int - 1]

# normaliza uma string removendo acentos
# removendo espaços extras e deixando tudo lowercase


def str_norm(string):
    # nfkd_form = unicodedata.normalize('NFKD', string)
    # ascii = nfkd_form.encode('ASCII', 'ignore')
    # ascii_string = ascii.decode()
    return " ".join(string.split()).strip().lower()


# compara se dois nomes são iguais
# depois de normaliza-los
def nomes_iguais(nome1, nome2):
    return str_norm(nome1) == str_norm(nome2)


def preenche_substitutos():
    planilhas = get_planilhas()
    pla = planilhas.getByName(PLA_SUP)

    matricula = None
    linha = 2
    while matricula != "":
        matricula = pla.getCellRangeByName(f"A{linha}").getString()

        if matricula != "":
            matricula_titular = pla.getCellRangeByName(f"L{linha}").getString()
            substituto = cria_substituto_da_planilha(
                pla, linha,
            )
            titulares[matricula_titular].substitutos.append(substituto)

        linha += 1

# cria um objeto Servidor a partir de uma linha da planilha
# serve tanto para titulares quanto para substitutos
# porém campos específicos são capturados nas funções
# específicas de criação de titulares e substitutos


def cria_servidor_da_planilha(
    planilha,
    linha,
    classe_objeto_servidor,
    col_matricula="A",
    col_nome="B",
    col_lotacao="C",
    col_funcao_confianca="D",
    col_categoria="E",
    col_ins="F",
    col_valor_275="G",
    col_valor_402="H",
    col_valor_404="I",
):
    servidor = classe_objeto_servidor()
    servidor.matricula = planilha.getCellRangeByName(
        f"{col_matricula}{linha}"
    ).getString()
    servidor.nome_formatado = planilha.getCellRangeByName(
        f"{col_nome}{linha}").getString()

    servidor.nome = str_norm(
        planilha.getCellRangeByName(f"{col_nome}{linha}").getString()
    )
    servidor.lotacao = planilha.getCellRangeByName(
        f"{col_lotacao}{linha}").getString()
    servidor.funcao_confianca = planilha.getCellRangeByName(
        f"{col_funcao_confianca}{linha}"
    ).getString()
    servidor.deduzir_insalubridade = (
        str_norm(planilha.getCellRangeByName(
            f"{col_ins}{linha}").getString()) == "sim"
    )
    servidor.categoria = (
        planilha.getCellRangeByName(f"{col_categoria}{linha}").getString()
    )

    servidor.valor_275 = (
        planilha.getCellRangeByName(f"{col_valor_275}{linha}").getValue()
    )
    servidor.valor_402 = (
        planilha.getCellRangeByName(f"{col_valor_402}{linha}").getValue()
    )
    servidor.valor_404 = (
        planilha.getCellRangeByName(f"{col_valor_404}{linha}").getValue()
    )

    return servidor


# cria um objeto Titular a partir de uma linha da planilha
def cria_titular_da_planilha(
    planilha,
    linha,
):
    return cria_servidor_da_planilha(planilha, linha, Titular)


# cria um objeto Substituto a partir de uma linha da planilha
def cria_substituto_da_planilha(
    planilha,
    linha,
    col_ordem_substituicao="J",
):
    substituto = cria_servidor_da_planilha(planilha, linha, Substituto)
    substituto.ordem_substituicao = int(
        planilha.getCellRangeByName(
            f"{col_ordem_substituicao}{linha}").getValue()
    )

    return substituto


# preenche a lista de titulares
# percorrendo todas as planilhas de titulares
def preenche_titulares():
    planilhas = get_planilhas()
    pla = planilhas.getByName(PLA_TIT)

    matricula = None
    linha = 2
    while matricula != "":
        matricula = pla.getCellRangeByName(f"A{linha}").getString()

        if matricula != "":
            titular = cria_titular_da_planilha(
                pla,
                linha,
            )

            titulares[matricula] = titular
        linha += 1


# preenche a lista de valores de GRs
def preenche_valores_grs():
    planilhas = get_planilhas()
    pla = planilhas.getByName("Tabela de GR")

    cargo = None
    linha = 3
    while cargo != "":
        cargo = str_norm(pla.getCellRangeByName(f"A{linha}").getString())

        if cargo != "":
            valores_grs[str_norm(cargo)] = pla.getCellRangeByName(
                f"D{linha}").getString()
        linha += 1


def ordena_substitutos():
    for matricula in titulares:
        titulares[matricula].substitutos = sorted(
            titulares[matricula].substitutos, key=lambda sub: sub.ordem_substituicao)


# gera as planilhas de substituição
# percorrendo a planilha de ocorrências
# e para cada titular presente na planilha de ocorrências
# cria uma planilha de substituição com os dados
# do titular e de seu substituto
def percorre_planilha_ocorrencias():
    planilhas = get_planilhas()
    pla_oco = planilhas.getByName(PLA_OCO)

    linha = 2
    matricula = None
    while matricula != "":
        matricula = pla_oco.getCellRangeByName(f"D{linha}").getString()

        # verifica se a matrícula é válida (evita #N/A e erros de digitação)
        if not matricula.replace("-", "").isnumeric():
            linha += 1
            continue

        ocorrencia = None
        if matricula in titulares:
            titular = titulares[matricula]
            ocorrencia = Ocorrencia(titular)
            ocorrencia.motivo_impedimento = pla_oco.getCellRangeByName(
                f"F{linha}").getString()
            ocorrencia.dias_ocorrencia = int(pla_oco.getCellRangeByName(
                f"I{linha}").getString())
            ocorrencia.mes_ocorrencia = pla_oco.getCellRangeByName(f"G{linha}").getString(
            ).split("/")[1]
            ocorrencia.ano_ocorrencia = pla_oco.getCellRangeByName(f"G{linha}").getString(
            ).split("/")[2]
            ocorrencia.periodo = "de " + pla_oco.getCellRangeByName(f"G{linha}").getString(
            ) + " a " + pla_oco.getCellRangeByName(f"H{linha}").getString()

            gera_planilha_substituicao(ocorrencia)

            # # verifica se o titular tem substitutos que também são titulares,
            # # e gera planilhas de substituição para eles, recursivamente
            # titulares_a_verificar = [titular]
            # while len(titulares_a_verificar) > 0:
            #     titular_atual = titulares_a_verificar.pop()
            #     for substituto in titular_atual.substitutos:
            #         if substituto.matricula in titulares:
            #             novo_titular = titulares[substituto.matricula]
            #             nova_ocorrencia = Ocorrencia(novo_titular)
            #             nova_ocorrencia.motivo_impedimento = "Substituição"
            #             nova_ocorrencia.dias_ocorrencia = ocorrencia.dias_ocorrencia
            #             nova_ocorrencia.mes_ocorrencia = ocorrencia.mes_ocorrencia
            #             nova_ocorrencia.ano_ocorrencia = ocorrencia.ano_ocorrencia
            #             nova_ocorrencia.periodo = ocorrencia.periodo
            #             nova_ocorrencia.titular = novo_titular

            #             titulares_a_verificar.append(novo_titular)
            #             # apso_utils.msgbox(novo_titular.nome_formatado + " - " + nova_ocorrencia.motivo_impedimento)
            #             arq_log.write(f"{novo_titular.nome_formatado}{nova_ocorrencia.motivo_impedimento}")
            #             arq_log.flush()

            #             gera_planilha_substituicao(nova_ocorrencia)
        linha += 1


def gera_planilha_substituicao(ocorrencia):
    planilhas = get_planilhas()

    titular = ocorrencia.titular

    # se o titular não tem substitutos, não é necessário gerar planilhas
    if len(titular.substitutos) == 0:
        return

    # insere uma planilha no final com o nome do primeiro substituto do titular
    if not planilhas.hasByName(titular.substitutos[0].nome_formatado):
        # planilhas.insertNewByName(titular.nome_formatado, planilhas.Count)
        planilhas.copyByName(
            PLA_MODL, titular.substitutos[0].nome_formatado, planilhas.Count)

    planilha_substituicao = planilhas.getByName(
        titular.substitutos[0].nome_formatado)

    # apaga o botão Gerar Planilhas
    colunas = planilha_substituicao.getColumns()
    colunas.removeByIndex(11, 2)

    pla_modelo = planilhas.getByName(PLA_MODL)

    cels_origem = pla_modelo.getCellRangeByName("A:J")
    cel_dest = planilha_substituicao.getCellRangeByName("A1")
    planilha_substituicao.copyRange(
        cel_dest.CellAddress, cels_origem.RangeAddress)

    primeiro_subs = titular.substitutos[0]
    planilha_substituicao.getCellRangeByName(
        "E11").setString(primeiro_subs.matricula)
    planilha_substituicao.getCellRangeByName(
        "F11").setString(primeiro_subs.categoria)
    planilha_substituicao.getCellRangeByName(
        "E12").setString(primeiro_subs.nome_formatado)

    # lista os demais substitutos na coluna I
    # para facilitar a alteração manual caso
    # seja necessária
    if len(titular.substitutos) > 1:
        for i in range(1, len(titular.substitutos)):
            subs = titular.substitutos[i]
            planilha_substituicao.getCellRangeByName(
                f"I{11+i}").setString(subs.nome_formatado)
            planilha_substituicao.getCellRangeByName(
                f"J{11+i}").setString(subs.matricula)

    planilha_substituicao.getCellRangeByName(
        "E16").setString(titular.matricula)
    planilha_substituicao.getCellRangeByName(
        "F16").setString(titular.categoria)
    planilha_substituicao.getCellRangeByName(
        "E17").setString(titular.nome_formatado)
    planilha_substituicao.getCellRangeByName(
        "E18").setString(titular.lotacao)
    planilha_substituicao.getCellRangeByName(
        "E19").setString(titular.funcao_confianca)

    planilha_substituicao.getCellRangeByName(
        "F22").setString(ocorrencia.motivo_impedimento)

    planilha_substituicao.getCellRangeByName(
        "F23").setString(ocorrencia.periodo)

    planilha_substituicao.getCellRangeByName("G28").setFormula(
        "=" + str(valores_grs[str_norm(primeiro_subs.funcao_confianca)])
    )

    planilha_substituicao.getCellRangeByName("G32").setValue(
        primeiro_subs.valor_275
    )
    planilha_substituicao.getCellRangeByName("G33").setValue(
        primeiro_subs.valor_402
    )
    planilha_substituicao.getCellRangeByName("G34").setValue(
        primeiro_subs.valor_404
    )

    planilha_substituicao.getCellRangeByName("F41").setValue(
        ocorrencia.dias_ocorrencia
    )

    planilha_substituicao.getCellRangeByName("C46").setString(
        get_nome_mes(ocorrencia.mes_ocorrencia) +
        "(" + str(ocorrencia.dias_ocorrencia) + ")"
    )

    planilha_substituicao.getCellRangeByName("E46").setFormula(
        f"=(G38/30)*{ocorrencia.dias_ocorrencia}"
    )

    if primeiro_subs.deduzir_insalubridade:
        planilha_substituicao.getCellRangeByName("G46").setString(
            "SIM ( X ) CÓDIGO 031- "
        )

        # mostra o número de dias a serem deduzidos se o SIM for marcado
        planilha_substituicao.getCellRangeByName("I46").setString(
            str(monthrange(int(ocorrencia.ano_ocorrencia), int(ocorrencia.mes_ocorrencia))[
                1] - ocorrencia.dias_ocorrencia)
        )
    else:
        planilha_substituicao.getCellRangeByName("G48").setString(
            "NÃO ( X )."
        )


def main(arg):
    preenche_titulares()
    preenche_substitutos()
    ordena_substitutos()
    preenche_valores_grs()
    percorre_planilha_ocorrencias()
