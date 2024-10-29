# coding: utf-8
from __future__ import unicode_literals

import uno
from com.sun.star.script.provider import XScript  # type: ignore
import unicodedata
from calendar import monthrange

# import apso_utils  # type: ignore

TIPO_TEC_ADM = "Técnico Administrativo"
TIPO_DOC = "Docente"

PLA_MODL = "Modelo"
PLA_OCOR = "Ocorrências"
PLA_CHEFE_DEP_TIT = "Chefes de Departamento - Titulares"
PLA_CHEFE_DEP_SUPL = "Chefes de Departamento - Suplentes"
PLA_COORD_GRAD_TIT = "Coordenadores Graduação - Titulares"
PLA_COORD_GRAD_SUPL = "Coordenadores Graduação - Suplentes"
PLA_COORD_POS_TIT = "Coordenadores Pós-Graduação - Titulares"
PLA_COORD_POS_SUPL = "Coordenadores Pós-Graduação - Suplentes"
PLA_TEC_ADM_TIT = "Tec. Adm. - Titulares"
PLA_TEC_ADM_SUB = "Tec. Adm. - Substitutos"

PLA_TIT = "Titulares"
PLA_SUP = "Suplentes"
PLA_OCO = "Ocorrências"

titulares = {}
valores_grs = {}


NOMES_PLANILHAS = {
    PLA_CHEFE_DEP_TIT: {"subs": PLA_CHEFE_DEP_SUPL, "tipo": TIPO_DOC},
    PLA_COORD_GRAD_TIT: {"subs": PLA_COORD_GRAD_SUPL, "tipo": TIPO_DOC},
    PLA_COORD_POS_TIT: {"subs": PLA_COORD_POS_SUPL, "tipo": TIPO_DOC},
    PLA_TEC_ADM_TIT: {"subs": PLA_TEC_ADM_SUB, "tipo": TIPO_TEC_ADM},
}


class Servidor:
    def __init__(self):
        super().__init__()
        self.nome_formatado = None
        self.nome = None
        self.tipo = None
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
        self.motivo_impedimento = None
        self.dias_ocorrencia = 0
        self.substitutos = []
        self.substitutos = []


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
    nfkd_form = unicodedata.normalize('NFKD', string)
    ascii = nfkd_form.encode('ASCII', 'ignore')
    ascii_string = ascii.decode()
    return " ".join(ascii_string.split()).strip().lower()


# compara se dois nomes são iguais
# depois de normaliza-los
def nomes_iguais(nome1, nome2):
    return str_norm(nome1) == str_norm(nome2)


def preenche_substitutos():
    planilhas = get_planilhas()
    for nome_planilha_titular in NOMES_PLANILHAS:

        pla = planilhas.getByName(
            NOMES_PLANILHAS[nome_planilha_titular]['subs'])

        nome_substituto = None
        linha = 2
        while nome_substituto != "":
            nome_substituto = str_norm(
                pla.getCellRangeByName(f"B{linha}").getString())

            if nome_substituto != "":
                nome_titular = str_norm(
                    pla.getCellRangeByName(f"L{linha}").getString())
                substituto = cria_substituto_da_planilha(
                    pla, linha, NOMES_PLANILHAS[nome_planilha_titular]['tipo'],
                    col_nome="B",
                    col_matricula="A",
                    col_lotacao="C",
                    col_funcao_titular="D",
                    col_funcao_confianca="E",

                    col_ins="G",
                )
                titulares[nome_titular].substitutos.append(substituto)

            linha += 1

# cria um objeto Servidor a partir de uma linha da planilha
# serve tanto para titulares quanto para substitutos
# porém campos específicos são capturados nas funções
# específicas de criação de titulares e substitutos


def cria_servidor_da_planilha(
    planilha,
    linha,
    tipo_servidor,
    classe_objeto_servidor,
    col_nome="B",
    col_lotacao="C",
    col_matricula="A",
    col_funcao_titular="D",
    col_funcao_confianca="E",
    col_ins="G",
    col_categoria="F",
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
    servidor.funcao_titular = planilha.getCellRangeByName(
        f"{col_funcao_titular}{linha}"
    ).getString()
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
    servidor.tipo = tipo_servidor
    return servidor


# cria um objeto Titular a partir de uma linha da planilha
def cria_titular_da_planilha(
    planilha,
    linha,
    tipo,
    col_nome="B",
    col_lotacao="C",
    col_matricula="A",
    col_funcao_titular="D",
    col_funcao_confianca="E",
    col_ins="F",
):
    return cria_servidor_da_planilha(planilha, linha, tipo, Titular)


# cria um objeto Substituto a partir de uma linha da planilha
def cria_substituto_da_planilha(
    planilha,
    linha,
    tipo,
    col_nome="B",
    col_lotacao="C",
    col_matricula="A",
    col_funcao_titular="D",
    col_funcao_confianca="E",
    col_ordem_substituicao="K",
    col_valor_275="H",
    col_valor_402="I",
    col_valor_404="J",
    col_ins="F",
):
    substituto = cria_servidor_da_planilha(planilha, linha, tipo, Substituto)
    substituto.ordem_substituicao = int(
        planilha.getCellRangeByName(
            f"{col_ordem_substituicao}{linha}").getValue()
    )
    substituto.valor_275 = planilha.getCellRangeByName(
        f"{col_valor_275}{linha}").getString()
    substituto.valor_402 = planilha.getCellRangeByName(
        f"{col_valor_402}{linha}").getString()
    substituto.valor_404 = planilha.getCellRangeByName(
        f"{col_valor_404}{linha}").getString()

    return substituto


# preenche a lista de titulares
# percorrendo todas as planilhas de titulares
def preenche_titulares():
    planilhas = get_planilhas()
    for nome_planilha_titular in NOMES_PLANILHAS:
        pla = planilhas.getByName(nome_planilha_titular)

        nome = None
        linha = 2
        while nome != "":
            nome = str_norm(pla.getCellRangeByName(f"B{linha}").getString())

            if nome != "":
                titular = cria_titular_da_planilha(
                    pla,
                    linha,
                    NOMES_PLANILHAS[nome_planilha_titular]["tipo"],
                )

                titulares[nome] = titular
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


# gera as planilhas de substituição
# percorrendo a planilha de ocorrências
# e para cada titular presente na planilha de ocorrências
# cria uma planilha de substituição com os dados
# do titular e de seu substituto
def gera_planilhas_substituicoes():
    planilhas = get_planilhas()
    pla_oco = planilhas.getByName(PLA_OCOR)

    linha = 2
    nome = None
    while nome != "":
        nome = str_norm(pla_oco.getCellRangeByName(f"C{linha}").getString())

        # sanity check:
        # verifica se o nome que consta na ocorrência está
        # na tabela de titulares
        if nome in titulares:
            titular = titulares[nome]
            titular.motivo_impedimento = pla_oco.getCellRangeByName(
                f"F{linha}").getString()

            titular.dias_ocorrencia = int(pla_oco.getCellRangeByName(
                f"I{linha}").getString())

            # insere uma planilha no final com o nome do titular
            if not planilhas.hasByName(titular.nome_formatado):
                #planilhas.insertNewByName(titular.nome_formatado, planilhas.Count)
                planilhas.copyByName(PLA_MODL, titular.nome_formatado, planilhas.Count)

            planilha_substituicao = planilhas.getByName(titular.nome)
            
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
            planilha_substituicao.getCellRangeByName(
                "E13").setString(primeiro_subs.funcao_titular)

            if len(titular.substitutos) > 1:
                for i in range(1, len(titular.substitutos)):
                    subs = titular.substitutos[i]
                    planilha_substituicao.getCellRangeByName(
                        f"I{11+i}").setString(subs.nome_formatado)

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

            periodo = "de " + pla_oco.getCellRangeByName(f"G{linha}").getString(
            ) + " a " + pla_oco.getCellRangeByName(f"H{linha}").getString()
            mes_ocorrencia = pla_oco.getCellRangeByName(f"G{linha}").getString(
            ).split("/")[1]
            ano_ocorrencia = pla_oco.getCellRangeByName(f"G{linha}").getString(
            ).split("/")[2]

            planilha_substituicao.getCellRangeByName(
                "F22").setString(titular.motivo_impedimento)

            planilha_substituicao.getCellRangeByName(
                "F23").setString(periodo)

            planilha_substituicao.getCellRangeByName("G28").setFormula(
                valores_grs[str_norm(primeiro_subs.funcao_confianca)]
            )

            planilha_substituicao.getCellRangeByName("G32").setFormula(
                primeiro_subs.valor_275
            )
            planilha_substituicao.getCellRangeByName("G33").setFormula(
                primeiro_subs.valor_402
            )
            planilha_substituicao.getCellRangeByName("G34").setFormula(
                primeiro_subs.valor_404
            )

            planilha_substituicao.getCellRangeByName("F41").setFormula(
                titular.dias_ocorrencia
            )

            planilha_substituicao.getCellRangeByName("C46").setString(
                f"{get_nome_mes(mes_ocorrencia)}({titular.dias_ocorrencia})"
            )

            planilha_substituicao.getCellRangeByName("E46").setFormula(
                f"=(G38/30)*{titular.dias_ocorrencia}"
            )
            

            if primeiro_subs.deduzir_insalubridade:
                planilha_substituicao.getCellRangeByName("G46").setFormula(
                    "SIM ( X ) CÓDIGO 031- "
                )

                # mostra o número de dias a serem deduzidos se o SIM for marcado
                planilha_substituicao.getCellRangeByName("I46").setFormula(
                    monthrange(int(ano_ocorrencia), int(mes_ocorrencia))[1] - titular.dias_ocorrencia
                )
            else:
                planilha_substituicao.getCellRangeByName("G48").setFormula(
                    "NÃO ( X )."
                )

        linha += 1


def main(arg):
    preenche_titulares()
    preenche_substitutos()
    preenche_valores_grs()
    gera_planilhas_substituicoes()