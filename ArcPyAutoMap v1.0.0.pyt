# -*- coding: utf-8 -*-

import sys  # type: ignore

reload(sys)  # type: ignore
sys.setdefaultencoding("utf-8")  # type: ignore

import io
import os
import re
import string
import getpass
import platform
import subprocess
import traceback
import unicodedata
import logging
import ctypes
import arcpy  # type: ignore
import math  # type: ignore
import yaml  # type: ignore
import datetime  # type: ignore
import psycopg2 # type: ignore

from openpyxl import load_workbook  # type: ignore
from docx import Document  # type: ignore


class Logger:

    def setup_logger(self):
        """Setup the logger."""
        logger = logging.getLogger("AAM")
        logger.setLevel(logging.DEBUG)

        # Create file handler which logs messages
        log_file_path = os.path.join(os.path.dirname(__file__), "log", "automap.log")
        fh = logging.FileHandler(log_file_path)
        fh.setLevel(logging.DEBUG)

        # Create formatter and add it to the handler
        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        fh.setFormatter(formatter)

        # Add handler to logger
        if not logger.handlers:
            logger.addHandler(fh)

        return logger


class ValoresConstantes:

    separador = "; "
    diretorio_raiz = os.path.dirname(os.path.abspath(__file__))

    diretorio_estilos = "{}\\estilos".format(diretorio_raiz)
    arquivo_xlsx_planilha_indv = "{}\\planilha\\modelo_planilha.xlsx".format(
        diretorio_raiz
    )
    arquivo_xlsx_controle = "{}\\planilha\\controle.xlsx".format(diretorio_raiz)
    arquivo_xlsx_controle_extra = "{}\\planilha\\controle_{}.xlsx".format(
        diretorio_raiz, datetime.datetime.now().strftime("%H%M%S")
    )
    arquivo_docx_despacho = "{}\\despacho\\modelo_despacho.docx".format(diretorio_raiz)
    arquivo_yaml = "{}\\dados\\formsfields.yaml".format(diretorio_raiz)
    camada_auxiliar_zee_2010 = r"AUXILIAR\ZEE_2010"
    camada_auxiliar_mzee_2008 = r"AUXILIAR\MZEE_2008"
    camada_auxiliar_area_limitante = r"AUXILIAR\AREA DE LIMITACAO"
    opcoes_municipios = [
        "ABAETETUBA",
        "ABEL FIGUEIREDO",
        "ACAR\xc1",
        "AFU\xc1",
        "\xc1GUA AZUL DO NORTE",
        "ALENQUER",
        "ALMEIRIM",
        "ALTAMIRA",
        "ANAJ\xc1S",
        "ANANINDEUA",
        "ANAPU",
        "AUGUSTO CORR\xcaA",
        "AURORA DO PAR\xc1",
        "AVEIRO",
        "BAGRE",
        "BAI\xc3O",
        "BANNACH",
        "BARCARENA",
        "BEL\xc9M",
        "BELTERRA",
        "BENEVIDES",
        "BOM JESUS DO TOCANTINS",
        "BONITO",
        "BRAGAN\xc7A",
        "BRASIL NOVO",
        "BREJO GRANDE DO ARAGUAIA",
        "BREU BRANCO",
        "BREVES",
        "BUJARU",
        "CACHOEIRA DO PIRI\xc1",
        "CACHOEIRA DO ARARI",
        "CAMET\xc1",
        "CANA\xc3 DOS CARAJ\xc1S",
        "CAPANEMA",
        "CAPIT\xc3O PO\xc7O",
        "CASTANHAL",
        "CHAVES",
        "COLARES",
        "CONCEI\xc7\xc3O DO ARAGUAIA",
        "CONC\xd3RDIA DO PAR\xc1",
        "CUMARU DO NORTE",
        "CURION\xd3POLIS",
        "CURRALINHO",
        "CURU\xc1",
        "CURU\xc7\xc1",
        "DOM ELISEU",
        "ELDORADO DO CARAJ\xc1S",
        "FARO",
        "FLORESTA DO ARAGUAIA",
        "GARRAF\xc3O DO NORTE",
        "GOIAN\xc9SIA DO PAR\xc1",
        "GURUP\xc1",
        "IGARAP\xc9-A\xc7U",
        "IGARAP\xc9-MIRI",
        "INHANGAPI",
        "IPIXUNA DO PAR\xc1",
        "IRITUIA",
        "ITAITUBA",
        "ITUPIRANGA",
        "JACAREACANGA",
        "JACUND\xc1",
        "JURUTI",
        "LIMOEIRO DO AJURU",
        "M\xc3E DO RIO",
        "MAGALH\xc3ES BARATA",
        "MARAB\xc1",
        "MARACAN\xc3",
        "MARAPANIM",
        "MARITUBA",
        "MEDICIL\xc2NDIA",
        "MELGA\xc7O",
        "MOCAJUBA",
        "MOJU",
        "MOJU\xcd DOS CAMPOS",
        "MONTE ALEGRE",
        "MUAN\xc1",
        "NOVA ESPERAN\xc7A DO PIRI\xc1",
        "NOVA IPIXUNA",
        "NOVA TIMBOTEUA",
        "NOVO PROGRESSO",
        "NOVO REPARTIMENTO",
        "\xd3BIDOS",
        "OEIRAS DO PAR\xc1",
        "ORIXIMIN\xc1",
        "OUR\xc9M",
        "OURIL\xc2NDIA DO NORTE",
        "PACAJ\xc1",
        "PALESTINA DO PAR\xc1",
        "PARAGOMINAS",
        "PARAUAPEBAS",
        "PAU D'ARCO",
        "PEIXE-BOI",
        "PI\xc7ARRA",
        "PLACAS",
        "PONTA DE PEDRAS",
        "PORTEL",
        "PORTO DE MOZ",
        "PRAINHA",
        "PRIMAVERA",
        "QUATIPURU",
        "REDEN\xc7\xc3O",
        "RIO MARIA",
        "RONDON DO PAR\xc1",
        "RUR\xd3POLIS",
        "SALIN\xd3POLIS",
        "SALVATERRA",
        "SANTA B\xc1RBARA DO PAR\xc1",
        "SANTA CRUZ DO ARARI",
        "SANTA IZABEL DO PAR\xc1",
        "SANTA LUZIA DO PAR\xc1",
        "SANTA MARIA DAS BARREIRAS",
        "SANTA MARIA DO PAR\xc1",
        "SANTANA DO ARAGUAIA",
        "SANTAR\xc9M",
        "SANTAR\xc9M NOVO",
        "SANTO ANT\xd4NIO DO TAU\xc1",
        "S\xc3O CAETANO DE ODIVELAS",
        "S\xc3O DOMINGOS DO ARAGUAIA",
        "S\xc3O DOMINGOS DO CAPIM",
        "S\xc3O F\xc9LIX DO XINGU",
        "S\xc3O FRANCISCO DO PAR\xc1",
        "S\xc3O GERALDO DO ARAGUAIA",
        "S\xc3O JO\xc3O DA PONTA",
        "S\xc3O JO\xc3O DE PIRABAS",
        "S\xc3O JO\xc3O DO ARAGUAIA",
        "S\xc3O MIGUEL DO GUAM\xc1",
        "S\xc3O SEBASTI\xc3O DA BOA VISTA",
        "SAPUCAIA",
        "SENADOR JOS\xc9 PORF\xcdRIO",
        "SOURE",
        "TAIL\xc2NDIA",
        "TERRA ALTA",
        "TERRA SANTA",
        "TOM\xc9-A\xc7U",
        "TRACUATEUA",
        "TRAIR\xc3O",
        "TUCUM\xc3",
        "TUCURU\xcd",
        "ULIAN\xd3POLIS",
        "URUAR\xc1",
        "VIGIA",
        "VISEU",
        "VIT\xd3RIA DO XINGU",
        "XINGUARA",
    ]
    opcoes_situacao = [
        "REGULARIZAÇÃO NÃO ONEROSA",
        "REGULARIZAÇÃO ONEROSA",
        "PROTOCOLO DE SOLICITAÇÃO DE ACESSO À INFORMAÇÃO",
        "CERTIDÃO",
        "AFORAMENTO",
        "QUILOMBO",
    ]
    opcoes_georreferenciamento = [
        "ITERPA",
        "SICARF",
        "SIGEF/INCRA",
        "SICAR/SEMAS",
        "TERCEIROS",
        "NÃO POSSUI",
        "NÃO CONSTA",
        "OUTROS",
    ]
    opcoes_ferramenta = [
        "Atualizar planilha Controle (...\\PLANILHA\\CONTROLE.xlsx)",
        "Exportar planilha Excel (.xlxs)",
        "Exportar Projeto (.mxd)",
        "Exportar Mapa (.pdf)",
        "Exportar Despacho (.docx)",
    ]


class Database:

    DBNAME = "dbteste"
    USER = "postgres"
    PASSWORD = "102045"
    HOST = "localhost"
    PORT = "5432"

    colunas_fc_database = [
        "ano",
        "numero",
        "interessad",
        "denominaca",
        "parcela",
        "georrefere",
        "municipio",
        "situacao",
        "carta",
        "complement",
        "data_atual",
    ]

    def __init__(self):
        diretorio_raiz = os.path.dirname(os.path.abspath(__file__))
        self.diretorio_database = "{}\\database".format(diretorio_raiz)
        self.nome_database = "database.gdb"
        self.nome_featureclass = "AREAS_DE_INTERESSE"
        self.db = "{}\\{}".format(self.diretorio_database, self.nome_database)
        self.featureclass_database = "{}\\{}".format(self.db, self.nome_featureclass)
        self.featureclass_local = "{}\\{}".format(
            self.diretorio_database, "{}.shp".format(self.nome_featureclass.lower())
        )
        self.logger = Logger().setup_logger()

    def adicionarFeatureclass(self, featureclass, params):
        def retornaDiaMesAnoAtual():
            """Retorna a Data atual de hoje me formato datetime, para a TDA."""
            diaHoje = datetime.datetime.now().day
            mesHoje = datetime.datetime.now().month
            anoHoje = datetime.datetime.now().year
            return datetime.datetime(anoHoje, mesHoje, diaHoje)

        arcpy.management.Append(featureclass, self.featureclass_database, "NO_TEST")
        arcpy.management.Append(featureclass, self.featureclass_local, "NO_TEST")

        valores_update = [
            params[Parametros.ano].valueAsText,
            params[Parametros.numero].valueAsText,
            params[Parametros.interessado].valueAsText,
            params[Parametros.denominacao].valueAsText,
            params[Parametros.parcela].valueAsText,
            params[Parametros.georreferenciamento].valueAsText,
            params[Parametros.municipio].valueAsText,
            params[Parametros.situacao].valueAsText,
            params[Parametros.carta].valueAsText,
            params[Parametros.complemento].valueAsText,
            retornaDiaMesAnoAtual(),
        ]
        empty = [0, 0, " ", " ", " ", " ", " ", " ", " ", " ", None]

        # GEODATABASE
        with arcpy.da.UpdateCursor(
            self.featureclass_database, self.colunas_fc_database
        ) as cursor:
            for row in cursor:
                if row == empty:
                    cursor.updateRow(valores_update)

        # LOCAL
        with arcpy.da.UpdateCursor(
            self.featureclass_local, self.colunas_fc_database
        ) as cursor:
            for row in cursor:
                if row == empty:
                    cursor.updateRow(valores_update)

    def atualizarBancoDeDadosPostgreSQL(self):
        conn = psycopg2.connect(
            dbname=self.DBNAME,
            user=self.USER,
            password=self.PASSWORD,
            host=self.HOST,
            port=self.PORT,
        )
        cursor = conn.cursor()

        try:
            with arcpy.da.SearchCursor(
                self.featureclass_local,
                ["SHAPE@"] + [self.colunas_fc_database],
            ) as cursor_shapefile:
                for row in cursor_shapefile:
                    geom = row[0]
                    ano = row[1]
                    numero = row[2]
                    interessad = row[3]
                    denominaca = row[4]
                    parcela = row[5]
                    georrefere = row[6]
                    municipio = row[7]
                    situacao = row[8]
                    carta = row[9]
                    complement = row[10]
                    data_atual = row[11]

                    if isinstance(data_atual, str):
                        data_atual = datetime.strptime(data_atual, "%Y-%m-%d")

                    # Verifica se a linha já existe no banco de dados
                    cursor.execute(
                        """
                        SELECT COUNT(*) FROM public.areas_de_interesse
                        WHERE ano = %s AND numero = %s
                    """,
                        (ano, numero),
                    )
                    count = cursor.fetchone()[0]

                    if count == 0:
                        # Insere a linha se não existir
                        self.logger.info(
                            "Executando Insercao de Feicao no banco PostgreSQL"
                        )
                        self.logger.debug("geom.WKT: {}".format(geom.WKT))
                        cursor.execute(
                            """
                            INSERT INTO public.areas_de_interesse (
                                ano, numero, interessad, denominaca, parcela,
                                georrefere, municipio, situacao, carta,
                                complement, data_atual, geom
                            ) VALUES (
                                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, ST_GeomFromText(%s, 4674)
                            )
                            """,
                            (
                                ano,
                                numero,
                                interessad,
                                denominaca,
                                parcela,
                                georrefere,
                                municipio,
                                situacao,
                                carta,
                                complement,
                                data_atual,
                                geom.WKT,
                            ),
                        )
                        self.logger.info("SQL Executando com sucesso!")

        except Exception as e:
            exc_info = traceback.format_exc()
            print(exc_info)
        finally:
            # Commit e fechamento da conexão
            conn.commit()
            cursor.close()
            conn.close()


class Parametros(ValoresConstantes):
    debug = False

    btn_carta = "btn_carta"
    btn_exportar_mapa = "btn_exportar_mapa"
    btn_exportar_despacho = "btn_exportar_despacho"
    btn_exportar_planilha = "btn_exportar_planilha"
    btn_alterar_layout = "btn_alterar_layout"
    btn_dados_carregados = "btn_dados_carregados"
    btn_incidencias = "btn_incidencias"
    btn_zoneamento = "btn_zoneamento"
    btn_macro_zoneamento = "btn_macro_zoneamento"

    shapefile = "shapefile"
    municipio = "municipio"
    situacao = "situacao"
    ano = "ano"
    numero = "numero"
    interessado = "interessado"
    denominacao = "denominacao"
    parcela = "parcela"
    georreferenciamento = "georreferenciamento"
    data_atual = datetime.datetime.now().strftime("%d/%m/%Y")
    complemento = "complemento"
    carta = "carta"
    zee_mzee = "zee_mzee"

    lista_incidencias = "lista_incidencias"
    lista_incidencias_colunas = "lista_incidencias_colunas"
    lista_municipios = "lista_municipios"
    lista_shapefiles = "lista_shapefiles"
    lista_zoneamento = "lista_zoneamento"

    modelo_despacho = "modelo_despacho"
    nome_exportar_mapa = "nome_exportar_mapa"
    pasta_exportar_mapa = "pasta_exportar_mapa"
    pasta_resultados = "pasta_resultados"
    shapefile_incidencias = "shapefile_incidencias"

    tester_comands = "tester"
    tester_list_helper = "helper"
    tester_helper = "helper2"

    parametros_lista = ["lista_zoneamento", "lista_incidencias"]
    salvar_parametros = [
        ano,
        carta,
        complemento,
        denominacao,
        georreferenciamento,
        interessado,
        lista_incidencias,
        # lista_incidencias_colunas,
        lista_zoneamento,
        modelo_despacho,
        municipio,
        numero,
        parcela,
        pasta_exportar_mapa,
        pasta_resultados,
        situacao,
    ]

    def __init__(self):
        parametros_incidencias = False
        parametros_auxiliares = True
        parametros_debug = self.debug

        _logger = Logger()
        self.logger = _logger.setup_logger()
        self.logger.debug("Criando Parametros da Ferramenta...")

        self.params = [
            arcpy.Parameter(  # type: ignore
                name=self.btn_dados_carregados,
                displayName="Dados Carregados",
                datatype="GPBoolean",
                parameterType="Optional",
                enabled=False,
            ),
            arcpy.Parameter(  # type: ignore
                name=self.shapefile,
                displayName="Shapefile ou Camada do Projeto",
                datatype="GPFeatureLayer",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.pasta_resultados,
                displayName="Exportar Resultados na Pasta abaixo",
                datatype="DEFolder",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.modelo_despacho,
                displayName="Modelo Despacho",
                datatype="DEFile",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.municipio,
                displayName="Município",
                datatype="GPString",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.situacao,
                displayName="Situação do Processo",
                datatype="GPString",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.georreferenciamento,
                displayName="Georreferenciamento",
                datatype="GPString",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.ano,
                displayName="Ano do Processo",
                datatype="GPString",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.numero,
                displayName="Número do Processo",
                datatype="GPString",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.interessado,
                displayName="Nome do Interessado",
                datatype="GPString",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.denominacao,
                displayName="Denominação do Imóvel",
                datatype="GPString",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.parcela,
                displayName="Parcela",
                datatype="GPLong",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.complemento,
                displayName="Complemento",
                datatype="GPString",
                parameterType="Optional",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.btn_carta,
                displayName="Clique para Obter Carta Cadastral",
                datatype="GPBoolean",
                parameterType="Optional",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.carta,
                displayName="Carta Cadastral",
                datatype="GPString",
                parameterType="Optional",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.btn_zoneamento,
                displayName="Clique para Obter Zoneamentos",
                datatype="GPBoolean",
                parameterType="Optional",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.lista_zoneamento,
                displayName="Tipo de Zoneamento encontrado",
                datatype="GPString",
                parameterType="Optional",
                multiValue=True,
            ),
            arcpy.Parameter(  # type: ignore
                name=self.btn_alterar_layout,
                displayName="Clique para Atualizar Layout",
                datatype="GPBoolean",
                parameterType="Optional",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.btn_exportar_mapa,
                displayName="Clique para Exportar o Mapa",
                datatype="GPBoolean",
                parameterType="Optional",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.btn_exportar_despacho,
                displayName="Clique para Exportar o Despacho",
                datatype="GPBoolean",
                parameterType="Optional",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.btn_exportar_planilha,
                displayName="Clique para Exportar a Planilha",
                datatype="GPBoolean",
                parameterType="Optional",
            ),
        ]
        lista_parametros_incidencias = [
            arcpy.Parameter(  # type: ignore
                name=self.shapefile_incidencias,
                displayName="Selecionar a Camada para Verificar Incidencias",
                datatype="GPFeatureLayer",
                parameterType="Optional",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.lista_incidencias_colunas,
                displayName="Selecione as Colunas a mostrar",
                datatype="GPString",
                parameterType="Optional",
                multiValue=True,
            ),
            arcpy.Parameter(  # type: ignore
                name=self.lista_incidencias,
                displayName="Incidencias encontradas",
                datatype="GPString",
                parameterType="Optional",
                multiValue=True,
            ),
            arcpy.Parameter(  # type: ignore
                name=self.btn_incidencias,
                displayName="Clique para Obter Incidencias",
                datatype="GPBoolean",
                parameterType="Optional",
            ),
        ]
        lista_parametros_auxiliares = [
            arcpy.Parameter(  # type: ignore
                name=self.lista_shapefiles,
                displayName="Shapefiles Já Selecionados",
                datatype="GPString",
                parameterType="Optional",
                multiValue=True,
                enabled=False,
                category="Parametros Auxiliares",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.lista_municipios,
                displayName="Municipios Já Selecionados",
                datatype="GPString",
                parameterType="Optional",
                multiValue=True,
                enabled=False,
                category="Parametros Auxiliares",
            ),
        ]
        lista_parametros_debug = [
            arcpy.Parameter(  # type: ignore
                name=self.lista_shapefiles,
                displayName="Shapefiles Já Selecionados",
                datatype="GPString",
                parameterType="Optional",
                multiValue=True,
                enabled=False,
                category="Parametros Auxiliares",
            ),
            arcpy.Parameter(  # type: ignore
                name=self.lista_municipios,
                displayName="Municipios Já Selecionados",
                datatype="GPString",
                parameterType="Optional",
                multiValue=True,
                enabled=False,
                category="Parametros Auxiliares",
            ),
        ]

        if parametros_incidencias:
            self.params += lista_parametros_incidencias

        if parametros_auxiliares:
            self.params += lista_parametros_auxiliares

        if parametros_debug:
            self.params += lista_parametros_debug

        self.formatarDisplayNames()
        self.adicionarValoresPadrao()

    @staticmethod
    def dicionarioDeParametros(parametros, **kwargs):
        """Retorna { nome_do_parametro : parametro }"""
        if kwargs.get("version") == 3:
            dict_params = {
                "{}".format(p.name): "{}".format(p.value)
                for p in parametros
                if p.dataType in ["String", "Long"]
            }
        elif kwargs.get("version") == 2:
            dict_params = {p.name: p.value for p in parametros}
        else:
            dict_params = {p.name: p for p in parametros}
        return dict_params

    @staticmethod
    def salvarDadosFerramenta(parametros, logger):
        dict_dados = {}

        salvar_campos = Parametros.salvar_parametros
        for campo in salvar_campos:
            for _nome_param, valor_param in parametros.items():
                if campo == _nome_param:
                    if campo in Parametros.parametros_lista:
                        valor_inserir = valor_param.filter.list
                    else:
                        if valor_param.value == None:
                            valor_inserir = ""
                        else:
                            valor_inserir = "{}".format(valor_param.value)

                    nome_param = "{}".format(_nome_param)
                    # Pulando dicionario
                    dict_dados[nome_param] = valor_inserir

        logger.debug("Criando arquivo YAML...")
        with open(ValoresConstantes.arquivo_yaml, "w") as arquivo:
            yaml.dump(
                dict_dados,
                arquivo,
                encoding="utf-8",
            )
        logger.debug("Arquivo YAML Salvo com Sucesso!")

    @staticmethod
    def carregarDadosFerramenta(parametros, logger):
        logger.debug("Carregando ultimos dados preenchidos...")
        logger.debug("Carregando arquivo YAML...")
        try:
            with open(ValoresConstantes.arquivo_yaml, "r") as arquivo:
                dados = yaml.load(arquivo, Loader=yaml.FullLoader)
                logger.debug("Arquivo YAML Carregado com Sucesso!")
                logger.debug("Arquivo YAML: {}".format(dados))
                if dados == None:
                    dados = {}
                return dados
        except Exception as e:
            logger.error("Erro ao carregar o arquivo YAML: {}".format(e))
            return None

    @staticmethod
    def atualizarCampos(parametros, logger):

        if parametros[Parametros.btn_dados_carregados].value == None:
            popup_result = Funcoes.mostrarPopUp(logger)
            # Sair de None garante que não aparecerá mais
            parametros[Parametros.btn_dados_carregados].value = True

            if popup_result in [1, 6]:  # Resultados de confirmação

                dict_dados_carregados = Parametros.carregarDadosFerramenta(
                    parametros, logger
                )

                for nome, valor in dict_dados_carregados.items():

                    _parametro = parametros.get(nome, False)

                    if _parametro:
                        if nome in Parametros.parametros_lista:
                            parametros[nome].filter.list = valor
                        else:
                            parametros[nome].value = valor

                logger.info("Dados carregando com Sucesso!")

            elif popup_result in [2]:
                logger.info("Cancelando operação")
            else:
                logger.info("Carregando ferramenta vazia")
                parametros[Parametros.btn_dados_carregados].value = False

    @staticmethod
    def atualizarDataFramePrincipal(parametros, logger):

        lista_aux_shp = parametros[Parametros.lista_shapefiles].filter.list

        logger.debug(
            "Parametros.atualizarDataFramePrincipal() -> parametros[Parametros.shapefile].valueAsText: {}".format(
                parametros[Parametros.shapefile].valueAsText
            )
        )
        logger.debug(
            "Parametros.atualizarDataFramePrincipal() -> parametros[Parametros.shapefile].value: {}".format(
                parametros[Parametros.shapefile].value
            )
        )

        logger.debug(
            "Parametros.atualizarDataFramePrincipal() -> dir(parametros[Parametros.shapefile].valueAsText): {}".format(
                dir(parametros[Parametros.shapefile].valueAsText)
            )
        )
        logger.debug(
            "Parametros.atualizarDataFramePrincipal() -> dir(parametros[Parametros.shapefile].value): {}".format(
                dir(parametros[Parametros.shapefile].value)
            )
        )

        if parametros[Parametros.shapefile].altered:

            shp_str = parametros[Parametros.shapefile].valueAsText
            shp_value = parametros[Parametros.shapefile].value

            # Se for adicionado Externamente

            camadas = Layout.carregarComponentes("cmds")

            camadas_ds = {}

            for nome, layer in camadas.items():
                if layer.supports("DATASOURCE"):
                    camadas_ds[layer.dataSource] = layer

            logger.debug("camadas_ds: {}".format(camadas_ds))

            if os.path.exists(shp_str):
                logger.debug("Existe")
                layer_ds = camadas_ds.get(shp_str)
                logger.debug("layer_ds: {}".format(layer_ds))

                if layer_ds:
                    logger.debug("layer_ds: {}".format(layer_ds))
                    shp_layer = layer_ds

                else:
                    df = Layout.carregarComponentes("df")["PRINCIPAL"]
                    _shp_layer = arcpy.mapping.Layer(shp_str)
                    logger.debug("Adicionando camada: {}".format(shp_str))
                    arcpy.mapping.AddLayer(df, _shp_layer, "TOP")
                    _shp_layer = Layout.carregarComponentes("cmds_list")[0]
                    shp_layer = _shp_layer

                shp_str = shp_layer.name
                shp_value = shp_layer

            logger.debug("lista_aux_shp.append camada: {}".format(shp_str))
            lista_aux_shp.append(shp_str)

            for _ in range(len(lista_aux_shp) - 2):
                lista_aux_shp.pop(0)

            parametros[Parametros.lista_shapefiles].filter.list = lista_aux_shp

        if len(lista_aux_shp) >= 2:
            if lista_aux_shp[-1] != lista_aux_shp[-2]:
                Layout.aplicarSimbologiaPadrao(shp_value, logger)
                Layout.zoomDataFramePrincipal(shp_value, logger)
                Layout.afastarTextosLayout(shp_value, logger)
                parametros[Parametros.lista_zoneamento].filter.list = []

        if not parametros[Parametros.shapefile].value:
            parametros[Parametros.btn_carta].enabled = False
            parametros[Parametros.btn_zoneamento].enabled = False
            # parametros[Parametros.btn_macro_zoneamento].enabled = False

            _btn_incidencias = parametros.get(Parametros.btn_incidencias)
            if _btn_incidencias:
                parametros[Parametros.btn_incidencias].enabled = False

            logger.debug(
                "Botoes desativados: btn_carta, btn_zoneamento, btn_macro_zoneamento, btn_incidencias"
            )

        else:
            parametros[Parametros.btn_carta].enabled = True
            parametros[Parametros.btn_zoneamento].enabled = True
            # parametros[Parametros.btn_macro_zoneamento].enabled = True

            _btn_incidencias = parametros.get(Parametros.btn_incidencias)
            if _btn_incidencias:
                parametros[Parametros.btn_incidencias].enabled = True

            logger.debug(
                "Botoes ativados: btn_carta, btn_zoneamento, btn_macro_zoneamento, btn_incidencias"
            )

    @staticmethod
    def atualizarMunicipio(parametros, logger):

        lista_aux_mun = parametros[Parametros.lista_municipios].filter.list

        if parametros[Parametros.municipio].altered:
            # Adicionando o valor na lista de comparação
            lista_aux_mun.append(parametros[Parametros.municipio].valueAsText)

            for _ in range(len(lista_aux_mun) - 2):
                lista_aux_mun.pop(0)
            parametros[Parametros.lista_municipios].filter.list = lista_aux_mun

        if len(lista_aux_mun) >= 2:
            if lista_aux_mun[-1] != lista_aux_mun[-2]:
                Layout.atualizarDataFrameSituacao(
                    parametros[Parametros.municipio].valueAsText,
                    logger,
                )

    @staticmethod
    def debugMode(parametros, logger):  # DEBUG MODE #
        if Parametros.debug:
            outputs = parametros[Parametros.tester_helper].filter.list
            if parametros[Parametros.tester_comands].altered:
                result = eval(parametros[Parametros.tester_comands].valueAsText)
                result_str = "{}".format(result)
                outputs.append(result_str)
                parametros[Parametros.tester_helper].value = result_str
                parametros[Parametros.tester_list_helper].filter.list = outputs

        Parametros.salvarDadosFerramenta(parametros, logger)

    def formatarDisplayNames(self):
        """Recebe os parametros e formata os display names
        ao padrão de encode 'cp1252' que é lido de maneira
        correta na interface arcmap"""
        for param in self.params:
            display_text = param.displayName.encode("cp1252")
            param.displayName = display_text

    def formatarStrings(self, lista_strings):
        """Forma uma lista de strings a encode 'cp1252' e
        retorna a lista formatada"""
        nova_lista_string = []
        for string in lista_strings:
            self.logger.debug("Formatando Strings: {}".format(string))
            try:
                string_fmt = string.encode("cp1252")
                nova_lista_string.append(string_fmt)
            except Exception as e:
                exc_info = traceback.format_exc()
                self.logger.debug(exc_info)

        return nova_lista_string

    def adicionarValoresPadrao(self):

        def verificarPasta(pasta):
            if not os.path.exists(pasta):
                os.makedirs(pasta)
            return pasta

        pasta_resultado = verificarPasta("{}\\resultados".format(self.diretorio_raiz))
        pasta_despacho = verificarPasta("{}\\despacho".format(self.diretorio_raiz))
        pasta_planilha = verificarPasta("{}\\planilha".format(self.diretorio_raiz))
        pasta_estilos = verificarPasta("{}\\estilos".format(self.diretorio_raiz))
        despacho = "{}\\modelo_despacho.docx".format(pasta_despacho)

        opts_georef = self.formatarStrings(self.opcoes_georreferenciamento)
        opts_situacao = self.formatarStrings(self.opcoes_situacao)
        dict_params = Parametros.dicionarioDeParametros(self.params)

        dict_params["modelo_despacho"].filter.list = ["docx", "doc"]
        dict_params["municipio"].filter.list = self.opcoes_municipios
        dict_params["situacao"].filter.list = opts_situacao
        dict_params["georreferenciamento"].filter.list = opts_georef
        dict_params["lista_shapefiles"].filter.list = ["Sem seleção".encode("cp1252")]
        dict_params["lista_municipios"].filter.list = ["Sem seleção".encode("cp1252")]
        dict_params["lista_zoneamento"].filter.list = []

        lista_incidencias = dict_params.get("lista_incidencias", None)
        if lista_incidencias:
            dict_params["lista_incidencias"].filter.list = []

        dict_params["modelo_despacho"].value = despacho
        dict_params["pasta_resultados"].value = pasta_resultado
        dict_params["complemento"].value = ""


class Planilha(ValoresConstantes):

    colunas = [
        "ID",
        "MUNICIPIO",
        "SITUACAO",
        "ANO",
        "NUMERO",
        "INTERESSADO",
        "IMOVEL",
        "PARCELA",
        "GEORREF",
        "DATA",
        "COMPLEMENTO",
        "CARTA",
        "ZEE",
        "PATH",
    ]

    def __init__(self):
        self.logger = Logger().setup_logger()
        self.logger.debug("Carregando arquivo modelo de Planilha")

    def buildXlsx(self, params):
        self.criarPlanilhaIndividual(params)
        self.adicionarParaControle(params)

    def criarDados(self, params, next_row):
        """Cria um dicionario com os resultados dos parametros
        preenchidos da ferramenta, e poem num dicionario que
        tem as colunas da tabela excel indicada."""

        ano = params[Parametros.ano].valueAsText
        numero = params[Parametros.numero].valueAsText
        municipio = params[Parametros.municipio].valueAsText
        carta = params[Parametros.carta].valueAsText
        situacao = params[Parametros.situacao].valueAsText
        interessado = params[Parametros.interessado].valueAsText
        denominacao = params[Parametros.denominacao].valueAsText
        complemento = params[Parametros.complemento].valueAsText
        parcela = params[Parametros.parcela].valueAsText
        georreferenciamento = params[Parametros.georreferenciamento].valueAsText
        data_atual = Parametros.data_atual
        lista_zoneamento = params[Parametros.lista_zoneamento].values

        zoneamentos = ", ".join(lista_zoneamento) if lista_zoneamento else ""

        param_shapefile = params[Parametros.shapefile].value
        param_shapefile_str = params[Parametros.shapefile].valueAsText

        if os.path.exists(param_shapefile_str):
            shapefile_datasource = param_shapefile_str
        else:
            shapefile_datasource = (
                params[Parametros.shapefile].value.dataSource if param_shapefile else ""
            )

        self.logger.debug(
            "Planilha.criarDados() -> shapefile_datasource: {}".format(
                shapefile_datasource
            )
        )

        dados = {
            "ID": next_row,
            "MUNICIPIO": municipio,
            "SITUACAO": situacao,
            "ANO": ano,
            "NUMERO": numero,
            "INTERESSADO": interessado,
            "IMOVEL": denominacao,
            "PARCELA": parcela,
            "GEORREF": georreferenciamento,
            "DATA": data_atual,
            "COMPLEMENTO": complemento,
            "CARTA": carta,
            "ZEE": zoneamentos,
            "PATH": shapefile_datasource,
        }
        return dados

    def adicionarParaControle(self, params):
        """
        Adiciona uma linha de dados à tabela.

        Args:
        tabela (str): Caminho para o arquivo da tabela do Excel.
        dados (dict): Dicionário contendo os valores das colunas.
        """
        # Carrega a planilha
        wb = load_workbook(Parametros.arquivo_xlsx_controle)
        ws = wb.active

        # Encontra a próxima linha vazia
        next_row = ws.max_row + 1

        dados = self.criarDados(params, next_row)

        # Preenche a linha com os dados do dicionário
        for col, header in enumerate(self.colunas, start=1):
            ws.cell(row=next_row, column=col, value=dados.get(header))

        # Salva as alterações no arquivo
        self.logger.info("Adiciona Dados em Planilha Controle.")
        try:
            wb.save(Parametros.arquivo_xlsx_controle)
        except:
            wb.save(Parametros.arquivo_xlsx_controle_extra)
        self.logger.info("Planilha Controle Atualizada com Sucesso!")
        # os.system(
        #     'start "" "{}"'.format(os.path.dirname(Parametros.arquivo_xlsx_controle))
        # )

    def criarPlanilhaIndividual(self, params):
        """
        Adiciona uma linha de dados à tabela.

        Args:
        tabela (str): Caminho para o arquivo da tabela do Excel.
        dados (dict): Dicionário contendo os valores das colunas.
        """
        # Carrega a planilha
        wb = load_workbook(Parametros.arquivo_xlsx_planilha_indv)
        ws = wb.active

        # Encontra a próxima linha vazia
        next_row = ws.max_row + 1

        dados = self.criarDados(params, next_row)

        # Preenche a linha com os dados do dicionário
        for col, header in enumerate(self.colunas, start=1):
            ws.cell(row=next_row, column=col, value=dados.get(header))

        ano = params[Parametros.ano].valueAsText
        numero = params[Parametros.numero].valueAsText
        interessado = params[Parametros.interessado].valueAsText

        pasta_saida = params[Parametros.pasta_resultados].valueAsText
        extensao = ".xlsx"

        nome_arquivo_saida = "{}_{}_{}".format(ano, numero, interessado)
        nome_arquivo_saida = Funcoes.normalizarString(nome_arquivo_saida)

        pasta_saida = Funcoes.criarSubPasta(pasta_saida, nome_arquivo_saida)

        caminho_saida_final = "{}\\{}{}".format(
            pasta_saida, nome_arquivo_saida, extensao
        )

        count = 1
        while os.path.exists(caminho_saida_final):
            caminho_saida_final = "{}\\{}_{}{}".format(
                pasta_saida, nome_arquivo_saida, count, extensao
            )
            count += 1

        caminho_saida_final

        # Salva as alterações no arquivo
        self.logger.info("Salvando Planilha em: {}...".format(caminho_saida_final))
        wb.save(caminho_saida_final)
        self.logger.info("Despacho Salvo com Sucesso!")
        # os.system('start "" "{}"'.format(pasta_saida))


class Despacho(ValoresConstantes):

    def __init__(self):
        self.logger = Logger().setup_logger()
        self.logger.debug("Carregando arquivo modelo de Despacho")
        self.despacho_docx = Document(self.arquivo_docx_despacho)
        self.logger.debug("Arquivo modelo de Despacho carregado com Sucesso!")

    def buildDocx(self, params):

        ano = params[Parametros.ano].valueAsText
        numero = params[Parametros.numero].valueAsText

        content_status = params[Parametros.municipio].valueAsText
        keywords = params[Parametros.carta].valueAsText
        category = "{}/{}".format(ano, numero)
        subject = params[Parametros.situacao].valueAsText
        author = params[Parametros.interessado].valueAsText
        title = params[Parametros.denominacao].valueAsText
        zees = params[Parametros.lista_zoneamento].values

        zonas = []
        if zees:
            for v in zees:
                if "ZEE: " in v:
                    nv = v.split("ZEE: ")[-1]
                else:
                    nv = v.split("MZEE: ")[-1]
                zonas.append(nv)

        zonas_str = ", ".join(zonas)
        comments = zonas_str

        self.despacho_docx.core_properties.content_status = content_status
        self.despacho_docx.core_properties.keywords = keywords
        self.despacho_docx.core_properties.category = category
        self.despacho_docx.core_properties.subject = subject
        self.despacho_docx.core_properties.author = author
        self.despacho_docx.core_properties.title = title
        self.despacho_docx.core_properties.comments = comments

        pasta_saida = params[Parametros.pasta_resultados].valueAsText

        extensao = ".docx"

        nome_arquivo_saida = "{}_{}_{}".format(ano, numero, author)
        nome_arquivo_saida = Funcoes.normalizarString(nome_arquivo_saida)

        pasta_saida = Funcoes.criarSubPasta(pasta_saida, nome_arquivo_saida)

        caminho_saida_final = "{}\\{}{}".format(
            pasta_saida, nome_arquivo_saida, extensao
        )

        count = 1
        while os.path.exists(caminho_saida_final):
            caminho_saida_final = "{}\\{}_{}{}".format(
                pasta_saida, nome_arquivo_saida, count, extensao
            )
            count += 1

        caminho_saida_final

        self.logger.info("Salvando Despacho em: {}...".format(caminho_saida_final))
        self.despacho_docx.save(caminho_saida_final)
        self.logger.info("Despacho Salvo com Sucesso!")
        # os.system('start "" "{}"'.format(pasta_saida))


class Shapefile:
    @staticmethod
    def inserirFeicoesEmShapefile(shapefile_input, shapefile_target):
        feicoes_do_shp = []

        with arcpy.da.SearchCursor(shapefile_input, "SHAPE@") as search_cursor:  # type: ignore
            for row in search_cursor:
                feicoes_do_shp.append(row[0])

        with arcpy.da.InsertCursor(shapefile_target, "SHAPE@") as insert_cursor:  # type: ignore
            for feicao in feicoes_do_shp:
                insert_cursor.insertRow([feicao])

    @staticmethod
    def apagarFeicoesEmShapefile(shapefile_input):
        with arcpy.da.UpdateCursor(shapefile_input, ["OID@"]) as cursor:  # type: ignore
            for _ in cursor:
                cursor.deleteRow()

    @staticmethod
    def obterValoresColunaIntersecao_v2(
        shapefile_a_selecionar,
        camada,
        colunas,
        logger,
    ):
        """Retorna uma string formatada com os valores da coluna
        de acordo com a interseção dos shapefiles

        Args:
            - shapefile_a_selecionar (Layer): camada pra seleção
            - camada (Layer): shapefile
            - coluna (list[str]): nome das colunas
        """
        selecao = arcpy.SelectLayerByLocation_management(
            shapefile_a_selecionar, "INTERSECT", camada
        )

        # TODO colocar no projeto com simbologia
        shapefile_incidencia_str_name = "INCIDENCIAS_{}".format(
            datetime.datetime.now().strftime("%H%M%S")
        )
        incidencia_cmd = arcpy.CopyFeatures_management(
            selecao, shapefile_incidencia_str_name
        )
        logger.debug("Adicionando Incidencias como Camada")
        nova_camada = Layout.carregarComponentes("cmds")[shapefile_incidencia_str_name]
        logger.debug("nova_camada: {}".format(nova_camada))
        logger.debug("Aplicando Simbologia Incidencias")
        Layout.aplicarSimbologiaIncidencias(nova_camada, logger)

        lista_valores = []
        with arcpy.da.SearchCursor(selecao, colunas) as tabela:  # type: ignore
            for linha in tabela:
                lista_valores.append(linha)

        arcpy.SelectLayerByAttribute_management(
            shapefile_a_selecionar, "CLEAR_SELECTION"
        )
        return lista_valores

    @staticmethod
    def obterValoresColunaIntersecao_v3(
        shapefile_a_selecionar,
        camada,
        colunas,
    ):
        """Retorna uma string formatada com os valores da coluna
        de acordo com a interseção dos shapefiles

        Args:
            - shapefile_a_selecionar (Layer): camada pra seleção
            - camada (Layer): shapefile
            - coluna (list[str]): nome das colunas
        """
        selecao = arcpy.SelectLayerByLocation_management(
            shapefile_a_selecionar, "INTERSECT", camada
        )
        lista_valores = []
        with arcpy.da.SearchCursor(selecao, colunas) as tabela:  # type: ignore
            for linha in tabela:
                nova_linha = list(linha)
                nova_linha_f = nova_linha[3:]
                lista_valores.append(nova_linha_f)

        arcpy.SelectLayerByAttribute_management(
            shapefile_a_selecionar, "CLEAR_SELECTION"
        )
        return lista_valores

    @staticmethod
    def obterValoresColunaIntersecao(
        shapefile_a_selecionar, camada, coluna, prefixo="", retorno="str"
    ):
        """Retorna uma string formatada com os valores da coluna
        de acordo com a interseção dos shapefiles

        Args:
            - shapefile_a_selecionar (Layer):
            - camada (Layer):
            - coluna (str):
        """
        selecao = arcpy.SelectLayerByLocation_management(
            shapefile_a_selecionar, "INTERSECT", camada
        )
        lista_valores = []
        with arcpy.da.SearchCursor(selecao, [coluna]) as tabela:  # type: ignore
            for linha in tabela:
                valor = linha[0]
                if valor:
                    valor_str = prefixo + linha[0]
                    lista_valores.append(valor_str)

        arcpy.SelectLayerByAttribute_management(
            shapefile_a_selecionar, "CLEAR_SELECTION"
        )
        if retorno == "str":
            valor_final = ", ".join(lista_valores)
        if retorno == "list":
            valor_final = lista_valores

        return valor_final

    @staticmethod
    def formatarColunasEmString(lista_tuplas_tda):
        inc_list = []
        for incid in lista_tuplas_tda:
            fields = []
            for _field in incid:
                if isinstance(_field, datetime.datetime):
                    field_fmt = _field.strftime("%d/%m/%Y")
                else:
                    field_fmt = str(_field)

                field_fmt_str = str("{}".format(field_fmt))

                fields.append(field_fmt_str)

            lista_campo_str = ValoresConstantes.separador.join(fields)
            inc_list.append(lista_campo_str)
        return inc_list

    @staticmethod
    def obterColunasShapefile(parametros, logger):
        if not parametros[Parametros.shapefile_incidencias].value:
            return

        if parametros[Parametros.shapefile_incidencias].altered:
            logger.debug("Calculando Colunas do shapefile incididor!")
            colunas = [
                i.name
                for i in arcpy.ListFields(
                    parametros[Parametros.shapefile_incidencias].value
                )
            ]
            logger.debug("Colunas: {}".format(colunas))
            parametros[Parametros.lista_incidencias_colunas].filter.list = colunas

    @staticmethod
    def encontrarIntersecoesComShapefile(parametros, logger):
        lista_incidencias_csv_fmt = []
        colunas = Shapefile.obterColunasShapefile(parametros, logger)

        if parametros[Parametros.btn_incidencias].altered:  # Se alterado
            logger.debug("Iniciando processo de obtencao de Incidencias...")

            if parametros[Parametros.btn_incidencias].value:  # Se houver valor
                colunas_ativas = parametros[Parametros.lista_incidencias_colunas].values
                logger.debug(
                    "Colunas atividas para formatar incidencias: {}".format(
                        colunas_ativas
                    )
                )

                resultados_incidencias = Shapefile.obterValoresColunaIntersecao_v2(
                    shapefile_a_selecionar=parametros[
                        Parametros.shapefile_incidencias
                    ].value,
                    camada=parametros[Parametros.shapefile].value,
                    colunas=colunas_ativas,
                    logger=logger,
                )

                logger.debug(
                    "RESULTADOS INCIDENCIAS: {}".format(resultados_incidencias)
                )

                lista_de_incidencias_formatado = Shapefile.formatarColunasEmString(
                    resultados_incidencias
                )

                logger.debug(
                    "RESULTADOS FORMATADAS: {}".format(lista_de_incidencias_formatado)
                )

                lista_incidencias_csv_fmt = [
                    ValoresConstantes.separador.join(colunas_ativas)
                ] + lista_de_incidencias_formatado

                parametros[Parametros.lista_incidencias].filter.list = (
                    lista_incidencias_csv_fmt
                )
                parametros[Parametros.btn_incidencias].value = False
        return lista_incidencias_csv_fmt

    @staticmethod
    def encontrarCartaCadastral(parametros, logger):
        if parametros[Parametros.btn_carta].altered:
            if parametros[Parametros.btn_carta].value:
                cartas_str = Funcoes.obterCartaCadastral(
                    parametros[Parametros.shapefile].value, logger
                )
                parametros[Parametros.carta].value = cartas_str
                parametros[Parametros.btn_carta].value = False

    @staticmethod
    def encontrarZoneamento(parametros, logger):
        if parametros[Parametros.btn_zoneamento].altered:
            if parametros[Parametros.btn_zoneamento].value:
                resultados_zoneamentos = Funcoes.obterZoneamentos(
                    parametros[Parametros.shapefile].value, logger
                )
                # Soma-se a lista zoneamento
                parametros[Parametros.lista_zoneamento].filter.list += list(
                    dict.fromkeys(resultados_zoneamentos).keys()
                )

                resultados_macro_zoneamentos = Funcoes.obterMacroZoneamentos(
                    parametros[Parametros.shapefile].value, logger
                )
                # Soma-se a lista zoneamento
                parametros[Parametros.lista_zoneamento].filter.list += list(
                    dict.fromkeys(resultados_macro_zoneamentos).keys()
                )

                parametros[Parametros.btn_zoneamento].value = False

    @staticmethod
    def encontrarZoneamentoMacro(parametros, logger):
        # if parametros[Parametros.btn_macro_zoneamento].altered:
        #     if parametros[Parametros.btn_macro_zoneamento].value:
        # CÓDIGO ANTERIOR
        # parametros[Parametros.btn_macro_zoneamento].value = False
        return


class Layout:

    @staticmethod
    def afastarTextosLayout(shapefile, logger):
        logger.info("Ajustando Textos Automaticos no Layout")
        area_de_limitacao_de_textos = ValoresConstantes.camada_auxiliar_area_limitante
        Shapefile.apagarFeicoesEmShapefile(area_de_limitacao_de_textos)
        Shapefile.inserirFeicoesEmShapefile(shapefile, area_de_limitacao_de_textos)

    @staticmethod
    def aumentarEscala(dataframe):
        dataframe.scale *= 1.38

    @staticmethod
    def atualizarDataFrameSituacao(municipio, logger):
        logger.info(
            "Atualizando Mapa de Situacao baseado no municipio: [{}]".format(municipio)
        )
        camadas = Layout.carregarComponentes("cmds_desc")
        df_situacao = Layout.carregarComponentes("df")["SITUACAO"]

        municipio_str = municipio.replace("'", "''")

        for desc, cmd in camadas.items():
            if desc.startswith("SEDES_MUNICIPAIS"):
                cmd.definitionQuery = r"nmSede = '{}'".format(municipio_str)

            if desc.startswith("LIMITES_MUNICIPIOS"):
                cmd.definitionQuery = r"nmMun <> '{}'".format(municipio_str)

            if desc.startswith("MUNICIPIO_DE_INTERESSE"):
                cmd.definitionQuery = r"nmMun = '{}'".format(municipio_str)
                df_situacao.extent = cmd.getExtent()
                df_situacao.scale *= 1.2

    @staticmethod
    def carregarComponentes(*args):
        """Retorna mxd, cmds e dfs"""
        mxd = arcpy.mapping.MapDocument("current")
        layers = arcpy.mapping.ListLayers
        data_frames = arcpy.mapping.ListDataFrames

        componentes = {
            "mxd": mxd,
            "cmds": {lyr.name: lyr for lyr in layers(mxd)},
            "cmd": {lyr.name: lyr for lyr in layers(mxd)},
            "cmds_list": [lyr for lyr in layers(mxd)],
            "cmd_list": [lyr for lyr in layers(mxd)],
            "cmds_desc": {lyr.description: lyr for lyr in layers(mxd)},
            "dfs": {df.name: df for df in data_frames(mxd)},
            "df": {df.name: df for df in data_frames(mxd)},
        }
        saida_comps = []
        for arg in args:
            comp = componentes.get(arg, None)
            if comp:
                saida_comps.append(comp)
        return saida_comps[0]

    @staticmethod
    def aplicarSimbologiaIncidencias(camada, logger):
        logger.info("Camda de entrada em simbologia: {}".format(camada))
        tipo_shapefile = arcpy.Describe(camada).shapeType
        simbologia = "{}\\shapefile_{}_incidencias.lyr".format(
            ValoresConstantes.diretorio_estilos, tipo_shapefile.lower()
        )
        logger.info("Aplicando Simbologia Incidencias")
        arcpy.ApplySymbologyFromLayer_management(camada, simbologia)

    @staticmethod
    def aplicarSimbologiaPadrao(camada, logger):
        tipo_shapefile = arcpy.Describe(camada).shapeType
        simbologia = "{}\\shapefile_{}.lyr".format(
            ValoresConstantes.diretorio_estilos, tipo_shapefile.lower()
        )
        logger.info("Aplicando Simbologia padrao")
        arcpy.ApplySymbologyFromLayer_management(camada, simbologia)

    @staticmethod
    def zoomDataFramePrincipal(camada, logger):
        def arredondarEscala(numero):
            if 1 <= numero <= 10000:
                return math.ceil(numero / 500) * 500
            elif 10000 < numero <= 100000:
                return math.ceil(numero / 2000) * 2000
            elif numero > 100000:
                return math.ceil(numero / 10000) * 10000

        logger.info("Aplicando Zoom controlado")
        df_principal = Layout.carregarComponentes("df")["PRINCIPAL"]
        camada_extent = camada.getSelectedExtent(camada)
        df_principal.extent = camada_extent

        df_escala = int(df_principal.scale)
        escala_base = int(str(df_escala)[:2]) * 3
        zeros = len(str(df_escala)[2:]) * "0"
        escala_nova = int(str(escala_base) + zeros)
        df_principal.scale = arredondarEscala(escala_nova)

    @staticmethod
    def atualizarVariaveisDinamicasProjeto(dict_params):

        def atualizarNomeAreaDeInteresseLegenda(credits):
            mxd = Layout.carregarComponentes("mxd")
            camadas = Layout.carregarComponentes("cmds_desc")
            legenda = arcpy.mapping.ListLayoutElements(mxd, "LEGEND_ELEMENT")[0]

            for desc, cmd in camadas.items():
                if desc.startswith("AREA_DE_INTERESSE"):
                    cmd.name = credits
                    camada_interesse = cmd

            if camada_interesse:
                legenda.updateItem(camada_interesse)

        mxd = Layout.carregarComponentes("mxd")

        autor = "User:{} PC:{}".format(getpass.getuser(), platform.node())
        ano_numero = "{}/{}".format(
            dict_params["ano"].value, dict_params["numero"].value
        )
        denominacao = "{}".format(dict_params["denominacao"].value)

        situacao_interessado = "{} - {}".format(
            dict_params["situacao"].value, dict_params["interessado"].value
        )

        zonas = []
        if dict_params["lista_zoneamento"].values:
            for v in dict_params["lista_zoneamento"].values:
                if "ZEE: " in v:
                    nv = v.split("ZEE: ")[-1]
                else:
                    nv = v.split("MZEE: ")[-1]
                zonas.append(nv)

        zonas_str = ", ".join(zonas)

        descricao_carta_zoneamento = "FOLHA/CIM / {} \nZEE: {}".format(
            dict_params["carta"].value, zonas_str
        )

        mxd.title = ano_numero.encode("cp1252").upper()
        mxd.summary = situacao_interessado.encode("cp1252").upper()
        mxd.description = descricao_carta_zoneamento.encode("cp1252").upper()
        mxd.author = autor.encode("cp1252").upper()
        mxd.credits = denominacao.encode("cp1252").upper()

        atualizarNomeAreaDeInteresseLegenda(denominacao.encode("cp1252").upper())

        arcpy.RefreshActiveView()

    @staticmethod
    def atualizarLayout(parametros, logger):
        if parametros[Parametros.btn_alterar_layout].altered:
            if parametros[Parametros.btn_alterar_layout].value:
                Layout.atualizarVariaveisDinamicasProjeto(parametros)
                parametros[Parametros.btn_alterar_layout].value = False

    @staticmethod
    def exportarMapa(params, logger):

        ano = params[Parametros.ano].valueAsText
        numero = params[Parametros.numero].valueAsText
        interessado = params[Parametros.interessado].valueAsText

        pasta_saida = params[Parametros.pasta_resultados].valueAsText
        extensao = ".pdf"

        nome_arquivo_saida = "{}_{}_{}".format(ano, numero, interessado)
        nome_arquivo_saida = Funcoes.normalizarString(nome_arquivo_saida)

        pasta_saida = Funcoes.criarSubPasta(pasta_saida, nome_arquivo_saida)

        caminho_saida_final = "{}\\{}{}".format(
            pasta_saida, nome_arquivo_saida, extensao
        )

        count = 1
        while os.path.exists(caminho_saida_final):
            caminho_saida_final = "{}\\{}_{}{}".format(
                pasta_saida, nome_arquivo_saida, count, extensao
            )
            count += 1

        caminho_saida_final

        mxd = Layout.carregarComponentes("mxd")
        logger.info("Exportando PDF...")
        arcpy.mapping.ExportToPDF(mxd, caminho_saida_final)
        logger.info("PDF Exportado com Sucesso!")
        # Funcoes.mostrarPopUpPDF()
        # os.system('start "" "{}"'.format(pasta_saida))

    @staticmethod
    def exportarPDF(parametros, logger, in_execute_mode=False):
        if in_execute_mode:
            logger.debug("Iniciando exportação do Mapa")
            Layout.exportarMapa(parametros, logger)
            logger.debug("Concluido processo de exportação")

        if parametros[Parametros.btn_exportar_mapa].altered:
            if parametros[Parametros.btn_exportar_mapa].value:

                logger.debug("Iniciando exportação do Mapa")
                Layout.exportarMapa(parametros, logger)
                logger.debug("Concluido processo de exportação")

                parametros[Parametros.btn_exportar_mapa].value = False


class Funcoes:

    @staticmethod
    def normalizarString(input_str):
        """Converter a string para Unicode normalizada (NFD).
        Remover os caracteres de acentuação e especiais.
        Substituir caracteres não alfanuméricos por underscore."""

        normalized_str = unicodedata.normalize("NFD", unicode(input_str, "utf-8"))  # type: ignore
        ascii_str = "".join(
            [c for c in normalized_str if unicodedata.category(c) != "Mn"]
        )
        ascii_str = re.sub(r"[^A-Za-z0-9]", "_", ascii_str)
        final_str = ascii_str.upper()
        return final_str

    @staticmethod
    def verificarPasta(pasta):
        if not os.path.exists(pasta):
            os.makedirs(pasta)
        return pasta

    @staticmethod
    def criarSubPasta(diretorio, sub):
        sub_pasta = os.path.join(diretorio, sub)
        Funcoes.verificarPasta(sub_pasta)
        return sub_pasta

    @staticmethod
    def criarYaml():
        dir_path = os.path.join(os.path.dirname(__file__), "dados")

        if not os.path.exists(dir_path):
            os.makedirs(dir_path)

        file_path = os.path.join(dir_path, "formsfields.yaml")

        if not os.path.exists(file_path):
            with io.open(file_path, "w", encoding="utf-8") as f:
                f.write("# YAML content goes here\n")

    @staticmethod
    def copyToClipboard(text):
        # Definições das funções e constantes da API do Windows
        CF_UNICODETEXT = 13

        # Funções da API do Windows
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        # Abre a área de transferência
        if not user32.OpenClipboard(None):
            raise ctypes.WinError()

        try:
            # Esvazia a área de transferência
            if not user32.EmptyClipboard():
                raise ctypes.WinError()

            # Aloca memória global
            hGlobal = kernel32.GlobalAlloc(
                0x2000, (len(text) + 1) * ctypes.sizeof(ctypes.c_wchar)
            )

            if not hGlobal:
                raise ctypes.WinError()

            # Bloqueia a memória global e copia o texto para ela
            locked_handle = kernel32.GlobalLock(hGlobal)

            if not locked_handle:
                kernel32.GlobalFree(hGlobal)
                raise ctypes.WinError()

            ctypes.cdll.msvcrt.wcscpy(ctypes.c_wchar_p(locked_handle), text)
            kernel32.GlobalUnlock(hGlobal)

            # Define os dados da área de transferência
            if not user32.SetClipboardData(CF_UNICODETEXT, hGlobal):
                kernel32.GlobalFree(hGlobal)
                raise ctypes.WinError()
        finally:
            # Fecha a área de transferência
            user32.CloseClipboard()

    @staticmethod
    def copiarParaAreaDeTransferencia(texto, logger):
        logger.info("Texto a copiar para cpliboard: {}{}".format(type(texto), texto))
        Funcoes.copyToClipboard(texto)
        # process = subprocess.Popen(['clip'], stdin=subprocess.PIPE, close_fds=True)
        # process.communicate(input=texto.encode('utf-8'))
        Funcoes.mostrarPopUpIncidenciasCopiadas(logger)

    @staticmethod
    def obterHorarioAtual(tipo=1):
        """Retorna o horário atual em varios formatos.
        - 0: "%H:%M",
        - 1: "%H:%M:%S",
        - 2: "%H_%M_%S",
        - 3: "%d%m%Y_%H%M%S",
        - 4: "%d/%m/%Y",
        """
        data = datetime.datetime.now().strftime
        tipos = {
            0: "%H:%M",
            1: "%H:%M:%S",
            2: "%H_%M_%S",
            3: "%d%m%Y_%H%M%S",
            4: "%d%m%y",
            5: "%d/%m/%Y",
        }
        return data(tipos.get(tipo, tipos[0]))

    @staticmethod
    def obterCartaCadastral(shapefile, logger):
        logger.info("Obtendo Cartas Cadastrais")
        cartas_cadastrais = Shapefile.obterValoresColunaIntersecao(
            shapefile_a_selecionar=r"AUXILIAR\CARTA_INDICE_IBGE_DSG",
            camada=shapefile,
            coluna="cint",
            prefixo="",
            retorno="str",
        )
        logger.info("Cartas: {}".format(cartas_cadastrais))
        return cartas_cadastrais

    @staticmethod
    def obterZoneamentos(shapefile, logger):
        logger.info("Obtendo Zoneamentos")
        zoneamentos = Shapefile.obterValoresColunaIntersecao(
            shapefile_a_selecionar=ValoresConstantes.camada_auxiliar_zee_2010,
            camada=shapefile,
            coluna="zona",
            prefixo="ZEE: ",
            retorno="list",
        )
        logger.info("Zoneamentos: {}".format(zoneamentos))
        return zoneamentos

    @staticmethod
    def obterMacroZoneamentos(shapefile, logger):
        logger.info("Obtendo Macro Zoneamentos")
        zoneamentos = Shapefile.obterValoresColunaIntersecao(
            shapefile_a_selecionar=ValoresConstantes.camada_auxiliar_mzee_2008,
            camada=shapefile,
            coluna="grupo",
            prefixo="MZEE: ",
            retorno="list",
        )
        logger.info("Macro Zoneamentos: {}".format(zoneamentos))
        return zoneamentos

    @staticmethod
    def obterIncidencias(camada_desejada, shapefile):
        # TODO Implementar funcionalidade
        return

    @staticmethod
    def mostrarPopUpPDF():

        hWnd = 0
        lpText = "Seu PDF foi exportado com sucesso!".encode("utf-8")
        lpCaption = "ArcpyAutoMap - Info PDF"
        uType = 0 | 64

        resultado = ctypes.windll.user32.MessageBoxA(hWnd, lpText, lpCaption, uType)

        return resultado

    @staticmethod
    def mostrarPopUp(logger):
        hWnd = 0
        lpText = "Deseja carregar os dados salvos do ultimo preenchimento da ferramenta?".encode(
            "utf-8"
        )
        lpCaption = "ArcpyAutoMap - Info Dados da Ferramenta"
        uType = 4 | 32

        logger.debug("mostrando PopUp!")
        resultado = ctypes.windll.user32.MessageBoxA(hWnd, lpText, lpCaption, uType)
        logger.info("Caixa de dialogo exibida.")
        message_box_returns = {
            1: "IDOK - Usuario clicou em OK",
            6: "IDYES - Usuario clicou em Sim",
            7: "IDNO - Usuario clicou em Não",
        }
        logger.info(
            "Caixa de dialogo retornou {}".format(message_box_returns[resultado])
        )

        return resultado

    @staticmethod
    def mostrarPopUpIncidenciasCopiadas(logger):
        hWnd = 0
        lpText = "Dados das incidencias copiados para sua Area de Transferencia (Crtl + V) com sucesso!".encode(
            "utf-8"
        )
        lpCaption = "ArcpyAutoMap - Dados Copiados"
        uType = 0 | 64

        logger.debug("mostrando PopUp!")
        resultado = ctypes.windll.user32.MessageBoxA(hWnd, lpText, lpCaption, uType)
        logger.info("Caixa de dialogo exibida.")
        message_box_returns = {
            1: "IDOK - Usuario clicou em OK",
            6: "IDYES - Usuario clicou em Sim",
            7: "IDNO - Usuario clicou em Não",
        }
        logger.info(
            "Caixa de dialogo retornou {}".format(message_box_returns[resultado])
        )

    @staticmethod
    def exportarParaGDB(parametros, logger, in_execute_mode=False):
        logger.info("Iniciando Exportacao para GeoDatabase")

        shp_str = parametros[Parametros.shapefile].valueAsText
        if os.path.exists(shp_str):
            fc = shp_str
        else:
            fc = parametros[Parametros.shapefile].value.dataSource

        database = Database()
        database.adicionarFeatureclass(fc, parametros)

        logger.info("Exportacao GeoDatabase concluida com Sucesso!")

    @staticmethod
    def exportarDespacho(parametros, logger, in_execute_mode=False):
        if in_execute_mode:
            logger.debug("Iniciando exportação do Despacho")
            despacho = Despacho()
            despacho.buildDocx(parametros)
            logger.debug("Concluido processo de exportação")

        if parametros[Parametros.btn_exportar_despacho].altered:
            if parametros[Parametros.btn_exportar_despacho].value:

                logger.debug("Iniciando exportação do Despacho")
                despacho = Despacho()
                despacho.buildDocx(parametros)
                logger.debug("Concluido processo de exportação")

                parametros[Parametros.btn_exportar_despacho].value = False

    @staticmethod
    def exportaPlanilhaExcel(parametros, logger, in_execute_mode=False):
        if in_execute_mode:
            logger.debug("Iniciando exportação da Planilha")
            planilha = Planilha()
            planilha.buildXlsx(parametros)
            logger.debug("Concluido processo de exportação")

        if parametros[Parametros.btn_exportar_planilha].altered:
            if parametros[Parametros.btn_exportar_planilha].value:

                logger.debug("Iniciando exportação da Planilha")
                planilha = Planilha()
                planilha.buildXlsx(parametros)
                logger.debug("Concluido processo de exportação")

                parametros[Parametros.btn_exportar_planilha].value = False
        return


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Ferramenta"
        self.alias = ""
        self.tools = [AutoMap]


class AutoMap(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Novo Processo"
        self.description = ""
        self.canRunInBackground = False
        Funcoes.verificarPasta(os.path.join(ValoresConstantes.diretorio_raiz, "log"))
        _logger = Logger()
        self.logger = _logger.setup_logger()
        Funcoes.criarYaml()
        self.logger.debug("INICIALIZANDO FERRAMENTA AUTOMAP")
        self.logger.debug("-" * 32)

    def getParameterInfo(self):
        """Define parameter definitions"""
        parametros = Parametros()
        for param in parametros.params:
            self.logger.debug('Criando Parametro: "{}"'.format(param.name))

        return parametros.params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""

        parametros = Parametros.dicionarioDeParametros(parameters)
        Parametros.atualizarCampos(parametros, self.logger)
        Parametros.salvarDadosFerramenta(parametros, self.logger)

        Parametros.atualizarDataFramePrincipal(parametros, self.logger)
        Parametros.atualizarMunicipio(parametros, self.logger)

        Shapefile.encontrarCartaCadastral(parametros, self.logger)
        Shapefile.encontrarZoneamento(parametros, self.logger)

        Layout.atualizarLayout(parametros, self.logger)
        Layout.exportarPDF(parametros, self.logger)
        Funcoes.exportarDespacho(parametros, self.logger)
        Funcoes.exportaPlanilhaExcel(parametros, self.logger)

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        parametros = Parametros.dicionarioDeParametros(parameters)
        Parametros.salvarDadosFerramenta(parametros, self.logger)
        Layout.exportarPDF(parametros, self.logger, in_execute_mode=True)
        Funcoes.exportarDespacho(parametros, self.logger, in_execute_mode=True)
        Funcoes.exportaPlanilhaExcel(parametros, self.logger, in_execute_mode=True)
        Funcoes.exportarParaGDB(parametros, self.logger, in_execute_mode=True)
        os.system(
            'start "" "{}"'.format(parametros[Parametros.pasta_resultados].valueAsText)
        )
