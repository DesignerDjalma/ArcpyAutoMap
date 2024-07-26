import arcpy  # type: ignore
import getpass
import os


def texto(string):
    return string


def formatarDpt(lista):
    return {i: i for i in lista}


def adicionarDataHoraEmString():
    return


def adicionarHoraMinutoSegundoEmString():
    return


def retornaDiaMesAnoAtual():
    return


def exportarResultados(dpt, messages):
    """Exporta os resultados no estilo padrao Analise Limites"""
    messages.addMessage(texto("Executando Exportar Resultados."))
    new_dpt = formatarDpt(dpt)
    messages.addMessage(texto("NEW Dpt: {}".format(new_dpt)))
    pasta_gdp = criarGDB(new_dpt["pasta_resultados"], messages)

    # subprocess.Popen('explorer {}'.format('\\'.join(pasta_gdp.split('\\')[:-1])))
    shp = criarFCdentroGDB(pasta_gdp, messages)
    messages.addMessage(texto("Pronto para executar replaceTDA"))
    copiarEcolarFCGDB(
        feature_shp=new_dpt["shapefile"], feature_alvo_gdb=shp, messages=messages
    )
    replaceTDA(feature=shp, dpt=new_dpt, messages=messages)
    arcpy.RefreshCatalog(pasta_gdp)


def verificaPasta(pasta):
    """Verifica se a pasta saida existe, se nao ele cria"""
    if not os.path.exists(pasta):
        os.makedirs(pasta)


def criarGDB(pasta_saida, messages, nome_gdb="Resultado", versao="10.0"):
    """Cria um .gdb Resultado para armazenar um Shapefile"""
    nome_gdb = "{}{}".format(adicionarHoraMinutoSegundoEmString(), nome_gdb)
    verificaPasta(pasta_saida)  # Verifica se pasta existe
    # Cria um GDB Resultado com versao compativel com arcMap 10x
    messages.addMessage(texto("Executando: Criar GeodataBase"))
    messages.addMessage(
        texto("pasta_saida: {}\nNome GDB: {}".format(pasta_saida, nome_gdb))
    )
    adicionarDataHoraEmString()
    arcpy.management.CreateFileGDB(
        "{}".format(pasta_saida), nome_gdb, versao  # Nome  # Versao de compatibilidade
    )
    messages.addMessage(
        texto("Retornando: {}".format(os.path.join(pasta_saida, nome_gdb + ".gdb")))
    )
    return os.path.join(pasta_saida, nome_gdb + ".gdb")


def criarFCdentroGDB(
    pasta_saida_gdb, messages, nome_fc="LimiteImovel", template_padrao="TemplatePadrao"
):
    """Cria o Shapefile dentro do .gdb com o Template adequado e padrao"""
    # Cria uma FeatureClass Vazia com Projecao Sirgas 2000
    messages.addMessage(texto("Executando Criar FC dentro GDB"))
    messages.addMessage(
        texto(
            "pasta_saida_gdb: {}\nnome_fc: {}".format(
                pasta_saida_gdb,
                nome_fc,
            )
        )
    )
    messages.addMessage(texto("Executando CreateFeatureclass"))

    arcpy.management.CreateFeatureclass(
        "{}".format(pasta_saida_gdb),
        nome_fc,
        "POLYGON",
        template_padrao,  # Template Padrao Pra TDA # QUE FICA NA
        "ENABLED",
        "ENABLED",
        arcpy.SpatialReference(4674),
    )

    # retorna a feature class dentro do banco GDB
    return os.path.join(pasta_saida_gdb, nome_fc)


def copiarEcolarFCGDB(feature_shp, feature_alvo_gdb, messages):
    messages.addMessage(texto("Executando copia"))
    temp_workspace = r"C:\Users\{}\Documents\ArcMapTempWorkspace".format(
        getpass.getuser()
    )
    verificaPasta(temp_workspace)
    arcpy.env.workspace = temp_workspace
    edit = arcpy.da.Editor(arcpy.env.workspace)
    edit.startEditing(True)
    edit.startOperation()
    arcpy.management.Append(feature_shp, feature_alvo_gdb, "NO_TEST")
    edit.stopOperation()
    edit.stopEditing(True)
    messages.addMessage(texto("Finalizado copia"))
    arcpy.RefreshActiveView()


def replaceTDA(feature, dpt, messages):
    dicionario = {
        "interessad": "interessado",
        "imovel": "denominacao",
        "ano": "ano",
        "processo": "numero",
        "municipio": "municipio",
        "parcela": "parcela",
        "situacao": "situacao",
        "georref": "georreferenciamento",
        "data": "data",
        "complement": "complemento",
    }

    parametros = [
        "OBJECTID",
        "Shape",  # ignoraveis
        # Informacoes Obrigatorias
        "interessad",
        "imovel",
        "ano",
        "processo",
        "municipio",
        "parcela",
        "situacao",
        "georref",
        "data",
        "complement",
        # detalhes
        "Shape_Length",
        "Shape_Area",
    ]

    skip = ["OBJECTID", "Shape", "Shape_Length", "Shape_Area"]

    messages.addMessage(texto("Exutando funcao UpdateCursor"))
    messages.addMessage(texto("feature: {}".format(feature)))

    # Adicionando Data Atual ao Dicionarios de Parametros
    dpt["data"] = retornaDiaMesAnoAtual()

    # Operacao de colocação das informações na Tabela de Atributos
    with arcpy.da.UpdateCursor(feature, "*") as cursor:

        for linha in cursor:
            messages.addMessage(texto("1st for: linha: {}".format(linha)))

            for index, parametro in enumerate(parametros):
                if parametro not in skip:
                    if dicionario[parametro] in ["situacao", "zoneamento"]:
                        resultado = dpt[dicionario[parametro]][1:-1]
                    else:
                        resultado = dpt[dicionario[parametro]]

                    linha[index] = resultado

            messages.addMessage(texto("updating linha".format()))
            cursor.updateRow(linha)
            messages.addMessage(texto("break"))
            break

    messages.addMessage(texto("Refreshing"))
    arcpy.RefreshActiveView()
