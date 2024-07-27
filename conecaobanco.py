import traceback
import psycopg2
from datetime import datetime

shp = r"C:\github\ArcpyAutoMap\database\areas_de_interesse.shp"


conn = psycopg2.connect(
    dbname="dbteste",
    user="postgres",
    password="102045",
    host="localhost",
    port="5432",
)
cursor = conn.cursor()


try:
    with arcpy.da.SearchCursor(
        shp,
        [
            "SHAPE@",
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
        ],
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
            print("data_atual: {}".format(data_atual))
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
                print("Executando Insercao de Feicao no banco PostgreSQL")
                print("geom.WKT: {}".format(geom.WKT))
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
                print("SQL executando")

except Exception as e:
    exc_info = traceback.format_exc()
    print(exc_info)
finally:
    # Commit e fechamento da conexão
    conn.commit()
    cursor.close()
    conn.close()


class ConexaoBanco(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Conexao Banco"
        self.description = ""
        self.canRunInBackground = False
        Funcoes.verificarPasta(os.path.join(ValoresConstantes.diretorio_raiz, "log"))
        _logger = Logger()
        self.logger = _logger.setup_logger()
        self.logger.debug("INICIALIZANDO FERRAMENTA CONEXAO BANCO")
        self.logger.debug("-" * 32)

    def getParameterInfo(self):
        """Define parameter definitions"""

        return []

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        database = Database()
        database.ler_dados_banco()

        return
