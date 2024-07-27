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
                print("Executado?")

except Exception as e:
    exc_info = traceback.format_exc()
    print(exc_info)
finally:
    # Commit e fechamento da conexão
    conn.commit()
    cursor.close()
    conn.close()
