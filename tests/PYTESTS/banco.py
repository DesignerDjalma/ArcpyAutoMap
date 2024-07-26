import psycopg2
import subprocess

def shp2portgres(db_host, db_name, db_user, db_password, db_port, schema, table, shapefile_path, srid):
    output_sql = 'output.sql'

    # Comando shp2pgsql
    shp2pgsql_cmd = [
        r'C:\Program Files\PostgreSQL\13\bin\shp2pgsql',  # Certifique-se de usar o caminho completo
        '-I',  # Cria um índice GiST no campo geometria
        '-s', str(srid),  # Define o SRID
        shapefile_path,
        '{}.{}'.format(schema, table)
    ]

    # Converte o shapefile em SQL usando shp2pgsql e salva em um arquivo
    with open(output_sql, 'w') as output_file:
        shp2pgsql_process = subprocess.Popen(shp2pgsql_cmd, stdout=output_file)
        shp2pgsql_process.wait()

    if shp2pgsql_process.returncode != 0:
        print('Erro ao converter shapefile')
        return False

    # Lê o conteúdo do arquivo SQL gerado
    with open(output_sql, 'r') as sql_file:
        sql_commands = sql_file.read()

    # Conecta ao banco de dados PostgreSQL
    try:
        conn = psycopg2.connect(
            dbname=db_name,
            user=db_user,
            password=db_password,
            host=db_host,
            port=db_port
        )
        cur = conn.cursor()

        # Executa os comandos SQL no PostgreSQL
        cur.execute(sql_commands)
        conn.commit()

        cur.close()
        conn.close()

        print('Shapefile importado com sucesso!')
        return True

    except psycopg2.DatabaseError as e:
        print('Erro ao executar comando no PostgreSQL: {}'.format(e))
        return False

# Configurações do banco de dados
DB_HOST = 'localhost'
DB_NAME = 'dbteste'
DB_USER = 'postgres'
DB_PASSWORD = '102045'
DB_PORT = '5432'
SCHEMA = 'public'
TABLE = 'tbteste'
SHAPEFILE_PATH = r"C:\pytscript\TESTS\Export_Output.shp"
SRID = 4674

# Chama a função para importar o shapefile
shp2portgres(DB_HOST, DB_NAME, DB_USER, DB_PASSWORD, DB_PORT, SCHEMA, TABLE, SHAPEFILE_PATH, SRID)
