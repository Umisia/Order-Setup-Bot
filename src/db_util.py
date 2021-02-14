import mysql.connector
import config
from logger import Logger

log = Logger(__name__).logger


class DB:
    def __init__(self):
        self.conn = mysql.connector.connect(
            host=config.db_host,
            port=config.db_port,
            user=config.db_user,
            password=config.db_pswd,
            database=config.db_database)

        self.cursor = self.conn.cursor(buffered=True)

        log.info("Connected to CC db.")

    def find_org(self, postcode, name):
        queries = []
        if postcode:
            queries.append(f"postcode = '{postcode}'")
            queries.append(f"""postcode = '{"".join(postcode.split())}'""")
        if name:
            queries.append(f"name = '{name}'")
        for query in queries:
            self.cursor.execute(f"SELECT name, id from lp_organizations WHERE {query}")
            result = self.cursor.fetchall()
            log.debug(f"{query}: {result}")
            if len(result) > 0:
                log.debug(f"result: {result}")
                return result

        return False

    def close_db(self):
        self.cursor.close()
        self.conn.close()
