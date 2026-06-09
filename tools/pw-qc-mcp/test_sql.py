import os
import pyodbc
from dotenv import load_dotenv

load_dotenv()

server = os.getenv("PWQC_SQL_SERVER", "localhost")
database = os.getenv("PWQC_SQL_DATABASE")
driver = os.getenv("PWQC_SQL_DRIVER", "ODBC Driver 18 for SQL Server")
trust_cert = os.getenv("PWQC_SQL_TRUST_CERT", "yes")

connection_string = (
    f"DRIVER={{{driver}}};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"Trusted_Connection=yes;"
    f"TrustServerCertificate={trust_cert};"
)

print(connection_string)

with pyodbc.connect(connection_string) as conn:
    cur = conn.cursor()
    cur.execute("SELECT DB_NAME() AS database_name, SYSDATETIME() AS server_time")
    row = cur.fetchone()
    print(row.database_name, row.server_time)