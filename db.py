import pyodbc
import pandas as pd


def get_connection():
    return pyodbc.connect(
        "DRIVER={ODBC Driver 18 for SQL Server};"
        "SERVER=192.168.0.13;"
        "DATABASE=mapexbp;"  # ⚠️ luego ajustaremos si hace falta
        "UID=FormacioSQL;"
        "PWD=Formacio2025;"
        "TrustServerCertificate=yes;"
    )


def get_escandallo_data(acabado: str):
    conn = get_connection()

    query = """
    SELECT TABLE_NAME 
    FROM INFORMATION_SCHEMA.TABLES
    """

    df = pd.read_sql(query, conn)
    conn.close()

    return df