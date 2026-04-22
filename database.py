import os
import mysql.connector
from dotenv import load_dotenv

load_dotenv()

MYSQL_HOST = os.getenv("MYSQL_HOST")
MYSQL_PORT = int(os.getenv("MYSQL_PORT", 3306))
MYSQL_DATABASE = os.getenv("MYSQL_DATABASE")
MYSQL_USER = os.getenv("MYSQL_USER")
MYSQL_PASSWORD = os.getenv("MYSQL_PASSWORD")


def get_connection():
    return mysql.connector.connect(
        host=MYSQL_HOST,
        port=MYSQL_PORT,
        user=MYSQL_USER,
        password=MYSQL_PASSWORD,
        database=MYSQL_DATABASE,
    )