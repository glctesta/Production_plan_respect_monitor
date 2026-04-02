# example.py
from config_manager import ConfigManager
from db_connection import DatabaseConnection
#import pyodbc

def main():
    # Prima volta: salva le credenziali (da eseguire una sola volta)
    # Stampa i driver disponibili
    # print("Driver ODBC disponibili:")
    # for driver in pyodbc.drivers():
    #     print(driver)
    config_manager = ConfigManager()
    # config_manager.save_config(
    #          driver='ODBC Driver 18 for SQL Server',
    #          server='roghipsql01.vandewiele.local\\emsreset',
    #          database='ResetServices',
    #          username='emsreset',
    #          password='E6QhqKUxHFXTbkB7eA8c9ya'
    #      )

    # Utilizzo della connessione
    db = DatabaseConnection(config_manager)
    try:
        conn = db.connect()
        cursor = conn.cursor()

        # Test della connessione
        cursor.execute("SELECT @@VERSION")
        row = cursor.fetchone()
        #print("Connessione riuscita!")
        #print("Versione SQL Server:", row[0])

        # Esegui altre query qui
        cursor.execute("SELECT top 1 * from tbsocieta;")
        results = cursor.fetchall()
        for row in results:
            print(row)

        # Chiudi esplicitamente il cursore
        cursor.close()

    except Exception as e:
        print(f"Errore durante l'esecuzione: {str(e)}")
    finally:
        # Chiudi esplicitamente la connessione
        db.disconnect()


if __name__ == "__main__":
    main()