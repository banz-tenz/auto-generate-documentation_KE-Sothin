import mysql.connector

# Database configuration
db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'user_result'
}

def insertIntoTranscript(name, file_name, file ):
    conn = None
    cursor = None
    try:
        conn = mysql.connector.connect(**db_config)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO transcript (name, file_name, file) VALUES(%s,%s,%s)",(name, file_name, file))
        conn.commit()
    except mysql.connector.Error as e:
        print(f'Error inserting into DB {e}')
    finally:
        if cursor:
            cursor.close()
        if conn and conn.is_connected():
            conn.close()


name = 'sothin'
file_name = 'sothin.docx'
file = 'Transcript/sothin.docx'


insertIntoTranscript(name, file_name, file)