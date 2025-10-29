import mysql.connector

# Database configuration
db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'user_result'
}

db_connection = None
cursor = None
try:
    # Corrected syntax: call the connect() function within the module
    db_connection = mysql.connector.connect(**db_config)
    
    # Check if the connection was successful
    if db_connection.is_connected():
        print("Connected to MySQL database successfully!")
        
        # Create a cursor object
        cursor = db_connection.cursor()
        
        # Execute the SQL query
        cursor.execute("SELECT * FROM certificates")
        
        # Fetch all results
        data = cursor.fetchall()
        
        # Print the data
        print("Fetched data from 'certificates' table:")
        print(data)

except mysql.connector.Error as err:
    print(f"Error: {err}")

finally:
    # Always ensure the cursor and connection are closed
    if cursor is not None:
        cursor.close()
    if db_connection is not None and db_connection.is_connected():
        db_connection.close()
        print("MySQL connection is closed.")

