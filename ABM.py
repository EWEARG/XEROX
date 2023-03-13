#Python Code
import sqlite3

#establish database connection
conn = sqlite3.connect('GTIC.db')


# Connect to the database
conn = sqlite3.connect('xrx_impresoras.db')

# Create a cursor
c = conn.cursor()

# Function to insert records
def insert_record(oblea, direccionip, modeloid, ubicacionid, responsableequipo, ubicacioncompleta, gerenciasector):
    with conn:
        c.execute("INSERT INTO xrx_impresoras VALUES (:oblea, :direccionip, :modeloid, :ubicacionid, :responsableequipo, :ubicacioncompleta, :gerenciasector)",
                  {'oblea': oblea, 'direccionip': direccionip, 'modeloid': modeloid, 'ubicacionid': ubicacionid, 'responsableequipo': responsableequipo, 'ubicacioncompleta': ubicacioncompleta, 'gerenciasector': gerenciasector})

# Function to view all records
def view_all_records():
    with conn:
        c.execute("SELECT * FROM xrx_impresoras")
        records = c.fetchall()
        return records

# Function to update a record
def update_record(oblea, direccionip, modeloid, ubicacionid, responsableequipo, ubicacioncompleta, gerenciasector):
    with conn:
        c.execute("UPDATE xrx_impresoras SET direccionip = :direccionip, modeloid = :modeloid, ubicacionid = :ubicacionid, responsableequipo = :responsableequipo, ubicacioncompleta = :ubicacioncompleta, gerenciasector = :gerenciasector WHERE oblea = :oblea",
                  {'oblea': oblea, 'direccionip': direccionip, 'modeloid': modeloid, 'ubicacionid': ubicacionid, 'responsableequipo': responsableequipo, 'ubicacioncompleta': ubicacioncompleta, 'gerenciasector': gerenciasector})

# Function to delete a record
def delete_record(oblea):
    with conn:
        c.execute("DELETE FROM xrx_impresoras WHERE oblea = :oblea",
                  {'oblea': oblea})

# Commit changes
conn.commit()

# Close connection
conn.close()

