import mysql.connector

conn = mysql.connector.connect(
    host="mysql-2ab1aba3-andymaycoperalescajamalqui-2892.j.aivencloud.com",
    port=14628,
    user="avnadmin",
    password="AVNS_GShM8Z7YgKhrgb7wKhr",
    database="defaultdb",
    ssl_disabled=False
)

cursor = conn.cursor()
cursor.execute("SHOW TABLES;")
print("Tablas:", cursor.fetchall())

cursor.execute("SELECT COUNT(*) FROM categorias_producto;")
print("Categorias:", cursor.fetchone())

cursor.execute("SELECT COUNT(*) FROM productos;")
print("Productos:", cursor.fetchone())

cursor.close()
conn.close()