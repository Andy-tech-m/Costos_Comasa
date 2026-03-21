"""
COMASA PRO — Backend Flask + MySQL
Rutas API para catálogo y cotizaciones
"""

from flask import Flask, jsonify, request, render_template
from flask_cors import CORS
import mysql.connector
from mysql.connector import Error
import os

app = Flask(__name__)
CORS(app)

# ─── CONFIGURACIÓN BASE DE DATOS ───
DB_CONFIG = {
    'host': 'localhost',
    'port': 3306,
    'database': 'comasa',
    'user': 'root',       # ← cambia si tu usuario es distinto
    'password': '',        # ← pon tu contraseña de MySQL aquí
    'charset': 'utf8mb4',
    'use_unicode': True,
}

def get_db():
    """Retorna una conexión activa a MySQL."""
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        return conn
    except Error as e:
        print(f"[DB ERROR] {e}")
        return None


# ════════════════════════════════════════
#  PÁGINA PRINCIPAL
# ════════════════════════════════════════
@app.route('/')
def index():
    return render_template('index.html')


# ════════════════════════════════════════
#  API — CATEGORÍAS
# ════════════════════════════════════════
@app.route('/api/categorias', methods=['GET'])
def get_categorias():
    """Lista todas las categorías activas."""
    conn = get_db()
    if not conn:
        return jsonify({'error': 'No se pudo conectar a la base de datos'}), 500
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute("""
            SELECT id_categoria, nombre
            FROM categorias_producto
            WHERE activo = 1
            ORDER BY nombre
        """)
        rows = cur.fetchall()
        return jsonify(rows)
    except Error as e:
        return jsonify({'error': str(e)}), 500
    finally:
        cur.close(); conn.close()


@app.route('/api/categorias', methods=['POST'])
def crear_categoria():
    """Crea una nueva categoría."""
    data = request.get_json()
    nombre = (data.get('nombre') or '').strip()
    if not nombre:
        return jsonify({'error': 'El nombre es requerido'}), 400

    conn = get_db()
    if not conn:
        return jsonify({'error': 'No se pudo conectar a la base de datos'}), 500
    try:
        cur = conn.cursor(dictionary=True)
        # Verificar duplicado
        cur.execute("SELECT id_categoria FROM categorias_producto WHERE nombre = %s", (nombre,))
        if cur.fetchone():
            return jsonify({'error': f'La categoría "{nombre}" ya existe'}), 409

        cur.execute("""
            INSERT INTO categorias_producto (nombre, descripcion, activo)
            VALUES (%s, %s, 1)
        """, (nombre, data.get('descripcion', '')))
        conn.commit()
        new_id = cur.lastrowid
        return jsonify({'id_categoria': new_id, 'nombre': nombre, 'message': 'Categoría creada'}), 201
    except Error as e:
        return jsonify({'error': str(e)}), 500
    finally:
        cur.close(); conn.close()


# ════════════════════════════════════════
#  API — PRODUCTOS
# ════════════════════════════════════════
@app.route('/api/productos', methods=['GET'])
def get_productos():
    """
    Retorna productos. 
    ?categoria_id=X  → filtra por categoría
    ?q=texto         → búsqueda libre en descripción/código
    """
    categoria_id = request.args.get('categoria_id')
    q = request.args.get('q', '').strip()

    conn = get_db()
    if not conn:
        return jsonify({'error': 'No se pudo conectar a la base de datos'}), 500
    try:
        cur = conn.cursor(dictionary=True)
        sql = """
            SELECT p.id_producto, p.codigo, p.descripcion,
                   p.peso_unit, p.unidad,
                   c.id_categoria, c.nombre AS categoria
            FROM productos p
            JOIN categorias_producto c ON p.id_categoria = c.id_categoria
            WHERE p.activo = 1 AND c.activo = 1
        """
        params = []
        if categoria_id:
            sql += " AND p.id_categoria = %s"
            params.append(categoria_id)
        if q:
            sql += " AND (p.descripcion LIKE %s OR p.codigo LIKE %s)"
            params += [f'%{q}%', f'%{q}%']
        sql += " ORDER BY c.nombre, p.codigo"

        cur.execute(sql, params)
        rows = cur.fetchall()
        # Convertir Decimal → float para JSON
        for r in rows:
            r['peso_unit'] = float(r['peso_unit'])
        return jsonify(rows)
    except Error as e:
        return jsonify({'error': str(e)}), 500
    finally:
        cur.close(); conn.close()


@app.route('/api/productos', methods=['POST'])
def crear_producto():
    """Crea un nuevo producto en el catálogo."""
    data = request.get_json()

    # Validaciones
    required = ['codigo', 'descripcion', 'peso_unit', 'unidad', 'id_categoria']
    missing = [f for f in required if not data.get(f)]
    if missing:
        return jsonify({'error': f'Campos requeridos: {", ".join(missing)}'}), 400

    try:
        peso = float(data['peso_unit'])
        if peso <= 0:
            raise ValueError
    except (ValueError, TypeError):
        return jsonify({'error': 'peso_unit debe ser un número mayor a 0'}), 400

    conn = get_db()
    if not conn:
        return jsonify({'error': 'No se pudo conectar a la base de datos'}), 500
    try:
        cur = conn.cursor(dictionary=True)
        # Verificar código duplicado
        cur.execute("SELECT id_producto FROM productos WHERE codigo = %s", (data['codigo'],))
        if cur.fetchone():
            return jsonify({'error': f'El código "{data["codigo"]}" ya existe'}), 409

        cur.execute("""
            INSERT INTO productos (id_categoria, codigo, descripcion, peso_unit, unidad, activo)
            VALUES (%s, %s, %s, %s, %s, 1)
        """, (
            data['id_categoria'],
            data['codigo'].strip(),
            data['descripcion'].strip(),
            peso,
            data['unidad'].strip()
        ))
        conn.commit()
        new_id = cur.lastrowid
        return jsonify({
            'id_producto': new_id,
            'codigo': data['codigo'],
            'descripcion': data['descripcion'],
            'peso_unit': peso,
            'unidad': data['unidad'],
            'message': 'Producto creado exitosamente'
        }), 201
    except Error as e:
        return jsonify({'error': str(e)}), 500
    finally:
        cur.close(); conn.close()


@app.route('/api/productos/<int:pid>', methods=['PUT'])
def editar_producto(pid):
    """Edita un producto existente."""
    data = request.get_json()
    conn = get_db()
    if not conn:
        return jsonify({'error': 'No se pudo conectar a la base de datos'}), 500
    try:
        cur = conn.cursor()
        fields, params = [], []
        for f in ['descripcion', 'peso_unit', 'unidad', 'id_categoria']:
            if f in data:
                fields.append(f"{f} = %s")
                params.append(data[f])
        if not fields:
            return jsonify({'error': 'Nada que actualizar'}), 400
        params.append(pid)
        cur.execute(f"UPDATE productos SET {', '.join(fields)} WHERE id_producto = %s", params)
        conn.commit()
        return jsonify({'message': 'Producto actualizado'})
    except Error as e:
        return jsonify({'error': str(e)}), 500
    finally:
        cur.close(); conn.close()

@app.route('/api/productos/<int:pid>', methods=['DELETE'])
def eliminar_producto(pid):
    """Desactiva (soft delete) un producto."""
    conn = get_db()
    if not conn:
        return jsonify({'error': 'No se pudo conectar a la base de datos'}), 500
    try:
        cur = conn.cursor()
        cur.execute("UPDATE productos SET activo = 0 WHERE id_producto = %s", (pid,))
        conn.commit()
        return jsonify({'message': 'Producto desactivado'})
    except Error as e:
        return jsonify({'error': str(e)}), 500
    finally:
        cur.close(); conn.close()


# ════════════════════════════════════════
#  HEALTH CHECK
# ════════════════════════════════════════
@app.route('/api/health')
def health():
    conn = get_db()
    if conn:
        conn.close()
        return jsonify({'status': 'ok', 'db': 'conectado'})
    return jsonify({'status': 'error', 'db': 'sin conexión'}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)
    