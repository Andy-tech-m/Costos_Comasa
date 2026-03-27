"""
COMASA PRO — Backend Flask + MySQL + OpenSearch (Aiven)
Rutas API para catálogo, cotizaciones y búsqueda avanzada
"""

from flask import Flask, jsonify, request, render_template
from flask_cors import CORS
import mysql.connector
from mysql.connector import Error
from opensearchpy import OpenSearch, helpers
import os
from dotenv import load_dotenv
import json

# Cargar variables de entorno
load_dotenv()

app = Flask(__name__)
CORS(app, origins=["https://costos-comasa.vercel.app", "http://localhost:5000"])

# ─── CONFIGURACIÓN BASE DE DATOS ───
DB_CONFIG = {
    'host': os.getenv('DB_HOST'),
    'port': int(os.getenv('DB_PORT', 14628)),
    'database': os.getenv('DB_NAME', 'defaultdb'),
    'user': os.getenv('DB_USER', 'avnadmin'),
    'password': os.getenv('DB_PASSWORD'),
    'charset': 'utf8mb4',
    'use_unicode': True,
    'ssl_disabled': False,
}

# ─── CONFIGURACIÓN OPENSEARCH (AIVEN) ───
def get_opensearch_client():
    """Crea y retorna un cliente conectado a Aiven OpenSearch"""
    try:
        host = os.getenv('OPENSEARCH_HOST', 'os-34be3cef-andymaycoperalescaajamalqui-2892.j.aivencloud.com')
        port = os.getenv('OPENSEARCH_PORT', '14628')
        user = os.getenv('OPENSEARCH_USER', 'avnadmin')
        password = os.getenv('OPENSEARCH_PASSWORD')
        
        if not password:
            print("[WARNING] OPENSEARCH_PASSWORD no está configurada en .env")
            return None
        
        # Construir URL para Aiven
        auth = f"{user}:{password}"
        url = f"https://{auth}@{host}:{port}"
        
        client = OpenSearch(
            url,
            use_ssl=True,
            verify_certs=True,
            timeout=30,
            max_retries=3,
            retry_on_timeout=True
        )
        
        # Probar conexión
        info = client.info()
        print(f"[OPENSEARCH] Conectado a versión: {info['version']['number']}")
        return client
        
    except Exception as e:
        print(f"[OPENSEARCH ERROR] {e}")
        return None

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
#  API — CATEGORÍAS (TUS RUTAS EXISTENTES)
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
#  API — PRODUCTOS E INVENTARIO (TUS RUTAS EXISTENTES)
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
        
        # 🔥 NUEVO: Indexar en OpenSearch automáticamente
        indexar_producto_en_opensearch(new_id)
        
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
        
        # 🔥 NUEVO: Actualizar en OpenSearch
        actualizar_producto_en_opensearch(pid)
        
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
        
        # 🔥 NUEVO: Eliminar de OpenSearch
        eliminar_producto_de_opensearch(pid)
        
        return jsonify({'message': 'Producto desactivado'})
    except Error as e:
        return jsonify({'error': str(e)}), 500
    finally:
        cur.close(); conn.close()


# ════════════════════════════════════════
#  🚀 NUEVAS RUTAS PARA OPENSEARCH (AIVEN)
# ════════════════════════════════════════

def indexar_producto_en_opensearch(producto_id):
    """Indexa un producto específico en OpenSearch"""
    conn = get_db()
    if not conn:
        return False
    
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute("""
            SELECT p.id_producto, p.codigo, p.descripcion,
                   p.peso_unit, p.unidad, p.activo,
                   c.id_categoria, c.nombre AS categoria_nombre
            FROM productos p
            JOIN categorias_producto c ON p.id_categoria = c.id_categoria
            WHERE p.id_producto = %s
        """, (producto_id,))
        
        producto = cur.fetchone()
        if not producto:
            return False
        
        # Convertir Decimal a float
        producto['peso_unit'] = float(producto['peso_unit'])
        
        # Conectar a OpenSearch
        os_client = get_opensearch_client()
        if not os_client:
            return False
        
        # Indexar documento
        response = os_client.index(
            index='comasa_productos',
            id=str(producto['id_producto']),
            body=producto,
            refresh=True
        )
        
        print(f"[OPENSEARCH] Producto {producto_id} indexado: {response['result']}")
        return True
        
    except Exception as e:
        print(f"[OPENSEARCH ERROR] Error indexando producto {producto_id}: {e}")
        return False
    finally:
        if conn:
            conn.close()

def actualizar_producto_en_opensearch(producto_id):
    """Actualiza un producto en OpenSearch"""
    return indexar_producto_en_opensearch(producto_id)  # Misma función, sobreescribe

def eliminar_producto_de_opensearch(producto_id):
    """Elimina un producto de OpenSearch"""
    try:
        os_client = get_opensearch_client()
        if not os_client:
            return False
        
        response = os_client.delete(
            index='comasa_productos',
            id=str(producto_id),
            ignore=[404]  # No error si no existe
        )
        
        print(f"[OPENSEARCH] Producto {producto_id} eliminado: {response['result']}")
        return True
        
    except Exception as e:
        print(f"[OPENSEARCH ERROR] Error eliminando producto {producto_id}: {e}")
        return False

@app.route('/api/opensearch/sincronizar', methods=['POST'])
def sincronizar_productos():
    """
    Sincroniza TODOS los productos activos de MySQL a OpenSearch
    Útil para la migración inicial
    """
    conn = get_db()
    if not conn:
        return jsonify({'error': 'No se pudo conectar a MySQL'}), 500
    
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute("""
            SELECT p.id_producto, p.codigo, p.descripcion,
                   p.peso_unit, p.unidad, p.activo,
                   c.id_categoria, c.nombre AS categoria_nombre
            FROM productos p
            JOIN categorias_producto c ON p.id_categoria = c.id_categoria
            WHERE p.activo = 1
        """)
        
        productos = cur.fetchall()
        
        if not productos:
            return jsonify({'message': 'No hay productos activos para sincronizar'}), 200
        
        # Convertir Decimal a float
        for p in productos:
            p['peso_unit'] = float(p['peso_unit'])
        
        # Conectar a OpenSearch
        os_client = get_opensearch_client()
        if not os_client:
            return jsonify({'error': 'No se pudo conectar a OpenSearch'}), 500
        
        # Preparar acciones para bulk insert
        actions = []
        for producto in productos:
            actions.append({
                '_index': 'comasa_productos',
                '_id': str(producto['id_producto']),
                '_source': producto
            })
        
        # Ejecutar bulk insert
        success, failed = helpers.bulk(
            os_client,
            actions,
            stats_only=True,
            raise_on_error=False
        )
        
        return jsonify({
            'message': 'Sincronización completada',
            'total_productos': len(productos),
            'exitos': success,
            'fallos': failed
        }), 200
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if conn:
            conn.close()

@app.route('/api/opensearch/buscar', methods=['POST'])
def buscar_productos_opensearch():
    """
    Búsqueda avanzada de productos usando OpenSearch
    Body JSON: {
        "query": "texto a buscar",
        "categoria_id": 1 (opcional),
        "limit": 20 (opcional)
    }
    """
    data = request.get_json()
    query_text = data.get('query', '').strip()
    categoria_id = data.get('categoria_id')
    limit = data.get('limit', 20)
    
    if not query_text:
        return jsonify({'error': 'Query de búsqueda requerido'}), 400
    
    os_client = get_opensearch_client()
    if not os_client:
        return jsonify({'error': 'No se pudo conectar a OpenSearch'}), 500
    
    try:
        # Construir query de OpenSearch
        must_conditions = []
        
        # Búsqueda multi-campo
        must_conditions.append({
            "multi_match": {
                "query": query_text,
                "fields": ["codigo^3", "descripcion^2", "categoria_nombre"],
                "fuzziness": "AUTO"
            }
        })
        
        # Filtrar por categoría si se especifica
        if categoria_id:
            must_conditions.append({
                "term": {"id_categoria": categoria_id}
            })
        
        search_query = {
            "size": limit,
            "query": {
                "bool": {
                    "must": must_conditions
                }
            },
            "sort": [
                {"_score": {"order": "desc"}},
                {"codigo": {"order": "asc"}}
            ]
        }
        
        # Ejecutar búsqueda
        response = os_client.search(
            index='comasa_productos',
            body=search_query
        )
        
        # Formatear resultados
        resultados = []
        for hit in response['hits']['hits']:
            producto = hit['_source']
            producto['score'] = hit['_score']
            resultados.append(producto)
        
        return jsonify({
            'total': response['hits']['total']['value'],
            'resultados': resultados,
            'tiempo_ms': response['took']
        }), 200
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/opensearch/verificar', methods=['GET'])
def verificar_opensearch():
    """Verifica el estado de la conexión a OpenSearch"""
    os_client = get_opensearch_client()
    if not os_client:
        return jsonify({
            'status': 'error',
            'message': 'No se pudo conectar a OpenSearch'
        }), 500
    
    try:
        info = os_client.info()
        indices = os_client.indices.get_alias()
        
        return jsonify({
            'status': 'ok',
            'version': info['version']['number'],
            'indices': list(indices.keys()),
            'message': 'Conexión exitosa a OpenSearch'
        })
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500

@app.route('/api/opensearch/indices', methods=['GET'])
def listar_indices_opensearch():
    """Lista todos los índices en OpenSearch"""
    os_client = get_opensearch_client()
    if not os_client:
        return jsonify({'error': 'No se pudo conectar a OpenSearch'}), 500
    
    try:
        indices = os_client.indices.get_alias()
        return jsonify({
            'indices': list(indices.keys())
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ════════════════════════════════════════
#  HEALTH CHECK (MEJORADO)
# ════════════════════════════════════════
@app.route('/api/health')
def health():
    conn = get_db()
    db_status = 'conectado' if conn else 'sin conexión'
    
    os_client = get_opensearch_client()
    os_status = 'conectado' if os_client else 'sin conexión'
    
    if conn:
        conn.close()
    
    return jsonify({
        'status': 'ok',
        'mysql': db_status,
        'opensearch': os_status
    })


if __name__ == '__main__':
    app.run(debug=True, port=5000)