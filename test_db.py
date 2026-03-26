# Crea un archivo test_opensearch.py
from dotenv import load_dotenv
import os
from opensearchpy import OpenSearch

load_dotenv()

try:
    auth = f"{os.getenv('OPENSEARCH_USER')}:{os.getenv('OPENSEARCH_PASSWORD')}"
    url = f"https://{auth}@{os.getenv('OPENSEARCH_HOST')}:{os.getenv('OPENSEARCH_PORT')}"
    
    client = OpenSearch(url, use_ssl=True, verify_certs=True)
    info = client.info()
    print(f" Conexión exitosa a OpenSearch ")
    print(f" Versión: {info['version']['number']}")
except Exception as e:
    print(f" Error: {e}")