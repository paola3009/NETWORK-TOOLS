"""
NETWORK TOOLS - Ponto de Entrada Principal
Executar:
    python main.py
Acessar:
    http://127.0.0.1:5000
"""

import sys
import os

# Adiciona o diretório raiz ao PYTHONPATH para resolução de módulos limpa
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from core.app import app, HOST, PORT

if __name__ == "__main__":
    app.run(host=HOST, port=PORT, debug=False)
