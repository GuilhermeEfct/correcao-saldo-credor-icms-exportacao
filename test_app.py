# -*- coding: utf-8 -*-
"""
test_app.py — mini servidor LOCAL só para testar a ferramenta fora do HUB.
NÃO vai para produção; o app real é o do HUB.

Uso:
    pip install flask openpyxl
    python test_app.py
    # abre http://127.0.0.1:5000

Monta a mesma tela (screen.html/css/js) e registra o blueprint de routes.py com
TESTING=True, que libera a permissão localmente — em produção quem autoriza é o HUB.

Inclui um STUB de `GET /api/cnpj/<cnpj>`: essa rota é do HUB, não da ferramenta. O
stub existe apenas para o fluxo "CNPJ primeiro" poder ser exercitado offline, e
devolve razão social sintética.
"""

import os

from flask import Flask, jsonify, render_template_string

import routes

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _ler(nome):
    with open(os.path.join(BASE_DIR, nome), encoding="utf-8") as fh:
        return fh.read()


def criar_app():
    app = Flask(__name__)
    app.config["TESTING"] = True          # libera requer_permissao no ambiente local
    # Teto local do corpo da requisição. Para testar o caminho de excesso:
    #   TESTE_MAX_MB=1 python test_app.py
    teto_mb = int(os.environ.get("TESTE_MAX_MB", "512"))
    app.config["MAX_CONTENT_LENGTH"] = teto_mb * 1024 * 1024

    routes.init_app(app)
    app.register_blueprint(routes.bp)

    @app.route("/api/cnpj/<cnpj>")
    def stub_cnpj(cnpj):
        """STUB da rota do HUB. Em produção quem responde é o próprio HUB."""
        digitos = "".join(c for c in cnpj if c.isdigit())
        if len(digitos) != 14:
            return jsonify({"erro": "CNPJ inválido"}), 400
        return jsonify({
            "cnpj": digitos,
            "razao_social": "Empresa Demonstração Indústria Ltda (stub local)",
            "uf": "RS",
        })

    @app.route("/")
    def index():
        pagina = """<!doctype html><html lang="pt-br"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Teste — Correção do Saldo Credor de ICMS proporcional às Exportações</title>
<link href="https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Exo+2:wght@400;600;700&display=swap" rel="stylesheet">
<style>body{margin:0;background:#eef2f4;padding:24px}</style>
<style>__CSS__</style></head><body>
__HTML__
<script>__JS__</script>
</body></html>"""
        pagina = (pagina.replace("__CSS__", _ler("screen.css"))
                        .replace("__HTML__", _ler("screen.html"))
                        .replace("__JS__", _ler("screen.js")))
        return render_template_string(pagina)

    return app


if __name__ == "__main__":
    criar_app().run(debug=True, port=5000)
