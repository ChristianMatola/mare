import io
import os
from hmac import compare_digest
from urllib.parse import urlsplit

from flask import Flask, Response, request

from app import Routes, download_header


app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = 256 * 1024


class WebRequest(Routes):
    hosted = True

    def __init__(self):
        self.path = request.full_path
        self.headers = request.headers
        self.rfile = io.BytesIO(request.get_data()) if request.method == "POST" else io.BytesIO()

    def send_bytes(self, data, content_type="text/html; charset=utf-8", status=200, download=""):
        self.response = Response(data, status=status, content_type=content_type)
        if download:
            self.response.headers["Content-Disposition"] = download_header(download)

    def redirect(self, path):
        self.response = Response(status=303, headers={"Location": path})

    def send_error_page(self, exc):
        app.logger.error("Falha na rota MARÉ (%s)", type(exc).__name__)
        super().send_error_page(exc)


@app.after_request
def prevent_cache(response):
    response.headers["Cache-Control"] = "private, no-store"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    return response


@app.route("/", defaults={"path": ""}, methods=["GET", "POST"])
@app.route("/<path:path>", methods=["GET", "POST"])
def route(path):
    username = os.environ.get("MARE_WEB_USER")
    password = os.environ.get("MARE_WEB_PASSWORD")
    if not (os.environ.get("MARE_DATABASE_URL") and username and password):
        return "Ambiente web pendente de configuração.", 503
    credentials = request.authorization
    if not (
        credentials and credentials.type == "basic"
        and compare_digest((credentials.username or "").encode(), username.encode())
        and compare_digest((credentials.password or "").encode(), password.encode())
    ):
        return Response("Acesso restrito à equipe.", 401,
                        {"WWW-Authenticate": 'Basic realm="MARE", charset="UTF-8"'})

    if request.method == "POST":
        origin = urlsplit(request.headers.get("Origin", ""))
        if origin.netloc != request.host or origin.scheme not in {"http", "https"}:
            return "Origem da solicitação inválida.", 403
    if path == "gerar" and request.headers.get("Sec-Fetch-Site") == "cross-site":
        return "Origem da solicitação inválida.", 403

    handler = WebRequest()
    if request.method == "POST":
        handler.do_POST()
    else:
        handler.do_GET()
    return handler.response
