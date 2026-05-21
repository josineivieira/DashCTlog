from __future__ import annotations

import hashlib
import hmac
import html
import os
from http import HTTPStatus
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

import build_dashboard


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
INDEX_PATH = ROOT / "index.html"
ORDERS_UPLOAD = DATA_DIR / "ordens.xlsx"
DETAIL_UPLOAD = DATA_DIR / "geral_ct_log.xlsx"
MAX_UPLOAD_BYTES = 25 * 1024 * 1024
SESSION_COOKIE = "dashboard_log_session"


def app_user() -> str:
    return os.environ.get("DASH_USER", "admin")


def app_password() -> str:
    return os.environ.get("DASH_PASSWORD", "admin")


def app_secret() -> str:
    return os.environ.get("DASH_SECRET", "dashboard-log-secret")


def signed_session(user: str) -> str:
    signature = hmac.new(app_secret().encode(), user.encode(), hashlib.sha256).hexdigest()
    return f"{user}:{signature}"


def valid_session(value: str) -> bool:
    user, separator, signature = value.partition(":")
    if not separator:
        return False
    expected = signed_session(user).split(":", 1)[1]
    return hmac.compare_digest(signature, expected)


LOGIN_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Login - Dashboard Log</title>
  <style>
    :root {
      --bg: #10232b;
      --panel: #ffffff;
      --ink: #16212d;
      --muted: #657282;
      --line: #d7e0e8;
      --teal: #0c7c83;
      --error-bg: #fff1ed;
      --error: #9a3412;
      --shadow: 0 24px 60px rgba(0, 0, 0, .22);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
      padding: 24px;
      background: var(--bg);
      color: var(--ink);
      font-family: Inter, Segoe UI, Roboto, Arial, sans-serif;
    }
    .login {
      width: min(100%, 420px);
      background: var(--panel);
      border: 1px solid rgba(255, 255, 255, .36);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 28px;
    }
    h1 { margin: 0; font-size: 34px; letter-spacing: 0; }
    p { margin: 9px 0 24px; color: var(--muted); line-height: 1.5; }
    form { display: grid; gap: 14px; }
    label {
      display: grid;
      gap: 8px;
      color: var(--muted);
      font-size: 13px;
      font-weight: 800;
      text-transform: uppercase;
    }
    input {
      width: 100%;
      min-height: 44px;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 10px 12px;
      color: var(--ink);
      font: inherit;
    }
    input:focus {
      outline: 3px solid rgba(12, 124, 131, .18);
      border-color: var(--teal);
    }
    button {
      min-height: 44px;
      border: 0;
      border-radius: 8px;
      background: var(--teal);
      color: #fff;
      cursor: pointer;
      font: inherit;
      font-weight: 900;
    }
    .error {
      margin-bottom: 14px;
      padding: 11px 12px;
      border-radius: 8px;
      background: var(--error-bg);
      color: var(--error);
      font-weight: 800;
    }
  </style>
</head>
<body>
  <main class="login">
    <h1>Dashboard Log</h1>
    <p>Acesse para visualizar o dashboard ou atualizar as planilhas.</p>
    {message}
    <form method="post" action="/login">
      <label>Usuario
        <input name="user" autocomplete="username" required>
      </label>
      <label>Senha
        <input name="password" type="password" autocomplete="current-password" required>
      </label>
      <button type="submit">Entrar</button>
    </form>
  </main>
</body>
</html>
"""


HOME_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Home - Dashboard Log</title>
  <style>
    :root {
      --bg: #10232b;
      --panel: #ffffff;
      --ink: #16212d;
      --muted: #657282;
      --line: #d7e0e8;
      --teal: #0c7c83;
      --green: #3c8c4f;
      --gold: #9f7a1c;
      --shadow: 0 18px 42px rgba(0, 0, 0, .18);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      background: var(--bg);
      color: var(--ink);
      font-family: Inter, Segoe UI, Roboto, Arial, sans-serif;
    }
    header {
      padding: 30px clamp(16px, 4vw, 44px) 22px;
      color: #fff;
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 16px;
    }
    h1 { margin: 0; font-size: clamp(30px, 4vw, 48px); letter-spacing: 0; }
    .subtitle { margin: 8px 0 0; color: #c8d6dc; }
    a { color: inherit; text-decoration: none; }
    .logout {
      display: inline-flex;
      align-items: center;
      min-height: 36px;
      padding: 7px 12px;
      border: 1px solid rgba(255, 255, 255, .28);
      border-radius: 8px;
      color: #fff;
      font-size: 13px;
      font-weight: 800;
      background: rgba(255, 255, 255, .08);
    }
    main { padding: 0 clamp(16px, 4vw, 44px) 44px; }
    .menu {
      display: grid;
      grid-template-columns: repeat(2, minmax(220px, 1fr));
      gap: 18px;
      max-width: 920px;
    }
    .card {
      display: grid;
      gap: 12px;
      min-height: 190px;
      padding: 24px;
      background: var(--panel);
      border: 1px solid var(--line);
      border-top: 5px solid var(--teal);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }
    .card:nth-child(2) { border-top-color: var(--green); }
    .card span {
      color: var(--muted);
      font-size: 12px;
      font-weight: 900;
      text-transform: uppercase;
    }
    .card strong {
      color: var(--ink);
      font-size: clamp(25px, 3vw, 34px);
      line-height: 1.05;
    }
    .card p {
      margin: 0;
      color: var(--muted);
      line-height: 1.5;
    }
    .button {
      align-self: end;
      justify-self: start;
      min-height: 42px;
      padding: 10px 15px;
      border-radius: 8px;
      background: var(--teal);
      color: #fff;
      font-weight: 900;
    }
    .card:nth-child(2) .button { background: var(--green); }
    @media (max-width: 680px) {
      header { flex-direction: column; }
      .menu { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <header>
    <div>
      <h1>Dashboard Log</h1>
      <p class="subtitle">Escolha uma area para continuar.</p>
    </div>
    <a class="logout" href="/logout">Sair</a>
  </header>
  <main>
    <section class="menu">
      <a class="card" href="/dashboard">
        <span>Visualizacao</span>
        <strong>Dashboard</strong>
        <p>Acompanhe viagens, placas, produtos e terminais.</p>
        <div class="button">Abrir dashboard</div>
      </a>
      <a class="card" href="/editar">
        <span>Dados</span>
        <strong>Editar dados</strong>
        <p>Envie novas planilhas e atualize os indicadores.</p>
        <div class="button">Atualizar planilhas</div>
      </a>
    </section>
  </main>
</body>
</html>
"""


EDIT_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Editar dados - Dashboard Log</title>
  <style>
    :root {
      --bg: #10232b;
      --panel: #ffffff;
      --ink: #16212d;
      --muted: #657282;
      --line: #d7e0e8;
      --teal: #0c7c83;
      --shadow: 0 18px 42px rgba(0, 0, 0, .18);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      background: var(--bg);
      color: var(--ink);
      font-family: Inter, Segoe UI, Roboto, Arial, sans-serif;
    }
    header {
      padding: 28px clamp(16px, 4vw, 44px) 22px;
      color: #fff;
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 16px;
    }
    h1 { margin: 0; font-size: clamp(28px, 3vw, 42px); letter-spacing: 0; }
    .subtitle { margin: 8px 0 0; color: #c8d6dc; }
    a { color: inherit; text-decoration: none; }
    .nav {
      display: flex;
      gap: 9px;
      flex-wrap: wrap;
      justify-content: flex-end;
    }
    .top-link {
      display: inline-flex;
      align-items: center;
      min-height: 36px;
      padding: 7px 12px;
      border: 1px solid rgba(255, 255, 255, .28);
      border-radius: 8px;
      color: #fff;
      font-size: 13px;
      font-weight: 800;
      background: rgba(255, 255, 255, .08);
    }
    main { padding: 0 clamp(16px, 4vw, 44px) 42px; }
    .panel {
      max-width: 980px;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 22px;
    }
    .message {
      margin-bottom: 16px;
      padding: 12px 14px;
      border-radius: 8px;
      background: #e7f4f2;
      color: #0d6268;
      font-weight: 800;
    }
    .error { background: #fff1ed; color: #9a3412; }
    form { display: grid; gap: 18px; }
    .upload-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
    }
    label {
      display: grid;
      gap: 9px;
      color: var(--muted);
      font-size: 13px;
      font-weight: 800;
      text-transform: uppercase;
    }
    input[type="file"] {
      width: 100%;
      min-height: 110px;
      border: 1px dashed #9fb2c1;
      border-radius: 8px;
      padding: 18px;
      background: #f8fafb;
      color: var(--ink);
    }
    .hint {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
      margin: 0;
    }
    .actions {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      align-items: center;
    }
    button, .button {
      border: 0;
      border-radius: 8px;
      min-height: 42px;
      padding: 10px 15px;
      background: var(--teal);
      color: #fff;
      font: inherit;
      font-weight: 900;
      cursor: pointer;
    }
    .button {
      display: inline-flex;
      align-items: center;
      background: #263645;
    }
    .status {
      margin-top: 18px;
      border-top: 1px solid var(--line);
      padding-top: 16px;
      display: grid;
      gap: 8px;
      color: var(--muted);
      font-size: 13px;
    }
    .status strong { color: var(--ink); }
    @media (max-width: 760px) {
      header { flex-direction: column; }
      .nav { justify-content: flex-start; }
      .upload-grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <header>
    <div>
      <h1>Editar dados</h1>
      <p class="subtitle">Envie as planilhas novas para recriar o dashboard.</p>
    </div>
    <nav class="nav">
      <a class="top-link" href="/home">Home</a>
      <a class="top-link" href="/dashboard">Dashboard</a>
      <a class="top-link" href="/logout">Sair</a>
    </nav>
  </header>
  <main>
    <section class="panel">
      {message}
      <form method="post" action="/editar" enctype="multipart/form-data">
        <div class="upload-grid">
          <label>Planilha de ordens
            <input type="file" name="orders_file" accept=".xlsx">
          </label>
          <label>Planilha geral CT LOG
            <input type="file" name="detail_file" accept=".xlsx">
          </label>
        </div>
        <p class="hint">Voce pode enviar uma ou as duas planilhas. A base de ordens usa a coluna <strong>Ident.Veiculo</strong>; a base geral usa <strong>Placa 1 Veiculo</strong>.</p>
        <div class="actions">
          <button type="submit">Atualizar dashboard</button>
          <a class="button" href="/dashboard">Ver dashboard</a>
        </div>
      </form>
      <div class="status">
        <div><strong>Ordens atual:</strong> {orders_name}</div>
        <div><strong>Geral atual:</strong> {detail_name}</div>
      </div>
    </section>
  </main>
</body>
</html>
"""


def current_file_label(path: Path, fallback: str) -> str:
    if path.exists():
        return f"{path.name} atualizado"
    return fallback


def parse_multipart(content_type: str, body: bytes) -> dict[str, tuple[str, bytes]]:
    marker = "boundary="
    if marker not in content_type:
        return {}
    boundary = content_type.split(marker, 1)[1].strip().strip('"')
    boundary_bytes = ("--" + boundary).encode()
    files: dict[str, tuple[str, bytes]] = {}

    for part in body.split(boundary_bytes):
        part = part.strip(b"\r\n")
        if not part or part == b"--" or b"\r\n\r\n" not in part:
            continue
        raw_headers, content = part.split(b"\r\n\r\n", 1)
        content = content.rstrip(b"\r\n")
        headers = raw_headers.decode("utf-8", errors="ignore").split("\r\n")
        disposition = next(
            (line for line in headers if line.lower().startswith("content-disposition:")),
            "",
        )
        if "filename=" not in disposition or "name=" not in disposition:
            continue
        field = disposition.split("name=", 1)[1].split(";", 1)[0].strip().strip('"')
        filename = disposition.split("filename=", 1)[1].split(";", 1)[0].strip().strip('"')
        if filename and content:
            files[field] = (Path(filename).name, content)
    return files


def rebuild_dashboard() -> None:
    build_dashboard.main()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt: str, *args: object) -> None:
        print(f"{self.address_string()} - {fmt % args}")

    def send_bytes(self, content: bytes, content_type: str, status: int = 200) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def redirect(self, location: str) -> None:
        self.send_response(HTTPStatus.SEE_OTHER)
        self.send_header("Location", location)
        self.end_headers()

    def body_params(self) -> dict[str, list[str]]:
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length).decode("utf-8", errors="ignore")
        return parse_qs(body)

    def session_value(self) -> str:
        cookie = SimpleCookie(self.headers.get("Cookie", ""))
        if SESSION_COOKIE not in cookie:
            return ""
        return cookie[SESSION_COOKIE].value

    def is_logged_in(self) -> bool:
        value = self.session_value()
        return bool(value and valid_session(value))

    def require_login(self) -> bool:
        if self.is_logged_in():
            return True
        self.redirect("/login")
        return False

    def set_session_cookie(self, user: str) -> None:
        self.send_header(
            "Set-Cookie",
            f"{SESSION_COOKIE}={signed_session(user)}; Path=/; HttpOnly; SameSite=Lax",
        )

    def clear_session_cookie(self) -> None:
        self.send_header(
            "Set-Cookie",
            f"{SESSION_COOKIE}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax",
        )

    def send_login(self, message: str = "") -> None:
        page = LOGIN_HTML.replace("{message}", message)
        self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")

    def send_dashboard(self) -> None:
        if not INDEX_PATH.exists():
            rebuild_dashboard()
        page = INDEX_PATH.read_text(encoding="utf-8")
        nav = """
  <nav class="nav">
    <a class="top-link" href="/home">Home</a>
    <a class="top-link" href="/editar">Editar dados</a>
    <a class="top-link" href="/logout">Sair</a>
  </nav>"""
        page = page.replace('<a class="top-link" href="/editar">Atualizar dados</a>', nav)
        self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")

    def send_edit(self) -> None:
        params = parse_qs(urlparse(self.path).query)
        message = ""
        if "ok" in params:
            message = '<div class="message">Dashboard atualizado com sucesso.</div>'
        if "erro" in params:
            message = '<div class="message error">' + html.escape(params["erro"][0]) + "</div>"
        page = (
            EDIT_HTML.replace("{message}", message)
            .replace(
                "{orders_name}",
                html.escape(current_file_label(ORDERS_UPLOAD, "planilha original do projeto")),
            )
            .replace(
                "{detail_name}",
                html.escape(current_file_label(DETAIL_UPLOAD, "planilha original do projeto")),
            )
        )
        self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.redirect("/home" if self.is_logged_in() else "/login")
            return
        if parsed.path == "/login":
            if self.is_logged_in():
                self.redirect("/home")
                return
            self.send_login()
            return
        if parsed.path == "/logout":
            self.send_response(HTTPStatus.SEE_OTHER)
            self.send_header("Location", "/login")
            self.clear_session_cookie()
            self.end_headers()
            return
        if parsed.path == "/home":
            if not self.require_login():
                return
            self.send_bytes(HOME_HTML.encode("utf-8"), "text/html; charset=utf-8")
            return
        if parsed.path == "/dashboard":
            if not self.require_login():
                return
            self.send_dashboard()
            return
        if parsed.path == "/editar":
            if not self.require_login():
                return
            self.send_edit()
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/login":
            params = self.body_params()
            user = params.get("user", [""])[0]
            password = params.get("password", [""])[0]
            if user == app_user() and password == app_password():
                self.send_response(HTTPStatus.SEE_OTHER)
                self.send_header("Location", "/home")
                self.set_session_cookie(user)
                self.end_headers()
                return
            self.send_login('<div class="error">Usuario ou senha invalidos.</div>')
            return

        if parsed.path != "/editar":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        if not self.require_login():
            return

        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0 or length > MAX_UPLOAD_BYTES:
            self.redirect("/editar?erro=Arquivo+muito+grande+ou+vazio")
            return

        body = self.rfile.read(length)
        files = parse_multipart(self.headers.get("Content-Type", ""), body)
        if not files:
            self.redirect("/editar?erro=Nenhuma+planilha+foi+enviada")
            return

        DATA_DIR.mkdir(exist_ok=True)
        saved = 0
        for field, (filename, content) in files.items():
            if not filename.lower().endswith(".xlsx"):
                continue
            if field == "orders_file":
                ORDERS_UPLOAD.write_bytes(content)
                saved += 1
            elif field == "detail_file":
                DETAIL_UPLOAD.write_bytes(content)
                saved += 1

        if not saved:
            self.redirect("/editar?erro=Envie+arquivos+.xlsx+validos")
            return

        try:
            rebuild_dashboard()
        except Exception as exc:
            self.redirect("/editar?erro=" + quote(str(exc)))
            return
        self.redirect("/editar?ok=1")


def main() -> None:
    DATA_DIR.mkdir(exist_ok=True)
    if not INDEX_PATH.exists():
        rebuild_dashboard()
    port = int(os.environ.get("PORT", "8000"))
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"Servidor rodando em http://0.0.0.0:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
