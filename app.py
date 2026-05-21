from __future__ import annotations

import html
import os
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import build_dashboard


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
INDEX_PATH = ROOT / "index.html"
ORDERS_UPLOAD = DATA_DIR / "ordens.xlsx"
DETAIL_UPLOAD = DATA_DIR / "geral_ct_log.xlsx"
MAX_UPLOAD_BYTES = 25 * 1024 * 1024


EDIT_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Atualizar dados - Dashboard CT LOG</title>
  <style>
    :root {
      --bg: #eef2f5;
      --top: #10232b;
      --panel: #ffffff;
      --ink: #16212d;
      --muted: #657282;
      --line: #d7e0e8;
      --teal: #0c7c83;
      --shadow: 0 18px 42px rgba(23, 32, 51, .10);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      background: linear-gradient(180deg, var(--top) 0, #173843 230px, var(--bg) 231px);
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
    a {
      color: inherit;
      text-decoration: none;
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
    .error {
      background: #fff1ed;
      color: #9a3412;
    }
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
      .upload-grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <header>
    <div>
      <h1>Atualizar Dados</h1>
      <p class="subtitle">Envie as planilhas novas para recriar o dashboard.</p>
    </div>
    <a class="top-link" href="/">Voltar ao dashboard</a>
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
        <p class="hint">Você pode enviar uma ou as duas planilhas. A base de ordens usa a coluna <strong>Ident.Veículo</strong>; a base geral usa <strong>Placa 1 Veículo</strong>.</p>
        <div class="actions">
          <button type="submit">Atualizar dashboard</button>
          <a class="button" href="/">Ver dashboard</a>
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

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            if not INDEX_PATH.exists():
                rebuild_dashboard()
            self.send_bytes(INDEX_PATH.read_bytes(), "text/html; charset=utf-8")
            return
        if parsed.path == "/editar":
            params = parse_qs(parsed.query)
            message = ""
            if "ok" in params:
                message = '<div class="message">Dashboard atualizado com sucesso.</div>'
            if "erro" in params:
                message = (
                    '<div class="message error">'
                    + html.escape(params["erro"][0])
                    + "</div>"
                )
            page = (
                EDIT_HTML.replace("{message}", message)
                .replace(
                    "{orders_name}",
                    html.escape(
                        current_file_label(ORDERS_UPLOAD, "planilha original do projeto")
                    ),
                )
                .replace(
                    "{detail_name}",
                    html.escape(
                        current_file_label(DETAIL_UPLOAD, "planilha original do projeto")
                    ),
                )
            )
            self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/editar":
            self.send_error(HTTPStatus.NOT_FOUND)
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
            self.redirect("/editar?erro=" + html.escape(str(exc)).replace(" ", "+"))
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
