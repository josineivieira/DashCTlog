from __future__ import annotations

import hashlib
import hmac
import html
import csv
import datetime as dt
import io
import json
import os
import re
import secrets
import unicodedata
from http import HTTPStatus
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

import build_dashboard


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
INDEX_PATH = ROOT / "index.html"
ORDERS_UPLOAD = DATA_DIR / "ordens.xlsx"
DETAIL_UPLOAD = DATA_DIR / "geral_ct_log.xlsx"
CT_CONTROL_UPLOAD = ROOT / "Controle de CT .xlsx"
CT_CONTROL_DATA = DATA_DIR / "controle_ct.json"
CONDUCTOR_DATA = DATA_DIR / "condutores.json"
MAX_UPLOAD_BYTES = 25 * 1024 * 1024
SESSION_COOKIE = "dashboard_log_session"
FAVICON_URL = "https://pages.greatpages.com.br/www.dislubequador.com.br/1777495651/imagens/mobile/3562683_1_177616861364933621_m.svg"

PERMISSIONS = [
    ("dashboard", "Dashboard", "Visualizar painel principal"),
    ("editar", "Editar dados", "Importar e editar base operacional"),
    ("capacidades", "Capacidades", "Cadastrar placas, capacidades e motoristas"),
    ("controle_ct", "Controle de CT", "Acompanhar fila, patio e saidas"),
    ("relatorio_diario", "Relatorio diario", "Visualizar relatorio operacional diario"),
    ("entrada_notas", "Entrada de notas", "Importar e acompanhar prazo das notas"),
    ("controle_medicao", "Controle Medicao", "Acompanhar prazo dos fechamentos de medicao"),
    ("exportar_ct", "Exportar CT", "Baixar planilha do Controle de CT"),
]
ALL_PERMISSION_KEYS = {key for key, _, _ in PERMISSIONS}


def app_user() -> str:
    return os.environ.get("DASH_USER", "admin").strip()


def app_password() -> str:
    return os.environ.get("DASH_PASSWORD", "admin").strip()


def app_secret() -> str:
    return os.environ.get("DASH_SECRET", "dashboard-log-secret")


def password_hash(password: str) -> str:
    salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 180000)
    return f"pbkdf2_sha256$180000${salt}${digest.hex()}"


def check_password(password: str, stored_hash: str) -> bool:
    parts = stored_hash.split("$")
    if len(parts) != 4 or parts[0] != "pbkdf2_sha256":
        return False
    _, rounds_text, salt, digest_hex = parts
    try:
        rounds = int(rounds_text)
    except ValueError:
        return False
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), rounds).hex()
    return hmac.compare_digest(digest, digest_hex)


def clean_username(value: str) -> str:
    return re.sub(r"[^a-zA-Z0-9_.-]+", "", value.strip())


def display_user_name(value: str) -> str:
    text = str(value or "").strip()
    text = re.sub(r"[_\.-]+", " ", text)
    text = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def load_user_records() -> list[dict[str, object]]:
    if build_dashboard.use_postgres():
        return postgres_user_records()
    return []


def save_user_records(records: list[dict[str, object]]) -> None:
    if build_dashboard.use_postgres():
        save_postgres_user_records(records)
        return
    raise RuntimeError("Cadastro de usuarios exige banco de dados Postgres configurado.")


def find_user_record(username: str) -> dict[str, object] | None:
    username = clean_username(username)
    for item in load_user_records():
        if item.get("username") == username:
            return item
    return None


def is_master_user(username: str) -> bool:
    return clean_username(username) == clean_username(app_user())


def authenticate_user(username: str, password: str) -> bool:
    raw_username = username.strip()
    username = clean_username(raw_username)
    master_matches = raw_username == app_user() or username == clean_username(app_user())
    if master_matches and password == app_password():
        return True
    record = find_user_record(username)
    if not record or not record.get("active"):
        return False
    stored_hash = str(record.get("password_hash", ""))
    return bool(stored_hash and check_password(password, stored_hash))


def user_permissions(username: str) -> set[str]:
    if is_master_user(username):
        return set(ALL_PERMISSION_KEYS)
    record = find_user_record(username)
    if not record or not record.get("active"):
        return set()
    permissions = record.get("permissions", [])
    if not isinstance(permissions, list):
        return set()
    return {str(key) for key in permissions if str(key) in ALL_PERMISSION_KEYS}


def signed_session(user: str) -> str:
    signature = hmac.new(app_secret().encode(), user.encode(), hashlib.sha256).hexdigest()
    return f"{quote(user, safe='')}:{signature}"


def valid_session(value: str) -> bool:
    cookie_user, separator, signature = value.partition(":")
    if not separator:
        return False
    user = unquote(cookie_user)
    expected = signed_session(user).split(":", 1)[1]
    return hmac.compare_digest(signature, expected)


LOGIN_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="icon" href="{favicon_url}" type="image/svg+xml">
  <title>Login - Dashboard</title>
  <style>
    :root {
      --bg: #34104f;
      --panel: #ffffff;
      --ink: #16212d;
      --muted: #657282;
      --line: #d7e0e8;
      --teal: #64248c;
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
      padding: 32px;
      overflow: hidden;
      background: var(--bg);
      color: var(--ink);
      font-family: Inter, Segoe UI, Roboto, Arial, sans-serif;
    }
    .login-shell {
      width: min(100%, 1040px);
      min-height: 590px;
      display: grid;
      grid-template-columns: 1.08fr .92fr;
      overflow: hidden;
      border: 1px solid rgba(255, 255, 255, .20);
      border-radius: 18px;
      background: rgba(255, 255, 255, .08);
      box-shadow: 0 34px 120px rgba(0, 0, 0, .36);
      backdrop-filter: blur(18px);
    }
    .welcome {
      position: relative;
      min-height: 590px;
      padding: 44px;
      color: #fff;
      background:
        radial-gradient(480px circle at 74% 64%, rgba(20, 153, 160, .26), transparent 62%),
        linear-gradient(145deg, rgba(255,255,255,.10), rgba(255,255,255,.02));
    }
    .welcome::after {
      content: "";
      position: absolute;
      inset: 28px;
      border: 1px solid rgba(255,255,255,.10);
      border-radius: 14px;
      pointer-events: none;
    }
    .brand-mark {
      width: 320px;
      min-height: 94px;
      display: flex;
      align-items: center;
      gap: 18px;
      color: #fff;
      font-weight: 900;
      text-transform: uppercase;
    }
    .brand-mark img {
      width: 86px;
      height: 86px;
      object-fit: contain;
      filter: drop-shadow(0 10px 18px rgba(0, 0, 0, .22));
    }
    .brand-mark span {
      line-height: 1.05;
      letter-spacing: .02em;
      font-size: 21px;
    }
    .welcome h1 {
      max-width: 430px;
      margin: 28px 0 12px;
      font-size: clamp(38px, 5vw, 64px);
      line-height: .96;
      letter-spacing: 0;
    }
    .welcome p {
      max-width: 390px;
      margin: 0;
      color: #c8d6dc;
      font-size: 16px;
      line-height: 1.65;
    }
    .stats {
      position: absolute;
      left: 44px;
      right: 44px;
      bottom: 40px;
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 10px;
      z-index: 2;
    }
    .stat {
      min-height: 76px;
      padding: 14px;
      border: 1px solid rgba(255,255,255,.14);
      border-radius: 12px;
      background: rgba(255,255,255,.08);
    }
    .stat span {
      display: block;
      color: #b9cbd1;
      font-size: 11px;
      font-weight: 900;
      text-transform: uppercase;
    }
    .stat strong {
      display: block;
      margin-top: 7px;
      color: #fff;
      font-size: 21px;
    }
    .login {
      position: relative;
      z-index: 3;
      align-self: center;
      width: min(100% - 52px, 430px);
      justify-self: center;
      background:
        radial-gradient(340px circle at 100% 0, rgba(12,124,131,.12), transparent 44%),
        linear-gradient(180deg, rgba(255, 255, 255, .98), rgba(246, 250, 251, .98));
      border: 1px solid rgba(255, 255, 255, .86);
      border-radius: 16px;
      box-shadow: 0 26px 80px rgba(0, 0, 0, .30), inset 0 1px 0 rgba(255, 255, 255, .9);
      padding: 34px;
      backdrop-filter: blur(10px);
    }
    .login h2 { margin: 0; font-size: 34px; letter-spacing: 0; color: var(--ink); }
    .login p { margin: 9px 0 26px; color: var(--muted); line-height: 1.55; }
    .eyebrow {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      margin-bottom: 16px;
      color: #64248c;
      font-size: 12px;
      font-weight: 900;
      text-transform: uppercase;
    }
    .eyebrow::before {
      content: "";
      width: 9px;
      height: 9px;
      border-radius: 50%;
      background: #2b84cb;
      box-shadow: 0 0 0 6px rgba(43, 132, 203, .12);
    }
    form { display: grid; gap: 15px; }
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
      min-height: 48px;
      border: 1px solid #ccd8e2;
      border-radius: 12px;
      padding: 11px 13px;
      color: var(--ink);
      font: inherit;
      font-weight: 700;
      background: #f7fafc;
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, .9);
    }
    input:focus {
      outline: 3px solid rgba(100, 36, 140, .18);
      border-color: var(--teal);
    }
    .password-field {
      position: relative;
    }
    .password-field input {
      padding-right: 70px;
    }
    .password-toggle {
      position: absolute;
      right: 8px;
      top: 50%;
      min-height: 34px;
      padding: 6px 10px;
      border: 0;
      border-radius: 8px;
      background: transparent;
      color: #64248c;
      cursor: pointer;
      font: inherit;
      font-size: 12px;
      font-weight: 900;
      box-shadow: none;
      transform: translateY(-50%);
      transition: background .16s ease, color .16s ease;
    }
    .password-toggle:hover {
      background: rgba(100, 36, 140, .10);
      color: #34104f;
      box-shadow: none;
      transform: translateY(-50%);
    }
    button {
      min-height: 50px;
      border: 0;
      border-radius: 12px;
      background: linear-gradient(135deg, #64248c, #2b84cb);
      color: #fff;
      cursor: pointer;
      font: inherit;
      font-weight: 900;
      box-shadow: 0 16px 32px rgba(100, 36, 140, .28);
      transition: transform .16s ease, box-shadow .16s ease;
    }
    button:hover {
      transform: translateY(-1px);
      box-shadow: 0 20px 40px rgba(100, 36, 140, .34);
    }
    button.is-loading,
    button:disabled {
      cursor: wait;
      opacity: .78;
      transform: none;
      box-shadow: none;
    }
    button.is-loading::after {
      content: "";
      width: 14px;
      height: 14px;
      margin-left: 9px;
      border: 2px solid rgba(255,255,255,.48);
      border-top-color: #fff;
      border-radius: 50%;
      animation: spin .75s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    .login-foot {
      margin-top: 18px;
      display: flex;
      justify-content: space-between;
      gap: 12px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
    }
    .login-foot span {
      display: inline-flex;
      align-items: center;
      gap: 6px;
    }
    .login-foot span::before {
      content: "";
      width: 7px;
      height: 7px;
      border-radius: 50%;
      background: #e2263c;
    }
    .error {
      margin-bottom: 14px;
      padding: 11px 12px;
      border-radius: 8px;
      background: var(--error-bg);
      color: var(--error);
      font-weight: 800;
    }
    @media (max-width: 860px) {
      body { place-items: start center; padding: 20px; overflow: auto; }
      .login-shell { min-height: auto; grid-template-columns: 1fr; }
      .welcome { min-height: 220px; padding: 28px; }
      .welcome h1 { font-size: 38px; }
      .welcome p { max-width: 100%; }
      .stats { position: static; margin-top: 22px; grid-template-columns: 1fr; }
      .brand-mark { width: 100%; min-height: 72px; }
      .brand-mark img { width: 64px; height: 64px; }
      .brand-mark span { font-size: 18px; }
      .login { width: calc(100% - 28px); margin: 14px 0 24px; padding: 26px; }
    }
  </style>
</head>
<body>
  <main class="login-shell">
    <section class="welcome">
      <div class="brand-mark"><img src="{favicon_url}" alt=""><span>Dislub<br>Equador</span></div>
      <h1>Controle logistico em tempo real.</h1>
      <p>Entre para acompanhar viagens, placas, produtos carregados e atualizar as planilhas do dashboard.</p>
      <div class="stats">
        <div class="stat"><span>Area</span><strong>CT LOG</strong></div>
        <div class="stat"><span>Acesso</span><strong>Seguro</strong></div>
        <div class="stat"><span>Dados</span><strong>Online</strong></div>
      </div>
    </section>
    <section class="login">
      <div class="eyebrow">Acesso ao sistema</div>
      <h2>Dashboard</h2>
      <p>Informe suas credenciais para entrar no painel.</p>
      {message}
      <form method="post" action="/login">
        <label>Usuario
          <input name="user" autocomplete="username" required>
        </label>
        <label>Senha
          <span class="password-field">
            <input id="passwordInput" name="password" type="password" autocomplete="current-password" required>
            <button id="passwordToggle" class="password-toggle" type="button" aria-label="Mostrar senha">Ver</button>
          </span>
        </label>
        <button type="submit">Entrar no dashboard</button>
      </form>
      <div class="login-foot">
        <span>Sessao protegida</span>
        <span>Dados privados</span>
      </div>
    </section>
  </main>
  <script>
    function setLoading(button, text) {
      if (!button) return;
      button.textContent = text;
      button.classList.add("is-loading");
      button.disabled = true;
    }
    const passwordInput = document.querySelector("#passwordInput");
    const passwordToggle = document.querySelector("#passwordToggle");
    passwordToggle.addEventListener("click", () => {
      const showing = passwordInput.type === "text";
      passwordInput.type = showing ? "password" : "text";
      passwordToggle.textContent = showing ? "Ver" : "Ocultar";
      passwordToggle.setAttribute("aria-label", showing ? "Mostrar senha" : "Ocultar senha");
      passwordInput.focus();
    });
    document.querySelector("form").addEventListener("submit", (event) => {
      setLoading(event.currentTarget.querySelector('button[type="submit"]'), "Entrando...");
    });
  </script>
</body>
</html>
"""


HOME_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="icon" href="{favicon_url}" type="image/svg+xml">
  <title>Home - Dashboard</title>
  <style>
    :root {
      --bg: #eef2f6;
      --side: #2f0d49;
      --side-2: #3d145f;
      --panel: #ffffff;
      --ink: #16212d;
      --muted: #657282;
      --line: #d7e0e8;
      --purple: #64248c;
      --blue: #2b84cb;
      --red: #e2263c;
      --green: #00856f;
      --navy: #1b255f;
      --shadow: 0 16px 36px rgba(23, 32, 51, .10);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      background: var(--bg);
      color: var(--ink);
      font-family: Inter, Segoe UI, Roboto, Arial, sans-serif;
    }
    a { color: inherit; text-decoration: none; }
    .app-shell {
      display: flex;
      min-height: 100vh;
    }
    .sidebar {
      width: 272px;
      flex: 0 0 272px;
      display: flex;
      flex-direction: column;
      gap: 18px;
      padding: 22px 16px;
      background:
        linear-gradient(180deg, var(--side), var(--side-2));
      color: #fff;
    }
    .brand {
      display: flex;
      align-items: center;
      gap: 13px;
      min-height: 74px;
      padding: 0 8px 18px;
      border-bottom: 1px solid rgba(255,255,255,.14);
    }
    .brand img {
      width: 88px;
      object-fit: contain;
    }
    .brand strong {
      display: block;
      font-size: 18px;
      line-height: 1.1;
    }
    .brand span {
      display: block;
      margin-top: 4px;
      color: #c8d6dc;
      font-size: 12px;
      font-weight: 800;
      text-transform: uppercase;
    }
    .nav-title {
      padding: 0 8px;
      color: #b9c8d2;
      font-size: 11px;
      font-weight: 900;
      text-transform: uppercase;
    }
    .side-nav {
      display: grid;
      gap: 6px;
    }
    .side-link {
      display: flex;
      align-items: center;
      gap: 11px;
      min-height: 42px;
      padding: 10px 11px;
      border-radius: 8px;
      color: #e8eef3;
      font-size: 13px;
      font-weight: 900;
      background: rgba(255,255,255,.04);
    }
    .side-link:hover, .side-link.active {
      background: rgba(255,255,255,.13);
      color: #fff;
    }
    .side-link svg {
      width: 18px;
      height: 18px;
      flex: 0 0 auto;
      stroke-width: 2.2;
    }
    .sidebar-footer {
      margin-top: auto;
      padding: 14px 8px 0;
      border-top: 1px solid rgba(255,255,255,.14);
      color: #c8d6dc;
      font-size: 12px;
      line-height: 1.45;
    }
    .content {
      flex: 1;
      min-width: 0;
      display: flex;
      flex-direction: column;
    }
    .topbar {
      min-height: 74px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
      padding: 18px clamp(18px, 3vw, 34px);
      border-bottom: 1px solid var(--line);
      background: #fff;
    }
    h1 { margin: 0; font-size: clamp(24px, 3vw, 34px); letter-spacing: 0; }
    .subtitle { margin: 4px 0 0; color: var(--muted); font-size: 14px; }
    .logout {
      min-height: 38px;
      display: inline-flex;
      align-items: center;
      padding: 8px 12px;
      border: 1px solid var(--line);
      border-radius: 8px;
      color: var(--ink);
      font-size: 13px;
      font-weight: 900;
      background: #f8fafc;
    }
    main {
      padding: 22px clamp(18px, 3vw, 34px) 34px;
    }
    .summary {
      display: grid;
      grid-template-columns: repeat(4, minmax(130px, 1fr));
      gap: 12px;
      margin-bottom: 18px;
    }
    .metric {
      min-height: 82px;
      padding: 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      box-shadow: var(--shadow);
    }
    .metric span {
      color: var(--muted);
      font-size: 11px;
      font-weight: 900;
      text-transform: uppercase;
    }
    .metric strong {
      display: block;
      margin-top: 8px;
      font-size: 24px;
    }
    .menu {
      display: grid;
      grid-template-columns: repeat(3, minmax(230px, 1fr));
      gap: 14px;
    }
    .card {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 10px 14px;
      min-height: 154px;
      padding: 18px;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }
    .card-icon {
      grid-column: 2;
      grid-row: 1 / span 2;
      width: 42px;
      height: 42px;
      display: grid;
      place-items: center;
      border-radius: 8px;
      color: #fff;
      background: var(--purple);
      box-shadow: inset 0 0 0 1px rgba(255,255,255,.32);
    }
    .card-icon svg { width: 22px; height: 22px; stroke-width: 2.2; }
    .card:nth-child(2) .card-icon { background: var(--blue); }
    .card:nth-child(3) .card-icon { background: var(--red); }
    .card:nth-child(4) .card-icon { background: var(--navy); }
    .card:nth-child(5) .card-icon { background: var(--green); }
    .card span {
      color: var(--muted);
      font-size: 11px;
      font-weight: 900;
      text-transform: uppercase;
    }
    .card strong {
      grid-column: 1;
      color: var(--ink);
      font-size: 24px;
      line-height: 1.1;
    }
    .card p {
      grid-column: 1 / -1;
      margin: 0;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.45;
    }
    .button {
      grid-column: 1 / -1;
      align-self: end;
      justify-self: start;
      min-height: 38px;
      padding: 9px 12px;
      border-radius: 8px;
      background: var(--purple);
      color: #fff;
      font-size: 13px;
      font-weight: 900;
    }
    .card:nth-child(2) .button { background: var(--blue); }
    .card:nth-child(3) .button { background: var(--red); }
    .card:nth-child(4) .button { background: var(--navy); }
    .card:nth-child(5) .button { background: var(--green); }
    @media (max-width: 1080px) {
      .menu { grid-template-columns: repeat(2, minmax(220px, 1fr)); }
      .summary { grid-template-columns: repeat(2, 1fr); }
    }
    @media (max-width: 760px) {
      .app-shell { display: block; }
      .sidebar { width: auto; min-height: 0; }
      .side-nav { grid-template-columns: repeat(2, 1fr); }
      .topbar { align-items: flex-start; flex-direction: column; }
      .menu, .summary { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <div class="app-shell">
    <aside class="sidebar">
      <div class="brand">
        <img src="{favicon_url}" alt="">
        <div><strong>Dislub Equador</strong><span>Operacao</span></div>
      </div>
      <div class="nav-title">Modulos</div>
      <nav class="side-nav">
        <a class="side-link active" href="/dashboard"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M4 13h6V4H4z"></path><path d="M14 20h6V4h-6z"></path><path d="M4 20h6v-3H4z"></path></svg>Dashboard</a>
        <a class="side-link" href="/editar"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M4 5h16"></path><path d="M4 12h10"></path><path d="M4 19h7"></path><path d="m15 18 5-5 2 2-5 5-3 1z"></path></svg>Editar dados</a>
        <a class="side-link" href="/capacidades"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M3 16V8h11v8"></path><path d="M14 11h4l3 3v2h-7"></path><circle cx="7" cy="18" r="2"></circle><circle cx="17" cy="18" r="2"></circle></svg>Capacidades</a>
        <a class="side-link" href="/relatorio-diario"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M7 3h8l4 4v14H7z"></path><path d="M15 3v5h5"></path><path d="M10 13h6"></path><path d="M10 17h4"></path></svg>Relatorio diario</a>
        <a class="side-link" href="/relatorio-entrada-notas"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M6 3h9l4 4v14H6z"></path><path d="M15 3v5h5"></path><path d="M9 12h7"></path><path d="M9 16h4"></path></svg>Entrada de notas</a>
        <a class="side-link" href="/controle-medicao"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M4 4h16v16H4z"></path><path d="M8 2v4"></path><path d="M16 2v4"></path><path d="M4 10h16"></path><path d="m9 15 2 2 4-4"></path></svg>Controle Medicao</a>
        <a class="side-link" href="/controle-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M4 7h16"></path><path d="M4 12h16"></path><path d="M4 17h10"></path><path d="M17 15l2 2 4-4"></path></svg>Controle de CT</a>
        __USER_LINK__
      </nav>
      <div class="sidebar-footer">Dislub Equador<br>Ambiente de acompanhamento logistico.</div>
    </aside>
    <div class="content">
      <header class="topbar">
        <div>
          <h1>__HOME_GREETING__, __HOME_USER__</h1>
          <p class="subtitle">Bom te ver por aqui. Escolha um modulo para acompanhar a operacao.</p>
        </div>
        <a class="logout" href="/logout">Sair</a>
      </header>
      <main>
        <section class="summary">
          <div class="metric"><span>Modulos ativos</span><strong>__MODULE_COUNT__</strong></div>
          <div class="metric"><span>Base operacional</span><strong>CT</strong></div>
          <div class="metric"><span>Terminais</span><strong>2</strong></div>
          <div class="metric"><span>Status</span><strong>Online</strong></div>
        </section>
        <section class="menu">
          <a class="card" href="/dashboard">
            <span>Visualizacao</span>
            <div class="card-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M4 13h6V4H4z"></path><path d="M14 20h6V4h-6z"></path><path d="M4 20h6v-3H4z"></path></svg></div>
            <strong>Dashboard</strong>
            <p>Acompanhe viagens, placas, produtos e terminais.</p>
            <div class="button">Abrir dashboard</div>
          </a>
          <a class="card" href="/editar">
            <span>Dados</span>
            <div class="card-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M4 5h16"></path><path d="M4 12h10"></path><path d="M4 19h7"></path><path d="m15 18 5-5 2 2-5 5-3 1z"></path></svg></div>
            <strong>Editar dados</strong>
            <p>Envie novas planilhas e atualize os indicadores.</p>
            <div class="button">Atualizar planilhas</div>
          </a>
          <a class="card" href="/capacidades">
            <span>Cadastro</span>
            <div class="card-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M3 16V8h11v8"></path><path d="M14 11h4l3 3v2h-7"></path><circle cx="7" cy="18" r="2"></circle><circle cx="17" cy="18" r="2"></circle></svg></div>
            <strong>Capacidades</strong>
            <p>Cadastre carretas, caminhoes, tanques e capacidades por placa.</p>
            <div class="button">Editar capacidades</div>
          </a>
          <a class="card" href="/relatorio-diario">
            <span>Relatorio</span>
            <div class="card-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M7 3h8l4 4v14H7z"></path><path d="M15 3v5h5"></path><path d="M10 13h6"></path><path d="M10 17h4"></path></svg></div>
            <strong>Diario</strong>
            <p>Resumo das viagens por dia, placa, terminal e volume carregado.</p>
            <div class="button">Abrir relatorio</div>
          </a>
          <a class="card" href="/controle-ct">
            <span>Operacao</span>
            <div class="card-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M4 7h16"></path><path d="M4 12h16"></path><path d="M4 17h10"></path><path d="M17 15l2 2 4-4"></path></svg></div>
            <strong>Controle de CT</strong>
            <p>Acompanhe chegada, patio, fila, saida e notas fiscais.</p>
            <div class="button">Abrir controle</div>
          </a>
          <a class="card" href="/relatorio-entrada-notas">
            <span>Notas</span>
            <div class="card-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M6 3h9l4 4v14H6z"></path><path d="M15 3v5h5"></path><path d="M9 12h7"></path><path d="M9 16h4"></path></svg></div>
            <strong>Entrada de notas</strong>
            <p>Confira se as notas fiscais foram dadas entrada dentro do prazo de 48 horas.</p>
            <div class="button">Abrir relatorio</div>
          </a>
          <a class="card" href="/controle-medicao">
            <span>Medicao</span>
            <div class="card-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M4 4h16v16H4z"></path><path d="M8 2v4"></path><path d="M16 2v4"></path><path d="M4 10h16"></path><path d="m9 15 2 2 4-4"></path></svg></div>
            <strong>Controle Medicao</strong>
            <p>Acompanhe fechamentos de medicao no prazo de 2 dias sem contar domingo.</p>
            <div class="button">Abrir controle</div>
          </a>
          __USER_CARD__
        </section>
      </main>
    </div>
  </div>
</body>
</html>
"""


USERS_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="icon" href="{favicon_url}" type="image/svg+xml">
  <title>Cadastro de usuarios - Dashboard</title>
  <style>
    :root {
      --purple: #64248c;
      --purple-dark: #34104f;
      --blue: #2b84cb;
      --green: #20a86b;
      --red: #e2263c;
      --bg: #eef3f6;
      --panel: #fff;
      --ink: #16212d;
      --muted: #657282;
      --line: #d3dee8;
      --shadow: 0 18px 42px rgba(23,32,51,.12);
    }
    * { box-sizing: border-box; }
    body { margin:0; min-height:100vh; background:var(--bg); color:var(--ink); font-family:Inter, Segoe UI, Roboto, Arial, sans-serif; }
    header { position:relative; overflow:hidden; padding:24px clamp(16px,4vw,42px) 30px; background:radial-gradient(720px circle at 76% 35%, rgba(43,132,203,.34), transparent 62%), linear-gradient(135deg,#34104f,#4c176d 58%,#1b255f); color:#fff; }
    header::after { content:""; position:absolute; right:clamp(20px,6vw,76px); bottom:-88px; width:min(46vw,520px); aspect-ratio:1.8; background:url("{favicon_url}") center / contain no-repeat; opacity:.18; pointer-events:none; }
    .topbar { position:relative; z-index:2; display:flex; justify-content:space-between; gap:18px; align-items:flex-start; }
    .brand-title { display:flex; align-items:center; gap:16px; }
    .brand-title img { width:98px; height:auto; object-fit:contain; filter:drop-shadow(0 10px 18px rgba(0,0,0,.24)); }
    h1 { margin:0; font-size:clamp(28px,4vw,48px); line-height:1; letter-spacing:0; }
    .subtitle { margin:9px 0 0; color:#d7e4ea; font-size:15px; }
    .nav { display:flex; flex-wrap:wrap; justify-content:flex-end; gap:9px; }
    a { color:inherit; text-decoration:none; }
    .top-link, button, .button { min-height:36px; display:inline-flex; align-items:center; justify-content:center; padding:8px 11px; border:1px solid rgba(255,255,255,.32); border-radius:8px; background:rgba(255,255,255,.10); color:#fff; font:inherit; font-size:13px; font-weight:900; cursor:pointer; }
    main { padding:22px clamp(16px,4vw,42px) 42px; }
    .toast-zone { position:fixed; right:clamp(14px,2vw,24px); bottom:clamp(14px,2vw,24px); z-index:30; display:grid; gap:10px; justify-items:end; pointer-events:none; }
    .message { width:min(440px,calc(100vw - 28px)); padding:13px 15px; border-radius:8px; border:1px solid #bbf7d0; background:#f0fdf4; color:#14532d; font-weight:900; box-shadow:0 18px 42px rgba(23,32,51,.18); pointer-events:auto; transition:opacity .22s ease, transform .22s ease; }
    .message.error { border-color:#fecaca; background:#fff1f2; color:#991b1b; }
    .message.is-hidden { opacity:0; transform:translateY(12px); pointer-events:none; }
    .layout { display:grid; grid-template-columns:minmax(320px, 420px) 1fr; gap:16px; align-items:start; }
    .panel { background:var(--panel); border:1px solid var(--line); border-radius:10px; box-shadow:var(--shadow); overflow:hidden; }
    .panel-head { padding:16px 18px; border-bottom:1px solid var(--line); display:flex; justify-content:space-between; gap:12px; align-items:center; }
    .panel-head h2 { margin:0; font-size:18px; }
    .panel-body { padding:18px; }
    .form-grid { display:grid; gap:13px; }
    label { display:grid; gap:6px; color:#506071; font-size:13px; font-weight:800; }
    input[type="text"], input[type="password"] { width:100%; border:1px solid var(--line); border-radius:8px; padding:10px 11px; background:#f8fafb; color:var(--ink); font:inherit; }
    .check-line { display:flex; gap:9px; align-items:center; color:var(--ink); font-weight:850; }
    .permissions { display:grid; gap:8px; margin-top:4px; }
    .permission { display:grid; grid-template-columns:22px 1fr; gap:9px; align-items:flex-start; padding:10px; border:1px solid var(--line); border-radius:8px; background:#f8fafb; }
    .permission strong { display:block; font-size:13px; }
    .permission span { display:block; margin-top:3px; color:var(--muted); font-size:12px; line-height:1.35; }
    .actions { display:flex; flex-wrap:wrap; gap:9px; margin-top:4px; }
    .primary { background:var(--purple); border-color:var(--purple); }
    .secondary { color:var(--ink); background:#fff; border-color:var(--line); }
    .danger { background:var(--red); border-color:var(--red); }
    .user-list { display:grid; gap:12px; }
    .user-card { display:grid; grid-template-columns:minmax(180px,.9fr) minmax(220px,1.35fr) auto; gap:16px; align-items:center; padding:14px; border:1px solid var(--line); border-radius:8px; background:linear-gradient(180deg,#fff,#fbfdff); }
    .user-main { min-width:0; }
    .user-pill { display:inline-flex; align-items:center; gap:7px; padding:6px 9px; border-radius:999px; background:#eef2ff; color:#3730a3; font-weight:900; }
    .user-pill::before { content:""; width:8px; height:8px; border-radius:50%; background:#64248c; }
    .user-name { margin-top:6px; color:var(--muted); font-size:12px; line-height:1.4; }
    .status { display:inline-flex; padding:5px 8px; border-radius:999px; font-weight:900; font-size:12px; background:#dcfce7; color:#166534; }
    .status.off { background:#fee2e2; color:#991b1b; }
    .user-meta { display:grid; gap:9px; }
    .perm-list { display:flex; flex-wrap:wrap; gap:6px; }
    .perm-tag { display:inline-flex; padding:5px 8px; border-radius:999px; background:#f1f5f9; color:#334155; font-size:12px; font-weight:850; }
    .row-actions { display:flex; flex-wrap:wrap; justify-content:flex-end; gap:7px; }
    .row-actions button { min-height:30px; padding:6px 8px; font-size:12px; }
    .empty-users { padding:22px; border:1px dashed var(--line); border-radius:8px; color:var(--muted); text-align:center; font-weight:850; }
    .hint { color:var(--muted); font-size:12px; line-height:1.45; }
    @media (max-width:980px) {
      .topbar, header { display:block; }
      .nav { justify-content:flex-start; margin-top:16px; }
      .layout { grid-template-columns:1fr; }
      .panel { overflow:auto; }
      .user-card { grid-template-columns:1fr; align-items:start; }
      .row-actions { justify-content:flex-start; }
    }
  </style>
</head>
<body>
  <div class="toast-zone" aria-live="polite">{message}</div>
  <header>
    <div class="topbar">
      <div class="brand-title"><img src="{favicon_url}" alt=""><div><h1>Cadastro de usuarios</h1><p class="subtitle">Controle quem acessa cada tela e funcao do sistema.</p></div></div>
      <nav class="nav">
        <a class="top-link" href="/home">Home</a>
        <a class="top-link" href="/dashboard">Dashboard</a>
        <a class="top-link" href="/editar">Editar dados</a>
        <a class="top-link" href="/controle-ct">Controle de CT</a>
        <a class="top-link" href="/relatorio-diario">Relatorio diario</a>
        <a class="top-link" href="/relatorio-entrada-notas">Entrada de notas</a>
        <a class="top-link" href="/capacidades">Capacidades</a>
        <a class="top-link" href="/logout">Sair</a>
      </nav>
    </div>
  </header>
  <main>
    <section class="layout">
      <form class="panel" method="post" action="/usuarios">
        <div class="panel-head"><h2>{form_title}</h2><span class="hint">Somente master</span></div>
        <div class="panel-body form-grid">
          <input type="hidden" name="original_username" value="{original_username}">
          <label>Usuario
            <input name="username" type="text" value="{username}" autocomplete="off" required>
          </label>
          <label>Nome
            <input name="name" type="text" value="{name}" autocomplete="off">
          </label>
          <label>Senha
            <input name="password" type="password" autocomplete="new-password" {password_required}>
            <span class="hint">{password_hint}</span>
          </label>
          <label class="check-line"><input type="checkbox" name="active" value="1" {active_checked}> Usuario ativo</label>
          <div>
            <label>Permissoes</label>
            <div class="permissions">{permission_checks}</div>
          </div>
          <div class="actions">
            <button class="primary" type="submit">Salvar usuario</button>
            <a class="button secondary" href="/usuarios">Limpar</a>
          </div>
        </div>
      </form>
      <section class="panel">
        <div class="panel-head"><h2>Usuarios cadastrados</h2><span class="hint">{user_count} usuarios</span></div>
        <div class="panel-body">
          <div class="user-list">{user_rows}</div>
        </div>
      </section>
    </section>
  </main>
  <script>
    document.querySelectorAll(".message.auto-dismiss").forEach((item) => {
      window.setTimeout(() => item.classList.add("is-hidden"), 4200);
      window.setTimeout(() => item.remove(), 4600);
    });
  </script>
</body>
</html>
"""


EDIT_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="icon" href="{favicon_url}" type="image/svg+xml">
  <title>Base editavel - Dashboard</title>
  <style>
    :root {
      --bg: #34104f;
      --panel: #ffffff;
      --ink: #16212d;
      --muted: #657282;
      --line: #d7e0e8;
      --teal: #64248c;
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
    .brand-title {
      display: flex;
      align-items: center;
      gap: 14px;
    }
    .brand-title img {
      width: 42px;
      height: 42px;
      object-fit: contain;
      filter: drop-shadow(0 8px 14px rgba(0, 0, 0, .22));
    }
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
      width: 100%;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      overflow: hidden;
    }
    .message {
      margin: 16px;
      padding: 12px 14px;
      border-radius: 8px;
      background: #e7f4f2;
      color: #0d6268;
      font-weight: 800;
    }
    .error { background: #fff1ed; color: #9a3412; }
    .toolbar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 14px;
      border-bottom: 1px solid var(--line);
      background: #f8fafb;
    }
    .actions {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      align-items: center;
    }
    .meta { color: var(--muted); font-size: 13px; font-weight: 800; }
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
      background: #1b255f;
    }
    .import-bar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 14px;
      border-bottom: 1px solid var(--line);
      background: #fff;
    }
    .import-bar form {
      display: flex;
      align-items: center;
      gap: 10px;
      flex-wrap: wrap;
    }
    .import-bar input[type="file"] {
      max-width: 320px;
      padding: 8px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #f8fafb;
      color: var(--ink);
      font: inherit;
      font-size: 13px;
    }
    .sheet-wrap {
      max-height: calc(100vh - 245px);
      overflow: auto;
      background: #fff;
    }
    table {
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      font-size: 13px;
    }
    th, td {
      min-width: 138px;
      border-right: 1px solid var(--line);
      border-bottom: 1px solid var(--line);
      padding: 0;
      text-align: left;
      vertical-align: top;
    }
    th:first-child, td:first-child {
      min-width: 54px;
      width: 54px;
      text-align: center;
      color: var(--muted);
      background: #f3f6f8;
      font-weight: 900;
      position: sticky;
      left: 0;
      z-index: 2;
    }
    th {
      position: sticky;
      top: 0;
      z-index: 3;
      padding: 10px;
      background: #eef3f6;
      color: #506071;
      font-size: 12px;
      text-transform: uppercase;
      white-space: nowrap;
    }
    th:first-child { z-index: 4; }
    td[contenteditable="true"] {
      padding: 9px 10px;
      outline: 0;
      background: #fff;
      color: var(--ink);
      min-height: 36px;
    }
    td[contenteditable="true"]:focus {
      background: #e7f4f2;
      box-shadow: inset 0 0 0 2px var(--teal);
    }
    tr:nth-child(even) td[contenteditable="true"] { background: #fbfcfd; }
    tr:nth-child(even) td[contenteditable="true"]:focus { background: #e7f4f2; }
    td.cell-selected,
    tr:nth-child(even) td.cell-selected {
      background: #dcecff;
      box-shadow: inset 0 0 0 1px #6aa3e8;
    }
    td.cell-anchor,
    tr:nth-child(even) td.cell-anchor {
      background: #c8e0ff;
      box-shadow: inset 0 0 0 2px #1f6feb;
    }
    td.cell-fill,
    tr:nth-child(even) td.cell-fill {
      background: #eaf3ff;
      box-shadow: inset 0 0 0 1px #9ac1f2;
    }
    body.selecting-cells {
      cursor: cell;
      user-select: none;
    }
    .hint {
      padding: 12px 14px;
      border-top: 1px solid var(--line);
      color: var(--muted);
      font-size: 13px;
      line-height: 1.45;
    }
    @media (max-width: 760px) {
      header { flex-direction: column; }
      .nav { justify-content: flex-start; }
      .toolbar { align-items: flex-start; flex-direction: column; }
    }
  </style>
</head>
<body>
  <header>
    <div>
      <div class="brand-title"><img src="{favicon_url}" alt=""><h1>Base editavel</h1></div>
      <p class="subtitle">Edite a base unica que alimenta o dashboard.</p>
    </div>
    <nav class="nav">
      <a class="top-link" href="/home">Home</a>
      <a class="top-link" href="/dashboard">Dashboard</a>
      <a class="top-link" href="/controle-ct">Controle de CT</a>
      <a class="top-link" href="/relatorio-diario">Relatorio diario</a>
      <a class="top-link" href="/relatorio-entrada-notas">Entrada de notas</a>
      <a class="top-link" href="/capacidades">Capacidades</a>
      <a class="top-link" href="/logout">Sair</a>
    </nav>
  </header>
  <main>
    <section class="panel">
      {message}
      <div class="import-bar">
        <div class="meta">Importe uma base em CSV ou XLSX. Colunas extras serao ignoradas automaticamente.</div>
        <form method="post" action="/importar" enctype="multipart/form-data">
          <a class="button" href="/template.csv">Baixar template</a>
          <input type="file" name="base_file" accept=".csv,.xlsx" required>
          <button type="submit">Importar base</button>
        </form>
      </div>
      <form id="sheetForm" method="post" action="/editar">
        <input type="hidden" name="rows_json" id="rowsJson">
        <div class="toolbar">
          <div class="meta"><span id="rowCount">__ROW_COUNT__</span> linhas na base <span id="draftStatus" class="draft-status"></span></div>
          <div class="actions">
            <button type="button" id="addRow">Adicionar linha</button>
            <button type="button" id="deleteRows" class="button">Excluir selecionadas</button>
            <button type="submit">Salvar e atualizar dashboard</button>
            <a class="button" href="/dashboard">Ver dashboard</a>
          </div>
        </div>
        <div class="sheet-wrap">
          <table id="sheet">
            <thead>__THEAD__</thead>
            <tbody>__TBODY__</tbody>
          </table>
        </div>
        <div class="hint">Colunas usadas pelo dashboard: data, placa, terminal, viagens, capacidade, nota fiscal, produto, cliente, municipio destino e quantidade. Terminal deve ser 10 para Equador ou 19 para Ipiranga.</div>
      </form>
    </section>
  </main>
  <script>
    const columns = [
      ["data", "Data"],
      ["placa", "Placa"],
      ["terminal", "Terminal"],
      ["viagens", "Viagens"],
      ["capacidade", "Capacidade"],
      ["motorista1", "Motorista 1"],
      ["notaFiscal", "Nota fiscal"],
      ["produto", "Produto"],
      ["cliente", "Cliente"],
      ["municipioDestino", "Municipio destino"],
      ["quantidade", "Quantidade"]
    ];
    const serverRows = __ROWS__;
    localStorage.removeItem("dashboard-edit-draft-v1");
    let rows = serverRows;
    const thead = document.querySelector("#sheet thead");
    const tbody = document.querySelector("#sheet tbody");
    const rowCount = document.querySelector("#rowCount");
    const draftStatus = document.querySelector("#draftStatus");
    let selectionStart = null;
    let selectionEnd = null;
    let isSelecting = false;

    function escapeHtml(value) {
      return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
    }

    function cleanRow(row = {}) {
      return Object.fromEntries(columns.map(([key]) => [key, row[key] ?? ""]));
    }

    function updateDraftStatus() {
      if (!draftStatus) return;
      draftStatus.textContent = "";
    }

    function render() {
      thead.innerHTML = `<tr><th><input type="checkbox" id="selectAllRows" aria-label="Selecionar todas as linhas"></th>${columns.map(([, label]) => `<th>${label}</th>`).join("")}</tr>`;
      tbody.innerHTML = rows.map((row, idx) => {
        const clean = cleanRow(row);
        return `<tr data-row="${idx}">
          <td><input type="checkbox" aria-label="Selecionar linha ${idx + 1}"><br>${idx + 1}</td>
          ${columns.map(([key], colIdx) => `<td contenteditable="true" data-key="${key}" data-col="${colIdx}">${escapeHtml(clean[key])}</td>`).join("")}
        </tr>`;
      }).join("");
      rowCount.textContent = rows.length.toLocaleString("pt-BR");
      updateDraftStatus();
    }

    function cellPosition(cell) {
      return {
        row: Number(cell.closest("tr").dataset.row),
        col: Number(cell.dataset.col)
      };
    }

    function clearSelection() {
      tbody.querySelectorAll(".cell-selected, .cell-anchor, .cell-fill").forEach((cell) => {
        cell.classList.remove("cell-selected", "cell-anchor", "cell-fill");
      });
    }

    function paintSelection() {
      clearSelection();
      if (!selectionStart || !selectionEnd) return;
      const rowMin = Math.min(selectionStart.row, selectionEnd.row);
      const rowMax = Math.max(selectionStart.row, selectionEnd.row);
      const colMin = Math.min(selectionStart.col, selectionEnd.col);
      const colMax = Math.max(selectionStart.col, selectionEnd.col);
      tbody.querySelectorAll("[data-key]").forEach((cell) => {
        const pos = cellPosition(cell);
        if (pos.row >= rowMin && pos.row <= rowMax && pos.col >= colMin && pos.col <= colMax) {
          cell.classList.add("cell-selected");
          if (pos.row === selectionStart.row && pos.col === selectionStart.col) {
            cell.classList.add("cell-anchor");
          }
        }
      });
    }

    function selectedCells() {
      return [...tbody.querySelectorAll(".cell-selected")];
    }

    function syncFromTable() {
      rows = [...tbody.querySelectorAll("tr")].map((tr) => {
        const row = {};
        tr.querySelectorAll("[data-key]").forEach((cell) => row[cell.dataset.key] = cell.textContent.trim());
        return cleanRow(row);
      });
    }

    document.querySelector("#addRow").addEventListener("click", () => {
      syncFromTable();
      rows.push(cleanRow({ terminal: "10", capacidade: "30000" }));
      render();
      const lastRow = tbody.querySelector(`tr[data-row="${rows.length - 1}"] [data-key="data"]`);
      lastRow?.focus();
    });
    document.querySelector("#deleteRows").addEventListener("click", () => {
      syncFromTable();
      const checkedRows = [...tbody.querySelectorAll("tr")]
        .filter((tr) => tr.querySelector("input")?.checked)
        .map((tr) => Number(tr.dataset.row));
      const cellRows = selectedCells().map((cell) => cellPosition(cell).row);
      const selected = new Set([...checkedRows, ...cellRows]);
      if (!selected.size) return;
      rows = rows.filter((_, idx) => !selected.has(idx));
      render();
    });
    thead.addEventListener("change", (event) => {
      if (event.target.id !== "selectAllRows") return;
      tbody.querySelectorAll('input[type="checkbox"]').forEach((input) => {
        input.checked = event.target.checked;
      });
    });
    tbody.addEventListener("input", (event) => {
      if (!event.target.closest("[data-key]")) return;
      syncFromTable();
    });
    tbody.addEventListener("blur", (event) => {
      if (!event.target.closest("[data-key]")) return;
      syncFromTable();
    }, true);
    tbody.addEventListener("pointerdown", (event) => {
      const cell = event.target.closest("[data-key]");
      if (!cell) return;
      selectionStart = cellPosition(cell);
      selectionEnd = selectionStart;
      isSelecting = true;
      document.body.classList.add("selecting-cells");
      paintSelection();
      cell.focus();
    });
    tbody.addEventListener("pointerover", (event) => {
      if (!isSelecting) return;
      const cell = event.target.closest("[data-key]");
      if (!cell) return;
      selectionEnd = cellPosition(cell);
      paintSelection();
    });
    document.addEventListener("pointerup", () => {
      isSelecting = false;
      document.body.classList.remove("selecting-cells");
    });
    tbody.addEventListener("paste", (event) => {
      const cell = event.target.closest("[data-key]");
      if (!cell) return;
      const text = event.clipboardData.getData("text");
      if (!text.includes("\\t") && !text.includes("\\n")) return;
      event.preventDefault();
      syncFromTable();
      const start = cellPosition(cell);
      text.trimEnd().split(/\\r?\\n/).forEach((line, rowOffset) => {
        line.split("\\t").forEach((value, colOffset) => {
          const row = rows[start.row + rowOffset];
          const column = columns[start.col + colOffset];
          if (row && column) row[column[0]] = value;
        });
      });
      render();
    });
    document.addEventListener("keydown", (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
        const cells = selectedCells();
        if (!cells.length) return;
        const positions = cells.map(cellPosition);
        const rowMin = Math.min(...positions.map((pos) => pos.row));
        const rowMax = Math.max(...positions.map((pos) => pos.row));
        const colMin = Math.min(...positions.map((pos) => pos.col));
        const colMax = Math.max(...positions.map((pos) => pos.col));
        const matrix = [];
        for (let row = rowMin; row <= rowMax; row++) {
          const values = [];
          for (let col = colMin; col <= colMax; col++) {
            const cell = tbody.querySelector(`tr[data-row="${row}"] [data-col="${col}"]`);
            values.push(cell ? cell.textContent.trim() : "");
          }
          matrix.push(values.join("\\t"));
        }
        navigator.clipboard?.writeText(matrix.join("\\n"));
      }
    });
    document.querySelector("#sheetForm").addEventListener("submit", () => {
      syncFromTable();
      document.querySelector("#rowsJson").value = JSON.stringify(rows);
      setLoading(document.querySelector('#sheetForm button[type="submit"]'), "Salvando...");
      lockForm(document.querySelector("#sheetForm"));
    });
    document.querySelector('.import-bar form').addEventListener("submit", (event) => {
      setLoading(event.currentTarget.querySelector('button[type="submit"]'), "Importando...");
      lockForm(event.currentTarget);
    });
    render();
  </script>
</body>
</html>
"""


CAPACITY_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="icon" href="{favicon_url}" type="image/svg+xml">
  <title>Capacidades - Dashboard</title>
  <style>
    :root {
      --bg: #eef2f5;
      --panel: #ffffff;
      --ink: #16212d;
      --muted: #657282;
      --line: #d7e0e8;
      --teal: #64248c;
      --blue: #2b84cb;
      --danger: #b91c1c;
      --shadow: 0 18px 42px rgba(23, 32, 51, .10);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font-family: Inter, Segoe UI, Roboto, Arial, sans-serif;
    }
    header {
      padding: 24px clamp(16px, 4vw, 38px);
      background: #34104f;
      color: #fff;
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 18px;
    }
    .brand-title { display: flex; align-items: center; gap: 14px; }
    .brand-title img { width: 78px; height: auto; object-fit: contain; }
    h1 { margin: 0; font-size: clamp(26px, 3vw, 38px); }
    .subtitle { margin: 7px 0 0; color: #c8d6dc; }
    .nav { display: flex; flex-wrap: wrap; gap: 9px; justify-content: flex-end; }
    .top-link, button, .button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 38px;
      padding: 9px 12px;
      border: 1px solid rgba(255,255,255,.28);
      border-radius: 8px;
      background: rgba(255,255,255,.08);
      color: #fff;
      font: inherit;
      font-size: 13px;
      font-weight: 900;
      text-decoration: none;
      cursor: pointer;
    }
    main { padding: 24px clamp(16px, 4vw, 38px) 38px; }
    .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      overflow: hidden;
    }
    .message {
      margin: 16px;
      padding: 12px 14px;
      border-radius: 8px;
      background: #e7f7ee;
      color: #166534;
      font-weight: 800;
    }
    .message.error { background: #fee2e2; color: var(--danger); }
    .toolbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      padding: 15px;
      border-bottom: 1px solid var(--line);
    }
    .meta { color: var(--muted); font-size: 13px; font-weight: 800; }
    .draft-status {
      display: inline-flex;
      margin-left: 8px;
      color: #64248c;
      font-size: 12px;
      font-weight: 900;
    }
    .tabs {
      display: flex;
      gap: 8px;
      padding: 15px 15px 0;
      background: #f8fafc;
      border-bottom: 1px solid var(--line);
    }
    .tab-button {
      border-color: var(--line);
      background: #fff;
      color: var(--ink);
    }
    .tab-button.active {
      border-color: transparent;
      background: var(--teal);
      color: #fff;
    }
    .tab-panel { display: none; }
    .tab-panel.active { display: block; }
    .actions { display: flex; flex-wrap: wrap; gap: 8px; justify-content: flex-end; }
    .actions button, .actions .button {
      border-color: transparent;
      background: var(--teal);
    }
    .actions .button { background: #f3f6f8; color: var(--ink); border-color: var(--line); }
    button.is-loading,
    button:disabled {
      cursor: wait;
      opacity: .78;
    }
    button.is-loading::after {
      content: "";
      width: 13px;
      height: 13px;
      margin-left: 8px;
      border: 2px solid rgba(255,255,255,.45);
      border-top-color: #fff;
      border-radius: 50%;
      animation: spin .75s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    .sheet-wrap { overflow: auto; max-height: calc(100vh - 230px); }
    table { width: 100%; border-collapse: collapse; min-width: 980px; }
    #conductorSheet { min-width: 520px; }
    th, td { border-bottom: 1px solid var(--line); border-right: 1px solid #eef2f5; text-align: left; }
    th {
      position: sticky;
      top: 0;
      z-index: 2;
      padding: 10px;
      background: #eef3f6;
      color: #506071;
      font-size: 12px;
      text-transform: uppercase;
      white-space: nowrap;
    }
    td[contenteditable="true"] {
      padding: 9px 10px;
      outline: 0;
      background: #fff;
      color: var(--ink);
      min-height: 36px;
    }
    td[contenteditable="true"]:focus {
      background: #e7f4f2;
      box-shadow: inset 0 0 0 2px var(--teal);
    }
    tr:nth-child(even) td[contenteditable="true"] { background: #fbfcfd; }
    .hint {
      padding: 12px 14px;
      border-top: 1px solid var(--line);
      color: var(--muted);
      font-size: 13px;
      line-height: 1.45;
    }
    @media (max-width: 760px) {
      header { flex-direction: column; }
      .toolbar { align-items: flex-start; flex-direction: column; }
    }
  </style>
</head>
<body>
  <header>
    <div>
      <div class="brand-title"><img src="{favicon_url}" alt=""><h1>Capacidades</h1></div>
      <p class="subtitle">Cadastro de capacidades por cavalo, carreta e caminhao.</p>
    </div>
    <nav class="nav">
        <a class="top-link" href="/home">Home</a>
        <a class="top-link" href="/dashboard">Dashboard</a>
        <a class="top-link" href="/controle-ct">Controle de CT</a>
        <a class="top-link" href="/relatorio-diario">Relatorio diario</a>
        <a class="top-link" href="/relatorio-entrada-notas">Entrada de notas</a>
        <a class="top-link" href="/editar">Editar dados</a>
        <a class="top-link" href="/logout">Sair</a>
    </nav>
  </header>
  <main>
    <section class="panel">
      {message}
      <form id="capacityForm" method="post" action="/capacidades">
        <input type="hidden" name="rows_json" id="rowsJson">
        <input type="hidden" name="conductors_json" id="conductorsJson">
        <div class="tabs" role="tablist" aria-label="Cadastros">
          <button type="button" class="tab-button active" data-tab="capacities">Capacidades</button>
          <button type="button" class="tab-button" data-tab="conductors">Condutores</button>
        </div>
        <div class="tab-panel active" data-panel="capacities">
          <div class="toolbar">
            <div class="meta"><span id="rowCount">__ROW_COUNT__</span> capacidades cadastradas</div>
            <div class="actions">
              <button type="button" id="addRow">Adicionar linha</button>
              <button type="button" id="deleteRows" class="button">Excluir selecionadas</button>
              <button type="submit">Salvar e atualizar dashboard</button>
            </div>
          </div>
          <div class="sheet-wrap">
            <table id="sheet">
              <thead></thead>
              <tbody></tbody>
            </table>
          </div>
          <div class="hint">A capacidade pode ser digitada em mil litros (30) ou litros (30000). O dashboard vincula pela placa do cavalo ou da carreta, ignorando hifen e espacos.</div>
        </div>
        <div class="tab-panel" data-panel="conductors">
          <div class="toolbar">
            <div class="meta"><span id="conductorCount">__CONDUCTOR_COUNT__</span> condutores cadastrados</div>
            <div class="actions">
              <button type="button" id="addConductor">Adicionar condutor</button>
              <button type="button" id="deleteConductors" class="button">Excluir selecionados</button>
              <button type="submit">Salvar e atualizar dashboard</button>
            </div>
          </div>
          <div class="sheet-wrap">
            <table id="conductorSheet">
              <thead><tr><th></th><th>Motorista</th><th>Tipo de Frete</th></tr></thead>
              <tbody></tbody>
            </table>
          </div>
          <div class="hint">Os nomes salvos aqui aparecem em ordem alfabetica no campo Motorista da tela Controle de CT.</div>
        </div>
      </form>
    </section>
  </main>
  <script>
    function setLoading(button, text) {
      if (!button) return;
      button.textContent = text;
      button.classList.add("is-loading");
      button.disabled = true;
    }

    function lockForm(form) {
      form.querySelectorAll("button").forEach((button) => {
        if (button.type !== "submit") button.disabled = true;
      });
    }

    const columns = [
      ["id", "ID"],
      ["tipo", "Tipo"],
      ["capacidade", "Capacidade"],
      ["placaCavalo", "Placa cavalo"],
      ["tanques", "Tanques"],
      ["carreta", "Carreta"],
      ["observacao", "Observacao"]
    ];
    let rows = __ROWS__;
    let conductors = __CONDUCTORS__;
    const thead = document.querySelector("#sheet thead");
    const tbody = document.querySelector("#sheet tbody");
    const conductorTbody = document.querySelector("#conductorSheet tbody");
    const rowCount = document.querySelector("#rowCount");
    const conductorCount = document.querySelector("#conductorCount");

    function escapeHtml(value) {
      return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
    }

    function cleanRow(row = {}) {
      return Object.fromEntries(columns.map(([key]) => [key, row[key] ?? ""]));
    }

    function syncFromTable() {
      rows = [...tbody.querySelectorAll("tr")].map((tr) => {
        const row = {};
        tr.querySelectorAll("[data-key]").forEach((cell) => {
          row[cell.dataset.key] = cell.textContent.trim().toUpperCase();
        });
        return row;
      });
    }

    function normalizeFreight(value = "") {
      const text = String(value ?? "").trim();
      const key = text.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
      if (key === "transferencia") return "Transferencia";
      if (key === "cif") return "CIF";
      if (key === "fob") return "FOB";
      if (key === "rzd") return "RZD";
      return text;
    }

    function cleanConductor(row = {}) {
      if (typeof row === "string") {
        return { nome: row.trim().toUpperCase(), tipoFrete: "" };
      }
      return {
        nome: String(row.nome ?? row.motorista ?? "").trim().toUpperCase(),
        tipoFrete: normalizeFreight(row.tipoFrete ?? row.tipo_frete ?? "")
      };
    }

    function sortConductors() {
      const cleaned = conductors.map(cleanConductor);
      const blankCount = cleaned.filter((row) => !row.nome).length;
      const unique = new Map();
      cleaned.filter((row) => row.nome).forEach((row) => {
        if (!unique.has(row.nome) || (!unique.get(row.nome).tipoFrete && row.tipoFrete)) unique.set(row.nome, row);
      });
      conductors = [
        ...[...unique.values()].sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR")),
        ...Array(blankCount).fill({ nome: "", tipoFrete: "" })
      ];
    }

    function syncConductorsFromTable() {
      conductors = [...conductorTbody.querySelectorAll("tr")].map((tr) => cleanConductor({
        nome: tr.querySelector("[data-key='nome']")?.textContent || "",
        tipoFrete: tr.querySelector("[data-key='tipoFrete']")?.value || ""
      }));
      sortConductors();
    }

    function render() {
      thead.innerHTML = `<tr><th></th>${columns.map(([, label]) => `<th>${label}</th>`).join("")}</tr>`;
      tbody.innerHTML = rows.map((row, idx) => {
        const clean = cleanRow(row);
        return `<tr>
          <td><input type="checkbox" aria-label="Selecionar linha ${idx + 1}"><br>${idx + 1}</td>
          ${columns.map(([key]) => `<td contenteditable="true" data-key="${key}">${escapeHtml(clean[key])}</td>`).join("")}
        </tr>`;
      }).join("");
      rowCount.textContent = rows.length.toLocaleString("pt-BR");
      sortConductors();
      conductorTbody.innerHTML = conductors.map((row, idx) => `
        <tr>
          <td><input type="checkbox" aria-label="Selecionar condutor ${idx + 1}"><br>${idx + 1}</td>
          <td contenteditable="true" data-key="nome">${escapeHtml(row.nome)}</td>
          <td>
            <select data-key="tipoFrete">
              <option value=""></option>
              <option value="CIF" ${row.tipoFrete === "CIF" ? "selected" : ""}>CIF</option>
              <option value="FOB" ${row.tipoFrete === "FOB" ? "selected" : ""}>FOB</option>
              <option value="Transferencia" ${row.tipoFrete === "Transferencia" ? "selected" : ""}>Transferencia</option>
              <option value="RZD" ${row.tipoFrete === "RZD" ? "selected" : ""}>RZD</option>
            </select>
          </td>
        </tr>
      `).join("");
      conductorCount.textContent = conductors.filter((row) => row.nome).length.toLocaleString("pt-BR");
    }

    document.querySelectorAll(".tab-button").forEach((button) => {
      button.addEventListener("click", () => {
        syncFromTable();
        syncConductorsFromTable();
        document.querySelectorAll(".tab-button").forEach((item) => item.classList.toggle("active", item === button));
        document.querySelectorAll(".tab-panel").forEach((panel) => panel.classList.toggle("active", panel.dataset.panel === button.dataset.tab));
        render();
      });
    });

    document.querySelector("#addRow").addEventListener("click", () => {
      syncFromTable();
      rows.push(cleanRow({ id: String(rows.length + 1), tipo: "CAMINHAO", capacidade: "25" }));
      render();
    });

    document.querySelector("#deleteRows").addEventListener("click", () => {
      syncFromTable();
      const checked = new Set([...tbody.querySelectorAll("input[type='checkbox']")]
        .map((input, idx) => input.checked ? idx : -1)
        .filter((idx) => idx >= 0));
      rows = rows.filter((_, idx) => !checked.has(idx));
      render();
    });

    document.querySelector("#addConductor").addEventListener("click", () => {
      syncConductorsFromTable();
      conductors.push({ nome: "", tipoFrete: "" });
      render();
      document.querySelector("#conductorSheet tbody tr:last-child [data-key='nome']")?.focus();
    });

    document.querySelector("#deleteConductors").addEventListener("click", () => {
      syncConductorsFromTable();
      const checked = new Set([...conductorTbody.querySelectorAll("input[type='checkbox']")]
        .map((input, idx) => input.checked ? idx : -1)
        .filter((idx) => idx >= 0));
      conductors = conductors.filter((_, idx) => !checked.has(idx));
      render();
    });

    document.querySelector("#capacityForm").addEventListener("submit", () => {
      syncFromTable();
      syncConductorsFromTable();
      document.querySelector("#rowsJson").value = JSON.stringify(rows);
      document.querySelector("#conductorsJson").value = JSON.stringify(conductors);
      setLoading(document.querySelector('#capacityForm button[type="submit"]'), "Salvando...");
      lockForm(document.querySelector("#capacityForm"));
    });

    render();
  </script>
</body>
</html>
"""


CT_CONTROL_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="icon" href="{favicon_url}" type="image/svg+xml">
  <title>Controle de CT - Dashboard</title>
  <style>
    :root {
      --bg: #eef2f5;
      --panel: #ffffff;
      --ink: #16212d;
      --muted: #657282;
      --line: #d7e0e8;
      --purple: #64248c;
      --blue: #2b84cb;
      --red: #e2263c;
      --green: #00856f;
      --shadow: 0 18px 42px rgba(23, 32, 51, .10);
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
      padding: 24px clamp(16px, 4vw, 38px);
      background: #34104f;
      color: #fff;
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 18px;
    }
    .brand-title { display: flex; align-items: center; gap: 14px; }
    .brand-title img { width: 78px; height: auto; object-fit: contain; }
    h1 { margin: 0; font-size: clamp(26px, 3vw, 38px); letter-spacing: 0; }
    .subtitle { margin: 7px 0 0; color: #c8d6dc; }
    .nav { display: flex; flex-wrap: wrap; gap: 9px; justify-content: flex-end; }
    a { color: inherit; text-decoration: none; }
    .top-link, button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 38px;
      padding: 9px 12px;
      border: 1px solid rgba(255,255,255,.28);
      border-radius: 8px;
      background: rgba(255,255,255,.08);
      color: #fff;
      font: inherit;
      font-size: 13px;
      font-weight: 900;
      cursor: pointer;
    }
    main { padding: 24px clamp(16px, 4vw, 38px) 38px; }
    .filters, .kpis, .content-grid { display: grid; gap: 14px; }
    .filters {
      grid-template-columns: repeat(5, minmax(150px, 1fr));
      align-items: end;
      margin-bottom: 16px;
      padding: 15px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
      box-shadow: var(--shadow);
    }
    label { display: grid; gap: 7px; color: var(--muted); font-size: 12px; font-weight: 900; text-transform: uppercase; }
    select, input {
      min-height: 40px;
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 8px 10px;
      color: var(--ink);
      background: #fff;
      font: inherit;
      font-weight: 800;
    }
    .filters button { border-color: transparent; background: var(--purple); }
    .kpis { grid-template-columns: repeat(4, minmax(150px, 1fr)); margin-bottom: 16px; }
    .kpi {
      min-height: 112px;
      padding: 16px;
      border: 1px solid var(--line);
      border-top: 5px solid var(--purple);
      border-radius: 8px;
      background: var(--panel);
      box-shadow: var(--shadow);
    }
    .kpi:nth-child(2) { border-top-color: var(--blue); }
    .kpi:nth-child(3) { border-top-color: var(--red); }
    .kpi:nth-child(4) { border-top-color: var(--green); }
    .kpi span { color: var(--muted); font-size: 12px; font-weight: 900; text-transform: uppercase; }
    .kpi strong { display: block; margin-top: 10px; font-size: clamp(28px, 4vw, 42px); line-height: 1; }
    .content-grid { grid-template-columns: minmax(280px, .72fr) minmax(0, 1.28fr); }
    .panel {
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
      box-shadow: var(--shadow);
      overflow: hidden;
    }
    .panel h2 { margin: 0; padding: 15px; border-bottom: 1px solid var(--line); font-size: 18px; }
    .bars { display: grid; gap: 12px; padding: 15px; }
    .bar-row { display: grid; gap: 7px; }
    .bar-head { display: flex; justify-content: space-between; gap: 12px; color: var(--muted); font-size: 13px; font-weight: 900; }
    .bar-track { height: 10px; border-radius: 999px; background: #edf2f5; overflow: hidden; }
    .bar-fill { height: 100%; border-radius: inherit; background: var(--purple); }
    .table-wrap { max-height: calc(100vh - 375px); min-height: 360px; overflow: auto; }
    table { width: 100%; min-width: 1180px; border-collapse: collapse; }
    th, td { padding: 10px 11px; border-bottom: 1px solid var(--line); text-align: left; vertical-align: top; }
    th { position: sticky; top: 0; z-index: 2; background: #eef3f6; color: #506071; font-size: 12px; text-transform: uppercase; white-space: nowrap; }
    td { font-size: 13px; }
    .status {
      display: inline-flex;
      padding: 5px 8px;
      border-radius: 999px;
      background: #edf2f5;
      color: var(--ink);
      font-size: 12px;
      font-weight: 900;
      white-space: nowrap;
    }
    .status.finalizado { background: #e7f7ee; color: #166534; }
    .status.aguardando { background: #fff4de; color: #92400e; }
    .status.fila { background: #e8f2ff; color: #1d4ed8; }
    .empty { padding: 26px; color: var(--muted); font-weight: 800; text-align: center; }
    @media (max-width: 980px) {
      header { flex-direction: column; }
      .nav { justify-content: flex-start; }
      .filters, .kpis, .content-grid { grid-template-columns: 1fr; }
      .table-wrap { max-height: none; }
    }
  </style>
</head>
<body>
  <header>
    <div>
      <div class="brand-title"><img src="{favicon_url}" alt=""><h1>Controle de CT</h1></div>
      <p class="subtitle">Entrada, fila, finalizados, frete, notas e tempos de carregamento.</p>
    </div>
    <nav class="nav">
      <a class="top-link" href="/home">Home</a>
      <a class="top-link" href="/dashboard">Dashboard</a>
      <a class="top-link" href="/editar">Editar dados</a>
      <a class="top-link" href="/capacidades">Capacidades</a>
      <a class="top-link" href="/relatorio-diario">Relatorio diario</a>
      <a class="top-link" href="/logout">Sair</a>
    </nav>
  </header>
  <main>
    <section class="filters">
      <label>Data <select id="dateFilter"></select></label>
      <label>Status <select id="statusFilter"></select></label>
      <label>Tipo de Frete <select id="freightFilter"></select></label>
      <label>Nota Fiscal <select id="invoiceFilter"></select></label>
      <label>Buscar <input id="searchFilter" type="search" placeholder="Motorista ou observacao"></label>
      <button type="button" id="clearFilters">Limpar filtros</button>
    </section>
    <section class="kpis">
      <div class="kpi"><span>Registros</span><strong id="kRecords">0</strong></div>
      <div class="kpi"><span>Viagens</span><strong id="kTrips">0</strong></div>
      <div class="kpi"><span>Finalizados</span><strong id="kDone">0</strong></div>
      <div class="kpi"><span>Tempo medio total</span><strong id="kAvg">-</strong></div>
    </section>
    <section class="content-grid">
      <div class="panel">
        <h2>Status por frete</h2>
        <div class="bars" id="bars"></div>
      </div>
      <div class="panel">
        <h2>Base da planilha</h2>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Data</th><th>Motorista</th><th>Frete</th><th>Status</th><th>Viagens</th>
                <th>Chegada</th><th>Entrada</th><th>Saida</th><th>Nota</th><th>Carregamento</th><th>Total</th><th>Observacao</th>
              </tr>
            </thead>
            <tbody id="rows"></tbody>
          </table>
        </div>
      </div>
    </section>
  </main>
  <script>
    const records = __CT_RECORDS__;
    const fmt = new Intl.NumberFormat("pt-BR");
    const $ = (id) => document.getElementById(id);

    function unique(key) {
      return [...new Set(records.map((row) => row[key]).filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), "pt-BR"));
    }
    function fillSelect(id, values, label) {
      $(id).innerHTML = [`<option value="">${label}</option>`, ...values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`)].join("");
    }
    function escapeHtml(value) {
      return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }
    function statusClass(value) {
      const text = String(value).toLowerCase();
      if (text.includes("finalizado")) return "finalizado";
      if (text.includes("aguardando")) return "aguardando";
      if (text.includes("fila")) return "fila";
      return "";
    }
    function filteredRows() {
      const date = $("dateFilter").value;
      const status = $("statusFilter").value;
      const freight = $("freightFilter").value;
      const invoice = $("invoiceFilter").value;
      const query = $("searchFilter").value.trim().toLowerCase();
      return records.filter((row) => {
        const haystack = `${row.motorista} ${row.observacao}`.toLowerCase();
        return (!date || row.data === date)
          && (!status || row.status === status)
          && (!freight || row.tipoFrete === freight)
          && (!invoice || row.notaFiscal === invoice)
          && (!query || haystack.includes(query));
      });
    }
    function trips(row) {
      const match = String(row.viagens || "").match(/\\d+/);
      return match ? Number(match[0]) : 0;
    }
    function averageTotal(rows) {
      const values = rows.map((row) => row.tempoTotalMinutos).filter((value) => Number.isFinite(value) && value > 0);
      if (!values.length) return "-";
      const minutes = values.reduce((total, value) => total + value, 0) / values.length;
      const h = Math.floor(minutes / 60);
      const m = Math.round(minutes % 60);
      return h ? `${h}h ${String(m).padStart(2, "0")}m` : `${m}m`;
    }
    function renderBars(rows) {
      const counts = new Map();
      rows.forEach((row) => {
        const key = `${row.status || "Sem status"} / ${row.tipoFrete || "Sem frete"}`;
        counts.set(key, (counts.get(key) || 0) + 1);
      });
      const entries = [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, 10);
      const max = Math.max(...entries.map(([, value]) => value), 1);
      $("bars").innerHTML = entries.length ? entries.map(([label, value]) => `
        <div class="bar-row">
          <div class="bar-head"><span>${escapeHtml(label)}</span><strong>${fmt.format(value)}</strong></div>
          <div class="bar-track"><div class="bar-fill" style="width:${Math.max(4, value / max * 100)}%"></div></div>
        </div>
      `).join("") : '<div class="empty">Nenhum registro para os filtros atuais.</div>';
    }
    function renderTable(rows) {
      $("rows").innerHTML = rows.length ? rows.map((row) => `
        <tr>
          <td>${escapeHtml(row.data)}</td>
          <td>${escapeHtml(row.motorista)}</td>
          <td>${escapeHtml(row.tipoFrete)}</td>
          <td><span class="status ${statusClass(row.status)}">${escapeHtml(row.status || "-")}</span></td>
          <td>${escapeHtml(row.viagens)}</td>
          <td>${escapeHtml(row.chegada)}</td>
          <td>${escapeHtml(row.entrada)}</td>
          <td>${escapeHtml(row.saida)}</td>
          <td>${escapeHtml(row.notaFiscal)}</td>
          <td>${escapeHtml(row.tempoCarregamento)}</td>
          <td>${escapeHtml(row.tempoTotal)}</td>
          <td>${escapeHtml(row.observacao)}</td>
        </tr>
      `).join("") : '<tr><td class="empty" colspan="12">Nenhum registro encontrado.</td></tr>';
    }
    function render() {
      const rows = filteredRows();
      $("kRecords").textContent = fmt.format(rows.length);
      $("kTrips").textContent = fmt.format(rows.reduce((total, row) => total + trips(row), 0));
      $("kDone").textContent = fmt.format(rows.filter((row) => String(row.status).toLowerCase().includes("finalizado")).length);
      $("kAvg").textContent = averageTotal(rows);
      renderBars(rows);
      renderTable(rows);
    }
    fillSelect("dateFilter", unique("data").reverse(), "Todas");
    fillSelect("statusFilter", unique("status"), "Todos");
    fillSelect("freightFilter", unique("tipoFrete"), "Todos");
    fillSelect("invoiceFilter", unique("notaFiscal"), "Todas");
    ["dateFilter", "statusFilter", "freightFilter", "invoiceFilter", "searchFilter"].forEach((id) => $(id).addEventListener("input", render));
    $("clearFilters").addEventListener("click", () => {
      ["dateFilter", "statusFilter", "freightFilter", "invoiceFilter", "searchFilter"].forEach((id) => $(id).value = "");
      render();
    });
    render();
  </script>
</body>
</html>
"""


CT_CONTROL_OPERATION_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="icon" href="{favicon_url}" type="image/svg+xml">
  <title>Controle de CT - Dashboard</title>
  <style>
    :root {
      --purple: #7027a8;
      --purple-dark: #4f167e;
      --panel: #ffffff;
      --ink: #18212f;
      --muted: #657282;
      --line: #cfd9e4;
      --green: #9de36e;
      --yellow: #ffe878;
      --red: #e2263c;
      --blue: #2b84cb;
      --shadow: 0 18px 42px rgba(23, 32, 51, .16);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      background: #eef2f5;
      color: var(--ink);
      font-family: Inter, Segoe UI, Roboto, Arial, sans-serif;
    }
    header {
      position: relative;
      overflow: hidden;
      padding: 24px clamp(16px, 4vw, 42px) 28px;
      background:
        radial-gradient(720px circle at 76% 35%, rgba(43,132,203,.34), transparent 62%),
        linear-gradient(135deg, #34104f 0%, #4c176d 58%, #1b255f 100%);
      color: #fff;
    }
    header::after {
      content: "";
      position: absolute;
      right: clamp(20px, 6vw, 76px);
      bottom: -88px;
      width: min(46vw, 520px);
      aspect-ratio: 1.8;
      background: url("{favicon_url}") center / contain no-repeat;
      opacity: .18;
      pointer-events: none;
    }
    .topbar {
      position: relative;
      z-index: 2;
      display: flex;
      justify-content: space-between;
      gap: 18px;
      align-items: flex-start;
    }
    .brand-title { display: flex; align-items: center; gap: 16px; }
    .brand-title img { width: 98px; height: auto; object-fit: contain; filter: drop-shadow(0 10px 18px rgba(0,0,0,.24)); }
    h1 { margin: 0; font-size: clamp(28px, 4vw, 50px); line-height: 1; letter-spacing: 0; }
    .subtitle { margin: 9px 0 0; color: #d7e4ea; font-size: 15px; }
    .nav { display: flex; flex-wrap: wrap; justify-content: flex-end; gap: 9px; }
    a { color: inherit; text-decoration: none; }
    .top-link, button, .button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 36px;
      padding: 8px 11px;
      border: 1px solid rgba(255,255,255,.32);
      border-radius: 6px;
      background: rgba(255,255,255,.10);
      color: #fff;
      font: inherit;
      font-size: 13px;
      font-weight: 900;
      cursor: pointer;
    }
    main { padding: 14px clamp(10px, 2.4vw, 24px) 30px; }
    .message {
      width: min(420px, calc(100vw - 28px));
      padding: 13px 15px;
      border-radius: 8px;
      border: 1px solid #bbf7d0;
      background: #f0fdf4;
      color: #14532d;
      font-weight: 900;
      box-shadow: 0 18px 42px rgba(23, 32, 51, .22);
    }
    .message-zone {
      position: fixed;
      right: clamp(14px, 2vw, 24px);
      bottom: clamp(14px, 2vw, 24px);
      z-index: 20;
      display: grid;
      gap: 10px;
      justify-items: end;
      pointer-events: none;
    }
    .message-zone .message {
      pointer-events: auto;
      animation: toast-in .18s ease-out both;
    }
    .message.error {
      border-color: #fecaca;
      background: #fff1f2;
      color: #991b1b;
    }
    .message.warning {
      border-color: #fde68a;
      background: #fffbeb;
      color: #92400e;
    }
    @keyframes toast-in {
      from { opacity: 0; transform: translateY(10px); }
      to { opacity: 1; transform: translateY(0); }
    }
    .control-board {
      display: grid;
      grid-template-columns: 1fr 1fr .8fr .8fr;
      gap: 12px;
      align-items: stretch;
      padding: 14px;
      border-radius: 8px 8px 0 0;
      background: var(--purple);
      box-shadow: var(--shadow);
    }
    .board-group {
      display: grid;
      grid-template-columns: repeat(2, minmax(90px, 1fr));
      gap: 9px;
      align-content: start;
    }
    .board-title {
      grid-column: 1 / -1;
      min-height: 30px;
      display: grid;
      place-items: center;
      border: 1px solid rgba(255,255,255,.4);
      background: linear-gradient(#ff6b78, #bb1232);
      color: #fff;
      font-size: 15px;
      font-weight: 1000;
      text-transform: uppercase;
      text-shadow: 0 1px 1px rgba(0,0,0,.35);
      box-shadow: inset 0 1px 0 rgba(255,255,255,.45);
    }
    .counter {
      display: grid;
      gap: 5px;
      align-content: start;
      text-align: center;
    }
    .counter-label {
      min-height: 24px;
      display: grid;
      place-items: center;
      padding: 3px 8px;
      background: linear-gradient(#fff, #d7dbe1);
      color: #263140;
      font-size: 12px;
      font-weight: 1000;
      text-transform: uppercase;
      box-shadow: 0 2px 8px rgba(0,0,0,.18);
    }
    .counter-value {
      min-height: 44px;
      display: grid;
      place-items: center;
      border-radius: 7px;
      background: linear-gradient(135deg, #ffffff, #cfd3dc);
      color: #444;
      font-size: 25px;
      font-weight: 1000;
      box-shadow: inset 0 1px 0 rgba(255,255,255,.9), 0 8px 16px rgba(0,0,0,.24);
    }
    .counter.green .counter-value { background: linear-gradient(135deg, #d8ffc4, var(--green)); color: #28521d; }
    .counter.yellow .counter-value { background: linear-gradient(135deg, #fff7ba, var(--yellow)); color: #735a00; }
    .toolbar {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
      justify-content: space-between;
      padding: 11px;
      border: 1px solid var(--line);
      border-top: 0;
      background: #fff;
    }
    .actions, .filters { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
    .actions button { border-color: transparent; background: var(--purple); }
    .actions .secondary { background: #f3f6f8; color: var(--ink); border-color: var(--line); }
    .export-link {
      min-height: 34px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 7px 9px;
      border: 1px solid var(--line);
      border-radius: 6px;
      color: var(--muted);
      background: #f8fafc;
      font-size: 12px;
      font-weight: 800;
      text-decoration: none;
    }
    .export-link:hover {
      color: var(--ink);
      background: #eef3f7;
    }
    .edit-toggle {
      gap: 7px;
    }
    .edit-toggle svg {
      width: 18px;
      height: 18px;
      flex: 0 0 auto;
    }
    .edit-toggle.is-editing {
      border-color: transparent;
      background: linear-gradient(180deg, #7c31b4, var(--purple));
      color: #fff;
    }
    .filters input, .filters select {
      min-height: 36px;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 7px 9px;
      color: var(--ink);
      font: inherit;
      font-weight: 800;
      background: #fff;
    }
    .multi-filter {
      position: relative;
      min-width: 210px;
    }
    .multi-filter summary {
      min-height: 36px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 7px 9px;
      color: var(--ink);
      background: #fff;
      font-weight: 900;
      cursor: pointer;
      list-style: none;
    }
    .multi-filter summary::-webkit-details-marker { display: none; }
    .multi-filter summary::after {
      content: "";
      width: 8px;
      height: 8px;
      border-right: 2px solid currentColor;
      border-bottom: 2px solid currentColor;
      transform: rotate(45deg) translateY(-2px);
      opacity: .8;
      flex: 0 0 auto;
    }
    .multi-filter[open] summary::after {
      transform: rotate(225deg) translateY(-1px);
    }
    .multi-filter-panel {
      position: absolute;
      top: calc(100% + 6px);
      left: 0;
      z-index: 10;
      width: min(320px, 88vw);
      max-height: 300px;
      overflow: auto;
      padding: 8px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      box-shadow: 0 18px 38px rgba(23,32,51,.18);
    }
    .multi-filter-actions {
      display: flex;
      justify-content: flex-end;
      padding-bottom: 6px;
      border-bottom: 1px solid #edf1f5;
      margin-bottom: 6px;
    }
    .multi-filter-clear {
      min-height: 28px;
      padding: 5px 8px;
      border: 1px solid var(--line);
      border-radius: 6px;
      color: var(--muted);
      background: #f8fafc;
      font-size: 12px;
      font-weight: 900;
      cursor: pointer;
    }
    .multi-filter-option {
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 6px 4px;
      color: var(--ink);
      font-size: 13px;
      font-weight: 800;
      cursor: pointer;
    }
    .multi-filter-option input {
      width: 14px;
      height: 14px;
      margin: 0;
    }
    .custom-date-filter {
      display: none;
      gap: 6px;
      align-items: center;
    }
    .custom-date-filter.is-visible { display: inline-flex; }
    .custom-date-filter input { width: 136px; }
    .sheet-wrap {
      overflow: auto;
      max-height: calc(100vh - 310px);
      border: 1px solid var(--line);
      border-top: 0;
      background: #fff;
      box-shadow: var(--shadow);
    }
    table { width: 100%; min-width: 1160px; border-collapse: collapse; }
    th, td { border-right: 1px solid #d7dbea; border-bottom: 1px solid #d7dbea; }
    th {
      position: sticky;
      top: 0;
      z-index: 2;
      padding: 7px 8px;
      background: var(--purple);
      color: #fff;
      font-size: 12px;
      text-align: left;
      white-space: nowrap;
    }
    td { padding: 0; background: #fff; }
    tr:nth-child(even) td { background: #fbfcfd; }
    tr.is-selected td { background: #e9f3ff; }
    td:first-child, th:first-child { width: 40px; text-align: center; }
    input.cell, select.cell {
      width: 100%;
      min-height: 31px;
      border: 0;
      padding: 5px 7px;
      color: var(--ink);
      background: transparent;
      font: inherit;
      font-size: 13px;
    }
    input.cell:focus, select.cell:focus {
      outline: 2px solid #2b84cb;
      outline-offset: -2px;
      background: #fff;
    }
    input.cell:disabled, select.cell:disabled {
      cursor: default;
      color: #243041;
      opacity: 1;
    }
    select.cell:disabled {
      appearance: none;
      -webkit-appearance: none;
      background-image: none;
      padding-right: 7px;
    }
    select.cell:disabled::-ms-expand { display: none; }
    body.editing-ct input.cell,
    body.editing-ct select.cell {
      background: #fffdf1;
      box-shadow: inset 0 0 0 1px rgba(226, 177, 38, .38);
    }
    body.editing-ct input.cell:focus,
    body.editing-ct select.cell:focus {
      background: #fff;
      box-shadow: inset 0 0 0 2px #2b84cb;
    }
    .status-finalizado { color: #008000; font-weight: 900; }
    .status-fila { color: #b27300; font-weight: 900; }
    .status-patio { color: #1d4ed8; font-weight: 900; }
    .hint {
      padding: 10px 12px;
      border: 1px solid var(--line);
      border-top: 0;
      color: var(--muted);
      background: #fff;
      font-size: 13px;
      line-height: 1.45;
    }
    @media (max-width: 980px) {
      .topbar { flex-direction: column; }
      .nav { justify-content: flex-start; }
      .control-board { grid-template-columns: 1fr; }
      .toolbar { align-items: flex-start; flex-direction: column; }
      .sheet-wrap { max-height: none; }
    }
  </style>
</head>
<body>
  <header>
    <div class="topbar">
      <div>
        <div class="brand-title"><img src="{favicon_url}" alt=""><div><h1>Controle de CT</h1><p class="subtitle">Fila de chegada, entrada de atendimento e saida dos motoristas.</p></div></div>
      </div>
      <nav class="nav">
        <a class="top-link" href="/home">Home</a>
        <a class="top-link" href="/dashboard">Dashboard</a>
        <a class="top-link" href="/editar">Editar dados</a>
        <a class="top-link" href="/capacidades">Capacidades</a>
        <a class="top-link" href="/relatorio-diario">Relatorio diario</a>
        <a class="top-link" href="/relatorio-entrada-notas">Entrada de notas</a>
        <a class="top-link" href="/logout">Sair</a>
      </nav>
    </div>
  </header>
  <main>
    <div class="message-zone" aria-live="polite">
      {message}
      <div id="ctNotice" class="message warning" hidden></div>
    </div>
    <section class="control-board">
      <div class="board-group">
        <div class="board-title">Painel CIF</div>
        <div class="counter"><div class="counter-label">Base</div><div class="counter-value" id="cifBase">0</div></div>
        <div class="counter"><div class="counter-label">Fila</div><div class="counter-value" id="cifFila">0</div></div>
      </div>
      <div class="board-group">
        <div class="board-title">Painel FOB</div>
        <div class="counter"><div class="counter-label">Base</div><div class="counter-value" id="fobBase">0</div></div>
        <div class="counter"><div class="counter-label">Fila</div><div class="counter-value" id="fobFila">0</div></div>
      </div>
      <div class="counter green"><div class="counter-label">Finalizados</div><div class="counter-value" id="finalizados">0</div></div>
      <div class="counter yellow"><div class="counter-label">Patio</div><div class="counter-value" id="patio">0</div></div>
    </section>
    <form id="ctForm" method="post" action="/controle-ct">
      <input type="hidden" name="rows_json" id="rowsJson">
      <div class="toolbar">
        <div class="actions">
          <button type="button" id="addArrival">Adicionar chegada</button>
          <button type="button" id="editModeToggle" class="secondary edit-toggle" title="Alternar modo de edicao" aria-label="Editar">
            <span class="edit-icon" aria-hidden="true"></span>
            <span class="edit-label">Editar</span>
          </button>
          <button type="button" id="markEntry" class="secondary">Marcar entrada</button>
          <button type="button" id="markExit" class="secondary">Marcar saida</button>
          <button type="button" id="deleteRows" class="secondary">Excluir</button>
        </div>
        <div class="filters">
          <a class="export-link" href="/controle-ct/exportar" title="Exportar dados para Excel">Exportar Excel</a>
          <select id="dateRangeFilter" title="Filtrar por data">
            <option value="">Todas as datas</option>
            <option value="today">Hoje</option>
            <option value="yesterday">Ontem</option>
            <option value="week">Esta semana</option>
            <option value="last7">Ultimos 7 dias</option>
            <option value="month">Este mes</option>
            <option value="custom">Periodo</option>
          </select>
          <span class="custom-date-filter" id="customDateFilter">
            <input id="dateFromFilter" type="date" title="Data inicial">
            <input id="dateToFilter" type="date" title="Data final">
          </span>
          <details class="multi-filter" id="driverMultiFilter">
            <summary><span id="driverFilterLabel">Motoristas</span></summary>
            <div class="multi-filter-panel">
              <div class="multi-filter-actions"><button type="button" class="multi-filter-clear" id="clearDriverFilter">Limpar</button></div>
              <div id="driverFilterOptions"></div>
            </div>
          </details>
          <input id="searchFilter" type="search" placeholder="Buscar motorista">
          <select id="statusFilter">
            <option value="">Todos os status</option>
            <option>Fila de Carregamento</option>
            <option>Patio</option>
            <option>Finalizado</option>
          </select>
          <select id="freightFilter">
            <option value="">Todos os fretes</option>
            <option>CIF</option>
            <option>FOB</option>
            <option>Transferencia</option>
            <option>RZD</option>
          </select>
        </div>
      </div>
      <div class="sheet-wrap">
        <table>
          <thead>
            <tr>
              <th><input type="checkbox" id="selectAll" aria-label="Selecionar todos"></th>
              <th>Data</th><th>Motorista</th><th>Tipo de Frete</th><th>Status</th><th>Viagens</th>
              <th>Chegada</th><th>Entrada</th><th>Saida</th><th>Nota Fiscal</th><th>Observacao</th>
            </tr>
          </thead>
          <tbody id="rows"></tbody>
        </table>
      </div>
      <div class="hint">Use Adicionar chegada para colocar o motorista na fila. Depois selecione a linha e marque entrada ou saida; o horario atual sera preenchido automaticamente.</div>
    </form>
  </main>
  <script>
    let rows = __ROWS__;
    const conductors = __CONDUCTORS__;
    let editMode = false;
    let selectedDriverFilters = new Set();
    const $ = (id) => document.getElementById(id);
    function dismissToast(element, delay = 4200) {
      if (!element) return;
      window.setTimeout(() => element.remove(), delay);
    }
    document.querySelectorAll(".message.auto-dismiss").forEach((item) => dismissToast(item, 4200));
    const statuses = ["", "Fila de Carregamento", "Patio", "Finalizado"];
    const freights = ["", "CIF", "FOB", "Transferencia", "RZD"];
    const invoices = ["", "Impresso", "Pendente"];
    const statusFlow = ["Fila de Carregamento", "Patio", "Finalizado"];
    const statusStepLabels = {
      "Fila de Carregamento": "adicionar chegada",
      "Patio": "marcar entrada",
      "Finalizado": "marcar saida"
    };
    const requiredStepMessages = {
      "Patio": "Para registrar a entrada, primeiro adicione a chegada do motorista na fila.",
      "Finalizado": "Para registrar a saida, primeiro marque a entrada do motorista no patio."
    };
    const conductorFreights = new Map(conductors.map((item) => {
      const row = typeof item === "string" ? { nome: item, tipoFrete: "" } : item;
      return [String(row.nome || row.motorista || "").trim(), String(row.tipoFrete || row.tipo_frete || "").trim()];
    }).filter(([name]) => name));

    function nowDateTimeLocal() {
      const date = new Date();
      date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
      return date.toISOString().slice(0, 16);
    }
    function todayDateLocal() {
      return nowDateTimeLocal().slice(0, 10);
    }
    function escapeHtml(value) {
      return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }
    function escapeAttr(value) {
      return escapeHtml(value).replace(/"/g, "&quot;");
    }
    function toDateInput(value) {
      const text = String(value || "").trim();
      if (/^\\d{4}-\\d{2}-\\d{2}/.test(text)) return text.slice(0, 10);
      const match = text.match(/^(\\d{2})\\/(\\d{2})\\/(\\d{4})$/);
      return match ? `${match[3]}-${match[2]}-${match[1]}` : "";
    }
    function localDate(value) {
      const dateText = toDateInput(value);
      if (!dateText) return null;
      const [year, month, day] = dateText.split("-").map(Number);
      return new Date(year, month - 1, day);
    }
    function addDays(date, days) {
      const next = new Date(date);
      next.setDate(next.getDate() + days);
      return next;
    }
    function dateOnlyToday() {
      return localDate(todayDateLocal());
    }
    function dateRangeBounds() {
      const mode = $("dateRangeFilter").value;
      const today = dateOnlyToday();
      if (!mode || !today) return null;
      if (mode === "today") return { start: today, end: today };
      if (mode === "yesterday") {
        const yesterday = addDays(today, -1);
        return { start: yesterday, end: yesterday };
      }
      if (mode === "last7") return { start: addDays(today, -6), end: today };
      if (mode === "week") {
        const mondayOffset = (today.getDay() + 6) % 7;
        return { start: addDays(today, -mondayOffset), end: addDays(today, 6 - mondayOffset) };
      }
      if (mode === "month") return { start: new Date(today.getFullYear(), today.getMonth(), 1), end: new Date(today.getFullYear(), today.getMonth() + 1, 0) };
      if (mode === "custom") {
        return {
          start: localDate($("dateFromFilter").value),
          end: localDate($("dateToFilter").value)
        };
      }
      return null;
    }
    function matchesDateRange(row) {
      const range = dateRangeBounds();
      if (!range) return true;
      const rowDate = localDate(row.data);
      if (!rowDate) return false;
      return (!range.start || rowDate >= range.start) && (!range.end || rowDate <= range.end);
    }
    function updateCustomDateFilter() {
      $("customDateFilter").classList.toggle("is-visible", $("dateRangeFilter").value === "custom");
    }
    function toDateTimeInput(value) {
      const text = String(value || "").trim();
      if (/^\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}/.test(text)) return text.slice(0, 16);
      const dateTime = text.match(/^(\\d{2})\\/(\\d{2})\\/(\\d{4})\\s+(\\d{2}):(\\d{2})$/);
      if (dateTime) return `${dateTime[3]}-${dateTime[2]}-${dateTime[1]}T${dateTime[4]}:${dateTime[5]}`;
      const time = text.match(/^(\\d{2}):(\\d{2})$/);
      if (time) return `${new Date().toISOString().slice(0, 10)}T${time[1]}:${time[2]}`;
      return "";
    }
    function cleanRow(row = {}) {
      const status = row.status === "Aguardando Entrada" ? "Fila de Carregamento" : (row.status || "");
      return {
        data: row.data || "",
        motorista: row.motorista || "",
        tipoFrete: row.tipoFrete || "",
        status,
        viagens: row.viagens || "",
        chegada: row.chegada || "",
        entrada: row.entrada || "",
        saida: row.saida || "",
        notaFiscal: row.notaFiscal || "",
        observacao: row.observacao || ""
      };
    }
    function tripKey(row) {
      const clean = cleanRow(row);
      const date = toDateInput(clean.data);
      const driver = clean.motorista.trim().toLowerCase();
      return date && driver ? `${date}||${driver}` : "";
    }
    function recalculateTrips() {
      const groups = new Map();
      rows.forEach((row, index) => {
        const key = tripKey(row);
        if (!key) return;
        const group = groups.get(key) || [];
        group.push({ index, row: cleanRow(row) });
        groups.set(key, group);
      });
      const sequenceByIndex = new Map();
      groups.forEach((group) => {
        group
          .sort((a, b) => arrivalSortValue(a.row) - arrivalSortValue(b.row) || a.index - b.index)
          .forEach((item, idx) => sequenceByIndex.set(item.index, idx + 1));
      });
      rows = rows.map((row, index) => {
        const clean = cleanRow(row);
        clean.viagens = sequenceByIndex.has(index) ? String(sequenceByIndex.get(index)) : "";
        return clean;
      });
    }
    function updateVisibleTrips() {
      document.querySelectorAll("#rows tr").forEach((tr) => {
        const index = Number(tr.dataset.index);
        const tripsInput = tr.querySelector('[data-key="viagens"]');
        if (tripsInput) tripsInput.value = cleanRow(rows[index]).viagens;
      });
    }
    function optionList(values, current) {
      return values.map((value) => `<option value="${escapeAttr(value)}" ${value === current ? "selected" : ""}>${escapeHtml(value || "-")}</option>`).join("");
    }
    function conductorOptions(current) {
      const conductorNames = conductors.map((item) => typeof item === "string" ? item : (item.nome || item.motorista || ""));
      const values = [...new Set(["", ...conductorNames, current].map((value) => String(value || "").trim()).filter((value, idx) => idx === 0 || value))];
      return values.map((value) => `<option value="${escapeAttr(value)}" ${value === current ? "selected" : ""}>${escapeHtml(value || "-")}</option>`).join("");
    }
    function driverFilterKey(value) {
      return String(value || "").trim().toUpperCase();
    }
    function availableDriverNames() {
      return [...new Set([
        ...rows.map((row) => cleanRow(row).motorista),
        ...conductors.map((item) => typeof item === "string" ? item : (item.nome || item.motorista || ""))
      ].map((value) => String(value || "").trim()).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, "pt-BR"));
    }
    function updateDriverFilterOptions() {
      const names = availableDriverNames();
      $("driverFilterOptions").innerHTML = names.length ? names.map((name) => {
        const key = driverFilterKey(name);
        return `<label class="multi-filter-option"><input type="checkbox" value="${escapeAttr(key)}" ${selectedDriverFilters.has(key) ? "checked" : ""}><span>${escapeHtml(name)}</span></label>`;
      }).join("") : '<div class="multi-filter-option">Sem motoristas</div>';
      $("driverFilterLabel").textContent = selectedDriverFilters.size ? `${selectedDriverFilters.size} motorista${selectedDriverFilters.size > 1 ? "s" : ""}` : "Motoristas";
    }
    function statusClass(status) {
      const normalized = String(status).toLowerCase();
      if (normalized.includes("finalizado")) return "status-finalizado";
      if (normalized.includes("fila")) return "status-fila";
      if (normalized.includes("patio")) return "status-patio";
      return "";
    }
    function showCtNotice(message, type = "warning") {
      const notice = $("ctNotice");
      notice.textContent = message;
      notice.className = `message ${type}`;
      notice.hidden = false;
      clearTimeout(showCtNotice.timer);
      showCtNotice.timer = setTimeout(() => {
        notice.hidden = true;
        notice.textContent = "";
      }, 4200);
    }
    function canMoveStatus(currentStatus, targetStatus) {
      if (currentStatus === targetStatus) return true;
      if (!targetStatus) return !currentStatus;
      const currentIndex = statusFlow.indexOf(currentStatus);
      const targetIndex = statusFlow.indexOf(targetStatus);
      if (targetIndex === 0) return currentIndex <= 0;
      return currentIndex === targetIndex - 1;
    }
    function blockedStatusMessage(currentStatus, targetStatus) {
      if (!targetStatus) return "O status nao pode ficar vazio. Siga o fluxo: chegada, entrada e saida.";
      if (requiredStepMessages[targetStatus]) return requiredStepMessages[targetStatus];
      if (statusFlow.includes(targetStatus)) return "Nao e possivel voltar etapas. Siga o fluxo: chegada, entrada e saida.";
      const current = statusStepLabels[currentStatus] || "a etapa atual";
      const target = statusStepLabels[targetStatus] || "a proxima etapa";
      return `Fluxo de status invalido. Depois de ${current}, use ${target}.`;
    }
    function applyStatusSideEffects(row, targetStatus) {
      if (targetStatus === "Patio") {
        return { ...row, status: targetStatus, entrada: row.entrada || nowDateTimeLocal() };
      }
      if (targetStatus === "Fila de Carregamento") {
        return { ...row, status: targetStatus, chegada: row.chegada || nowDateTimeLocal() };
      }
      if (targetStatus === "Finalizado") {
        return { ...row, status: targetStatus, saida: row.saida || nowDateTimeLocal(), notaFiscal: row.notaFiscal || "Impresso" };
      }
      return { ...row, status: targetStatus };
    }
    function editIcon() {
      if (editMode) {
        return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round">
          <path d="M20 6 9 17l-5 1 1-5L16 2z"></path>
          <path d="m15 3 6 6"></path>
        </svg>`;
      }
      return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round">
        <path d="M2 12s3.6-6 10-6 10 6 10 6-3.6 6-10 6S2 12 2 12Z"></path>
        <circle cx="8.5" cy="12" r="2.5"></circle>
        <circle cx="15.5" cy="12" r="2.5"></circle>
        <path d="M11 12h2"></path>
        <path d="m17 19 4-4 1 1-4 4-2 .5z"></path>
      </svg>`;
    }
    function updateEditModeButton() {
      const button = $("editModeToggle");
      if (!button) return;
      button.classList.toggle("is-editing", editMode);
      button.title = editMode ? "Visualizar" : "Alternar modo de edicao";
      button.setAttribute("aria-label", editMode ? "Visualizar" : "Editar");
      button.querySelector(".edit-icon").innerHTML = editIcon();
      button.querySelector(".edit-label").textContent = editMode ? "Salvar" : "Editar";
      document.body.classList.toggle("editing-ct", editMode);
    }
    function visibleRows() {
      const query = $("searchFilter").value.trim().toLowerCase();
      const status = $("statusFilter").value;
      const freight = $("freightFilter").value;
      return rows.map((row, index) => ({ row: cleanRow(row), index })).filter(({ row }) => {
        return (!query || row.motorista.toLowerCase().includes(query))
          && (!selectedDriverFilters.size || selectedDriverFilters.has(driverFilterKey(row.motorista)))
          && (!status || row.status === status)
          && (!freight || row.tipoFrete === freight)
          && matchesDateRange(row);
      }).sort((a, b) => {
        const aDone = a.row.status === "Finalizado";
        const bDone = b.row.status === "Finalizado";
        if (aDone !== bDone) return aDone ? 1 : -1;
        if (aDone && bDone) return a.index - b.index;
        return arrivalSortValue(b.row) - arrivalSortValue(a.row);
      });
    }
    function arrivalSortValue(row) {
      const value = toDateTimeInput(row.chegada);
      const time = value ? new Date(value).getTime() : 0;
      return Number.isFinite(time) && time > 0 ? time : 0;
    }
    function renderCounters() {
      const active = visibleRows().map(({ row }) => row);
      const count = (freight, status) => active.filter((row) => row.tipoFrete === freight && row.status === status).length;
      $("cifBase").textContent = count("CIF", "Patio");
      $("cifFila").textContent = count("CIF", "Fila de Carregamento");
      $("fobBase").textContent = count("FOB", "Patio");
      $("fobFila").textContent = count("FOB", "Fila de Carregamento");
      $("finalizados").textContent = active.filter((row) => row.status === "Finalizado").length;
      $("patio").textContent = active.filter((row) => row.status === "Patio").length;
    }
    function render() {
      recalculateTrips();
      const data = visibleRows();
      const disabled = editMode ? "" : " disabled";
      $("rows").innerHTML = data.map(({ row, index }) => `
        <tr data-index="${index}">
          <td><input type="checkbox" aria-label="Selecionar linha"></td>
          <td><input class="cell" type="date" data-key="data" value="${escapeAttr(toDateInput(row.data))}"${disabled}></td>
          <td><select class="cell" data-key="motorista"${disabled}>${conductorOptions(row.motorista)}</select></td>
          <td><select class="cell" data-key="tipoFrete"${disabled}>${optionList(freights, row.tipoFrete)}</select></td>
          <td><select class="cell ${statusClass(row.status)}" data-key="status"${disabled}>${optionList(statuses, row.status)}</select></td>
          <td><input class="cell" data-key="viagens" value="${escapeAttr(row.viagens)}" readonly${disabled}></td>
          <td><input class="cell" type="datetime-local" data-key="chegada" value="${escapeAttr(toDateTimeInput(row.chegada))}"${disabled}></td>
          <td><input class="cell" type="datetime-local" data-key="entrada" value="${escapeAttr(toDateTimeInput(row.entrada))}"${disabled}></td>
          <td><input class="cell" type="datetime-local" data-key="saida" value="${escapeAttr(toDateTimeInput(row.saida))}"${disabled}></td>
          <td><select class="cell" data-key="notaFiscal"${disabled}>${optionList(invoices, row.notaFiscal)}</select></td>
          <td><input class="cell" data-key="observacao" value="${escapeAttr(row.observacao)}"${disabled}></td>
        </tr>
      `).join("");
      renderCounters();
      updateDriverFilterOptions();
      updateEditModeButton();
    }
    function syncFromTableIfReady() {
      if (!$("rows").children.length) return;
      const visible = [...document.querySelectorAll("#rows tr")];
      visible.forEach((tr) => {
        const index = Number(tr.dataset.index);
        rows[index] = cleanRow({
          data: tr.querySelector('[data-key="data"]').value.trim(),
          motorista: tr.querySelector('[data-key="motorista"]').value.trim(),
          tipoFrete: tr.querySelector('[data-key="tipoFrete"]').value,
          status: tr.querySelector('[data-key="status"]').value,
          viagens: tr.querySelector('[data-key="viagens"]').value.trim(),
          chegada: tr.querySelector('[data-key="chegada"]').value.trim(),
          entrada: tr.querySelector('[data-key="entrada"]').value.trim(),
          saida: tr.querySelector('[data-key="saida"]').value.trim(),
          notaFiscal: tr.querySelector('[data-key="notaFiscal"]').value,
          observacao: tr.querySelector('[data-key="observacao"]').value.trim()
        });
      });
    }
    function selectedIndexes() {
      syncFromTableIfReady();
      recalculateTrips();
      return [...document.querySelectorAll("#rows tr")]
        .filter((tr) => tr.querySelector('input[type="checkbox"]').checked)
        .map((tr) => Number(tr.dataset.index));
    }
    function updateSelected(updater) {
      const indexes = selectedIndexes();
      if (!indexes.length) {
        showCtNotice("Selecione ao menos uma linha para atualizar o status.");
        return;
      }
      indexes.forEach((index) => rows[index] = cleanRow(updater(rows[index])));
      render();
    }
    function moveSelectedToStatus(targetStatus) {
      const indexes = selectedIndexes();
      if (!indexes.length) {
        showCtNotice("Selecione ao menos uma linha para atualizar o status.");
        return false;
      }
      const blocked = indexes.map((index) => cleanRow(rows[index])).find((row) => !canMoveStatus(row.status, targetStatus));
      if (blocked) {
        showCtNotice(blockedStatusMessage(blocked.status, targetStatus));
        return false;
      }
      indexes.forEach((index) => rows[index] = cleanRow(applyStatusSideEffects(cleanRow(rows[index]), targetStatus)));
      render();
      return true;
    }
    function moveSelectedToStatusAndSave(targetStatus) {
      if (moveSelectedToStatus(targetStatus)) {
        $("ctForm").requestSubmit();
      }
    }
    $("addArrival").addEventListener("click", () => {
      syncFromTableIfReady();
      rows.unshift(cleanRow({
        data: todayDateLocal(),
        status: "Fila de Carregamento",
        chegada: nowDateTimeLocal()
      }));
      $("dateRangeFilter").value = "today";
      $("dateFromFilter").value = "";
      $("dateToFilter").value = "";
      $("searchFilter").value = "";
      $("statusFilter").value = "";
      $("freightFilter").value = "";
      updateCustomDateFilter();
      render();
      document.querySelector('[data-key="motorista"]')?.focus();
    });
    $("editModeToggle").addEventListener("click", () => {
      if (editMode) {
        syncFromTableIfReady();
        $("ctForm").requestSubmit();
        return;
      }
      editMode = true;
      render();
    });
    $("markEntry").addEventListener("click", () => moveSelectedToStatusAndSave("Patio"));
    $("markExit").addEventListener("click", () => moveSelectedToStatusAndSave("Finalizado"));
    $("deleteRows").addEventListener("click", () => {
      const remove = new Set(selectedIndexes());
      if (!remove.size) {
        showCtNotice("Selecione ao menos uma linha para excluir.");
        return;
      }
      rows = rows.filter((_, index) => !remove.has(index));
      render();
      $("ctForm").requestSubmit();
    });
    $("driverFilterOptions").addEventListener("change", (event) => {
      if (event.target?.type !== "checkbox") return;
      if (event.target.checked) {
        selectedDriverFilters.add(event.target.value);
      } else {
        selectedDriverFilters.delete(event.target.value);
      }
      render();
    });
    $("clearDriverFilter").addEventListener("click", () => {
      selectedDriverFilters.clear();
      render();
    });
    $("selectAll").addEventListener("change", (event) => {
      document.querySelectorAll("#rows input[type='checkbox']").forEach((input) => input.checked = event.target.checked);
    });
    $("rows").addEventListener("input", () => {
      syncFromTableIfReady();
      recalculateTrips();
      updateVisibleTrips();
      renderCounters();
    });
    $("rows").addEventListener("change", (event) => {
      if (event.target?.dataset?.key === "motorista") {
        const tr = event.target.closest("tr");
        const freight = conductorFreights.get(event.target.value.trim()) || "";
        const freightSelect = tr?.querySelector('[data-key="tipoFrete"]');
        if (freight && freightSelect) freightSelect.value = freight;
      }
      if (event.target?.dataset?.key === "status") {
        const tr = event.target.closest("tr");
        const index = Number(tr?.dataset?.index);
        const currentRow = cleanRow(rows[index]);
        const targetStatus = event.target.value;
        if (!canMoveStatus(currentRow.status, targetStatus)) {
          event.target.value = currentRow.status;
          showCtNotice(blockedStatusMessage(currentRow.status, targetStatus));
        } else {
          const nextRow = applyStatusSideEffects(currentRow, targetStatus);
          tr.querySelector('[data-key="entrada"]').value = toDateTimeInput(nextRow.entrada);
          tr.querySelector('[data-key="saida"]').value = toDateTimeInput(nextRow.saida);
          tr.querySelector('[data-key="notaFiscal"]').value = nextRow.notaFiscal;
        }
      }
      syncFromTableIfReady();
      recalculateTrips();
      updateVisibleTrips();
      renderCounters();
    });
    ["dateRangeFilter", "dateFromFilter", "dateToFilter", "searchFilter", "statusFilter", "freightFilter"].forEach((id) => $(id).addEventListener("input", () => {
      updateCustomDateFilter();
      render();
    }));
    $("ctForm").addEventListener("submit", () => {
      syncFromTableIfReady();
      recalculateTrips();
      $("rowsJson").value = JSON.stringify(rows);
    });
    $("dateRangeFilter").value = "today";
    updateCustomDateFilter();
    render();
  </script>
</body>
</html>
"""


DAILY_REPORT_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="icon" href="{favicon_url}" type="image/svg+xml">
  <title>Relatorio Diario - Dashboard</title>
  <style>
    :root {
      --bg: #f2f5f8;
      --top: #34104f;
      --ink: #16212d;
      --muted: #657282;
      --line: #d7e0e8;
      --panel: #ffffff;
      --teal: #64248c;
      --blue: #2b84cb;
      --red: #e2263c;
      --navy: #1b255f;
      --shadow: 0 18px 42px rgba(23, 32, 51, .10);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font-family: Inter, Segoe UI, Roboto, Arial, sans-serif;
    }
    a { color: inherit; text-decoration: none; }
    header {
      position: relative;
      overflow: hidden;
      padding: 24px clamp(16px, 4vw, 42px) 28px;
      background:
        radial-gradient(720px circle at 76% 35%, rgba(43,132,203,.34), transparent 62%),
        linear-gradient(135deg, #34104f 0%, #4c176d 58%, #1b255f 100%);
      color: #fff;
    }
    header::after {
      content: "";
      position: absolute;
      right: clamp(20px, 6vw, 76px);
      bottom: -88px;
      width: min(46vw, 520px);
      aspect-ratio: 1.8;
      background: url("{favicon_url}") center / contain no-repeat;
      opacity: .18;
      pointer-events: none;
    }
    .topbar {
      position: relative;
      z-index: 2;
      display: flex;
      justify-content: space-between;
      gap: 18px;
      align-items: flex-start;
    }
    .brand-title { display: flex; align-items: center; gap: 16px; }
    .brand-title img { width: 98px; height: auto; filter: drop-shadow(0 10px 18px rgba(0,0,0,.24)); }
    h1 { margin: 0; font-size: clamp(28px, 4vw, 50px); line-height: 1; letter-spacing: 0; }
    .subtitle { margin: 9px 0 0; color: #d7e4ea; font-size: 15px; }
    .nav { display: flex; flex-wrap: wrap; justify-content: flex-end; gap: 9px; }
    .top-link, button {
      min-height: 38px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 9px 12px;
      border: 1px solid rgba(255,255,255,.30);
      border-radius: 8px;
      background: rgba(255,255,255,.10);
      color: #fff;
      font: inherit;
      font-size: 13px;
      font-weight: 900;
      cursor: pointer;
    }
    button.is-loading,
    button:disabled {
      cursor: wait;
      opacity: .78;
    }
    button.is-loading::after {
      content: "";
      width: 13px;
      height: 13px;
      margin-left: 8px;
      border: 2px solid rgba(255,255,255,.45);
      border-top-color: #fff;
      border-radius: 50%;
      animation: spin .75s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    button.secondary {
      background: #ffffff;
      border-color: #ffffff;
      color: var(--navy);
    }
    .hero-grid {
      position: relative;
      z-index: 2;
      display: grid;
      grid-template-columns: minmax(260px, 1fr) auto;
      gap: 18px;
      align-items: end;
      margin-top: 34px;
    }
    .report-date { font-size: clamp(24px, 3vw, 40px); font-weight: 950; }
    .filters {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      justify-content: flex-end;
      align-items: end;
    }
    .custom-date-filter {
      display: none;
      gap: 10px;
      align-items: end;
    }
    .custom-date-filter.is-visible { display: flex; }
    label { display: grid; gap: 7px; color: #d7e4ea; font-size: 12px; font-weight: 900; text-transform: uppercase; }
    select, input {
      min-height: 42px;
      min-width: 190px;
      border: 1px solid rgba(255,255,255,.44);
      border-radius: 8px;
      padding: 9px 10px;
      background: rgba(255,255,255,.96);
      color: var(--ink);
      font: inherit;
      font-weight: 800;
    }
    textarea {
      width: 100%;
      min-height: 72px;
      resize: vertical;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 9px 10px;
      color: var(--ink);
      background: #fff;
      font: inherit;
      font-size: 13px;
      line-height: 1.35;
    }
    textarea:focus, select:focus, input:focus {
      outline: 2px solid rgba(43,132,203,.38);
      outline-offset: 1px;
    }
    main { padding: 22px clamp(16px, 4vw, 42px) 40px; }
    .kpis {
      display: grid;
      grid-template-columns: repeat(4, minmax(150px, 1fr));
      gap: 12px;
    }
    .kpi, .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }
    .kpi {
      min-height: 112px;
      padding: 18px;
      border-top: 5px solid var(--teal);
    }
    .kpi:nth-child(2) { border-top-color: var(--blue); }
    .kpi:nth-child(3) { border-top-color: var(--red); }
    .kpi:nth-child(4) { border-top-color: var(--navy); }
    .kpi span { color: var(--muted); font-size: 12px; font-weight: 900; text-transform: uppercase; }
    .kpi strong { display: block; margin-top: 10px; font-size: clamp(25px, 3vw, 34px); line-height: 1; }
    .grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 14px;
      margin-top: 14px;
    }
    .panel { overflow: hidden; }
    .panel h2 {
      margin: 0;
      padding: 16px 18px 0;
      font-size: 18px;
    }
    .panel h2 .panel-count {
      margin-left: 8px;
      color: var(--muted);
      font-size: 13px;
      font-weight: 850;
    }
    .panel-body { padding: 14px 18px 18px; }
    .bars { display: grid; gap: 10px; }
    .bar-row {
      display: grid;
      grid-template-columns: minmax(100px, 150px) 1fr auto;
      gap: 10px;
      align-items: center;
      font-size: 13px;
      font-weight: 850;
    }
    .track { height: 10px; border-radius: 999px; background: #edf2f6; overflow: hidden; }
    .fill { height: 100%; border-radius: inherit; background: linear-gradient(90deg, var(--teal), var(--blue)); }
    .terminal-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; }
    .terminal {
      padding: 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #f8fafb;
    }
    .terminal span { color: var(--muted); font-size: 12px; font-weight: 900; text-transform: uppercase; }
    .terminal strong { display: block; margin-top: 7px; font-size: 24px; }
    .wide { grid-column: 1 / -1; }
    .table-wrap { overflow: auto; max-height: calc(100vh - 420px); }
    table { width: 100%; border-collapse: collapse; min-width: 1280px; }
    th, td { padding: 11px 12px; border-bottom: 1px solid var(--line); text-align: left; vertical-align: top; }
    th {
      position: sticky;
      top: 0;
      z-index: 2;
      background: #eef3f6;
      color: #506071;
      font-size: 12px;
      text-transform: uppercase;
    }
    td.num, th.num { text-align: right; }
    tr.needs-note td { background: #fffaf0; }
    .observation-cell { min-width: 260px; }
    .observation-cell textarea {
      min-height: 56px;
      font-size: 12px;
    }
    .observation-dash {
      color: var(--muted);
      font-weight: 850;
    }
    .pill {
      display: inline-flex;
      min-height: 28px;
      align-items: center;
      padding: 5px 9px;
      border-radius: 999px;
      background: #edf2ff;
      color: var(--navy);
      font-weight: 950;
      white-space: nowrap;
    }
    .save-state {
      min-height: 18px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
      text-align: right;
    }
    .empty { padding: 26px; color: var(--muted); font-weight: 800; text-align: center; }
    .share-panel {
      margin-top: 14px;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      overflow: hidden;
    }
    .share-head {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      padding: 16px 18px;
      border-bottom: 1px solid var(--line);
    }
    .share-head h2 { margin: 0; font-size: 18px; }
    .share-actions { display: flex; flex-wrap: wrap; gap: 8px; justify-content: flex-end; }
    .share-actions button {
      background: var(--teal);
      border-color: transparent;
    }
    .share-actions button.secondary {
      background: #f3f6f8;
      border-color: var(--line);
      color: var(--ink);
    }
    .canvas-wrap {
      padding: 18px;
      background:
        linear-gradient(90deg, rgba(52,16,79,.06), rgba(43,132,203,.07)),
        #f8fafb;
      overflow: auto;
    }
    #shareCanvas {
      display: block;
      width: min(100%, 720px);
      height: auto;
      margin: 0 auto;
      border-radius: 8px;
      box-shadow: 0 18px 42px rgba(23, 32, 51, .18);
      background: #fff;
    }
    @media (max-width: 900px) {
      .topbar, .hero-grid { grid-template-columns: 1fr; flex-direction: column; }
      .filters { justify-content: flex-start; }
      .kpis, .grid { grid-template-columns: 1fr; }
      .custom-date-filter { width: 100%; }
      .wide { grid-column: auto; }
    }
    @media print {
      body { background: #fff; }
      header { padding: 18px 22px; }
      .nav, .filters button { display: none; }
      main { padding: 18px 22px; }
      .kpi, .panel { box-shadow: none; break-inside: avoid; }
      .table-wrap { max-height: none; overflow: visible; }
      th { position: static; }
    }
  </style>
</head>
<body>
  <header>
    <div class="topbar">
      <div>
        <div class="brand-title"><img src="{favicon_url}" alt=""><h1>Relatorio Diario</h1></div>
        <p class="subtitle">Viagens, capacidade e volume carregado por placa e terminal.</p>
      </div>
      <nav class="nav">
        <a class="top-link" href="/home">Home</a>
        <a class="top-link" href="/dashboard">Dashboard</a>
        <a class="top-link" href="/controle-ct">Controle de CT</a>
        <a class="top-link" href="/editar">Editar dados</a>
        <a class="top-link" href="/capacidades">Capacidades</a>
        <a class="top-link" href="/relatorio-entrada-notas">Entrada de notas</a>
        <a class="top-link" href="/logout">Sair</a>
      </nav>
    </div>
    <div class="hero-grid">
      <div class="report-date" id="reportDate">-</div>
      <div class="filters">
        <label>Filtro
          <select id="dateModeSelect">
            <option value="selected">Data selecionada</option>
            <option value="today">Hoje</option>
            <option value="yesterday">Ontem</option>
            <option value="week">Esta semana</option>
            <option value="last7">Ultimos 7 dias</option>
            <option value="month">Este mes</option>
            <option value="custom">Periodo</option>
          </select>
        </label>
        <label>Data
          <select id="dateSelect"></select>
        </label>
        <span class="custom-date-filter" id="customDateFilter">
          <label>Inicio
            <input id="dateStart" type="date">
          </label>
          <label>Fim
            <input id="dateEnd" type="date">
          </label>
        </span>
        <label>Terminal
          <select id="terminalSelect">
            <option value="">Todos</option>
            <option value="10">Equador</option>
            <option value="19">Ipiranga</option>
          </select>
        </label>
        <label>Municipio
          <select id="municipioSelect">
            <option value="">Todos</option>
          </select>
        </label>
        <button type="button" id="generateImage">Gerar imagem</button>
      </div>
    </div>
  </header>
  <main>
    <section class="kpis">
      <div class="kpi"><span>Viagens</span><strong id="kTrips">0</strong></div>
      <div class="kpi"><span>Volume</span><strong id="kVolume">0</strong></div>
      <div class="kpi"><span>Placas</span><strong id="kPlates">0</strong></div>
      <div class="kpi"><span>Notas fiscais</span><strong id="kNotes">0</strong></div>
    </section>
    <section class="grid">
      <div class="panel">
        <h2>Terminais</h2>
        <div class="panel-body terminal-grid" id="terminalSummary"></div>
      </div>
      <div class="panel wide">
        <h2>Detalhamento das viagens <span class="save-state" id="observationSaveState"></span></h2>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Data</th>
                <th>Placa</th>
                <th>Motorista</th>
                <th>Terminal</th>
                <th>Municipio destino</th>
                <th class="num">Viagens</th>
                <th class="num">Capacidade</th>
                <th class="num">Volume</th>
                <th class="num">Notas</th>
                <th>Observacao</th>
              </tr>
            </thead>
            <tbody id="reportRows"></tbody>
          </table>
        </div>
      </div>
    </section>
    <section class="share-panel">
      <div class="share-head">
        <h2>Imagem para WhatsApp</h2>
        <div class="share-actions">
          <button type="button" id="shareImage">Compartilhar</button>
          <button type="button" id="downloadImage" class="secondary">Baixar PNG</button>
        </div>
      </div>
      <div class="canvas-wrap">
        <canvas id="shareCanvas" width="1080" height="1600"></canvas>
      </div>
    </section>
  </main>
  <script>
    const dataset = __DATA__;
    const rows = dataset.dailyPlateRows || [];
    const observations = dataset.dailyObservations || {};
    const fmt = new Intl.NumberFormat("pt-BR");
    const volume = (value) => `${fmt.format(Math.round(value / 1000))} mil`;
    const $ = (id) => document.getElementById(id);
    const logoUrl = "{favicon_url}";
    let saveObservationTimer = null;

    async function withButtonLoading(button, text, action) {
      if (!button) return action();
      const original = button.textContent;
      button.textContent = text;
      button.classList.add("is-loading");
      button.disabled = true;
      try {
        return await action();
      } finally {
        button.textContent = original;
        button.classList.remove("is-loading");
        button.disabled = false;
      }
    }

    function parseDate(value) {
      const [day, month, year] = String(value).split("/");
      return new Date(Number(year), Number(month) - 1, Number(day));
    }

    function inputDate(value) {
      if (!value) return "";
      const date = parseDate(value);
      return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
    }

    function fromInputDate(value) {
      if (!value) return null;
      const [year, month, day] = value.split("-").map(Number);
      return new Date(year, month - 1, day);
    }

    function formatDate(date) {
      return `${String(date.getDate()).padStart(2, "0")}/${String(date.getMonth() + 1).padStart(2, "0")}/${date.getFullYear()}`;
    }

    function addDays(date, days) {
      const next = new Date(date);
      next.setDate(next.getDate() + days);
      return next;
    }

    function escapeHtml(value) {
      return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }

    function escapeAttr(value) {
      return escapeHtml(value).replace(/"/g, "&quot;");
    }

    function dateLabel(value) {
      return parseDate(value).toLocaleDateString("pt-BR", {
        weekday: "long",
        day: "2-digit",
        month: "long",
        year: "numeric"
      });
    }

    function uniqueDates() {
      return [...new Set(rows.map((row) => row.data))]
        .sort((a, b) => parseDate(b) - parseDate(a));
    }

    function selectedDateRange() {
      const mode = $("dateModeSelect").value;
      const selected = $("dateSelect").value;
      const latest = selected ? parseDate(selected) : new Date();
      if (mode === "selected") return { start: selected ? parseDate(selected) : null, end: selected ? parseDate(selected) : null };
      if (mode === "today") {
        const today = new Date();
        return { start: new Date(today.getFullYear(), today.getMonth(), today.getDate()), end: new Date(today.getFullYear(), today.getMonth(), today.getDate()) };
      }
      if (mode === "yesterday") {
        const yesterday = addDays(new Date(), -1);
        const day = new Date(yesterday.getFullYear(), yesterday.getMonth(), yesterday.getDate());
        return { start: day, end: day };
      }
      if (mode === "last7") return { start: addDays(latest, -6), end: latest };
      if (mode === "week") {
        const mondayOffset = (latest.getDay() + 6) % 7;
        return { start: addDays(latest, -mondayOffset), end: addDays(latest, 6 - mondayOffset) };
      }
      if (mode === "month") return { start: new Date(latest.getFullYear(), latest.getMonth(), 1), end: new Date(latest.getFullYear(), latest.getMonth() + 1, 0) };
      if (mode === "custom") return { start: fromInputDate($("dateStart").value), end: fromInputDate($("dateEnd").value) };
      return { start: null, end: null };
    }

    function matchesDateRange(row) {
      const range = selectedDateRange();
      const date = parseDate(row.data);
      return (!range.start || date >= range.start) && (!range.end || date <= range.end);
    }

    function updateDateInputs() {
      const custom = $("dateModeSelect").value === "custom";
      $("customDateFilter").classList.toggle("is-visible", custom);
      $("dateSelect").disabled = !["selected", "week", "last7", "month"].includes($("dateModeSelect").value);
    }

    function reportDateLabel() {
      const mode = $("dateModeSelect").value;
      const range = selectedDateRange();
      if (mode === "selected") return $("dateSelect").value ? dateLabel($("dateSelect").value) : "-";
      if (range.start && range.end && formatDate(range.start) === formatDate(range.end)) return dateLabel(formatDate(range.start));
      if (range.start && range.end) return `${formatDate(range.start)} ate ${formatDate(range.end)}`;
      if (range.start) return `A partir de ${formatDate(range.start)}`;
      if (range.end) return `Ate ${formatDate(range.end)}`;
      return "Todas as datas";
    }

    function filteredRows() {
      const terminal = $("terminalSelect").value;
      const municipio = $("municipioSelect").value;
      return rows.filter((row) => {
        const municipios = String(row.municipioDestino || "").split("/").map((name) => name.trim()).filter(Boolean);
        return matchesDateRange(row)
          && (!terminal || row.terminal === terminal)
          && (!municipio || municipios.includes(municipio));
      });
    }

    function groupedRowsByPlate(data) {
      const terminalOrder = { Equador: 1, Ipiranga: 2 };
      const groups = new Map();
      const terminalFilter = $("terminalSelect").value;
      data.forEach((row) => {
        const terminalKey = terminalFilter ? row.terminal : "todos";
        const driverKey = String(row.motorista || "").trim().toUpperCase();
        const key = `${row.data}||${row.placa}||${driverKey}||${terminalKey}`;
        const current = groups.get(key) || {
          ...row,
          terminal: terminalKey,
          terminalNome: "",
          terminalShort: "",
          viagens: 0,
          viagensOrdens: 0,
          viagensCarga: 0,
          capacidade: 0,
          quantidade: 0,
          notas: 0,
          clientes: 0,
          motoristas: [],
          motorista: "",
          municipios: [],
          produtos: [],
          mixProdutos: ""
        };
        current.viagens += Number(row.viagens) || 0;
        current.viagensOrdens += Number(row.viagensOrdens) || 0;
        current.viagensCarga += Number(row.viagensCarga) || 0;
        current.capacidade = Math.max(current.capacidade, Number(row.capacidade) || 0);
        current.quantidade += Number(row.quantidade) || 0;
        current.notas += Number(row.notas) || 0;
        current.clientes += Number(row.clientes) || 0;
        if (row.motorista) current.motoristas.push(row.motorista);
        if (row.municipioDestino) current.municipios.push(row.municipioDestino);
        current.produtos.push(...(row.produtos || []));
        current._terminals = current._terminals || new Map();
        current._terminals.set(row.terminalNome, row.terminal);
        groups.set(key, current);
      });
      return [...groups.values()].map((row) => {
        const terminals = [...row._terminals.entries()]
          .sort((a, b) => (terminalOrder[a[0]] || 99) - (terminalOrder[b[0]] || 99));
        const productTotals = new Map();
        row.produtos.forEach((item) => {
          const name = item.produto || "Sem produto detalhado";
          productTotals.set(name, (productTotals.get(name) || 0) + (Number(item.quantidade) || 0));
        });
        const products = [...productTotals.entries()]
          .sort((a, b) => b[1] - a[1])
          .map(([produto, quantidade]) => ({ produto, quantidade }));
        const motoristas = [...new Set(row.motoristas.flatMap((name) => String(name).split("/")).map((name) => name.trim()).filter(Boolean))]
          .sort((a, b) => a.localeCompare(b, "pt-BR"));
        const municipios = [...new Set(row.municipios.map((name) => String(name).trim()).filter(Boolean))]
          .sort((a, b) => a.localeCompare(b, "pt-BR"));
        return {
          ...row,
          motorista: motoristas.join(" / "),
          municipioDestino: municipios.join(" / ") || row.municipioDestino || "-",
          terminal: terminals.map(([, code]) => code).join("/"),
          terminalNome: terminals.map(([name]) => name).join("/"),
          terminalShort: terminals.map(([name]) => name.slice(0, 3)).join("/"),
          produtos: products,
          mixProdutos: products.slice(0, 4).map((item) => `${item.produto} (${fmt.format(Math.round(item.quantidade / 1000))}k)`).join(", ") || "Sem produto detalhado",
          _terminals: undefined
        };
      });
    }

    function renderTerminals(data) {
      const totals = new Map();
      data.forEach((row) => {
        const current = totals.get(row.terminalNome) || { viagens: 0, quantidade: 0 };
        current.viagens += row.viagens;
        current.quantidade += row.quantidade;
        totals.set(row.terminalNome, current);
      });
      const terminalRows = ["Equador", "Ipiranga"].map((name) => [name, totals.get(name) || { viagens: 0, quantidade: 0 }]);
      $("terminalSummary").innerHTML = terminalRows.map(([name, item]) => `
        <div class="terminal">
          <span>${name}</span>
          <strong>${fmt.format(item.viagens)}</strong>
          <div>${volume(item.quantidade)} carregados</div>
        </div>
      `).join("");
    }

    function renderTable(data) {
      const detailData = groupedRowsByPlate(data);
      if (!detailData.length) {
        $("reportRows").innerHTML = `<tr><td colspan="10" class="empty">Sem dados para o filtro.</td></tr>`;
        return;
      }
      $("reportRows").innerHTML = detailData
        .slice()
        .sort((a, b) => parseDate(b.data) - parseDate(a.data) || b.viagens - a.viagens || b.quantidade - a.quantidade || a.placa.localeCompare(b.placa))
        .map((row) => {
          const key = observationKey(row);
          const needsNote = row.viagens < 2;
          const note = observations[key] || "";
          return `
          <tr class="${needsNote && !note ? "needs-note" : ""}">
            <td>${escapeHtml(row.data)}</td>
            <td><span class="pill">${escapeHtml(row.placa)}</span></td>
            <td>${escapeHtml(row.motorista || "-")}</td>
            <td>${escapeHtml(row.terminalNome)}</td>
            <td>${escapeHtml(row.municipioDestino || "-")}</td>
            <td class="num">${fmt.format(row.viagens)}</td>
            <td class="num">${volume(row.capacidade)}</td>
            <td class="num">${volume(row.quantidade)}</td>
            <td class="num">${fmt.format(row.notas)}</td>
            <td class="observation-cell">${
              needsNote
                ? `<textarea data-observation-key="${escapeAttr(key)}" placeholder="Informe o motivo">${escapeHtml(note)}</textarea>`
                : '<span class="observation-dash">-</span>'
            }</td>
          </tr>
        `;
        }).join("");
    }

    function observationKey(row) {
      return [row.data, row.placa, row.motorista || "", row.terminal || "todos"].join("||");
    }

    async function saveObservations() {
      $("observationSaveState").textContent = "Salvando observacoes...";
      const response = await fetch("/relatorio-diario/observacoes", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ observations })
      });
      if (!response.ok) throw new Error("Falha ao salvar");
      $("observationSaveState").textContent = "Observacoes salvas.";
      window.setTimeout(() => {
        if ($("observationSaveState").textContent === "Observacoes salvas.") $("observationSaveState").textContent = "";
      }, 2600);
    }

    function queueObservationSave() {
      clearTimeout(saveObservationTimer);
      saveObservationTimer = window.setTimeout(() => {
        saveObservations().catch(() => {
          $("observationSaveState").textContent = "Nao foi possivel salvar as observacoes.";
        });
      }, 650);
    }

    function wrapText(ctx, text, maxWidth) {
      const words = String(text || "").split(/\\s+/).filter(Boolean);
      const lines = [];
      let line = "";
      words.forEach((word) => {
        const test = line ? `${line} ${word}` : word;
        if (ctx.measureText(test).width > maxWidth && line) {
          lines.push(line);
          line = word;
        } else {
          line = test;
        }
      });
      if (line) lines.push(line);
      return lines.length ? lines : [""];
    }

    function roundRect(ctx, x, y, width, height, radius) {
      const r = Math.min(radius, width / 2, height / 2);
      ctx.beginPath();
      ctx.moveTo(x + r, y);
      ctx.arcTo(x + width, y, x + width, y + height, r);
      ctx.arcTo(x + width, y + height, x, y + height, r);
      ctx.arcTo(x, y + height, x, y, r);
      ctx.arcTo(x, y, x + width, y, r);
      ctx.closePath();
    }

    function drawCard(ctx, x, y, width, height, accent) {
      ctx.save();
      ctx.fillStyle = "#ffffff";
      roundRect(ctx, x, y, width, height, 18);
      ctx.fill();
      ctx.fillStyle = accent;
      roundRect(ctx, x, y, width, 8, 18);
      ctx.fill();
      ctx.restore();
    }

    function selectedReport() {
      const data = filteredRows()
        .slice()
        .sort((a, b) => b.viagens - a.viagens || b.quantidade - a.quantidade || a.placa.localeCompare(b.placa));
      const detailData = groupedRowsByPlate(data)
        .sort((a, b) => b.viagens - a.viagens || b.quantidade - a.quantidade || a.placa.localeCompare(b.placa));
      const trips = data.reduce((total, row) => total + row.viagens, 0);
      const qty = data.reduce((total, row) => total + row.quantidade, 0);
      const notes = data.reduce((total, row) => total + row.notas, 0);
      const plates = new Set(data.map((row) => row.placa)).size;
      return { data, detailData, trips, qty, notes, plates };
    }

    function drawBrand(ctx) {
      ctx.save();
      ctx.translate(58, 50);
      ctx.fillStyle = "#2b84cb";
      ctx.beginPath();
      ctx.moveTo(0, 20);
      ctx.quadraticCurveTo(36, 2, 82, 12);
      ctx.quadraticCurveTo(78, 64, 34, 84);
      ctx.quadraticCurveTo(8, 58, 0, 20);
      ctx.fill();
      ctx.fillStyle = "#e2263c";
      ctx.globalAlpha = .82;
      ctx.beginPath();
      ctx.moveTo(44, 16);
      ctx.quadraticCurveTo(88, 28, 76, 58);
      ctx.quadraticCurveTo(56, 76, 34, 84);
      ctx.closePath();
      ctx.fill();
      ctx.globalAlpha = 1;
      ctx.fillStyle = "#ffffff";
      ctx.font = "900 16px Arial";
      ctx.fillText("GRUPO", 104, 24);
      ctx.font = "900 22px Arial";
      ctx.fillText("DISLUB", 104, 50);
      ctx.fillText("EQUADOR", 104, 76);
      ctx.restore();
    }

    async function drawShareImage() {
      const report = selectedReport();
      const canvas = $("shareCanvas");
      canvas.width = 1080;
      let ctx = canvas.getContext("2d");
      ctx.font = "700 17px Arial";
      const observationWidth = 320;
      const rowMetrics = report.detailData.map((row) => {
        const driverLines = wrapText(ctx, row.motorista || "-", 210).slice(0, 2);
        const note = observations[observationKey(row)] || (row.viagens < 2 ? "Sem observacao informada" : "-");
        const noteLines = wrapText(ctx, note, observationWidth).slice(0, 3);
        return {
          row,
          driverLines,
          note,
          noteLines,
          height: Math.max(76, 42 + Math.max(driverLines.length, noteLines.length) * 21)
        };
      });
      const detailTop = 720;
      const tableTop = detailTop + 140;
      const rowsHeight = rowMetrics.reduce((total, item) => total + item.height, 0);
      canvas.height = Math.max(1240, tableTop + rowsHeight + 86);
      ctx = canvas.getContext("2d");

      ctx.fillStyle = "#f2f5f8";
      ctx.fillRect(0, 0, canvas.width, canvas.height);

      const gradient = ctx.createLinearGradient(0, 0, canvas.width, 330);
      gradient.addColorStop(0, "#34104f");
      gradient.addColorStop(.62, "#4c176d");
      gradient.addColorStop(1, "#1b255f");
      ctx.fillStyle = gradient;
      ctx.fillRect(0, 0, canvas.width, 360);

      ctx.globalAlpha = .18;
      ctx.fillStyle = "#2b84cb";
      ctx.beginPath();
      ctx.arc(910, 90, 230, 0, Math.PI * 2);
      ctx.fill();
      ctx.globalAlpha = 1;

      drawBrand(ctx);
      ctx.fillStyle = "#ffffff";
      ctx.font = "900 58px Arial";
      ctx.fillText("Relatorio Diario", 54, 200);
      ctx.font = "700 28px Arial";
      ctx.fillStyle = "#d7e4ea";
      ctx.fillText($("reportDate").textContent || "-", 58, 252);
      ctx.font = "800 24px Arial";
      const terminalLabel = $("terminalSelect").selectedOptions[0]?.textContent || "Todos";
      ctx.fillText(`Terminal: ${terminalLabel}`, 58, 298);
      const municipioLabel = $("municipioSelect").selectedOptions[0]?.textContent || "Todos";
      if ($("municipioSelect").value) {
        ctx.fillText(`Municipio: ${municipioLabel}`, 360, 298);
      }

      const kpiY = 302;
      const kpiW = 228;
      const kpis = [
        ["Viagens", fmt.format(report.trips), "#64248c"],
        ["Volume", volume(report.qty), "#2b84cb"],
        ["Placas", fmt.format(report.plates), "#e2263c"],
        ["Notas", fmt.format(report.notes), "#1b255f"]
      ];
      kpis.forEach(([label, value, accent], idx) => {
        const x = 54 + idx * (kpiW + 18);
        drawCard(ctx, x, kpiY, kpiW, 142, accent);
        ctx.fillStyle = "#657282";
        ctx.font = "900 22px Arial";
        ctx.fillText(label.toUpperCase(), x + 22, kpiY + 46);
        ctx.fillStyle = "#16212d";
        ctx.font = "900 38px Arial";
        ctx.fillText(value, x + 22, kpiY + 102);
      });

      let y = 492;
      drawCard(ctx, 54, y, 972, 188, "#2b84cb");
      ctx.fillStyle = "#16212d";
      ctx.font = "900 30px Arial";
      ctx.fillText("Resumo por terminal", 82, y + 52);
      const terminalTotals = new Map();
      report.data.forEach((row) => {
        const item = terminalTotals.get(row.terminalNome) || { viagens: 0, quantidade: 0 };
        item.viagens += row.viagens;
        item.quantidade += row.quantidade;
        terminalTotals.set(row.terminalNome, item);
      });
      [["Equador", 84], ["Ipiranga", 560]].forEach(([name, x]) => {
        const item = terminalTotals.get(name) || { viagens: 0, quantidade: 0 };
        ctx.fillStyle = "#657282";
        ctx.font = "900 20px Arial";
        ctx.fillText(name.toUpperCase(), x, y + 96);
        ctx.fillStyle = "#16212d";
        ctx.font = "900 34px Arial";
        ctx.fillText(`${fmt.format(item.viagens)} viagens`, x, y + 138);
        ctx.fillStyle = "#657282";
        ctx.font = "700 22px Arial";
        ctx.fillText(`${volume(item.quantidade)} carregados`, x, y + 168);
      });

      y += 228;
      drawCard(ctx, 54, y, 972, canvas.height - y - 54, "#e2263c");
      ctx.fillStyle = "#16212d";
      ctx.font = "900 30px Arial";
      ctx.fillText("Detalhamento", 82, y + 52);
      ctx.font = "900 18px Arial";
      ctx.fillStyle = "#657282";
      ctx.fillText("PLACA", 82, y + 94);
      ctx.fillText("MOTORISTA", 224, y + 94);
      ctx.fillText("VIAG.", 500, y + 94);
      ctx.fillText("VOL.", 584, y + 94);
      ctx.fillText("OBSERVACAO", 690, y + 94);
      ctx.strokeStyle = "#d7e0e8";
      ctx.beginPath();
      ctx.moveTo(82, y + 112);
      ctx.lineTo(998, y + 112);
      ctx.stroke();

      let tableY = y + 140;
      rowMetrics.forEach(({ row, driverLines, noteLines, height }, idx) => {
        if (idx % 2 === 0) {
          ctx.fillStyle = "#f8fafb";
          roundRect(ctx, 78, tableY, 916, height - 8, 8);
          ctx.fill();
        }
        ctx.fillStyle = "#16212d";
        ctx.font = "900 20px Arial";
        ctx.fillText(row.placa, 88, tableY + 28);
        ctx.font = "800 19px Arial";
        driverLines.forEach((line, lineIdx) => {
          ctx.fillText(line, 224, tableY + 20 + lineIdx * 20);
        });
        ctx.fillText(fmt.format(row.viagens), 516, tableY + 28);
        ctx.fillText(volume(row.quantidade).replace(" mil", "k"), 584, tableY + 28);
        ctx.fillStyle = "#657282";
        ctx.font = "800 16px Arial";
        ctx.fillText(row.terminalShort || row.terminalNome.slice(0, 3), 88, tableY + 52);
        ctx.font = "700 16px Arial";
        ctx.fillStyle = row.viagens < 2 ? "#92400e" : "#657282";
        noteLines.forEach((line, lineIdx) => {
          ctx.fillText(line, 690, tableY + 20 + lineIdx * 20);
        });
        tableY += height;
      });

      ctx.fillStyle = "#657282";
      ctx.font = "700 18px Arial";
      ctx.fillText("Dashboard Log - Grupo Dislub Equador", 54, canvas.height - 24);
    }

    function canvasToBlob(canvas) {
      return new Promise((resolve) => canvas.toBlob(resolve, "image/png", .95));
    }

    async function downloadImage() {
      await drawShareImage();
      const canvas = $("shareCanvas");
      const link = document.createElement("a");
      link.download = `relatorio-diario-${reportDateLabel().replaceAll("/", "-").replaceAll(" ", "-")}.png`;
      link.href = canvas.toDataURL("image/png");
      link.click();
    }

    async function shareImage() {
      await drawShareImage();
      const blob = await canvasToBlob($("shareCanvas"));
      if (!blob) return downloadImage();
      const file = new File([blob], `relatorio-diario-${reportDateLabel().replaceAll("/", "-").replaceAll(" ", "-")}.png`, { type: "image/png" });
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: "Relatorio diario" });
      } else {
        await downloadImage();
      }
    }

    function render() {
      const data = filteredRows();
      const trips = data.reduce((total, row) => total + row.viagens, 0);
      const qty = data.reduce((total, row) => total + row.quantidade, 0);
      const notes = data.reduce((total, row) => total + row.notas, 0);
      updateDateInputs();
      $("reportDate").textContent = reportDateLabel();
      $("kTrips").textContent = fmt.format(trips);
      $("kVolume").textContent = volume(qty);
      $("kPlates").textContent = fmt.format(new Set(data.map((row) => row.placa)).size);
      $("kNotes").textContent = fmt.format(notes);
      renderTerminals(data);
      renderTable(data);
      drawShareImage();
    }

    const dates = uniqueDates();
    $("dateSelect").innerHTML = dates.map((date) => `<option value="${date}">${date}</option>`).join("");
    $("dateSelect").value = dates[0] || "";
    $("dateStart").value = dates.length ? inputDate(dates[dates.length - 1]) : "";
    $("dateEnd").value = dates.length ? inputDate(dates[0]) : "";
    $("municipioSelect").innerHTML = ['<option value="">Todos</option>', ...[...new Set(rows.flatMap((row) => String(row.municipioDestino || "").split("/").map((name) => name.trim())).filter((name) => name && name !== "-"))].sort((a, b) => a.localeCompare(b, "pt-BR")).map((name) => `<option value="${escapeAttr(name)}">${escapeHtml(name)}</option>`)].join("");
    ["dateModeSelect", "dateSelect", "dateStart", "dateEnd"].forEach((id) => $(id).addEventListener("change", render));
    $("terminalSelect").addEventListener("change", render);
    $("municipioSelect").addEventListener("change", render);
    $("reportRows").addEventListener("input", (event) => {
      const key = event.target?.dataset?.observationKey;
      if (!key) return;
      observations[key] = event.target.value;
      event.target.closest("tr")?.classList.toggle("needs-note", !event.target.value.trim());
      queueObservationSave();
    });
    $("generateImage").addEventListener("click", (event) => withButtonLoading(event.currentTarget, "Gerando...", drawShareImage));
    $("downloadImage").addEventListener("click", (event) => withButtonLoading(event.currentTarget, "Baixando...", downloadImage));
    $("shareImage").addEventListener("click", (event) => withButtonLoading(event.currentTarget, "Compartilhando...", shareImage));
    render();
  </script>
</body>
</html>
"""


MEASUREMENT_CONTROL_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="icon" href="{favicon_url}" type="image/svg+xml">
  <title>Controle Medicao - Dashboard</title>
  <style>
    :root { --bg:#eef2f5; --ink:#16212d; --muted:#657282; --line:#d7e0e8; --panel:#fff; --panel-soft:#f8fafb; --purple:#64248c; --blue:#2b84cb; --red:#e2263c; --green:#00856f; --amber:#f59e0b; --shadow:0 18px 42px rgba(23,32,51,.10); }
    * { box-sizing:border-box; }
    body { margin:0; background:var(--bg); color:var(--ink); font-family:Inter, Segoe UI, Roboto, Arial, sans-serif; }
    a { color:inherit; text-decoration:none; }
    header { position:relative; overflow:hidden; padding:24px clamp(16px,4vw,42px) 30px; background:radial-gradient(720px circle at 76% 35%, rgba(43,132,203,.34), transparent 62%), linear-gradient(135deg,#34104f,#4c176d 58%,#1b255f); color:#fff; }
    header::after { content:""; position:absolute; right:clamp(20px,6vw,76px); bottom:-88px; width:min(46vw,520px); aspect-ratio:1.8; background:url("{favicon_url}") center / contain no-repeat; opacity:.18; pointer-events:none; }
    .topbar { position:relative; z-index:2; display:flex; justify-content:space-between; gap:18px; align-items:flex-start; }
    .brand-title { display:flex; align-items:center; gap:16px; }
    .brand-title img { width:90px; height:auto; filter:drop-shadow(0 10px 18px rgba(0,0,0,.24)); }
    h1 { margin:0; font-size:clamp(28px,4vw,46px); line-height:1; letter-spacing:0; }
    .subtitle { margin:8px 0 0; color:#d7e4ea; }
    .nav { display:flex; flex-wrap:wrap; justify-content:flex-end; gap:9px; }
    .top-link, button, .button { min-height:38px; display:inline-flex; align-items:center; justify-content:center; padding:9px 12px; border:1px solid rgba(255,255,255,.30); border-radius:8px; background:rgba(255,255,255,.10); color:#fff; font:inherit; font-size:13px; font-weight:900; cursor:pointer; }
    main { padding:22px clamp(16px,4vw,42px) 40px; }
    .panel, .kpi { background:var(--panel); border:1px solid var(--line); border-radius:8px; box-shadow:var(--shadow); }
    .message { position:fixed; left:50%; bottom:18px; z-index:30; min-width:min(520px,calc(100vw - 32px)); transform:translateX(-50%); padding:12px 14px; border-radius:8px; border:1px solid #b8e6c8; background:#f0fff5; color:#166534; font-weight:850; box-shadow:0 18px 38px rgba(23,32,51,.18); transition:opacity .2s ease, transform .2s ease; }
    .message.error { border-color:#fecaca; background:#fff1f2; color:#991b1b; }
    .message.is-hidden { opacity:0; transform:translate(-50%, 12px); pointer-events:none; }
    .tabs { display:flex; gap:8px; margin-bottom:14px; }
    .tabs button { background:#f3f6f8; border:1px solid var(--line); color:var(--ink); }
    .tabs button.active { background:var(--purple); color:#fff; border-color:transparent; }
    .tab-view[hidden] { display:none; }
    .filters { display:flex; flex-wrap:wrap; gap:10px; align-items:end; margin-bottom:14px; padding:12px; border:1px solid var(--line); border-radius:10px; background:rgba(255,255,255,.94); box-shadow:var(--shadow); }
    .custom-date-filter { display:none; gap:10px; align-items:end; }
    .custom-date-filter.is-visible { display:flex; }
    label { display:grid; gap:6px; color:var(--muted); font-size:12px; font-weight:900; text-transform:uppercase; }
    select, input { min-height:38px; border:1px solid var(--line); border-radius:8px; padding:8px 10px; font:inherit; font-weight:800; background:#fff; color:var(--ink); }
    .import-panel { margin-bottom:14px; padding:14px; display:flex; flex-wrap:wrap; gap:12px; align-items:end; justify-content:space-between; }
    .import-panel form { display:flex; flex-wrap:wrap; gap:10px; align-items:end; }
    .import-panel button { background:var(--purple); border-color:transparent; color:#fff; }
    .meta { color:var(--muted); font-size:13px; font-weight:800; }
    .kpis { display:grid; grid-template-columns:repeat(5,minmax(180px,1fr)); gap:14px; margin-bottom:14px; }
    .kpi { min-height:144px; padding:20px; display:grid; grid-template-columns:64px 1fr; gap:18px; align-items:center; }
    .kpi-icon { width:58px; height:58px; border-radius:12px; display:grid; place-items:center; color:#fff; background:linear-gradient(135deg,#64248c,#3f48cc); }
    .kpi-icon svg { width:24px; height:24px; stroke:currentColor; stroke-width:2.4; fill:none; stroke-linecap:round; stroke-linejoin:round; }
    .kpi:nth-child(2) .kpi-icon { background:var(--green); }
    .kpi:nth-child(3) .kpi-icon { background:var(--red); }
    .kpi:nth-child(4) .kpi-icon { background:var(--blue); }
    .kpi:nth-child(5) .kpi-icon { background:#1b255f; }
    .kpi span { color:var(--muted); font-size:12px; font-weight:900; text-transform:uppercase; }
    .kpi strong { display:block; margin-top:8px; font-size:34px; line-height:1; color:var(--ink); }
    .kpi small { display:block; margin-top:6px; color:var(--muted); font-size:12px; font-weight:800; }
    .delta { margin-top:12px; display:inline-flex; padding:6px 9px; border-radius:8px; background:#e7f7ee; color:#166534; font-size:12px; font-weight:900; }
    .delta.bad { background:#fff1f2; color:#991b1b; }
    .sla-bar { height:12px; border-radius:999px; background:#e8eef5; overflow:hidden; margin-top:14px; }
    .sla-fill { height:100%; border-radius:inherit; background:linear-gradient(90deg,#ff8a1c,#ffc22e); }
    .ops-grid { display:grid; grid-template-columns:1fr; gap:14px; margin-bottom:14px; }
    .lower-grid { display:grid; grid-template-columns:minmax(360px,.82fr) minmax(520px,1.18fr); grid-template-areas:"branch heatmap" "reasons heatmap" "alerts alerts"; gap:14px; margin-bottom:14px; align-items:stretch; }
    .branch-panel { grid-area:branch; }
    .heatmap-panel { grid-area:heatmap; }
    .reasons-panel { grid-area:reasons; }
    .alerts-panel { grid-area:alerts; }
    .heatmap-panel .panel-body { height:calc(100% - 44px); display:flex; flex-direction:column; }
    .panel h2 { margin:0; padding:16px 18px 0; font-size:18px; color:var(--ink); }
    .panel-head { display:flex; align-items:center; justify-content:space-between; gap:14px; padding:16px 18px 0; }
    .panel .panel-head h2 { padding:0; }
    .panel.compact h2 { font-size:17px; padding-top:14px; }
    .panel-body { padding:12px 18px 16px; }
    .panel.compact .panel-body { padding:10px 18px 14px; }
    .evolution-legend { display:flex; flex-wrap:wrap; justify-content:flex-end; gap:14px; color:var(--muted); font-size:12px; font-weight:900; }
    .evolution-legend span { display:inline-flex; align-items:center; gap:6px; white-space:nowrap; }
    .evolution-legend span::before { content:""; width:10px; height:10px; border-radius:50%; background:#00856f; }
    .evolution-legend .late::before { background:#e2263c; }
    .status-card { display:grid; grid-template-columns:160px 1fr; gap:18px; align-items:center; }
    .donut { position:relative; width:156px; height:156px; border-radius:50%; display:grid; place-items:center; background:conic-gradient(var(--green) 0deg, var(--green) var(--okDeg), var(--red) var(--okDeg), var(--red) 360deg); }
    .donut::before { content:""; width:104px; height:104px; border-radius:50%; background:#fff; box-shadow:inset 0 0 0 1px var(--line); }
    .donut-label { position:absolute; display:grid; gap:3px; text-align:center; font-weight:950; }
    .donut-label span { color:var(--muted); font-size:11px; text-transform:uppercase; }
    .legend { display:grid; gap:12px; }
    .legend-row { display:grid; grid-template-columns:12px 1fr auto; gap:9px; align-items:center; font-size:13px; font-weight:900; }
    .dot { width:12px; height:12px; border-radius:50%; background:var(--green); }
    .dot.late { background:var(--red); }
    .bars { display:grid; gap:11px; }
    .bar-row { display:grid; grid-template-columns:150px 1fr 80px; gap:10px; align-items:center; font-weight:850; font-size:13px; }
    .track { height:12px; border-radius:999px; background:#edf2f6; overflow:hidden; display:flex; }
    .fill-ok { height:100%; background:var(--green); }
    .fill-late { height:100%; background:var(--red); }
    .table-wrap { overflow:auto; max-height:620px; }
    table { width:100%; border-collapse:collapse; min-width:980px; font-size:13px; }
    th, td { padding:11px 10px; border-bottom:1px solid var(--line); text-align:left; vertical-align:middle; }
    th { background:#f3f6f8; color:var(--muted); text-transform:uppercase; font-size:12px; position:sticky; top:0; }
    td.num, th.num { text-align:right; }
    .badge { display:inline-flex; min-height:27px; align-items:center; padding:5px 8px; border-radius:999px; font-weight:950; white-space:nowrap; }
    .badge.ok { background:#e7f7ee; color:#166534; }
    .badge.bad { background:#fff1f2; color:#991b1b; }
    .chart-wrap { overflow:hidden; padding-bottom:6px; }
    .chart-wrap svg { width:100%; height:auto; display:block; }
    .detail-trigger { cursor:pointer; }
    .detail-trigger:hover { filter:brightness(.96); outline:2px solid rgba(43,132,203,.28); outline-offset:2px; }
    .alert-list { display:grid; gap:8px; }
    .alert-item { display:grid; grid-template-columns:34px 1fr; gap:10px; padding:10px; border-radius:8px; background:#f8fafb; border:1px solid var(--line); }
    .alert-icon { width:32px; height:32px; border-radius:8px; display:grid; place-items:center; background:rgba(242,56,78,.16); color:#ff6b78; font-weight:950; }
    .alert-item.warning .alert-icon { background:rgba(245,158,11,.16); color:#ffb13b; }
    .alert-item strong { display:block; margin-bottom:4px; }
    .branch-table { width:100%; min-width:0; }
    .mini-progress { width:74px; height:12px; border-radius:999px; background:#e8eef5; overflow:hidden; display:inline-flex; vertical-align:middle; }
    .mini-progress span { display:block; background:var(--red); }
    .mini-progress.ok span { background:var(--green); }
    .heatmap-note { margin-bottom:8px; color:var(--muted); font-size:12px; font-weight:850; }
    .heatmap { flex:1; display:grid; grid-template-columns:42px repeat(6,minmax(54px,1fr)); gap:5px; align-items:stretch; font-size:12px; }
    .heatmap > strong { display:flex; align-items:center; justify-content:center; }
    .heat-cell { min-height:48px; border-radius:4px; background:#e8eef5; display:flex; flex-direction:column; align-items:center; justify-content:center; color:#16212d; font-weight:900; line-height:1.05; }
    .heat-cell small { margin-top:3px; font-size:10px; font-weight:950; color:rgba(22,33,45,.78); }
    .reason-list { display:grid; gap:10px; font-size:13px; font-weight:850; }
    .reason-row { display:grid; grid-template-columns:78px minmax(120px,1fr) 44px 52px; gap:10px; align-items:center; }
    .reason-label { display:flex; align-items:center; gap:8px; white-space:nowrap; }
    .reason-dot { width:10px; height:10px; border-radius:50%; background:var(--red); flex:0 0 auto; }
    .reason-track { height:10px; border-radius:999px; background:#edf2f6; overflow:hidden; }
    .reason-fill { display:block; height:100%; border-radius:inherit; background:var(--red); }
    .reason-value { text-align:right; font-weight:950; }
    .reason-percent { color:var(--muted); text-align:right; font-weight:850; }
    .reason-total { display:grid; grid-template-columns:1fr auto; gap:12px; margin-top:4px; padding-top:10px; border-top:1px solid var(--line); font-weight:950; }
    .empty { padding:26px; color:var(--muted); font-weight:800; text-align:center; }
    .detail-modal { position:fixed; inset:0; z-index:50; display:grid; place-items:center; padding:24px; background:rgba(15,23,42,.44); }
    .detail-modal[hidden] { display:none; }
    .detail-dialog { width:min(1180px,96vw); max-height:88vh; display:grid; grid-template-rows:auto 1fr; background:#fff; border-radius:8px; box-shadow:0 24px 70px rgba(15,23,42,.30); overflow:hidden; }
    .detail-head { display:flex; align-items:center; justify-content:space-between; gap:16px; padding:16px 18px; border-bottom:1px solid var(--line); }
    .detail-head h3 { margin:0; font-size:18px; }
    .detail-head button { width:36px; height:36px; border-radius:8px; border:1px solid var(--line); background:#fff; color:var(--ink); font-size:20px; font-weight:950; cursor:pointer; }
    .detail-count { color:var(--muted); font-size:13px; font-weight:900; }
    .detail-table-wrap { overflow:auto; padding:0 18px 18px; }
    .detail-table { min-width:1120px; }
    @media (max-width:1280px) { .kpis { grid-template-columns:repeat(3,minmax(170px,1fr)); } .lower-grid { grid-template-columns:1fr; grid-template-areas:"branch" "reasons" "heatmap" "alerts"; } }
    @media (max-width:1100px) { .kpis { grid-template-columns:repeat(2,minmax(150px,1fr)); } .grid { grid-template-columns:1fr; } }
    @media (max-width:760px) { .topbar { flex-direction:column; } .nav { justify-content:flex-start; } .kpis, .status-card, .lower-grid { grid-template-columns:1fr; } .bar-row { grid-template-columns:1fr; } .panel-head { align-items:flex-start; flex-direction:column; } .evolution-legend { justify-content:flex-start; } }
  </style>
</head>
<body>
  <header>
    <div class="topbar">
      <div>
        <div class="brand-title"><img src="{favicon_url}" alt=""><h1>Controle Medicao</h1></div>
      </div>
      <nav class="nav">
        <a class="top-link" href="/home">Home</a>
        <a class="top-link" href="/dashboard">Dashboard</a>
        <a class="top-link" href="/controle-ct">Controle de CT</a>
        <a class="top-link" href="/relatorio-diario">Relatorio diario</a>
        <a class="top-link" href="/relatorio-entrada-notas">Entrada de notas</a>
        <button type="button" id="refreshDashboard">Atualizar</button>
        <a class="top-link" href="/logout">Sair</a>
      </nav>
    </div>
  </header>
  <main>
    {message}
    <div class="filters">
      <label>Periodo <select id="dateModeFilter"><option value="month">Mes atual</option><option value="today">Hoje</option><option value="week">Semana atual</option><option value="last7">Ultimos 7 dias</option><option value="custom">Periodo</option><option value="all">Todas</option></select></label>
      <span class="custom-date-filter" id="customDateFilter"><label>Inicio <input id="dateStartFilter" type="date"></label><label>Fim <input id="dateEndFilter" type="date"></label></span>
      <label>Terminal <select id="terminalFilter"><option value="">Todos</option></select></label>
      <label>Filial <select id="branchFilter"><option value="">Todas</option></select></label>
      <label>Usuario <select id="userChangeFilter"><option value="">Todos</option></select></label>
      <label>Status <select id="statusFilter"><option value="">Todos</option><option value="ok">No prazo</option><option value="late">Fora do prazo</option></select></label>
      <label>Tempo fechamento <select id="timeFilter"><option value="">Todos</option><option value="same">Mesmo dia</option><option value="one">Ate 1 dia</option><option value="two">Ate 2 dias</option><option value="late">Acima de 2 dias</option></select></label>
    </div>
    <div class="tabs"><button type="button" class="active" data-tab="dashboard">Dashboard</button><button type="button" data-tab="data">Dados</button></div>
    <section id="dashboard" class="tab-view">
      <div class="kpis">
        <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24"><path d="M4 4h16v16H4z"></path><path d="M8 2v4"></path><path d="M16 2v4"></path><path d="M4 10h16"></path></svg></div><div><span>Total medicoes</span><strong id="kTotal">0</strong><small>Fechamentos no filtro</small></div></div>
        <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"></circle><path d="m8 12 3 3 5-6"></path></svg></div><div><span>No prazo</span><strong id="kOk">0</strong><small id="kOkHint">0% do total</small></div></div>
        <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"></circle><path d="m15 9-6 6"></path><path d="m9 9 6 6"></path></svg></div><div><span>Fora do prazo</span><strong id="kLate">0</strong><small id="kLateHint">0% do total</small><div class="delta bad" id="kLateDelta">Acompanhar atrasos</div></div></div>
        <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"></circle><path d="M12 7v5l3 2"></path></svg></div><div><span>Tempo medio</span><strong id="kAvg">-</strong><small>Fechamento medio</small></div></div>
        <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"></circle><path d="M12 8v4l3-2"></path><path d="m8 16 2-2"></path></svg></div><div><span>SLA (Meta)</span><strong id="kSla">0%</strong><small>Meta: 100% no prazo</small><div class="sla-bar"><div id="kSlaFill" class="sla-fill" style="width:0%"></div></div></div></div>
      </div>
      <div class="ops-grid">
        <section class="panel"><div class="panel-head"><h2>Evolucao dos Fechamentos</h2><div class="evolution-legend"><span>No prazo</span><span class="late">Fora do prazo</span></div></div><div class="panel-body"><div id="evolutionChart" class="chart-wrap"></div></div></section>
      </div>
      <div class="lower-grid">
        <section class="panel compact branch-panel"><h2>Desempenho por Filial</h2><div class="panel-body"><div id="branchPerformance"></div></div></section>
        <section class="panel compact reasons-panel"><h2>Faixas de Atraso</h2><div class="panel-body"><div id="reasonPanel"></div></div></section>
        <section class="panel compact heatmap-panel"><h2>Fechamentos por Dia/Hora</h2><div class="panel-body"><div class="heatmap-note">Total por horario; abaixo, fora do prazo.</div><div id="heatmapPanel" class="heatmap"></div></div></section>
        <section class="panel compact alerts-panel"><h2>Alertas Criticos</h2><div class="panel-body"><div id="alertsPanel" class="alert-list"></div></div></section>
      </div>
    </section>
    <section id="data" class="tab-view" hidden>
      <section class="panel import-panel">
        <div><strong>Importar medicoes</strong><div class="meta">Envie a planilha atualizada para renovar o relatorio.</div></div>
        <form method="post" action="/controle-medicao/importar" enctype="multipart/form-data"><input type="file" name="measurement_file" accept=".xlsx" required><button type="submit">Importar medicoes</button></form>
      </section>
      <div class="panel"><div class="table-wrap"><table><thead><tr><th>Seq.</th><th>Filial</th><th>Terminal</th><th>Usuario</th><th>Dt. Medicao</th><th>Fechamento</th><th>Prazo limite</th><th>Status</th><th class="num">Tempo</th><th class="num">Horas fora</th></tr></thead><tbody id="rows"></tbody></table></div></div>
    </section>
  </main>
  <div class="detail-modal" id="detailModal" hidden>
    <div class="detail-dialog" role="dialog" aria-modal="true" aria-labelledby="detailTitle">
      <div class="detail-head">
        <div><h3 id="detailTitle">Detalhes</h3><div class="detail-count" id="detailCount"></div></div>
        <button type="button" id="detailClose" aria-label="Fechar">x</button>
      </div>
      <div class="detail-table-wrap"><table class="detail-table"><thead><tr><th>Seq.</th><th>Filial</th><th>Terminal</th><th>Usuario</th><th>Dt. Medicao</th><th>Fechamento</th><th>Prazo limite</th><th>Status</th><th class="num">Tempo</th><th class="num">Horas fora</th></tr></thead><tbody id="detailRows"></tbody></table></div>
    </div>
  </div>
  <script>
    const rows = __ROWS__;
    const fmt = new Intl.NumberFormat("pt-BR");
    const $ = (id) => document.getElementById(id);
    function escapeHtml(value) { return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
    function parseBrDate(value) { const [d,m,y]=String(value||"").split("/").map(Number); return d&&m&&y ? new Date(y,m-1,d) : null; }
    function parseBrDateTime(value) { const [datePart,timePart="00:00"]=String(value||"").split(" "); const [d,m,y]=datePart.split("/").map(Number); const [hh=0,mm=0]=timePart.split(":").map(Number); return d&&m&&y ? new Date(y,m-1,d,hh||0,mm||0) : null; }
    function isoDate(date) { return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,"0")}-${String(date.getDate()).padStart(2,"0")}`; }
    function startOfDay(date) { return new Date(date.getFullYear(), date.getMonth(), date.getDate()); }
    function endOfDay(date) { return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23,59,59,999); }
    function durationLabel(hours) { if (!Number.isFinite(hours)) return "-"; const h=Math.max(0,Math.round(hours)); const d=Math.floor(h/24); const r=h%24; return d ? `${d}d ${r}h` : `${r}h`; }
    function percent(part,total) { return total ? `${(part/total*100).toFixed(1).replace(".", ",")}%` : "0%"; }
    function periodBounds() {
      const mode=$("dateModeFilter").value; const today=startOfDay(new Date());
      if (mode==="all") return null;
      if (mode==="today") return [today,endOfDay(today)];
      if (mode==="last7") { const s=new Date(today); s.setDate(s.getDate()-6); return [s,endOfDay(today)]; }
      if (mode==="week") { const s=new Date(today); const day=s.getDay()||7; s.setDate(s.getDate()-day+1); const e=new Date(s); e.setDate(e.getDate()+6); return [s,endOfDay(e)]; }
      if (mode==="custom") { const sv=$("dateStartFilter").value; const ev=$("dateEndFilter").value; const s=sv?new Date(`${sv}T00:00:00`):null; const e=ev?new Date(`${ev}T23:59:59`):null; if(s&&e)return[s,e]; if(s)return[s,endOfDay(today)]; if(e)return[new Date(1900,0,1),e]; return null; }
      return [new Date(today.getFullYear(), today.getMonth(), 1), endOfDay(new Date(today.getFullYear(), today.getMonth()+1, 0))];
    }
    function updateCustomDateFilter(){ $("customDateFilter").classList.toggle("is-visible", $("dateModeFilter").value==="custom"); }
    function matchesPeriod(row){ const b=periodBounds(); if(!b)return true; const d=parseBrDate(row.medicao); return d && d>=b[0] && d<=b[1]; }
    function matchesTime(row){ const f=$("timeFilter").value; const h=Number(row.horasFechamento); if(!f)return true; if(f==="same")return h<24; if(f==="one")return h<=24; if(f==="two")return h<=48; return h>48; }
    function branchName(value){ const names = { "171": "MANAUS", "182": "BOA VISTA", "178": "ITACOATIARA" }; return names[String(value || "").trim()] || value || "-"; }
    function filteredRows(){ const terminal=$("terminalFilter").value, filial=$("branchFilter").value, usuario=$("userChangeFilter").value, status=$("statusFilter").value; return rows.filter(r=>matchesPeriod(r)&&matchesTime(r)&&(!terminal||r.terminal===terminal)&&(!filial||r.filial===filial)&&(!usuario||r.usuarioAlteracao===usuario)&&(!status||r.status===status)); }
    function fillFilters(){ const terminals=[...new Set(rows.map(r=>r.terminal).filter(Boolean))].sort(); const branches=[...new Set(rows.map(r=>r.filial).filter(Boolean))].sort(); const users=[...new Set(rows.map(r=>r.usuarioAlteracao).filter(Boolean))].sort(); $("terminalFilter").innerHTML='<option value="">Todos</option>'+terminals.map(v=>`<option>${escapeHtml(v)}</option>`).join(""); $("branchFilter").innerHTML='<option value="">Todas</option>'+branches.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(branchName(v))}</option>`).join(""); $("userChangeFilter").innerHTML='<option value="">Todos</option>'+users.map(v=>`<option>${escapeHtml(v)}</option>`).join(""); }
    function grouped(data,key){ const map=new Map(); data.forEach(r=>{ const label=r[key]||"-"; const item=map.get(label)||{ok:0,late:0,total:0}; item.total++; if(r.status==="ok")item.ok++; else item.late++; map.set(label,item); }); return [...map.entries()].sort((a,b)=>b[1].total-a[1].total); }
    function rowMarkup(r){ return `<tr><td>${escapeHtml(r.seq)}</td><td>${escapeHtml(branchName(r.filial))}</td><td>${escapeHtml(r.terminal)}</td><td>${escapeHtml(r.usuarioAlteracao)}</td><td>${escapeHtml(r.medicao)}</td><td>${escapeHtml(r.fechamento)}</td><td>${escapeHtml(r.prazo)}</td><td><span class="badge ${r.status==="ok"?"ok":"bad"}">${r.status==="ok"?"No prazo":"Fora do prazo"}</span></td><td class="num">${durationLabel(Number(r.horasFechamento))}</td><td class="num">${r.horasFora?durationLabel(Number(r.horasFora)):"-"}</td></tr>`; }
    function showDetails(title, data){ $("detailTitle").textContent=title; $("detailCount").textContent=`${fmt.format(data.length)} medicoes`; $("detailRows").innerHTML=data.map(rowMarkup).join("")||'<tr><td colspan="10" class="empty">Sem dados para este item.</td></tr>'; $("detailModal").hidden=false; }
    function closeDetails(){ $("detailModal").hidden=true; }
    function renderBars(id, entries){ const max=Math.max(1,...entries.map(([,v])=>v.total)); $(id).innerHTML=entries.length?entries.map(([label,item])=>`<div class="bar-row"><span>${escapeHtml(label)}</span><div class="track"><div class="fill-ok" style="width:${item.ok/max*100}%"></div><div class="fill-late" style="width:${item.late/max*100}%"></div></div><strong>${fmt.format(item.total)}</strong></div>`).join(""):'<div class="empty">Sem dados para o filtro.</div>'; }
    function dailyStats(data){
      const map = new Map();
      data.forEach((row) => {
        const item = map.get(row.medicao) || { ok: 0, late: 0, total: 0, rows: [] };
        item.total++;
        item.rows.push(row);
        if (row.status === "ok") item.ok++; else item.late++;
        map.set(row.medicao, item);
      });
      return [...map.entries()].sort((a,b) => (parseBrDate(a[0]) || 0) - (parseBrDate(b[0]) || 0));
    }
    function renderEvolution(data){
      const stats = dailyStats(data).slice(-31);
      if (!stats.length) { $("evolutionChart").innerHTML = '<div class="empty">Sem dados para o filtro.</div>'; return; }
      const containerWidth = $("evolutionChart").clientWidth || 980;
      const width = Math.max(520, Math.round(containerWidth));
      const height = Math.max(260, Math.min(330, Math.round(width * .24)));
      const pad = { left: 42, right: 24, top: 24, bottom: 48 };
      const plotW = width - pad.left - pad.right, plotH = height - pad.top - pad.bottom;
      const max = Math.max(1, ...stats.map(([, item]) => item.total));
      const x = (idx) => pad.left + (stats.length === 1 ? plotW / 2 : idx * (plotW / (stats.length - 1)));
      const barArea = plotW / Math.max(1, stats.length);
      const labelStep = width < 700 ? Math.ceil(stats.length / 12) : width < 980 ? Math.ceil(stats.length / 18) : 1;
      const bars = stats.map(([date, item], idx) => {
        const barW = Math.max(10, Math.min(34, barArea * .62));
        const okH = item.ok / max * plotH;
        const lateH = item.late / max * plotH;
        const bx = x(idx) - barW / 2;
        const base = pad.top + plotH;
        const fontSize = barW < 17 ? 9 : 11;
        const okLabel = item.ok && okH >= 22 && barW >= 14 ? `<text x="${x(idx)}" y="${base-okH/2+4}" text-anchor="middle" fill="#fff" font-size="${fontSize}" font-weight="900">${fmt.format(item.ok)}</text>` : "";
        const lateLabel = item.late && lateH >= 22 && barW >= 14 ? `<text x="${x(idx)}" y="${base-okH-lateH/2+4}" text-anchor="middle" fill="#fff" font-size="${fontSize}" font-weight="900">${fmt.format(item.late)}</text>` : "";
        const dateLabel = idx % labelStep === 0 || idx === stats.length - 1 ? `<text x="${x(idx)}" y="${height-18}" text-anchor="middle" fill="#334155" font-size="11" font-weight="850">${date.slice(0,5)}</text>` : "";
        return `<g class="detail-trigger" data-evolution-idx="${idx}"><rect x="${bx}" y="${base-okH}" width="${barW}" height="${okH}" fill="#00856f" rx="2"></rect><rect x="${bx}" y="${base-okH-lateH}" width="${barW}" height="${lateH}" fill="#e2263c" rx="2"></rect>${okLabel}${lateLabel}${dateLabel}</g>`;
      }).join("");
      $("evolutionChart").innerHTML = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Evolucao dos fechamentos">
        <g>${[0,.25,.5,.75,1].map((v)=>`<line x1="${pad.left}" x2="${width-pad.right}" y1="${pad.top+plotH*v}" y2="${pad.top+plotH*v}" stroke="#d7e0e8"/>`).join("")}</g>
        ${bars}
      </svg>`;
      $("evolutionChart").querySelectorAll("[data-evolution-idx]").forEach((node) => node.addEventListener("click", () => {
        const [date, item] = stats[Number(node.dataset.evolutionIdx)];
        showDetails(`Fechamentos em ${date}`, item.rows);
      }));
    }
    function topEntry(entries){ return entries[0] || ["-", { total: 0, ok: 0, late: 0 }]; }
    function renderAlerts(data){
      const terminals = grouped(data, "terminal");
      const branches = grouped(data, "filial");
      const [topTerminal, topTerminalItem] = topEntry(terminals.filter(([, item]) => item.late > 0));
      const [topBranch, topBranchItem] = topEntry(branches.filter(([, item]) => item.late > 0));
      const lateAfter18 = data.filter((row) => row.status === "late" && Number(row.horasFechamento) > 18).length;
      const alerts = [
        { icon: "!", cls: "", title: `${topTerminal || "Terminal"} acima da meta`, text: `${percent(topTerminalItem.late, topTerminalItem.total)} fora do prazo (${fmt.format(topTerminalItem.late)} atrasos)`, rows: data.filter(r=>r.terminal===topTerminal && r.status==="late") },
        { icon: "!", cls: "warning", title: `${branchName(topBranch)} com maior atraso`, text: `${fmt.format(topBranchItem.late)} fechamentos fora do prazo no filtro`, rows: data.filter(r=>r.filial===topBranch && r.status==="late") },
        { icon: "!", cls: "warning", title: "Atrasos acima de 18h", text: `${fmt.format(lateAfter18)} fechamentos em atraso com mais de 18h`, rows: data.filter((row) => row.status === "late" && Number(row.horasFechamento) > 18) }
      ];
      $("alertsPanel").innerHTML = alerts.map((item, idx) => `<div class="alert-item detail-trigger ${item.cls}" data-alert-idx="${idx}"><div class="alert-icon">${item.icon}</div><div><strong>${escapeHtml(item.title)}</strong><span class="meta">${escapeHtml(item.text)}</span></div></div>`).join("");
      $("alertsPanel").querySelectorAll("[data-alert-idx]").forEach((node) => node.addEventListener("click", () => {
        const item = alerts[Number(node.dataset.alertIdx)];
        showDetails(item.title, item.rows);
      }));
    }
    function renderBranchPerformance(data){
      const entries = grouped(data, "filial").slice(0, 8);
      $("branchPerformance").innerHTML = entries.length ? `<table class="branch-table"><thead><tr><th>Filial</th><th>Total</th><th>No prazo</th><th>Fora</th><th>% No prazo</th><th>SLA</th></tr></thead><tbody>${entries.map(([label,item], idx) => {
        const okPct = item.total ? item.ok / item.total * 100 : 0;
        return `<tr class="detail-trigger" data-branch-idx="${idx}"><td><strong>${escapeHtml(branchName(label))}</strong></td><td>${fmt.format(item.total)}</td><td>${fmt.format(item.ok)}</td><td>${fmt.format(item.late)}</td><td>${percent(item.ok,item.total)}</td><td><span class="mini-progress ${okPct>=85?"ok":""}"><span style="width:${okPct}%"></span></span></td></tr>`;
      }).join("")}</tbody></table>` : '<div class="empty">Sem filiais no filtro.</div>';
      $("branchPerformance").querySelectorAll("[data-branch-idx]").forEach((node) => node.addEventListener("click", () => {
        const branch = entries[Number(node.dataset.branchIdx)][0];
        showDetails(`Fechamentos de ${branchName(branch)}`, data.filter(r=>r.filial===branch));
      }));
    }
    function renderHeatmap(data){
      const days = ["Seg","Ter","Qua","Qui","Sex","Sab","Dom"];
      const buckets = ["08h","10h","12h","14h","16h","18h"];
      const map = new Map();
      data.forEach((row) => {
        const date = parseBrDateTime(row.fechamento) || parseBrDate(row.medicao);
        if (!date) return;
        const day = (date.getDay() + 6) % 7;
        const hour = date.getHours();
        const bucket = Math.min(5, Math.max(0, Math.floor(Math.max(8, Math.min(19, hour)) / 2) - 4));
        const key = `${day}-${bucket}`;
        const item = map.get(key) || { total: 0, late: 0, rows: [] };
        item.total++;
        item.rows.push(row);
        if (row.status === "late") item.late++;
        map.set(key, item);
      });
      $("heatmapPanel").innerHTML = `<span></span>${buckets.map(b=>`<strong>${b}</strong>`).join("")}${days.map((day, d) => `<strong>${day}</strong>${buckets.map((bucketLabel, b) => { const item = map.get(`${d}-${b}`) || { total:0, late:0, rows:[] }; const ratio = item.total ? item.late / item.total : 0; const hue = 145 - ratio * 145; const detail = `${day} ${bucketLabel}: ${fmt.format(item.total)} fechamentos, ${fmt.format(item.late)} fora do prazo`; return item.total ? `<span class="heat-cell detail-trigger" data-heat-key="${d}-${b}" title="${detail}" style="background:hsl(${hue} 76% 52%)"><strong>${fmt.format(item.total)}</strong><small>${fmt.format(item.late)} fora</small></span>` : `<span class="heat-cell" title="${detail}"></span>`; }).join("")}`).join("")}`;
      $("heatmapPanel").querySelectorAll("[data-heat-key]").forEach((node) => node.addEventListener("click", () => {
        const item = map.get(node.dataset.heatKey);
        showDetails(`Fechamentos ${node.title.split(":")[0]}`, item?.rows || []);
      }));
    }
    function renderReasons(data){
      const lateData = data.filter(r=>r.status==="late");
      const ranges = [
        ["1 dia", "#4fbf7a", lateData.filter(r=>Number(r.horasFora)<=24)],
        ["2 dias", "#ffcc3d", lateData.filter(r=>Number(r.horasFora)>24 && Number(r.horasFora)<=48)],
        ["3 dias", "#ff7a2d", lateData.filter(r=>Number(r.horasFora)>48 && Number(r.horasFora)<=72)],
        ["4 dias", "#ff453a", lateData.filter(r=>Number(r.horasFora)>72 && Number(r.horasFora)<=96)],
        ["5+ dias", "#d72630", lateData.filter(r=>Number(r.horasFora)>96)]
      ];
      const total = ranges.reduce((sum, [, , items]) => sum + items.length, 0);
      const max = Math.max(1, ...ranges.map(([, , items]) => items.length));
      $("reasonPanel").innerHTML = `<div class="reason-list">
        ${ranges.map(([label,color,items], idx)=>`
          <div class="reason-row detail-trigger" data-range-idx="${idx}">
            <span class="reason-label"><span class="reason-dot" style="background:${color}"></span>${label}</span>
            <span class="reason-track"><span class="reason-fill" style="width:${items.length / max * 100}%; background:${color}"></span></span>
            <span class="reason-value">${fmt.format(items.length)}</span>
            <span class="reason-percent">${percent(items.length,total)}</span>
          </div>
        `).join("")}
        <div class="reason-total"><span>Total em atraso</span><strong>${fmt.format(total)}</strong></div>
      </div>`;
      $("reasonPanel").querySelectorAll("[data-range-idx]").forEach((node) => node.addEventListener("click", () => {
        const [label,, items] = ranges[Number(node.dataset.rangeIdx)];
        showDetails(`Faixa de atraso: ${label}`, items);
      }));
    }
    function renderTable(data){ $("rows").innerHTML=data.map(rowMarkup).join("")||'<tr><td colspan="10" class="empty">Sem dados para o filtro.</td></tr>'; }
    function render(){
      updateCustomDateFilter();
      const data=filteredRows();
      const ok=data.filter(r=>r.status==="ok").length;
      const late=data.length-ok;
      const avg=data.length?data.reduce((t,r)=>t+Number(r.horasFechamento||0),0)/data.length:null;
      const sla=data.length ? ok/data.length*100 : 0;
      $("kTotal").textContent=fmt.format(data.length);
      $("kOk").textContent=fmt.format(ok);
      $("kLate").textContent=fmt.format(late);
      $("kOkHint").textContent=`${percent(ok,data.length)} do total`;
      $("kLateHint").textContent=`${percent(late,data.length)} do total`;
      $("kAvg").textContent=avg===null?"-":durationLabel(avg);
      $("kSla").textContent=`${sla.toFixed(0)}%`;
      $("kSlaFill").style.width=`${Math.min(100,sla)}%`;
      $("kLateDelta").textContent = "";
      renderEvolution(data);
      renderAlerts(data);
      renderBranchPerformance(data);
      renderHeatmap(data);
      renderReasons(data);
      renderTable(data);
    }
    document.querySelectorAll(".message").forEach((item)=>{ setTimeout(()=>item.classList.add("is-hidden"),4200); setTimeout(()=>item.remove(),4600); });
    document.querySelectorAll(".tabs button").forEach(btn=>btn.addEventListener("click",()=>{ document.querySelectorAll(".tabs button").forEach(b=>b.classList.toggle("active",b===btn)); document.querySelectorAll(".tab-view").forEach(view=>view.hidden=view.id!==btn.dataset.tab); }));
    ["dateModeFilter","dateStartFilter","dateEndFilter","terminalFilter","branchFilter","userChangeFilter","statusFilter","timeFilter"].forEach(id=>$(id).addEventListener("change",render));
    $("refreshDashboard").addEventListener("click", render);
    let resizeTimer = null;
    window.addEventListener("resize", () => { clearTimeout(resizeTimer); resizeTimer = setTimeout(() => { if (!$("dashboard").hidden) renderEvolution(filteredRows()); }, 120); });
    $("detailClose").addEventListener("click", closeDetails);
    $("detailModal").addEventListener("click", (event)=>{ if(event.target===$("detailModal")) closeDetails(); });
    document.addEventListener("keydown", (event)=>{ if(event.key==="Escape" && !$("detailModal").hidden) closeDetails(); });
    fillFilters(); render();
  </script>
</body>
</html>
"""


NOTE_ENTRY_REPORT_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="icon" href="{favicon_url}" type="image/svg+xml">
  <title>Relatorio de entrada de notas - Dashboard</title>
  <style>
    :root { --bg:#eef2f5; --top:#34104f; --ink:#16212d; --muted:#657282; --line:#d7e0e8; --panel:#fff; --purple:#64248c; --blue:#2b84cb; --red:#e2263c; --green:#00856f; --shadow:0 18px 42px rgba(23,32,51,.10); }
    * { box-sizing: border-box; }
    body { margin:0; background:var(--bg); color:var(--ink); font-family:Inter, Segoe UI, Roboto, Arial, sans-serif; }
    a { color:inherit; text-decoration:none; }
    header { position:relative; overflow:hidden; padding:24px clamp(16px,4vw,42px) 30px; background:radial-gradient(720px circle at 76% 35%, rgba(43,132,203,.34), transparent 62%), linear-gradient(135deg,#34104f,#4c176d 58%,#1b255f); color:#fff; }
    header::after { content:""; position:absolute; right:clamp(20px,6vw,76px); bottom:-88px; width:min(46vw,520px); aspect-ratio:1.8; background:url("{favicon_url}") center / contain no-repeat; opacity:.18; pointer-events:none; }
    .topbar { position:relative; z-index:2; display:flex; justify-content:space-between; gap:18px; align-items:flex-start; }
    .brand-title { display:flex; align-items:center; gap:16px; }
    .brand-title img { width:90px; height:auto; filter:drop-shadow(0 10px 18px rgba(0,0,0,.24)); }
    h1 { margin:0; font-size:clamp(28px,4vw,46px); line-height:1; letter-spacing:0; }
    .subtitle { margin:8px 0 0; color:#d7e4ea; }
    .nav { display:flex; flex-wrap:wrap; justify-content:flex-end; gap:9px; }
    .top-link, button, .button { min-height:38px; display:inline-flex; align-items:center; justify-content:center; padding:9px 12px; border:1px solid rgba(255,255,255,.30); border-radius:8px; background:rgba(255,255,255,.10); color:#fff; font:inherit; font-size:13px; font-weight:900; cursor:pointer; }
    main { padding:22px clamp(16px,4vw,42px) 40px; }
    .panel, .kpi { background:var(--panel); border:1px solid var(--line); border-radius:8px; box-shadow:var(--shadow); }
    .message { position:fixed; left:50%; bottom:18px; z-index:30; min-width:min(520px,calc(100vw - 32px)); transform:translateX(-50%); padding:12px 14px; border-radius:8px; border:1px solid #b8e6c8; background:#f0fff5; color:#166534; font-weight:850; box-shadow:0 18px 38px rgba(23,32,51,.18); transition:opacity .2s ease, transform .2s ease; }
    .message.error { border-color:#fecaca; background:#fff1f2; color:#991b1b; }
    .message.is-hidden { opacity:0; transform:translate(-50%, 12px); pointer-events:none; }
    .import-panel { margin-bottom:14px; padding:14px; display:flex; flex-wrap:wrap; gap:12px; align-items:end; justify-content:space-between; }
    .import-panel form { display:flex; flex-wrap:wrap; gap:10px; align-items:end; }
    .import-panel input, select { min-height:38px; border:1px solid var(--line); border-radius:8px; padding:8px 10px; font:inherit; font-weight:800; background:#fff; color:var(--ink); }
    .import-panel button, .tabs button { background:var(--purple); border-color:transparent; color:#fff; }
    .meta { color:var(--muted); font-size:13px; font-weight:800; }
    .filters { display:flex; flex-wrap:wrap; gap:10px; align-items:end; margin-bottom:14px; }
    .custom-date-filter { display:none; gap:10px; align-items:end; }
    .custom-date-filter.is-visible { display:flex; }
    label { display:grid; gap:6px; color:var(--muted); font-size:12px; font-weight:900; text-transform:uppercase; }
    .kpis { display:grid; grid-template-columns:repeat(4,minmax(180px,1fr)); gap:14px; margin-bottom:14px; }
    .kpi { min-height:112px; padding:18px; display:grid; grid-template-columns:52px 1fr; gap:14px; align-items:center; border:0; }
    .kpi-icon { width:46px; height:46px; border-radius:12px; display:grid; place-items:center; color:#fff; background:#3f48cc; }
    .kpi-icon svg { width:24px; height:24px; stroke:currentColor; stroke-width:2.4; fill:none; stroke-linecap:round; stroke-linejoin:round; }
    .kpi:nth-child(2) .kpi-icon { background:#0b66d8; }
    .kpi:nth-child(3) .kpi-icon { background:var(--red); }
    .kpi:nth-child(4) .kpi-icon { background:var(--green); }
    .kpi span { color:var(--muted); font-size:12px; font-weight:900; text-transform:uppercase; }
    .kpi strong { display:block; margin-top:6px; font-size:30px; line-height:1; }
    .kpi small { display:block; margin-top:6px; color:var(--muted); font-size:12px; font-weight:800; }
    .tabs { display:flex; gap:8px; margin-bottom:14px; }
    .tabs button { background:#f3f6f8; border:1px solid var(--line); color:var(--ink); }
    .tabs button.active { background:var(--purple); color:#fff; border-color:transparent; }
    .tab-view[hidden] { display:none; }
    .branch-grid { display:grid; grid-template-columns:repeat(3,minmax(260px,1fr)); gap:14px; margin-bottom:14px; }
    .branch-card { padding:16px; background:#fff; border:1px solid var(--line); border-radius:8px; box-shadow:var(--shadow); }
    .branch-title { display:flex; align-items:center; gap:8px; margin:0 0 12px; color:#17377d; font-size:16px; font-weight:950; }
    .branch-title svg { width:22px; height:22px; color:#0b66d8; stroke:currentColor; fill:none; stroke-width:2; stroke-linecap:round; stroke-linejoin:round; flex:0 0 auto; }
    .branch-main { display:grid; grid-template-columns:128px 1fr; gap:14px; align-items:center; }
    .city-ring { position:relative; width:118px; height:118px; border-radius:50%; display:grid; place-items:center; background:conic-gradient(#0b66d8 0deg, #0b66d8 var(--okDeg), var(--red) var(--okDeg), var(--red) 360deg); }
    .city-ring::before { content:""; width:74px; height:74px; border-radius:50%; background:#fff; box-shadow:inset 0 0 0 1px var(--line); }
    .city-ring-label { position:absolute; text-align:center; font-weight:950; }
    .city-ring-label strong { display:block; font-size:21px; }
    .city-ring-label span { color:var(--muted); font-size:11px; }
    .city-metrics { display:grid; gap:10px; font-size:13px; font-weight:900; }
    .city-metric { display:grid; grid-template-columns:10px auto 1fr; gap:8px; align-items:center; }
    .city-metric::before { content:""; width:9px; height:9px; border-radius:50%; background:#0b66d8; }
    .city-metric.late::before { background:var(--red); }
    .city-total { padding-top:8px; color:var(--muted); font-weight:950; text-align:right; }
    .late-notes { margin-top:16px; padding-top:12px; border-top:1px solid #edf1f5; }
    .late-notes h3 { margin:0 0 8px; color:var(--red); font-size:13px; }
    .mini-table { width:100%; min-width:0; border-collapse:collapse; }
    .mini-table th, .mini-table td { padding:7px 0; border-bottom:1px solid #edf1f5; font-size:12px; }
    .mini-table th { position:static; background:transparent; color:var(--muted); }
    .mini-table td:last-child, .mini-table th:last-child { text-align:right; }
    .view-late { margin:10px auto 0; display:flex; width:max-content; border:0; background:transparent; color:#0b66d8; padding:4px; min-height:auto; font-size:12px; }
    .branch-summary { padding:16px 24px 20px; }
    .summary-legend { display:flex; justify-content:center; gap:22px; margin:4px 0 16px; color:var(--muted); font-size:12px; font-weight:900; }
    .summary-legend span { display:inline-flex; gap:7px; align-items:center; }
    .summary-legend span::before { content:""; width:12px; height:12px; border-radius:3px; background:#0b66d8; }
    .summary-legend span.late::before { background:var(--red); }
    .summary-chart { display:grid; grid-template-columns:56px 1fr; grid-template-rows:260px auto; gap:0 14px; align-items:stretch; }
    .summary-axis { grid-row:1; display:flex; flex-direction:column; justify-content:space-between; align-items:flex-end; padding:0 0 18px; color:var(--muted); font-size:12px; font-weight:800; }
    .summary-plot { grid-row:1; position:relative; display:grid; grid-auto-flow:column; grid-auto-columns:minmax(130px,1fr); align-items:end; gap:44px; padding:22px 18px 18px; border-left:1px solid var(--line); border-bottom:1px solid var(--line); background:repeating-linear-gradient(to top, transparent 0, transparent 51px, #edf1f5 52px); overflow-x:auto; }
    .summary-group { height:100%; min-width:130px; display:flex; align-items:end; justify-content:center; gap:10px; }
    .summary-bar { width:48px; min-height:3px; border-radius:6px 6px 0 0; background:#0b66d8; position:relative; cursor:pointer; box-shadow:0 10px 20px rgba(11,102,216,.18); }
    .summary-bar.late { background:var(--red); }
    .summary-bar span { position:absolute; left:50%; bottom:calc(100% + 5px); transform:translateX(-50%); color:var(--ink); font-size:13px; font-weight:950; }
    .summary-labels { grid-column:2; display:grid; grid-auto-flow:column; grid-auto-columns:minmax(130px,1fr); gap:44px; padding:8px 18px 0; overflow:hidden; }
    .summary-name { color:var(--ink); font-size:13px; font-weight:950; text-align:center; }
    .panel h2 { margin:0; padding:16px 18px 0; font-size:18px; }
    .panel-body { padding:14px 18px 18px; }
    .bars { display:grid; gap:10px; }
    .bar-row { display:grid; grid-template-columns:150px 1fr auto; gap:10px; align-items:center; font-weight:850; font-size:13px; }
    .track { height:10px; border-radius:999px; background:#edf2f6; overflow:hidden; }
    .fill { height:100%; border-radius:inherit; background:linear-gradient(90deg,var(--purple),var(--blue)); }
    .status-card { display:grid; grid-template-columns:160px 1fr; gap:18px; align-items:center; }
    .donut { position:relative; width:156px; height:156px; border-radius:50%; display:grid; place-items:center; background:conic-gradient(var(--green) 0deg, var(--green) var(--okDeg), var(--red) var(--okDeg), var(--red) 360deg); }
    .donut::before { content:""; width:104px; height:104px; border-radius:50%; background:#fff; box-shadow:inset 0 0 0 1px var(--line); }
    .donut-label { position:absolute; display:grid; gap:3px; text-align:center; font-weight:950; }
    .donut-label span { color:var(--muted); font-size:11px; text-transform:uppercase; }
    .status-legend { display:grid; gap:12px; }
    .legend-row { display:grid; grid-template-columns:12px 1fr auto; gap:9px; align-items:center; font-size:13px; font-weight:900; }
    .legend-dot { width:12px; height:12px; border-radius:50%; background:var(--green); }
    .legend-dot.late { background:var(--red); }
    .daily-chart { display:grid; gap:11px; }
    .chart-legend { display:flex; flex-wrap:wrap; justify-content:flex-end; gap:16px; margin:-4px 0 4px; color:var(--muted); font-size:12px; font-weight:950; text-transform:uppercase; }
    .chart-legend span { display:inline-flex; align-items:center; gap:6px; }
    .chart-legend span::before { content:""; width:10px; height:10px; border-radius:50%; background:var(--green); }
    .chart-legend span.late::before { background:var(--red); }
    .day-row { display:grid; grid-template-columns:90px 1fr 150px; gap:10px; align-items:center; font-size:13px; font-weight:850; }
    .day-track { height:18px; display:flex; border-radius:999px; background:#edf2f6; overflow:hidden; box-shadow:inset 0 0 0 1px rgba(0,0,0,.03); }
    .day-stack { height:100%; display:flex; min-width:10px; border-radius:inherit; overflow:hidden; }
    .day-ok, .day-late { flex:0 0 auto; min-width:0; cursor:pointer; transition:filter .15s ease; }
    .day-ok:hover, .day-late:hover { filter:brightness(.9); }
    .day-ok { background:var(--green); }
    .day-late { background:var(--red); }
    .day-row strong { text-align:right; }
    .day-counts { display:flex; justify-content:flex-end; gap:8px; align-items:center; white-space:nowrap; }
    .mini-pill { display:inline-flex; align-items:center; gap:5px; font-size:12px; font-weight:950; color:var(--muted); }
    .mini-pill::before { content:""; width:8px; height:8px; border-radius:50%; background:var(--green); }
    .mini-pill.late::before { background:var(--red); }
    .drilldown { margin-top:14px; }
    .drilldown-head { display:flex; justify-content:space-between; gap:10px; align-items:center; padding:14px 18px 0; }
    .drilldown-head h2 { padding:0; }
    .drilldown-head button { min-height:32px; padding:6px 10px; background:#f3f6f8; color:var(--ink); border-color:var(--line); }
    .drill-list { padding:12px 18px 18px; display:grid; gap:8px; }
    .drill-item { display:grid; grid-template-columns:110px 110px minmax(150px,1fr) minmax(150px,1fr) auto auto; gap:12px; align-items:center; padding:10px 12px; border:1px solid var(--line); border-radius:8px; background:#fbfcfd; font-size:13px; font-weight:850; }
    .drill-item strong { font-size:15px; }
    .table-wrap { overflow:auto; max-height:calc(100vh - 390px); }
    table { width:100%; min-width:980px; border-collapse:collapse; }
    .mini-table { min-width:0; }
    th, td { padding:10px 11px; border-bottom:1px solid var(--line); text-align:left; vertical-align:top; font-size:13px; }
    th { position:sticky; top:0; z-index:2; background:#eef3f6; color:#506071; font-size:12px; text-transform:uppercase; white-space:nowrap; }
    td.num, th.num { text-align:right; }
    .badge { display:inline-flex; min-height:27px; align-items:center; padding:5px 8px; border-radius:999px; font-weight:950; white-space:nowrap; }
    .badge.ok { background:#e7f7ee; color:#166534; }
    .badge.bad { background:#fff1f2; color:#991b1b; }
    .share-panel { margin-top:14px; background:var(--panel); border:1px solid var(--line); border-radius:8px; box-shadow:var(--shadow); overflow:hidden; }
    .share-head { display:flex; justify-content:space-between; align-items:center; gap:12px; padding:16px 18px; border-bottom:1px solid var(--line); }
    .share-head h2 { margin:0; font-size:18px; }
    .share-actions { display:flex; flex-wrap:wrap; gap:8px; justify-content:flex-end; }
    .share-actions button { background:var(--purple); border-color:transparent; color:#fff; }
    .share-actions button.secondary { background:#f3f6f8; border-color:var(--line); color:var(--ink); }
    .canvas-wrap { padding:18px; background:linear-gradient(90deg, rgba(52,16,79,.06), rgba(43,132,203,.07)), #f8fafb; overflow:auto; }
    #noteShareCanvas { display:block; width:min(100%,760px); height:auto; margin:0 auto; border-radius:8px; box-shadow:0 18px 42px rgba(23,32,51,.18); background:#fff; }
    .empty { padding:26px; color:var(--muted); font-weight:800; text-align:center; }
    @media (max-width:1100px) { .kpis { grid-template-columns:repeat(2,minmax(150px,1fr)); } .branch-grid { grid-template-columns:repeat(2,minmax(260px,1fr)); } }
    @media (max-width:900px) { .topbar { flex-direction:column; } .kpis, .branch-grid, .status-card, .drill-item { grid-template-columns:1fr; } .nav { justify-content:flex-start; } .branch-main { grid-template-columns:1fr; } }
  </style>
</head>
<body>
  <header>
    <div class="topbar">
      <div>
        <div class="brand-title"><img src="{favicon_url}" alt=""><h1>Relatorio de entrada de notas</h1></div>
        <p class="subtitle">Acompanhamento da entrada de notas fiscais.</p>
      </div>
      <nav class="nav">
        <a class="top-link" href="/home">Home</a>
        <a class="top-link" href="/dashboard">Dashboard</a>
        <a class="top-link" href="/controle-ct">Controle de CT</a>
        <a class="top-link" href="/relatorio-diario">Relatorio diario</a>
        <a class="top-link" href="/logout">Sair</a>
      </nav>
    </div>
  </header>
  <main>
    {message}
    <div class="filters">
      <label>Periodo <select id="dateModeFilter"><option value="month">Mes atual</option><option value="today">Hoje</option><option value="week">Semana atual</option><option value="last7">Ultimos 7 dias</option><option value="custom">Periodo</option><option value="all">Todas</option></select></label>
      <span class="custom-date-filter" id="customDateFilter">
        <label>Inicio <input id="dateStartFilter" type="date"></label>
        <label>Fim <input id="dateEndFilter" type="date"></label>
      </span>
      <label>Cidade <select id="cityFilter"><option value="">Todas</option></select></label>
      <label>Status <select id="statusFilter"><option value="">Todos</option><option value="ok">No prazo</option><option value="late">Fora do prazo</option></select></label>
    </div>
    <div class="tabs">
      <button type="button" class="active" data-tab="dashboard">Dashboard</button>
      <button type="button" data-tab="data">Dados</button>
    </div>
    <section id="dashboard" class="tab-view">
      <div class="kpis">
        <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 4h6"></path><path d="M9 2h6v4H9z"></path><path d="M7 4H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2h-2"></path><path d="M8 12h8"></path><path d="M8 16h6"></path></svg></div><div><span>Total de notas</span><strong id="kTotal">0</strong><small id="kTotalHint">Todas as cidades</small></div></div>
        <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"></circle><path d="m8 12 3 3 5-6"></path></svg></div><div><span>No prazo</span><strong id="kOk">0</strong><small id="kOkHint">0% do total</small></div></div>
        <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"></circle><path d="m15 9-6 6"></path><path d="m9 9 6 6"></path></svg></div><div><span>Fora do prazo</span><strong id="kLate">0</strong><small id="kLateHint">0% do total</small></div></div>
        <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"></circle><path d="M12 7v5l3 2"></path></svg></div><div><span>Tempo medio</span><strong id="kAvgEntry">-</strong><small>Tempo medio geral</small></div></div>
      </div>
      <div class="branch-grid" id="cityCards"></div>
      <div class="panel branch-summary">
        <h2>Resumo por filial</h2>
        <div class="summary-legend"><span>No prazo</span><span class="late">Fora do prazo</span></div>
        <div class="summary-chart" id="citySummary"></div>
      </div>
      <div class="panel drilldown" id="drilldownPanel" hidden></div>
      <section class="share-panel">
        <div class="share-head">
          <h2>Imagem para WhatsApp</h2>
          <div class="share-actions">
            <button type="button" id="noteShareImage">Compartilhar</button>
            <button type="button" id="noteDownloadImage" class="secondary">Baixar PNG</button>
          </div>
        </div>
        <div class="canvas-wrap">
          <canvas id="noteShareCanvas" width="1080" height="1350"></canvas>
        </div>
      </section>
    </section>
    <section id="data" class="tab-view" hidden>
      <section class="panel import-panel">
        <div>
          <strong>Importar notas fiscais</strong>
          <div class="meta">Envie a planilha atualizada para renovar os indicadores.</div>
        </div>
        <form method="post" action="/relatorio-entrada-notas/importar" enctype="multipart/form-data">
          <input type="file" name="note_file" accept=".xlsx" required>
          <button type="submit">Importar notas</button>
        </form>
      </section>
      <div class="panel">
        <div class="table-wrap">
          <table>
            <thead><tr><th>Nota fiscal</th><th>Cidade</th><th>Emissao</th><th>Entrada</th><th>Prazo limite</th><th>Status</th><th class="num">Tempo entrada</th><th class="num">Horas fora</th></tr></thead>
            <tbody id="rows"></tbody>
          </table>
        </div>
      </div>
    </section>
  </main>
  <script>
    const rows = __ROWS__;
    const fmt = new Intl.NumberFormat("pt-BR");
    const $ = (id) => document.getElementById(id);
    let activeDrilldown = null;
    function escapeHtml(value) { return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
    function parseBrDate(value) {
      const [day, month, year] = String(value || "").split("/").map(Number);
      if (!day || !month || !year) return null;
      return new Date(year, month - 1, day);
    }
    function parseBrDateTime(value) {
      const [datePart, timePart = "00:00"] = String(value || "").split(" ");
      const [day, month, year] = datePart.split("/").map(Number);
      const [hour = 0, minute = 0] = timePart.split(":").map(Number);
      return day && month && year ? new Date(year, month - 1, day, hour || 0, minute || 0) : null;
    }
    function emissionSortValue(row) {
      const date = parseBrDateTime(row.emissao);
      return date ? date.getTime() : Number.POSITIVE_INFINITY;
    }
    function sortByEmission(items) {
      return items.slice().sort((a, b) => emissionSortValue(a) - emissionSortValue(b) || String(a.nota).localeCompare(String(b.nota), "pt-BR", { numeric: true }));
    }
    function isoDate(date) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, "0");
      const day = String(date.getDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    }
    function brDate(date) {
      return `${String(date.getDate()).padStart(2, "0")}/${String(date.getMonth() + 1).padStart(2, "0")}/${date.getFullYear()}`;
    }
    function startOfDay(date) { return new Date(date.getFullYear(), date.getMonth(), date.getDate()); }
    function endOfDay(date) { return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59, 999); }
    function periodBounds() {
      const mode = $("dateModeFilter").value;
      const today = startOfDay(new Date());
      if (mode === "all") return null;
      if (mode === "today") return [today, endOfDay(today)];
      if (mode === "last7") {
        const start = new Date(today);
        start.setDate(start.getDate() - 6);
        return [start, endOfDay(today)];
      }
      if (mode === "week") {
        const start = new Date(today);
        const day = start.getDay() || 7;
        start.setDate(start.getDate() - day + 1);
        const end = new Date(start);
        end.setDate(end.getDate() + 6);
        return [start, endOfDay(end)];
      }
      if (mode === "custom") {
        const startValue = $("dateStartFilter").value;
        const endValue = $("dateEndFilter").value;
        const start = startValue ? new Date(`${startValue}T00:00:00`) : null;
        const end = endValue ? new Date(`${endValue}T23:59:59`) : null;
        if (start && end) return [start, end];
        if (start) return [start, endOfDay(today)];
        if (end) return [new Date(1900, 0, 1), end];
        return null;
      }
      return [new Date(today.getFullYear(), today.getMonth(), 1), endOfDay(new Date(today.getFullYear(), today.getMonth() + 1, 0))];
    }
    function updateCustomDateFilter() {
      $("customDateFilter").classList.toggle("is-visible", $("dateModeFilter").value === "custom");
    }
    function matchesPeriod(row) {
      const bounds = periodBounds();
      if (!bounds) return true;
      const date = parseBrDate(row.emissao);
      return date && date >= bounds[0] && date <= bounds[1];
    }
    function filteredRows() {
      const city = $("cityFilter").value;
      const status = $("statusFilter").value;
      return rows.filter((row) => matchesPeriod(row) && (!city || row.cidade === city) && (!status || row.status === status));
    }
    function bars(id, entries, color) {
      const max = Math.max(1, ...entries.map(([, value]) => value));
      $(id).innerHTML = entries.length ? entries.map(([label, value]) => `
        <div class="bar-row"><span>${escapeHtml(label)}</span><div class="track"><div class="fill" style="width:${Math.max(4, value / max * 100)}%; background:${color}"></div></div><strong>${fmt.format(value)}</strong></div>
      `).join("") : '<div class="empty">Sem dados para o filtro.</div>';
    }
    function renderStatusPanel(ok, late) {
      const total = ok + late;
      const okDeg = total ? ok / total * 360 : 0;
      $("statusBars").innerHTML = `
        <div class="status-card">
          <div class="donut" style="--okDeg:${okDeg}deg"><div class="donut-label"><strong>${percentLabel(ok, total)}</strong><span>no prazo</span></div></div>
          <div class="status-legend">
            <div class="legend-row"><span class="legend-dot"></span><span>No prazo</span><strong>${fmt.format(ok)}</strong></div>
            <div class="legend-row"><span class="legend-dot late"></span><span>Fora do prazo</span><strong>${fmt.format(late)}</strong></div>
          </div>
        </div>
      `;
    }
    function durationLabel(hours) {
      if (!Number.isFinite(hours)) return "-";
      const rounded = Math.max(0, Math.round(hours));
      const days = Math.floor(rounded / 24);
      const rest = rounded % 24;
      return days ? `${days}d ${rest}h` : `${rest}h`;
    }
    function percentLabel(ok, total) {
      if (!total) return "0%";
      return `${(ok / total * 100).toFixed(1).replace(".", ",")}%`;
    }
    function roundRect(ctx, x, y, width, height, radius) {
      const r = Math.min(radius, width / 2, height / 2);
      ctx.beginPath();
      ctx.moveTo(x + r, y);
      ctx.arcTo(x + width, y, x + width, y + height, r);
      ctx.arcTo(x + width, y + height, x, y + height, r);
      ctx.arcTo(x, y + height, x, y, r);
      ctx.arcTo(x, y, x + width, y, r);
      ctx.closePath();
    }
    function drawNoteCard(ctx, x, y, width, height, accent) {
      ctx.save();
      ctx.fillStyle = "#ffffff";
      roundRect(ctx, x, y, width, height, 18);
      ctx.fill();
      ctx.fillStyle = accent;
      roundRect(ctx, x, y, width, 8, 18);
      ctx.fill();
      ctx.restore();
    }
    function selectedLabel() {
      const labels = {
        month: "Mes atual",
        today: "Hoje",
        week: "Semana atual",
        last7: "Ultimos 7 dias",
        custom: "Periodo",
        all: "Todas as datas"
      };
      const bounds = periodBounds();
      const mode = $("dateModeFilter").value;
      const date = bounds ? `${labels[mode]} (${brDate(bounds[0])} a ${brDate(bounds[1])})` : labels[mode];
      const city = $("cityFilter").value || "Todas as cidades";
      const status = $("statusFilter").selectedOptions[0]?.textContent || "Todos";
      return `${date} | ${city} | ${status}`;
    }
    function drawNoteBrand(ctx) {
      ctx.save();
      ctx.translate(58, 48);
      ctx.fillStyle = "#2b84cb";
      ctx.beginPath();
      ctx.moveTo(0, 20);
      ctx.quadraticCurveTo(36, 2, 82, 12);
      ctx.quadraticCurveTo(78, 64, 34, 84);
      ctx.quadraticCurveTo(8, 58, 0, 20);
      ctx.fill();
      ctx.fillStyle = "#e2263c";
      ctx.globalAlpha = .82;
      ctx.beginPath();
      ctx.moveTo(44, 16);
      ctx.quadraticCurveTo(88, 28, 76, 58);
      ctx.quadraticCurveTo(56, 76, 34, 84);
      ctx.closePath();
      ctx.fill();
      ctx.globalAlpha = 1;
      ctx.fillStyle = "#ffffff";
      ctx.font = "900 16px Arial";
      ctx.fillText("GRUPO", 104, 24);
      ctx.font = "900 22px Arial";
      ctx.fillText("DISLUB", 104, 50);
      ctx.fillText("EQUADOR", 104, 76);
      ctx.restore();
    }
    function drawCanvasCircleIcon(ctx, x, y, color, text) {
      ctx.save();
      ctx.fillStyle = color;
      ctx.beginPath();
      ctx.arc(x, y, 18, 0, Math.PI * 2);
      ctx.fill();
      ctx.fillStyle = "#ffffff";
      ctx.font = "900 20px Arial";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(text, x, y + 1);
      ctx.restore();
    }
    function fillFittedText(ctx, text, x, y, maxWidth, size, weight = "900", color = "#16212d", minSize = 11) {
      let fontSize = size;
      ctx.save();
      ctx.fillStyle = color;
      do {
        ctx.font = `${weight} ${fontSize}px Arial`;
        if (ctx.measureText(text).width <= maxWidth || fontSize <= minSize) break;
        fontSize -= 1;
      } while (fontSize >= minSize);
      ctx.fillText(text, x, y);
      ctx.restore();
    }
    function drawCanvasRing(ctx, x, y, radius, ok, total) {
      const okRad = total ? ok / total * Math.PI * 2 : 0;
      ctx.save();
      ctx.lineWidth = 28;
      ctx.strokeStyle = "#e2263c";
      ctx.beginPath();
      ctx.arc(x, y, radius, 0, Math.PI * 2);
      ctx.stroke();
      ctx.strokeStyle = "#0b66d8";
      ctx.beginPath();
      ctx.arc(x, y, radius, -Math.PI / 2, -Math.PI / 2 + okRad);
      ctx.stroke();
      ctx.fillStyle = "#16212d";
      ctx.font = "900 26px Arial";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(percentLabel(ok, total), x, y - 4);
      ctx.fillStyle = "#657282";
      ctx.font = "800 13px Arial";
      ctx.fillText("No prazo", x, y + 24);
      ctx.restore();
    }
    function drawNoteShareImage() {
      const data = filteredRows();
      const ok = data.filter((row) => row.status === "ok").length;
      const late = data.filter((row) => row.status === "late").length;
      const entryHours = data.map((row) => Number(row.horasEntrada)).filter((value) => Number.isFinite(value));
      const avgEntry = entryHours.length ? entryHours.reduce((total, value) => total + value, 0) / entryHours.length : null;
      const stats = cityStats(data);
      const canvas = $("noteShareCanvas");
      canvas.width = 1080;
      canvas.height = 1560;
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#f2f5f8";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      const gradient = ctx.createLinearGradient(0, 0, canvas.width, 330);
      gradient.addColorStop(0, "#34104f");
      gradient.addColorStop(.62, "#4c176d");
      gradient.addColorStop(1, "#1b255f");
      ctx.fillStyle = gradient;
      ctx.fillRect(0, 0, canvas.width, 330);
      ctx.globalAlpha = .18;
      ctx.fillStyle = "#2b84cb";
      ctx.beginPath();
      ctx.arc(900, 70, 220, 0, Math.PI * 2);
      ctx.fill();
      ctx.globalAlpha = 1;
      drawNoteBrand(ctx);
      ctx.fillStyle = "#fff";
      ctx.font = "900 52px Arial";
      ctx.fillText("Entrada de Notas", 58, 188);
      ctx.font = "700 24px Arial";
      ctx.fillStyle = "#d7e4ea";
      ctx.fillText(selectedLabel(), 58, 234);

      const kpiY = 292;
      const kpiW = 224;
      [
        ["Total de notas", fmt.format(data.length), "#3f48cc", "N"],
        ["No prazo", fmt.format(ok), "#0b66d8", "OK"],
        ["Fora do prazo", fmt.format(late), "#e2263c", "X"],
        ["Tempo medio", avgEntry === null ? "-" : durationLabel(avgEntry), "#00856f", "T"]
      ].forEach(([label, value, accent, icon], idx) => {
        const x = 54 + idx * (kpiW + 18);
        drawNoteCard(ctx, x, kpiY, kpiW, 124, accent);
        drawCanvasCircleIcon(ctx, x + 34, kpiY + 64, accent, icon);
        fillFittedText(ctx, label.toUpperCase(), x + 68, kpiY + 42, kpiW - 86, 16, "900", "#657282", 12);
        fillFittedText(ctx, value, x + 68, kpiY + 88, kpiW - 86, 31, "900", "#16212d", 20);
      });

      const cardY = 466;
      const cardW = 312;
      stats.slice(0, 3).forEach((item, idx) => {
        const x = 54 + idx * (cardW + 18);
        drawNoteCard(ctx, x, cardY, cardW, 410, "#0b66d8");
        fillFittedText(ctx, item.city, x + 22, cardY + 46, cardW - 44, 24, "900", "#16212d", 16);
        drawCanvasRing(ctx, x + 84, cardY + 150, 58, item.ok, item.total);
        drawCanvasCircleIcon(ctx, x + 174, cardY + 114, "#0b66d8", "");
        drawCanvasCircleIcon(ctx, x + 174, cardY + 172, "#e2263c", "");
        fillFittedText(ctx, fmt.format(item.ok), x + 204, cardY + 109, 82, 22, "900", "#16212d", 16);
        fillFittedText(ctx, "no prazo", x + 204, cardY + 130, 82, 15, "900", "#16212d", 12);
        fillFittedText(ctx, fmt.format(item.late), x + 204, cardY + 167, 82, 22, "900", "#16212d", 16);
        fillFittedText(ctx, "fora", x + 204, cardY + 188, 82, 15, "900", "#16212d", 12);
        ctx.fillStyle = "#657282";
        ctx.font = "900 17px Arial";
        ctx.fillText(`Total: ${fmt.format(item.total)}`, x + 204, cardY + 224);
        ctx.strokeStyle = "#e3e9ef";
        ctx.lineWidth = 2;
        ctx.beginPath();
        ctx.moveTo(x + 22, cardY + 250);
        ctx.lineTo(x + cardW - 22, cardY + 250);
        ctx.stroke();
        ctx.fillStyle = "#e2263c";
        ctx.font = "900 16px Arial";
        ctx.fillText("Notas fora do prazo", x + 22, cardY + 282);
        ctx.fillStyle = "#657282";
        ctx.font = "900 13px Arial";
        ctx.fillText("NOTA", x + 22, cardY + 310);
        ctx.fillText("EMISSAO", x + 112, cardY + 310);
        ctx.fillText("ENTRADA", x + 216, cardY + 310);
        ctx.fillStyle = "#16212d";
        ctx.font = "800 14px Arial";
        item.lateRows.slice(0, 4).forEach((row, rowIdx) => {
          const lineY = cardY + 338 + rowIdx * 24;
          ctx.fillText(row.nota, x + 22, lineY);
          ctx.fillText(String(row.emissao).split(" ")[0] || "-", x + 112, lineY);
          ctx.fillText(String(row.entrada).split(" ")[0] || "-", x + 216, lineY);
        });
      });

      const y = 936;
      drawNoteCard(ctx, 54, y, 972, 420, "#0b66d8");
      ctx.fillStyle = "#16212d";
      ctx.font = "900 30px Arial";
      ctx.fillText("Resumo por filial", 82, y + 54);
      drawCanvasCircleIcon(ctx, 420, y + 48, "#0b66d8", "");
      drawCanvasCircleIcon(ctx, 600, y + 48, "#e2263c", "");
      ctx.fillStyle = "#657282";
      ctx.font = "900 16px Arial";
      ctx.fillText("No prazo", 448, y + 54);
      ctx.fillText("Fora do prazo", 628, y + 54);
      const max = Math.max(1, ...stats.map((item) => Math.max(item.ok, item.late)));
      const axisMax = Math.max(4, Math.ceil(max / 4) * 4);
      const plotX = 130;
      const plotY = y + 98;
      const plotW = 840;
      const plotH = 250;
      ctx.strokeStyle = "#d7e0e8";
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.moveTo(plotX, plotY);
      ctx.lineTo(plotX, plotY + plotH);
      ctx.lineTo(plotX + plotW, plotY + plotH);
      ctx.stroke();
      ctx.fillStyle = "#657282";
      ctx.font = "800 14px Arial";
      [4, 3, 2, 1, 0].forEach((tick) => {
        const value = axisMax / 4 * tick;
        const ty = plotY + plotH - (plotH * tick / 4);
        ctx.fillText(fmt.format(value), 76, ty + 4);
        ctx.strokeStyle = "#edf1f5";
        ctx.beginPath();
        ctx.moveTo(plotX, ty);
        ctx.lineTo(plotX + plotW, ty);
        ctx.stroke();
      });
      stats.slice(0, 3).forEach((item, idx) => {
        const groupX = plotX + 120 + idx * 250;
        const okH = Math.max(3, item.ok / axisMax * plotH);
        const lateH = Math.max(3, item.late / axisMax * plotH);
        ctx.fillStyle = "#0b66d8";
        roundRect(ctx, groupX, plotY + plotH - okH, 56, okH, 8);
        ctx.fill();
        ctx.fillStyle = "#e2263c";
        roundRect(ctx, groupX + 70, plotY + plotH - lateH, 56, lateH, 8);
        ctx.fill();
        ctx.fillStyle = "#16212d";
        ctx.font = "900 16px Arial";
        ctx.textAlign = "center";
        ctx.fillText(fmt.format(item.ok), groupX + 28, plotY + plotH - okH - 10);
        ctx.fillText(fmt.format(item.late), groupX + 98, plotY + plotH - lateH - 10);
        ctx.fillText(item.city, groupX + 64, plotY + plotH + 34);
        ctx.textAlign = "left";
      });
      ctx.fillStyle = "#657282";
      ctx.font = "700 18px Arial";
      ctx.fillText("Dashboard Log - Grupo Dislub Equador", 54, canvas.height - 28);
    }
    function canvasToBlob(canvas) {
      return new Promise((resolve) => canvas.toBlob(resolve, "image/png", .95));
    }
    async function downloadNoteImage() {
      drawNoteShareImage();
      const link = document.createElement("a");
      link.download = "relatorio-entrada-notas.png";
      link.href = $("noteShareCanvas").toDataURL("image/png");
      link.click();
    }
    async function shareNoteImage() {
      drawNoteShareImage();
      const blob = await canvasToBlob($("noteShareCanvas"));
      if (!blob) return downloadNoteImage();
      const file = new File([blob], "relatorio-entrada-notas.png", { type: "image/png" });
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: "Relatorio de entrada de notas" });
      } else {
        await downloadNoteImage();
      }
    }
    function cityName(row) {
      return row.cidade || "Sem cidade";
    }
    function cityOrder(name) {
      const order = { "BOA VISTA": 1, "MANAUS": 2, "ITACOATIARA": 3 };
      return order[name] || 99;
    }
    function cityStats(data) {
      const map = new Map();
      data.forEach((row) => {
        const city = cityName(row);
        const item = map.get(city) || { city, total: 0, ok: 0, late: 0, lateRows: [], okRows: [] };
        item.total += 1;
        if (row.status === "ok") {
          item.ok += 1;
          item.okRows.push(row);
        }
        if (row.status === "late") {
          item.late += 1;
          item.lateRows.push(row);
        }
        map.set(city, item);
      });
      return [...map.values()].sort((a, b) => cityOrder(a.city) - cityOrder(b.city) || a.city.localeCompare(b.city));
    }
    function renderCityCards(data) {
      const stats = cityStats(data);
      $("cityCards").innerHTML = stats.length ? stats.map((item) => {
        const okDeg = item.total ? item.ok / item.total * 360 : 0;
        const lateRows = sortByEmission(item.lateRows).slice(0, 5);
        return `
          <section class="branch-card">
            <h2 class="branch-title"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 21V5a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v16"></path><path d="M16 9h2a2 2 0 0 1 2 2v10"></path><path d="M8 7h4"></path><path d="M8 11h4"></path><path d="M8 15h4"></path><path d="M3 21h18"></path></svg>${escapeHtml(item.city)}</h2>
            <div class="branch-main">
              <div class="city-ring" style="--okDeg:${okDeg}deg">
                <div class="city-ring-label"><strong>${percentLabel(item.ok, item.total)}</strong><span>No prazo</span></div>
              </div>
              <div class="city-metrics">
                <div class="city-metric"><strong>${fmt.format(item.ok)}</strong><span>No prazo</span></div>
                <div class="city-metric late"><strong>${fmt.format(item.late)}</strong><span>Fora do prazo</span></div>
                <div class="city-total">Total: ${fmt.format(item.total)}</div>
              </div>
            </div>
            <div class="late-notes">
              <h3>Notas fora do prazo</h3>
              ${lateRows.length ? `
                <table class="mini-table">
                  <thead><tr><th>Nota</th><th>Emissao</th><th>Entrada</th></tr></thead>
                  <tbody>${lateRows.map((row) => `<tr><td>${escapeHtml(row.nota)}</td><td>${escapeHtml(String(row.emissao).split(" ")[0] || "-")}</td><td>${escapeHtml(String(row.entrada).split(" ")[0] || "-")}</td></tr>`).join("")}</tbody>
                </table>
                <button type="button" class="view-late" data-drill-city="${escapeHtml(item.city)}" data-drill-status="late">Ver todas (${fmt.format(item.late)})</button>
              ` : '<div class="empty">Sem notas fora do prazo.</div>'}
            </div>
          </section>
        `;
      }).join("") : '<div class="empty">Sem dados para o filtro.</div>';
      document.querySelectorAll("[data-drill-city]").forEach((button) => button.addEventListener("click", () => {
        activeDrilldown = { city: button.dataset.drillCity, status: button.dataset.drillStatus };
        render();
        $("drilldownPanel").scrollIntoView({ behavior: "smooth", block: "nearest" });
      }));
    }
    function renderCitySummary(data) {
      const stats = cityStats(data);
      const max = Math.max(1, ...stats.map((item) => Math.max(item.ok, item.late)));
      const step = Math.max(1, Math.ceil(max / 4));
      const axisMax = step * 4;
      $("citySummary").innerHTML = stats.length ? `
        <div class="summary-axis">${[4, 3, 2, 1, 0].map((idx) => `<span>${fmt.format(step * idx)}</span>`).join("")}</div>
        <div class="summary-plot">
          ${stats.map((item) => `
            <div class="summary-group">
              <div class="summary-bar" data-drill-city="${escapeHtml(item.city)}" data-drill-status="ok" style="height:${Math.max(2, item.ok / axisMax * 100)}%"><span>${fmt.format(item.ok)}</span></div>
              <div class="summary-bar late" data-drill-city="${escapeHtml(item.city)}" data-drill-status="late" style="height:${Math.max(2, item.late / axisMax * 100)}%"><span>${fmt.format(item.late)}</span></div>
            </div>
          `).join("")}
        </div>
        <div class="summary-labels">${stats.map((item) => `<div class="summary-name">${escapeHtml(item.city)}</div>`).join("")}</div>
      ` : '<div class="empty">Sem dados para o filtro.</div>';
      $("citySummary").onclick = (event) => {
        const bar = event.target.closest("[data-drill-city]");
        if (bar) {
          activeDrilldown = { city: bar.dataset.drillCity, status: bar.dataset.drillStatus };
          render();
          $("drilldownPanel").scrollIntoView({ behavior: "smooth", block: "nearest" });
          return;
        }
        if (activeDrilldown) {
          activeDrilldown = null;
          render();
        }
      };
    }
    function renderDrilldown(data) {
      const panel = $("drilldownPanel");
      if (!activeDrilldown) {
        panel.hidden = true;
        panel.innerHTML = "";
        return;
      }
      const statusLabel = activeDrilldown.status === "late" ? "fora do prazo" : "no prazo";
      const selected = sortByEmission(data.filter((row) => cityName(row) === activeDrilldown.city && row.status === activeDrilldown.status));
      if (!selected.length) {
        panel.hidden = true;
        panel.innerHTML = "";
        return;
      }
      panel.hidden = false;
      panel.innerHTML = `
        <div class="drilldown-head">
          <h2>${fmt.format(selected.length)} notas ${statusLabel} em ${escapeHtml(activeDrilldown.city)}</h2>
          <button type="button" id="clearDrilldown">Fechar</button>
        </div>
        <div class="drill-list">
          ${selected.map((row) => `
            <div class="drill-item">
              <strong>${escapeHtml(row.nota)}</strong>
              <span>${escapeHtml(row.cidade || "Sem cidade")}</span>
              <span>Emissao: ${escapeHtml(row.emissao)}</span>
              <span>Entrada: ${escapeHtml(row.entrada)}</span>
              <span>Tempo: ${durationLabel(Number(row.horasEntrada))}</span>
              <span>${row.status === "late" ? `${fmt.format(row.horasFora)}h fora` : "No prazo"}</span>
            </div>
          `).join("")}
        </div>
      `;
      $("clearDrilldown").addEventListener("click", () => {
        activeDrilldown = null;
        render();
      });
    }
    function render() {
      const data = filteredRows();
      const ok = data.filter((row) => row.status === "ok").length;
      const late = data.filter((row) => row.status === "late").length;
      const entryHours = data.map((row) => Number(row.horasEntrada)).filter((value) => Number.isFinite(value));
      const avgEntry = entryHours.length ? entryHours.reduce((total, value) => total + value, 0) / entryHours.length : null;
      $("kTotal").textContent = fmt.format(data.length);
      $("kOk").textContent = fmt.format(ok);
      $("kLate").textContent = fmt.format(late);
      $("kTotalHint").textContent = $("cityFilter").value || "Todas as cidades";
      $("kOkHint").textContent = `${percentLabel(ok, data.length)} do total`;
      $("kLateHint").textContent = `${percentLabel(late, data.length)} do total`;
      $("kAvgEntry").textContent = avgEntry === null ? "-" : durationLabel(avgEntry);
      renderCityCards(data);
      renderCitySummary(data);
      renderDrilldown(data);
      drawNoteShareImage();
      $("rows").innerHTML = data.length ? data.map((row) => `
        <tr>
          <td>${escapeHtml(row.nota)}</td><td>${escapeHtml(row.cidade || "Sem cidade")}</td><td>${escapeHtml(row.emissao)}</td><td>${escapeHtml(row.entrada)}</td><td>${escapeHtml(row.prazo)}</td>
          <td><span class="badge ${row.status === "ok" ? "ok" : "bad"}">${row.status === "ok" ? "No prazo" : "Fora do prazo"}</span></td>
          <td class="num">${durationLabel(Number(row.horasEntrada))}</td>
          <td class="num">${row.horasFora ? fmt.format(row.horasFora) : "-"}</td>
        </tr>
      `).join("") : '<tr><td class="empty" colspan="8">Sem dados para os filtros atuais.</td></tr>';
    }
    const today = new Date();
    $("dateStartFilter").value = isoDate(new Date(today.getFullYear(), today.getMonth(), 1));
    $("dateEndFilter").value = isoDate(new Date(today.getFullYear(), today.getMonth() + 1, 0));
    $("cityFilter").innerHTML = ['<option value="">Todas</option>', ...[...new Set(rows.map((row) => row.cidade).filter(Boolean))].sort().map((city) => `<option value="${escapeHtml(city)}">${escapeHtml(city)}</option>`)].join("");
    updateCustomDateFilter();
    ["dateModeFilter", "dateStartFilter", "dateEndFilter", "cityFilter", "statusFilter"].forEach((id) => $(id).addEventListener("input", () => {
      activeDrilldown = null;
      updateCustomDateFilter();
      render();
    }));
    document.querySelectorAll(".message").forEach((item) => {
      setTimeout(() => item.classList.add("is-hidden"), 3600);
      setTimeout(() => item.remove(), 4000);
    });
    if (window.location.search.includes("ok=") || window.location.search.includes("erro=")) {
      window.history.replaceState({}, document.title, window.location.pathname);
    }
    $("noteDownloadImage").addEventListener("click", downloadNoteImage);
    $("noteShareImage").addEventListener("click", shareNoteImage);
    document.querySelectorAll("[data-tab]").forEach((button) => button.addEventListener("click", () => {
      document.querySelectorAll("[data-tab]").forEach((item) => item.classList.toggle("active", item === button));
      document.querySelectorAll(".tab-view").forEach((view) => view.hidden = view.id !== button.dataset.tab);
    }));
    render();
  </script>
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
    os.environ.pop("BUILDING_DASHBOARD", None)
    build_dashboard.main()


def editable_rows() -> list[dict[str, object]]:
    return build_dashboard.ensure_editable_data()


def save_editable_rows(rows: list[dict[str, object]]) -> None:
    clean_rows = []
    for row in rows:
        clean_rows.append({key: row.get(key, "") for key in build_dashboard.EDITABLE_COLUMNS})
    build_dashboard.save_editable_data(clean_rows)


def capacity_rows() -> list[dict[str, object]]:
    return build_dashboard.ensure_capacity_rows()


def save_capacity_rows(rows: list[dict[str, object]]) -> None:
    clean_rows = []
    for row in rows:
        clean_rows.append({key: row.get(key, "") for key in build_dashboard.CAPACITY_COLUMNS})
    build_dashboard.save_capacity_rows(clean_rows)
    save_editable_rows(editable_rows())


def clean_conductor_name(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).upper()


def clean_conductor_freight(value: object) -> str:
    text = re.sub(r"\s+", " ", str(value or "").strip())
    normalized = normalize_header(text)
    if normalized == "cif":
        return "CIF"
    if normalized == "fob":
        return "FOB"
    if normalized == "rzd":
        return "RZD"
    if normalized == "transferencia":
        return "Transferencia"
    return text


def normalize_header(value: str) -> str:
    text = unicodedata.normalize("NFKD", str(value).strip().lower())
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", "", text)


def clean_conductor_row(row: object) -> dict[str, str]:
    if isinstance(row, dict):
        name = row.get("nome", row.get("motorista", ""))
        freight = row.get("tipoFrete", row.get("tipo_frete", ""))
    else:
        name = row
        freight = ""
    return {"nome": clean_conductor_name(name), "tipoFrete": clean_conductor_freight(freight)}


def sort_conductor_rows(rows: list[object]) -> list[dict[str, str]]:
    unique: dict[str, dict[str, str]] = {}
    for row in rows:
        clean = clean_conductor_row(row)
        if clean["nome"]:
            key = normalize_header(clean["nome"])
            if key not in unique or (not unique[key]["tipoFrete"] and clean["tipoFrete"]):
                unique[key] = clean
    return sorted(unique.values(), key=lambda item: normalize_header(item["nome"]))


HEADER_ALIASES = {
    "data": "data",
    "dtemissao": "data",
    "dataemissao": "data",
    "emissao": "data",
    "placa": "placa",
    "placa1veiculo": "placa",
    "placa1": "placa",
    "placaveiculo": "placa",
    "terminal": "terminal",
    "codterminal": "terminal",
    "codigoterminal": "terminal",
    "viagens": "viagens",
    "viagem": "viagens",
    "capacidade": "capacidade",
    "motorista1": "motorista1",
    "motorista": "motorista1",
    "nomemotorista1": "motorista1",
    "nomedomotorista1": "motorista1",
    "motoristaum": "motorista1",
    "condutor1": "motorista1",
    "condutor": "motorista1",
    "notafiscal": "notaFiscal",
    "nrnotafiscal": "notaFiscal",
    "numnotafiscal": "notaFiscal",
    "numeronotafiscal": "notaFiscal",
    "codnotafiscal": "notaFiscal",
    "codigonotafiscal": "notaFiscal",
    "nf": "notaFiscal",
    "produto": "produto",
    "descricaodoproduto": "produto",
    "descricaoproduto": "produto",
    "cliente": "cliente",
    "nomefantasiacliente": "cliente",
    "razaosocialcliente": "cliente",
    "nomecliente": "cliente",
    "municipiodestino": "municipioDestino",
    "nomemunicipiodestino": "municipioDestino",
    "municipiodedestino": "municipioDestino",
    "cidadedestino": "municipioDestino",
    "quantidade": "quantidade",
    "qtd": "quantidade",
    "cfop": "cfopDescricao",
    "cfopdescricao": "cfopDescricao",
    "descricaocfop": "cfopDescricao",
    "cfopdescrio": "cfopDescricao",
}

RETURN_CFOP_DESCRIPTIONS = {
    "devolucao de venda de combustivel ou lubrificante destinado a comercializacao",
    "devolucao de venda-cons final",
}


def row_from_import(raw: dict[str, object]) -> dict[str, object]:
    mapped = {key: "" for key in build_dashboard.EDITABLE_COLUMNS}
    for header, value in raw.items():
        key = HEADER_ALIASES.get(normalize_header(str(header)))
        if key:
            mapped[key] = "" if value is None else str(value).strip()
    if mapped["data"]:
        mapped["data"] = build_dashboard.day(str(mapped["data"])) or str(mapped["data"]).strip()
    return mapped


def is_return_row(row: dict[str, object]) -> bool:
    if build_dashboard.num(str(row.get("quantidade", 0))) < 0:
        return True
    description = normalize_header(str(row.get("cfopDescricao", "")))
    if description in {normalize_header(item) for item in RETURN_CFOP_DESCRIPTIONS}:
        return True
    if "devolu" not in description or "venda" not in description:
        return False
    return (
        "consfinal" in description
        or "combustivel" in description
        or "lubrificante" in description
        or "comercializacao" in description
    )


def import_key(row: dict[str, object]) -> tuple[str, ...]:
    note = str(row.get("notaFiscal", "")).strip()
    if note:
        return (
            "nf",
            str(row.get("data", "")).strip(),
            str(row.get("terminal", "")).strip(),
            str(row.get("placa", "")).strip().upper(),
            note,
            str(row.get("produto", "")).strip().upper(),
            str(row.get("cliente", "")).strip().upper(),
        )
    return tuple(
        ["row"]
        + [
            str(row.get(key, "")).strip().upper()
            for key in build_dashboard.EDITABLE_COLUMNS
            if key != "cfopDescricao"
        ]
    )


def import_group_key(row: dict[str, object]) -> tuple[str, str, str, str]:
    return (
        str(row.get("data", "")).strip(),
        str(row.get("terminal", "")).strip(),
        str(row.get("placa", "")).strip().upper(),
        str(row.get("motorista1", "")).strip().upper(),
    )


def unique_join(values: list[object]) -> str:
    seen: set[str] = set()
    output: list[str] = []
    for value in values:
        text = str(value or "").strip()
        if not text:
            continue
        normalized = normalize_header(text)
        if normalized in seen:
            continue
        seen.add(normalized)
        output.append(text)
    return " / ".join(output)


def collapse_import_rows(imported_rows: list[dict[str, object]]) -> list[dict[str, object]]:
    grouped: dict[tuple[str, str, str, str], dict[str, object]] = {}
    grouped_rows: dict[tuple[str, str, str, str], list[dict[str, object]]] = {}
    for row in imported_rows:
        if is_return_row(row):
            continue
        key = import_group_key(row)
        grouped_rows.setdefault(key, []).append(row)
        if key not in grouped:
            grouped[key] = dict(row)

    collapsed: list[dict[str, object]] = []
    for key, rows in grouped_rows.items():
        row = dict(grouped[key])
        row["quantidade"] = sum(build_dashboard.num(str(item.get("quantidade", 0))) for item in rows)

        trips = [int(build_dashboard.num(str(item.get("viagens", 0))) or 0) for item in rows]
        row["viagens"] = max(trips) if any(trips) else ""

        for field in ("notaFiscal", "produto", "cliente", "municipioDestino", "cfopDescricao"):
            row[field] = unique_join([item.get(field, "") for item in rows])

        capacity_values = [str(item.get("capacidade", "")).strip() for item in rows if str(item.get("capacidade", "")).strip()]
        if capacity_values:
            row["capacidade"] = capacity_values[0]
        collapsed.append(row)
    return collapsed


def merge_import_rows(current_rows: list[dict[str, object]], imported_rows: list[dict[str, object]]) -> list[dict[str, object]]:
    collapsed_imports = collapse_import_rows(imported_rows)
    imported_keys = {import_group_key(row) for row in collapsed_imports}
    merged: list[dict[str, object]] = []
    index: dict[tuple[str, ...], int] = {}
    for row in current_rows:
        if is_return_row(row):
            continue
        if import_group_key(row) in imported_keys:
            continue
        key = import_key(row)
        if key in index:
            merged[index[key]] = dict(row)
        else:
            index[key] = len(merged)
            merged.append(dict(row))
    for row in collapsed_imports:
        key = import_group_key(row)
        if key in index:
            merged[index[key]] = dict(row)
        else:
            index[key] = len(merged)
            merged.append(dict(row))
    return merged


def parse_import_file(filename: str, content: bytes) -> list[dict[str, object]]:
    suffix = Path(filename).suffix.lower()
    if suffix == ".xlsx":
        temp_path = DATA_DIR / "_import_base.xlsx"
        temp_path.write_bytes(content)
        try:
            records, _ = build_dashboard.read_xlsx(temp_path)
        finally:
            temp_path.unlink(missing_ok=True)
        return [row for row in (row_from_import(item) for item in records) if not is_return_row(row)]

    text = content.decode("utf-8-sig", errors="ignore")
    sample = text[:2048]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t") if sample.strip() else csv.excel
    except csv.Error:
        dialect = csv.excel
    reader = csv.DictReader(io.StringIO(text), dialect=dialect)
    return [row for row in (row_from_import(item) for item in reader) if not is_return_row(row)]


def template_csv() -> bytes:
    output = io.StringIO()
    writer = csv.writer(output, delimiter=";")
    writer.writerow(
        [
            "data",
            "placa",
            "terminal",
            "viagens",
            "capacidade",
            "motorista1",
            "notaFiscal",
            "produto",
            "cliente",
            "municipioDestino",
            "quantidade",
        ]
    )
    writer.writerow(["16/03/2026", "ABC1D23", "10", "1", "30000", "MOTORISTA EXEMPLO", "123456", "DIESEL S10", "CLIENTE EXEMPLO", "MANAUS", "5000"])
    return output.getvalue().encode("utf-8-sig")


def json_for_script(value: object) -> str:
    return (
        json.dumps(value, ensure_ascii=False)
        .replace("</", "<\\/")
        .replace("\u2028", "\\u2028")
        .replace("\u2029", "\\u2029")
    )


def html_escape(value: object) -> str:
    return html.escape(str(value if value is not None else ""), quote=False)


def permission_label(key: str) -> str:
    for perm_key, label, _description in PERMISSIONS:
        if perm_key == key:
            return label
    return key


def render_permission_checks(selected: set[str]) -> str:
    return "".join(
        f"""<label class="permission"><input type="checkbox" name="permissions" value="{html.escape(key)}" {"checked" if key in selected else ""}><span><strong>{html.escape(label)}</strong><span>{html.escape(description)}</span></span></label>"""
        for key, label, description in PERMISSIONS
    )


def render_user_rows(records: list[dict[str, object]]) -> str:
    if not records:
        return '<div class="empty-users">Nenhum usuario cadastrado ainda.</div>'
    rows = []
    for record in sorted(records, key=lambda item: str(item.get("username", "")).lower()):
        username = str(record.get("username", ""))
        name = str(record.get("name", "")).strip()
        active = bool(record.get("active", True))
        permissions = record.get("permissions", [])
        if not isinstance(permissions, list):
            permissions = []
        perm_tags = "".join(
            f'<span class="perm-tag">{html.escape(permission_label(str(key)))}</span>'
            for key in permissions
        ) or '<span class="hint">Sem acessos</span>'
        status_class = " off" if not active else ""
        status_text = "Ativo" if active else "Inativo"
        toggle_text = "Desativar" if active else "Ativar"
        rows.append(f"""
          <article class="user-card">
            <div class="user-main">
              <span class="user-pill">{html.escape(username)}</span>
              <div class="user-name">{html.escape(name or "Sem nome informado")}</div>
            </div>
            <div class="user-meta">
              <span class="status{status_class}">{status_text}</span>
              <div class="perm-list">{perm_tags}</div>
            </div>
            <div class="row-actions">
              <a class="button secondary" href="/usuarios?editar={quote(username)}">Editar</a>
              <form method="post" action="/usuarios" style="display:inline">
                <input type="hidden" name="action" value="toggle">
                <input type="hidden" name="username" value="{html.escape(username)}">
                <button class="secondary" type="submit">{toggle_text}</button>
              </form>
              <form method="post" action="/usuarios" style="display:inline" onsubmit="return confirm('Excluir este usuario?')">
                <input type="hidden" name="action" value="delete">
                <input type="hidden" name="username" value="{html.escape(username)}">
                <button class="danger" type="submit">Excluir</button>
              </form>
            </div>
          </article>
        """)
    return "".join(rows)


def clean_user_record(record: dict[str, object]) -> dict[str, object]:
    username = clean_username(str(record.get("username", "")))
    permissions = record.get("permissions", [])
    if not isinstance(permissions, list):
        permissions = []
    return {
        "username": username,
        "name": str(record.get("name", "")).strip(),
        "password_hash": str(record.get("password_hash", "")),
        "active": bool(record.get("active", True)),
        "permissions": sorted({str(key) for key in permissions if str(key) in ALL_PERMISSION_KEYS}),
    }


def ensure_postgres_user_table() -> None:
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS sistema_usuarios (
                    username TEXT PRIMARY KEY,
                    name TEXT NOT NULL DEFAULT '',
                    password_hash TEXT NOT NULL,
                    active BOOLEAN NOT NULL DEFAULT TRUE,
                    permissions JSONB NOT NULL DEFAULT '[]'::jsonb,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )


def postgres_user_records() -> list[dict[str, object]]:
    ensure_postgres_user_table()
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT username, name, password_hash, active, permissions::text
                FROM sistema_usuarios
                ORDER BY username
                """
            )
            records = []
            for username, name, stored_hash, active, permissions_text in cur.fetchall():
                try:
                    permissions = json.loads(permissions_text or "[]")
                except json.JSONDecodeError:
                    permissions = []
                records.append(clean_user_record({
                    "username": username,
                    "name": name,
                    "password_hash": stored_hash,
                    "active": active,
                    "permissions": permissions,
                }))
            return records


def save_postgres_user_records(records: list[dict[str, object]]) -> None:
    ensure_postgres_user_table()
    clean_records = []
    seen = set()
    for record in records:
        item = clean_user_record(record)
        username = str(item["username"])
        if not username or username in seen or username == clean_username(app_user()):
            continue
        seen.add(username)
        clean_records.append(item)
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM sistema_usuarios")
            for item in clean_records:
                cur.execute(
                    """
                    INSERT INTO sistema_usuarios (
                        username, name, password_hash, active, permissions, updated_at
                    )
                    VALUES (%s, %s, %s, %s, %s::jsonb, now())
                    """,
                    (
                        item["username"],
                        item["name"],
                        item["password_hash"],
                        item["active"],
                        json.dumps(item["permissions"], ensure_ascii=False),
                    ),
                )


def ensure_postgres_audit_table() -> None:
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS sistema_auditoria (
                    id SERIAL PRIMARY KEY,
                    occurred_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                    username TEXT NOT NULL DEFAULT '',
                    action TEXT NOT NULL DEFAULT '',
                    target TEXT NOT NULL DEFAULT '',
                    details JSONB NOT NULL DEFAULT '{}'::jsonb,
                    ip TEXT NOT NULL DEFAULT '',
                    ok BOOLEAN NOT NULL DEFAULT TRUE
                )
                """
            )


def audit_log(username: str, action: str, target: str = "", details: dict[str, object] | None = None, ip: str = "", ok: bool = True) -> None:
    safe_details = details or {}
    event = {
        "ts": dt.datetime.now(dt.timezone.utc).isoformat(),
        "user": clean_username(username) or "-",
        "action": action,
        "target": target,
        "details": safe_details,
        "ip": ip,
        "ok": ok,
    }
    print("AUDIT_LOG " + json.dumps(event, ensure_ascii=False, default=str), flush=True)
    if not build_dashboard.use_postgres():
        return
    try:
        ensure_postgres_audit_table()
        with build_dashboard.postgres_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO sistema_auditoria (username, action, target, details, ip, ok)
                    VALUES (%s, %s, %s, %s::jsonb, %s, %s)
                    """,
                    (
                        event["user"],
                        action,
                        target,
                        json.dumps(safe_details, ensure_ascii=False, default=str),
                        ip,
                        ok,
                    ),
                )
    except Exception as exc:
        print("AUDIT_LOG_ERROR " + json.dumps({"error": str(exc)}, ensure_ascii=False), flush=True)


XLSX_NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def column_number(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref.upper())
    if not letters:
        return 0
    number = 0
    for char in letters.group(0):
        number = number * 26 + ord(char) - 64
    return number


def xlsx_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    value = cell.find("a:v", XLSX_NS)
    inline = cell.find("a:is", XLSX_NS)
    if cell_type == "s" and value is not None:
        return shared_strings[int(value.text or "0")]
    if cell_type == "inlineStr" and inline is not None:
        return "".join(text.text or "" for text in inline.findall(".//a:t", XLSX_NS))
    if value is not None:
        return value.text or ""
    return ""


def xlsx_rows(path: Path, sheet_name: str) -> list[list[str]]:
    if not path.exists():
        return []
    with ZipFile(path) as xlsx:
        return xlsx_rows_from_zip(xlsx, sheet_name)


def xlsx_rows_from_zip(xlsx: ZipFile, sheet_name: str) -> list[list[str]]:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in xlsx.namelist():
            shared_root = ET.fromstring(xlsx.read("xl/sharedStrings.xml"))
            for item in shared_root.findall("a:si", XLSX_NS):
                shared_strings.append("".join(text.text or "" for text in item.findall(".//a:t", XLSX_NS)))

        workbook = ET.fromstring(xlsx.read("xl/workbook.xml"))
        rels = ET.fromstring(xlsx.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        sheet_path = ""
        for sheet in workbook.findall("a:sheets/a:sheet", XLSX_NS):
            if sheet.attrib.get("name") == sheet_name:
                rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                target = rel_map.get(rel_id, "")
                sheet_path = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
                break
        if not sheet_path:
            return []

        root = ET.fromstring(xlsx.read(sheet_path))
        rows: list[list[str]] = []
        for row in root.findall("a:sheetData/a:row", XLSX_NS):
            values: list[str] = []
            last_column = 0
            for cell in row.findall("a:c", XLSX_NS):
                column = column_number(cell.attrib.get("r", ""))
                while last_column + 1 < column:
                    values.append("")
                    last_column += 1
                values.append(xlsx_text(cell, shared_strings))
                last_column = column
            rows.append(values)
        return rows


def xlsx_rows_from_bytes(content: bytes, sheet_name: str) -> list[list[str]]:
    with ZipFile(io.BytesIO(content)) as xlsx:
        return xlsx_rows_from_zip(xlsx, sheet_name)


def xlsx_sheet_names_from_zip(xlsx: ZipFile) -> list[str]:
    workbook = ET.fromstring(xlsx.read("xl/workbook.xml"))
    return [sheet.attrib.get("name", "") for sheet in workbook.findall("a:sheets/a:sheet", XLSX_NS) if sheet.attrib.get("name")]


def excel_number(value: object) -> float | None:
    try:
        text = str(value).strip().replace(",", ".")
        return float(text) if text else None
    except ValueError:
        return None


def excel_date(value: object) -> str:
    number = excel_number(value)
    if number is None:
        return str(value or "").strip()
    return (dt.datetime(1899, 12, 30) + dt.timedelta(days=number)).strftime("%d/%m/%Y")


def excel_time(value: object) -> str:
    number = excel_number(value)
    if number is None:
        return str(value or "").strip()
    total_minutes = round((number % 1) * 24 * 60)
    hours, minutes = divmod(total_minutes, 60)
    return f"{hours:02d}:{minutes:02d}"


def excel_duration(value: object) -> tuple[str, float | None]:
    number = excel_number(value)
    if number is None:
        return str(value or "").strip(), None
    total_minutes = round(number * 24 * 60)
    hours, minutes = divmod(total_minutes, 60)
    return f"{hours:02d}:{minutes:02d}", float(total_minutes)


def parse_ct_control_rows() -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    rows = xlsx_rows(CT_CONTROL_UPLOAD, "Controle de CT")
    for raw_row in rows:
        values = raw_row + [""] * 24
        data, motorista, tipo_frete, status, viagens = [str(item).strip() for item in values[10:15]]
        chegada, entrada, saida, nota_fiscal = [str(item).strip() for item in values[15:19]]
        tempo_carregamento, carregamento_minutos = excel_duration(values[19])
        tempo_total, total_minutos = excel_duration(values[20])
        observacao = str(values[21]).strip()
        if not motorista:
            continue
        if not any([tipo_frete, status, viagens, chegada, entrada, saida, nota_fiscal, observacao]):
            continue
        if normalize_header(data) == "data" or normalize_header(motorista) == "motorista":
            continue
        records.append(
            {
                "data": excel_date(data),
                "motorista": motorista,
                "tipoFrete": tipo_frete,
                "status": status,
                "viagens": viagens,
                "chegada": excel_time(chegada),
                "entrada": excel_time(entrada),
                "saida": excel_time(saida),
                "notaFiscal": nota_fiscal,
                "tempoCarregamento": tempo_carregamento,
                "tempoCarregamentoMinutos": carregamento_minutos,
                "tempoTotal": tempo_total,
                "tempoTotalMinutos": total_minutos,
                "observacao": observacao,
            }
        )
    return records


def parse_conductor_base_rows(path: Path = CT_CONTROL_UPLOAD) -> list[dict[str, str]]:
    rows = xlsx_rows(path, "Base")
    if not rows:
        return []
    headers = [normalize_header(item) for item in rows[0]]
    try:
        name_idx = headers.index("motorista")
    except ValueError:
        return []
    freight_idx = headers.index("tipodefrete") if "tipodefrete" in headers else None
    conductors: list[dict[str, str]] = []
    for raw_row in rows[1:]:
        values = raw_row + [""] * 5
        name = values[name_idx] if name_idx < len(values) else ""
        freight = values[freight_idx] if freight_idx is not None and freight_idx < len(values) else ""
        conductors.append({"nome": name, "tipoFrete": freight})
    return sort_conductor_rows(conductors)


CT_CONTROL_COLUMNS = [
    "data",
    "motorista",
    "tipoFrete",
    "status",
    "viagens",
    "chegada",
    "entrada",
    "saida",
    "notaFiscal",
    "observacao",
]


def clean_ct_control_row(row: dict[str, object]) -> dict[str, str]:
    clean = {key: str(row.get(key, "") or "").strip() for key in CT_CONTROL_COLUMNS}
    return clean


def ensure_postgres_ct_control_table() -> None:
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS controle_ct (
                    id SERIAL PRIMARY KEY,
                    row_order INTEGER NOT NULL DEFAULT 0,
                    data TEXT NOT NULL DEFAULT '',
                    motorista TEXT NOT NULL DEFAULT '',
                    tipo_frete TEXT NOT NULL DEFAULT '',
                    status TEXT NOT NULL DEFAULT '',
                    viagens TEXT NOT NULL DEFAULT '',
                    chegada TEXT NOT NULL DEFAULT '',
                    entrada TEXT NOT NULL DEFAULT '',
                    saida TEXT NOT NULL DEFAULT '',
                    nota_fiscal TEXT NOT NULL DEFAULT '',
                    observacao TEXT NOT NULL DEFAULT '',
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )


def postgres_ct_control_rows() -> list[dict[str, str]]:
    ensure_postgres_ct_control_table()
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT data, motorista, tipo_frete, status, viagens,
                       chegada, entrada, saida, nota_fiscal, observacao
                FROM controle_ct
                ORDER BY row_order, id
                """
            )
            return [
                clean_ct_control_row(
                    {
                        "data": item[0],
                        "motorista": item[1],
                        "tipoFrete": item[2],
                        "status": item[3],
                        "viagens": item[4],
                        "chegada": item[5],
                        "entrada": item[6],
                        "saida": item[7],
                        "notaFiscal": item[8],
                        "observacao": item[9],
                    }
                )
                for item in cur.fetchall()
            ]


def save_postgres_ct_control_rows(rows: list[dict[str, object]]) -> None:
    ensure_postgres_ct_control_table()
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("TRUNCATE TABLE controle_ct RESTART IDENTITY")
            for idx, row in enumerate(rows, start=1):
                item = clean_ct_control_row(row)
                cur.execute(
                    """
                    INSERT INTO controle_ct (
                        row_order, data, motorista, tipo_frete, status, viagens,
                        chegada, entrada, saida, nota_fiscal, observacao
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        idx,
                        item["data"],
                        item["motorista"],
                        item["tipoFrete"],
                        item["status"],
                        item["viagens"],
                        item["chegada"],
                        item["entrada"],
                        item["saida"],
                        item["notaFiscal"],
                        item["observacao"],
                    ),
                )


def ct_control_rows() -> list[dict[str, str]]:
    if build_dashboard.use_postgres():
        return postgres_ct_control_rows()
    if CT_CONTROL_DATA.exists():
        try:
            rows = json.loads(CT_CONTROL_DATA.read_text(encoding="utf-8"))
            if isinstance(rows, list):
                return [clean_ct_control_row(row) for row in rows if isinstance(row, dict)]
        except json.JSONDecodeError:
            return []
    return []


def save_ct_control_rows(rows: list[dict[str, object]]) -> None:
    if build_dashboard.use_postgres():
        save_postgres_ct_control_rows(rows)
        return
    DATA_DIR.mkdir(exist_ok=True)
    clean_rows = [clean_ct_control_row(row) for row in rows if isinstance(row, dict)]
    CT_CONTROL_DATA.write_text(json.dumps(clean_rows, ensure_ascii=False, indent=2), encoding="utf-8")


def ensure_postgres_daily_observation_table() -> None:
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS relatorio_diario_observacoes (
                    observation_key TEXT PRIMARY KEY,
                    observacao TEXT NOT NULL DEFAULT '',
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )


def daily_report_observations() -> dict[str, str]:
    if not build_dashboard.use_postgres():
        return {}
    ensure_postgres_daily_observation_table()
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT observation_key, observacao
                FROM relatorio_diario_observacoes
                ORDER BY observation_key
                """
            )
            return {str(key): str(value or "") for key, value in cur.fetchall()}


def save_daily_report_observations(observations: dict[str, object]) -> None:
    if not build_dashboard.use_postgres():
        raise RuntimeError("Banco de dados nao configurado para salvar observacoes.")
    ensure_postgres_daily_observation_table()
    clean = {
        str(key): str(value or "").strip()
        for key, value in observations.items()
        if str(key).strip()
    }
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            for key, value in clean.items():
                if not value:
                    cur.execute("DELETE FROM relatorio_diario_observacoes WHERE observation_key = %s", (key,))
                    continue
                cur.execute(
                    """
                    INSERT INTO relatorio_diario_observacoes (observation_key, observacao, updated_at)
                    VALUES (%s, %s, now())
                    ON CONFLICT (observation_key)
                    DO UPDATE SET observacao = EXCLUDED.observacao, updated_at = now()
                    """,
                    (key, value),
                )


def parse_note_entry_datetime(value: object) -> dt.datetime | None:
    text = str(value or "").strip()
    if not text:
        return None
    number = excel_number(text)
    if number is not None:
        return dt.datetime(1899, 12, 30) + dt.timedelta(days=number)
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def note_entry_deadline(emissao: dt.datetime) -> dt.datetime:
    deadline_date = emissao.date() + dt.timedelta(days=2)
    return dt.datetime.combine(deadline_date, dt.time.max).replace(microsecond=0)


def format_note_date(value: dt.datetime) -> str:
    return value.strftime("%d/%m/%Y")


def format_note_datetime(value: dt.datetime) -> str:
    return value.strftime("%d/%m/%Y %H:%M")


def note_entry_view_row(row: dict[str, str]) -> dict[str, object]:
    emissao = parse_note_entry_datetime(row.get("emissao_iso", ""))
    entrada = parse_note_entry_datetime(row.get("entrada_iso", ""))
    if not emissao or not entrada:
        return {
            "nota": row.get("nota", ""),
            "emissao": "",
            "entrada": "",
            "prazo": "",
            "status": "late",
            "horasEntrada": None,
            "horasFora": 0,
        }
    prazo = note_entry_deadline(emissao)
    late = entrada.date() > prazo.date()
    horas_entrada = max(0, (entrada.date() - emissao.date()).days * 24)
    horas_fora = max(0, (entrada.date() - prazo.date()).days * 24) if late else 0
    return {
        "nota": row.get("nota", ""),
        "cidade": row.get("cidade", ""),
        "emissao": format_note_date(emissao),
        "entrada": format_note_datetime(entrada),
        "prazo": format_note_date(prazo),
        "status": "late" if late else "ok",
        "horasEntrada": horas_entrada,
        "horasFora": horas_fora,
    }


NOTE_ENTRY_CITY_BY_BRANCH = {
    "178": "ITACOATIARA",
    "171": "MANAUS",
    "182": "BOA VISTA",
}


def note_entry_city_from_branch(value: object) -> str:
    text = str(value or "").strip()
    if text.endswith(".0"):
        text = text[:-2]
    return NOTE_ENTRY_CITY_BY_BRANCH.get(text, text)


def clean_note_entry_import_row(nota: object, emissao: object, entrada: object, filial: object) -> dict[str, str] | None:
    note = str(nota or "").strip()
    doc_date = parse_note_entry_datetime(emissao)
    entry_date = parse_note_entry_datetime(entrada)
    if not note or not doc_date or not entry_date:
        return None
    return {
        "nota": note,
        "cidade": note_entry_city_from_branch(filial),
        "emissao_iso": doc_date.isoformat(timespec="minutes"),
        "entrada_iso": entry_date.isoformat(timespec="minutes"),
    }


def note_entry_action_is_413(value: object) -> bool:
    text = str(value or "").strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text == "413"


def parse_note_entry_file(content: bytes) -> list[dict[str, str]]:
    required = {
        "nrdocumento": "Nr.Documento",
        "dtdocumento": "Dt.Documento",
        "dtinclusao": "Dt.Inclusao",
        "acao": "Acao",
        "filial": "Filial",
    }
    with ZipFile(io.BytesIO(content)) as xlsx:
        for sheet_name in xlsx_sheet_names_from_zip(xlsx):
            rows = xlsx_rows_from_zip(xlsx, sheet_name)
            if not rows:
                continue
            headers = [normalize_header(item) for item in rows[0]]
            if not all(key in headers for key in required):
                continue
            indexes = {key: headers.index(key) for key in required}
            imported: list[dict[str, str]] = []
            for raw in rows[1:]:
                values = raw + [""] * (len(headers) + 1)
                if not note_entry_action_is_413(values[indexes["acao"]]):
                    continue
                item = clean_note_entry_import_row(
                    values[indexes["nrdocumento"]],
                    values[indexes["dtdocumento"]],
                    values[indexes["dtinclusao"]],
                    values[indexes["filial"]],
                )
                if item:
                    imported.append(item)
            return imported
    raise ValueError("A planilha importada nao possui as colunas necessarias.")


def ensure_postgres_note_entry_table() -> None:
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS entrada_notas (
                    id SERIAL PRIMARY KEY,
                    row_order INTEGER NOT NULL DEFAULT 0,
                    nota_fiscal TEXT NOT NULL DEFAULT '',
                    cidade TEXT NOT NULL DEFAULT '',
                    emissao_iso TEXT NOT NULL DEFAULT '',
                    entrada_iso TEXT NOT NULL DEFAULT '',
                    imported_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )
            cur.execute("ALTER TABLE entrada_notas ADD COLUMN IF NOT EXISTS cidade TEXT NOT NULL DEFAULT ''")
            cur.execute(
                """
                DELETE FROM entrada_notas newer
                USING entrada_notas older
                WHERE newer.nota_fiscal = older.nota_fiscal
                  AND newer.id > older.id
                """
            )
            cur.execute(
                """
                CREATE UNIQUE INDEX IF NOT EXISTS entrada_notas_nota_fiscal_idx
                ON entrada_notas (nota_fiscal)
                """
            )


def save_note_entry_rows(rows: list[dict[str, str]]) -> None:
    if not build_dashboard.use_postgres():
        raise RuntimeError("Banco de dados nao configurado para salvar entrada de notas.")
    ensure_postgres_note_entry_table()
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            for idx, row in enumerate(rows, start=1):
                cur.execute(
                    """
                    INSERT INTO entrada_notas (row_order, nota_fiscal, cidade, emissao_iso, entrada_iso)
                    VALUES (%s, %s, %s, %s, %s)
                    ON CONFLICT (nota_fiscal) DO UPDATE SET
                        row_order = EXCLUDED.row_order,
                        cidade = EXCLUDED.cidade,
                        emissao_iso = EXCLUDED.emissao_iso,
                        entrada_iso = EXCLUDED.entrada_iso,
                        imported_at = now()
                    """,
                    (idx, row["nota"], row.get("cidade", ""), row["emissao_iso"], row["entrada_iso"]),
                )


def note_entry_rows() -> list[dict[str, object]]:
    if not build_dashboard.use_postgres():
        return []
    ensure_postgres_note_entry_table()
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT nota_fiscal, cidade, emissao_iso, entrada_iso
                FROM entrada_notas
                ORDER BY row_order, id
                """
            )
            rows = [
                {"nota": item[0], "cidade": item[1], "emissao_iso": item[2], "entrada_iso": item[3]}
                for item in cur.fetchall()
            ]
    return [note_entry_view_row(row) for row in rows]


def measurement_deadline(start: dt.datetime) -> dt.datetime:
    deadline_date = start.date() + dt.timedelta(days=2)
    return dt.datetime.combine(deadline_date, dt.time.max).replace(microsecond=0)


def measurement_view_row(row: dict[str, str]) -> dict[str, object]:
    medicao = parse_note_entry_datetime(row.get("medicao_iso", ""))
    fechamento = parse_note_entry_datetime(row.get("fechamento_iso", ""))
    if not medicao or not fechamento:
        return {
            "seq": row.get("seq", ""),
            "filial": row.get("filial", ""),
            "terminal": row.get("terminal", ""),
            "usuarioAlteracao": row.get("usuario_alteracao", row.get("usuarioAlteracao", "")),
            "medicao": "",
            "fechamento": "",
            "prazo": "",
            "status": "late",
            "horasFechamento": None,
            "horasFora": 0,
        }
    prazo = measurement_deadline(medicao)
    late = fechamento.date() > prazo.date()
    hours_to_close = max(0, (fechamento.date() - medicao.date()).days * 24)
    hours_late = max(0, (fechamento.date() - prazo.date()).days * 24) if late else 0
    return {
        "seq": row.get("seq", ""),
        "filial": row.get("filial", ""),
        "terminal": row.get("terminal", ""),
        "usuarioAlteracao": row.get("usuario_alteracao", row.get("usuarioAlteracao", "")),
        "medicao": format_note_date(medicao),
        "fechamento": format_note_datetime(fechamento),
        "prazo": format_note_date(prazo),
        "status": "late" if late else "ok",
        "horasFechamento": hours_to_close,
        "horasFora": hours_late,
    }


def clean_measurement_import_row(seq: object, filial: object, terminal: object, nome_terminal: object, medicao: object, fechamento: object, usuario_alteracao: object = "") -> dict[str, str] | None:
    measurement_date = parse_note_entry_datetime(medicao)
    close_date = parse_note_entry_datetime(fechamento)
    if not measurement_date or not close_date:
        return None
    branch = str(filial or "").strip()
    if branch.endswith(".0"):
        branch = branch[:-2]
    terminal_code = str(terminal or "").strip()
    if terminal_code.endswith(".0"):
        terminal_code = terminal_code[:-2]
    terminal_name = str(nome_terminal or "").strip()
    terminal_label = f"{terminal_code} - {terminal_name}" if terminal_name else terminal_code
    return {
        "seq": str(seq or "").strip(),
        "filial": branch,
        "terminal": terminal_label,
        "usuario_alteracao": str(usuario_alteracao or "").strip(),
        "medicao_iso": measurement_date.isoformat(timespec="minutes"),
        "fechamento_iso": close_date.isoformat(timespec="minutes"),
    }


def parse_measurement_file(content: bytes) -> list[dict[str, str]]:
    required = {
        "seqlancamento": "Seq.Lancamento",
        "filial": "Filial",
        "dtmedicao": "Dt.Medicao",
        "terminal": "Terminal",
        "nometerminal": "Nome Terminal",
        "dtalteracao": "Dt.Alteracao",
    }
    optional_type = "tipodamedicao"
    optional_user_change = "nomeusuarioalteracao"
    with ZipFile(io.BytesIO(content)) as xlsx:
        for sheet_name in xlsx_sheet_names_from_zip(xlsx):
            rows = xlsx_rows_from_zip(xlsx, sheet_name)
            if not rows:
                continue
            headers = [normalize_header(item) for item in rows[0]]
            if not all(key in headers for key in required):
                continue
            indexes = {key: headers.index(key) for key in required}
            type_index = headers.index(optional_type) if optional_type in headers else None
            user_change_index = headers.index(optional_user_change) if optional_user_change in headers else None
            imported: list[dict[str, str]] = []
            for raw in rows[1:]:
                values = raw + [""] * (len(headers) + 1)
                if type_index is not None and "fechamento" not in normalize_header(values[type_index]):
                    continue
                item = clean_measurement_import_row(
                    values[indexes["seqlancamento"]],
                    values[indexes["filial"]],
                    values[indexes["terminal"]],
                    values[indexes["nometerminal"]],
                    values[indexes["dtmedicao"]],
                    values[indexes["dtalteracao"]],
                    values[user_change_index] if user_change_index is not None else "",
                )
                if item:
                    imported.append(item)
            return imported
    raise ValueError("A planilha importada nao possui as colunas necessarias.")


def ensure_postgres_measurement_table() -> None:
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS controle_medicao (
                    id SERIAL PRIMARY KEY,
                    row_order INTEGER NOT NULL DEFAULT 0,
                    seq_lancamento TEXT NOT NULL DEFAULT '',
                    filial TEXT NOT NULL DEFAULT '',
                    terminal TEXT NOT NULL DEFAULT '',
                    usuario_alteracao TEXT NOT NULL DEFAULT '',
                    medicao_iso TEXT NOT NULL DEFAULT '',
                    fechamento_iso TEXT NOT NULL DEFAULT '',
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )
            cur.execute("ALTER TABLE controle_medicao ADD COLUMN IF NOT EXISTS usuario_alteracao TEXT NOT NULL DEFAULT ''")
            cur.execute(
                """
                CREATE UNIQUE INDEX IF NOT EXISTS controle_medicao_seq_idx
                ON controle_medicao (seq_lancamento)
                """
            )


def save_measurement_rows(rows: list[dict[str, str]]) -> None:
    if not build_dashboard.use_postgres():
        raise RuntimeError("Controle Medicao exige banco de dados Postgres configurado.")
    ensure_postgres_measurement_table()
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("TRUNCATE TABLE controle_medicao RESTART IDENTITY")
            for idx, row in enumerate(rows, start=1):
                cur.execute(
                    """
                    INSERT INTO controle_medicao (
                        row_order, seq_lancamento, filial, terminal, usuario_alteracao, medicao_iso, fechamento_iso, updated_at
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s, now())
                    """,
                    (idx, row["seq"], row["filial"], row["terminal"], row.get("usuario_alteracao", ""), row["medicao_iso"], row["fechamento_iso"]),
                )


def measurement_rows() -> list[dict[str, object]]:
    if not build_dashboard.use_postgres():
        return []
    ensure_postgres_measurement_table()
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT seq_lancamento, filial, terminal, usuario_alteracao, medicao_iso, fechamento_iso
                FROM controle_medicao
                ORDER BY row_order, id
                """
            )
            rows = [
                {
                    "seq": item[0],
                    "filial": item[1],
                    "terminal": item[2],
                    "usuario_alteracao": item[3],
                    "medicao_iso": item[4],
                    "fechamento_iso": item[5],
                }
                for item in cur.fetchall()
            ]
    return [measurement_view_row(row) for row in rows]


CT_CONTROL_EXPORT_COLUMNS = [
    ("data", "Data"),
    ("motorista", "Motorista"),
    ("tipoFrete", "Tipo de Frete"),
    ("status", "Status"),
    ("viagens", "Viagens"),
    ("chegada", "Chegada"),
    ("entrada", "Entrada"),
    ("saida", "Saida"),
    ("notaFiscal", "Nota Fiscal"),
    ("observacao", "Observacao"),
]


def xlsx_cell_ref(row: int, column: int) -> str:
    letters = ""
    current = column
    while current:
        current, remainder = divmod(current - 1, 26)
        letters = chr(65 + remainder) + letters
    return f"{letters}{row}"


def xlsx_inline_cell(row: int, column: int, value: object) -> str:
    text = html.escape(str(value if value is not None else ""), quote=False)
    return f'<c r="{xlsx_cell_ref(row, column)}" t="inlineStr"><is><t>{text}</t></is></c>'


def ct_control_export_xlsx(rows: list[dict[str, str]]) -> bytes:
    sheet_rows = [[label for _, label in CT_CONTROL_EXPORT_COLUMNS]]
    sheet_rows.extend([[clean_ct_control_row(row).get(key, "") for key, _ in CT_CONTROL_EXPORT_COLUMNS] for row in rows])
    last_row = max(1, len(sheet_rows))
    last_col = xlsx_cell_ref(1, len(CT_CONTROL_EXPORT_COLUMNS)).rstrip("1")
    xml_rows = []
    for row_idx, values in enumerate(sheet_rows, start=1):
        cells = "".join(xlsx_inline_cell(row_idx, col_idx, value) for col_idx, value in enumerate(values, start=1))
        xml_rows.append(f'<row r="{row_idx}">{cells}</row>')
    worksheet = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:{last_col}{last_row}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>{''.join(xml_rows)}</sheetData>
</worksheet>"""
    workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Controle de CT" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""
    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
    output = io.BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as xlsx:
        xlsx.writestr("[Content_Types].xml", content_types)
        xlsx.writestr("_rels/.rels", root_rels)
        xlsx.writestr("xl/workbook.xml", workbook)
        xlsx.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        xlsx.writestr("xl/worksheets/sheet1.xml", worksheet)
    return output.getvalue()


def ensure_postgres_conductor_table() -> None:
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS condutores (
                    id SERIAL PRIMARY KEY,
                    nome TEXT NOT NULL DEFAULT '',
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )
            cur.execute("ALTER TABLE condutores ADD COLUMN IF NOT EXISTS tipo_frete TEXT NOT NULL DEFAULT ''")


def postgres_conductor_rows() -> list[dict[str, str]]:
    ensure_postgres_conductor_table()
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT nome, tipo_frete FROM condutores ORDER BY nome")
            conductors = sort_conductor_rows([{"nome": item[0], "tipoFrete": item[1]} for item in cur.fetchall()])
    return conductors


def save_postgres_conductor_rows(rows: list[object]) -> None:
    ensure_postgres_conductor_table()
    clean_rows = sort_conductor_rows(rows)
    with build_dashboard.postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("TRUNCATE TABLE condutores RESTART IDENTITY")
            for row in clean_rows:
                cur.execute("INSERT INTO condutores (nome, tipo_frete) VALUES (%s, %s)", (row["nome"], row["tipoFrete"]))


def conductor_rows() -> list[dict[str, str]]:
    if build_dashboard.use_postgres():
        return postgres_conductor_rows()
    elif CONDUCTOR_DATA.exists():
        try:
            data = json.loads(CONDUCTOR_DATA.read_text(encoding="utf-8"))
            conductors = sort_conductor_rows(data if isinstance(data, list) else [])
        except json.JSONDecodeError:
            conductors = []
    else:
        conductors = []
    if conductors:
        merged = sort_conductor_rows([*conductors, *parse_conductor_base_rows()])
        return merged
    conductors = parse_conductor_base_rows()
    if conductors:
        return conductors
    return sort_conductor_rows([{"nome": row.get("motorista", ""), "tipoFrete": row.get("tipoFrete", "")} for row in ct_control_rows()])


def save_conductor_rows(rows: list[object]) -> None:
    clean_rows = sort_conductor_rows(rows)
    if build_dashboard.use_postgres():
        save_postgres_conductor_rows(clean_rows)
        return
    DATA_DIR.mkdir(exist_ok=True)
    CONDUCTOR_DATA.write_text(json.dumps(clean_rows, ensure_ascii=False, indent=2), encoding="utf-8")


def database_status() -> dict[str, object]:
    url = build_dashboard.database_url()
    parsed = urlparse(url) if url else None
    status: dict[str, object] = {
        "database_url_defined": bool(url),
        "postgres_driver": build_dashboard.postgres_driver_available(),
        "postgres_driver_name": build_dashboard.postgres_driver_name(),
        "postgres_driver_error": build_dashboard.postgres_driver_error(),
        "python_executable": build_dashboard.python_executable(),
        "use_postgres": build_dashboard.use_postgres(),
        "host": parsed.hostname if parsed else "",
        "database": parsed.path.lstrip("/") if parsed else "",
        "table_exists": False,
        "row_count": None,
        "invalid_terminal_count": None,
        "empty_key_count": None,
        "latest_created_at": None,
        "error": "",
    }
    if not status["use_postgres"]:
        return status

    try:
        build_dashboard.ensure_postgres_table()
        with build_dashboard.postgres_connection() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT current_database(), current_user")
                current_database, current_user = cur.fetchone()
                status["current_database"] = current_database
                status["current_user"] = current_user
                cur.execute(
                    """
                    SELECT EXISTS (
                        SELECT 1
                        FROM information_schema.tables
                        WHERE table_schema = 'public'
                          AND table_name = 'dashboard_base'
                    )
                    """
                )
                status["table_exists"] = bool(cur.fetchone()[0])
                cur.execute("SELECT COUNT(*) FROM dashboard_base")
                status["row_count"] = cur.fetchone()[0]
                cur.execute("SELECT COUNT(*) FROM dashboard_base WHERE terminal NOT IN ('10', '19')")
                status["invalid_terminal_count"] = cur.fetchone()[0]
                cur.execute(
                    """
                    SELECT COUNT(*)
                    FROM dashboard_base
                    WHERE btrim(data) = '' OR btrim(placa) = '' OR btrim(terminal) = ''
                    """
                )
                status["empty_key_count"] = cur.fetchone()[0]
                cur.execute("SELECT MAX(created_at) FROM dashboard_base")
                latest = cur.fetchone()[0]
                status["latest_created_at"] = str(latest) if latest else ""
    except Exception as exc:
        status["error"] = f"{type(exc).__name__}: {exc}"
    return status


def render_sheet_parts(rows: list[dict[str, object]]) -> tuple[str, str]:
    columns = [
        ("data", "Data"),
        ("placa", "Placa"),
        ("terminal", "Terminal"),
        ("viagens", "Viagens"),
        ("capacidade", "Capacidade"),
        ("motorista1", "Motorista 1"),
        ("notaFiscal", "Nota fiscal"),
        ("produto", "Produto"),
        ("cliente", "Cliente"),
        ("municipioDestino", "Municipio destino"),
        ("quantidade", "Quantidade"),
    ]
    thead = (
        '<tr><th><input type="checkbox" id="selectAllRows" '
        'aria-label="Selecionar todas as linhas"></th>'
        + "".join(f"<th>{label}</th>" for _, label in columns)
        + "</tr>"
    )
    body_rows = []
    for idx, row in enumerate(rows):
        cells = []
        for col_idx, (key, _) in enumerate(columns):
            cells.append(
                f'<td contenteditable="true" data-key="{key}" data-col="{col_idx}">'
                f"{html_escape(row.get(key, ''))}</td>"
            )
        body_rows.append(
            f'<tr data-row="{idx}">'
            f'<td><input type="checkbox" aria-label="Selecionar linha {idx + 1}"><br>{idx + 1}</td>'
            + "".join(cells)
            + "</tr>"
        )
    return thead, "".join(body_rows)


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt: str, *args: object) -> None:
        print(f"{self.address_string()} - {fmt % args}")

    def audit(self, action: str, target: str = "", details: dict[str, object] | None = None, ok: bool = True, username: str | None = None) -> None:
        audit_log(
            username if username is not None else self.current_user(),
            action,
            target,
            details,
            self.client_address[0] if self.client_address else "",
            ok,
        )

    def filter_access_links(self, page: str) -> str:
        if not self.is_logged_in():
            return page
        permissions = user_permissions(self.current_user())
        href_permissions = {
            "/dashboard": "dashboard",
            "/editar": "editar",
            "/capacidades": "capacidades",
            "/controle-ct": "controle_ct",
            "/controle-ct/exportar": "exportar_ct",
            "/relatorio-diario": "relatorio_diario",
            "/relatorio-entrada-notas": "entrada_notas",
            "/controle-medicao": "controle_medicao",
        }
        for href, permission in href_permissions.items():
            if permission in permissions:
                continue
            page = re.sub(rf'\s*<a\b[^>]*class="[^"]*(?:top-link|side-link|card|export-link)[^"]*"[^>]*href="{re.escape(href)}"[\s\S]*?</a>', "", page)
        if not is_master_user(self.current_user()):
            page = re.sub(r'\s*<a\b[^>]*class="[^"]*(?:top-link|side-link|card|export-link)[^"]*"[^>]*href="/usuarios"[\s\S]*?</a>', "", page)
        return page

    def send_bytes(self, content: bytes, content_type: str, status: int = 200) -> None:
        if status == 200 and content_type.startswith("text/html"):
            try:
                content = self.filter_access_links(content.decode("utf-8")).encode("utf-8")
            except UnicodeDecodeError:
                pass
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        if content_type.startswith("text/html"):
            self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
            self.send_header("Pragma", "no-cache")
            self.send_header("Expires", "0")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def send_download(self, content: bytes, filename: str, content_type: str) -> None:
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
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

    def current_user(self) -> str:
        value = self.session_value()
        cookie_user, separator, _signature = value.partition(":")
        if not separator:
            return ""
        return clean_username(unquote(cookie_user))

    def is_logged_in(self) -> bool:
        value = self.session_value()
        if not (value and valid_session(value)):
            return False
        username = self.current_user()
        return is_master_user(username) or bool(find_user_record(username) and user_permissions(username))

    def require_login(self) -> bool:
        if self.is_logged_in():
            return True
        self.redirect("/login")
        return False

    def require_permission(self, permission: str) -> bool:
        if not self.require_login():
            return False
        if permission in user_permissions(self.current_user()):
            return True
        self.send_error(HTTPStatus.FORBIDDEN, "Acesso nao autorizado")
        return False

    def require_master(self) -> bool:
        if not self.require_login():
            return False
        if is_master_user(self.current_user()):
            return True
        self.send_error(HTTPStatus.FORBIDDEN, "Apenas o usuario master pode acessar esta tela")
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
        page = LOGIN_HTML.replace("{message}", message).replace("{favicon_url}", FAVICON_URL)
        self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")

    def send_home(self) -> None:
        is_master = is_master_user(self.current_user())
        visible_modules = ["dashboard", "editar", "capacidades", "relatorio_diario", "entrada_notas", "controle_medicao", "controle_ct"]
        username = self.current_user()
        allowed = user_permissions(username)
        user_record = find_user_record(username)
        display_name = str(user_record.get("name", "")).strip() if user_record else ""
        if not display_name:
            display_name = username or "usuario"
        display_name = display_user_name(display_name)
        user_link = ""
        user_card = ""
        module_count = str(sum(1 for key in visible_modules if key in allowed))
        if is_master:
            module_count = str(int(module_count) + 1)
            user_link = '<a class="side-link" href="/usuarios"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M19 8v6"></path><path d="M16 11h6"></path></svg>Usuarios</a>'
            user_card = """
          <a class="card" href="/usuarios">
            <span>Acesso</span>
            <div class="card-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M19 8v6"></path><path d="M16 11h6"></path></svg></div>
            <strong>Usuarios</strong>
            <p>Cadastre usuarios e defina quais telas e funcoes cada um pode acessar.</p>
            <div class="button">Gerenciar acessos</div>
          </a>"""
        page = (
            HOME_HTML.replace("{favicon_url}", FAVICON_URL)
            .replace("__USER_LINK__", user_link)
            .replace("__USER_CARD__", user_card)
            .replace("__MODULE_COUNT__", module_count)
            .replace("__HOME_GREETING__", "Ola")
            .replace("__HOME_USER__", html.escape(display_name))
        )
        self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")

    def send_users(self) -> None:
        params = parse_qs(urlparse(self.path).query)
        message = ""
        if not build_dashboard.use_postgres():
            message = '<div class="message error auto-dismiss">Cadastro de usuarios exige banco de dados Postgres configurado. Nenhum usuario sera salvo localmente.</div>'
        if "ok" in params:
            message = '<div class="message auto-dismiss">Usuario salvo com sucesso.</div>'
        if "removido" in params:
            message = '<div class="message auto-dismiss">Usuario removido com sucesso.</div>'
        if "erro" in params:
            message = '<div class="message error auto-dismiss">' + html.escape(params["erro"][0]) + "</div>"
        records = load_user_records()
        edit_username = clean_username(params.get("editar", [""])[0])
        edit_record = next((item for item in records if item.get("username") == edit_username), None)
        selected = set()
        username = ""
        name = ""
        active_checked = "checked"
        form_title = "Novo usuario"
        password_required = "required"
        password_hint = "Informe a senha inicial do usuario."
        original_username = ""
        if edit_record:
            username = str(edit_record.get("username", ""))
            name = str(edit_record.get("name", ""))
            active_checked = "checked" if edit_record.get("active", True) else ""
            permissions = edit_record.get("permissions", [])
            if isinstance(permissions, list):
                selected = {str(key) for key in permissions}
            form_title = "Editar usuario"
            password_required = ""
            password_hint = "Deixe em branco para manter a senha atual."
            original_username = username
        page = (
            USERS_HTML.replace("{favicon_url}", FAVICON_URL)
            .replace("{message}", message)
            .replace("{form_title}", form_title)
            .replace("{original_username}", html.escape(original_username))
            .replace("{username}", html.escape(username))
            .replace("{name}", html.escape(name))
            .replace("{password_required}", password_required)
            .replace("{password_hint}", password_hint)
            .replace("{active_checked}", active_checked)
            .replace("{permission_checks}", render_permission_checks(selected))
            .replace("{user_count}", str(len(records)))
            .replace("{user_rows}", render_user_rows(records))
        )
        self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")

    def send_dashboard(self) -> None:
        if build_dashboard.use_postgres() or not INDEX_PATH.exists():
            rebuild_dashboard()
        page = INDEX_PATH.read_text(encoding="utf-8")
        nav = """
  <nav class="nav">
    <a class="top-link" href="/home">Home</a>
    <a class="top-link" href="/editar">Editar dados</a>
    <a class="top-link" href="/controle-ct">Controle de CT</a>
    <a class="top-link" href="/relatorio-diario">Relatorio diario</a>
    <a class="top-link" href="/relatorio-entrada-notas">Entrada de notas</a>
    <a class="top-link" href="/controle-medicao">Controle Medicao</a>
    <a class="top-link" href="/capacidades">Capacidades</a>
    __USER_LINK__
    <a class="top-link" href="/logout">Sair</a>
  </nav>"""
        nav = nav.replace("__USER_LINK__", '<a class="top-link" href="/usuarios">Usuarios</a>' if is_master_user(self.current_user()) else "")
        page = page.replace('<a class="top-link" href="/editar">Atualizar dados</a>', nav)
        self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")

    def send_edit(self) -> None:
        params = parse_qs(urlparse(self.path).query)
        message = ""
        if "ok" in params:
            message = '<div class="message">Dashboard atualizado com sucesso.</div>'
        if "erro" in params:
            message = '<div class="message error">' + html.escape(params["erro"][0]) + "</div>"
        rows = editable_rows()
        thead, tbody = render_sheet_parts(rows)
        page = (
            EDIT_HTML.replace("{message}", message)
            .replace("{favicon_url}", FAVICON_URL)
            .replace("__THEAD__", thead)
            .replace("__TBODY__", tbody)
            .replace("__ROW_COUNT__", f"{len(rows):,}".replace(",", "."))
            .replace("__ROWS__", json_for_script(rows))
        )
        self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")

    def send_capacities(self) -> None:
        params = parse_qs(urlparse(self.path).query)
        message = ""
        if "ok" in params:
            message = '<div class="message">Capacidades salvas e dashboard atualizado com sucesso.</div>'
        if "erro" in params:
            message = '<div class="message error">' + html.escape(params["erro"][0]) + "</div>"
        rows = capacity_rows()
        conductors = conductor_rows()
        page = (
            CAPACITY_HTML.replace("{message}", message)
            .replace("{favicon_url}", FAVICON_URL)
            .replace("__ROW_COUNT__", f"{len(rows):,}".replace(",", "."))
            .replace("__CONDUCTOR_COUNT__", f"{len(conductors):,}".replace(",", "."))
            .replace("__ROWS__", json_for_script(rows))
            .replace("__CONDUCTORS__", json_for_script(conductors))
        )
        self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")

    def send_daily_report(self) -> None:
        data = build_dashboard.build_data()
        data["dailyObservations"] = daily_report_observations()
        page = (
            DAILY_REPORT_HTML.replace("{favicon_url}", FAVICON_URL)
            .replace("__DATA__", json_for_script(data))
        )
        self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")

    def send_note_entry_report(self) -> None:
        params = parse_qs(urlparse(self.path).query)
        message = ""
        if "ok" in params:
            message = '<div class="message">Planilha importada com sucesso.</div>'
        if "erro" in params:
            message = '<div class="message error">' + html.escape(params["erro"][0]) + "</div>"
        page = (
            NOTE_ENTRY_REPORT_HTML.replace("{message}", message)
            .replace("{favicon_url}", FAVICON_URL)
            .replace("__ROWS__", json_for_script(note_entry_rows()))
        )
        self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")

    def send_measurement_control(self) -> None:
        params = parse_qs(urlparse(self.path).query)
        message = ""
        if not build_dashboard.use_postgres():
            message = '<div class="message error">Controle Medicao exige banco de dados Postgres configurado.</div>'
        if "ok" in params:
            message = '<div class="message">Planilha importada com sucesso.</div>'
        if "erro" in params:
            message = '<div class="message error">' + html.escape(params["erro"][0]) + "</div>"
        page = (
            MEASUREMENT_CONTROL_HTML.replace("{message}", message)
            .replace("{favicon_url}", FAVICON_URL)
            .replace("__ROWS__", json_for_script(measurement_rows()))
        )
        self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")

    def send_ct_control(self) -> None:
        params = parse_qs(urlparse(self.path).query)
        message = ""
        if "ok" in params:
            message = '<div class="message auto-dismiss">Controle de CT salvo com sucesso.</div>'
        if "erro" in params:
            message = '<div class="message error auto-dismiss">Nao foi possivel salvar o Controle de CT: ' + html.escape(params["erro"][0]) + "</div>"
        rows = ct_control_rows()
        conductors = conductor_rows()
        page = (
            CT_CONTROL_OPERATION_HTML.replace("{message}", message)
            .replace("{favicon_url}", FAVICON_URL)
            .replace("__ROWS__", json_for_script(rows))
            .replace("__CONDUCTORS__", json_for_script(conductors))
        )
        self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")

    def send_ct_control_export(self) -> None:
        today = dt.datetime.now().strftime("%Y%m%d")
        self.send_download(
            ct_control_export_xlsx(ct_control_rows()),
            f"controle_ct_{today}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    def send_db_status(self) -> None:
        status = database_status()
        ok = (
            status.get("database_url_defined")
            and status.get("postgres_driver")
            and status.get("use_postgres")
            and status.get("table_exists")
            and not status.get("error")
        )
        rows = "".join(
            f"<tr><th>{html.escape(str(key))}</th><td>{html.escape(str(value))}</td></tr>"
            for key, value in status.items()
        )
        page = f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Status do Banco - Dashboard</title>
  <style>
    body {{ margin: 0; padding: 28px; background: #34104f; color: #16212d; font-family: Inter, Segoe UI, Roboto, Arial, sans-serif; }}
    main {{ max-width: 980px; margin: 0 auto; background: #fff; border-radius: 8px; padding: 24px; box-shadow: 0 18px 42px rgba(0,0,0,.18); }}
    h1 {{ margin: 0 0 8px; }}
    .badge {{ display: inline-flex; margin: 8px 0 18px; padding: 7px 10px; border-radius: 8px; color: #fff; font-weight: 900; background: {"#15803d" if ok else "#b91c1c"}; }}
    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{ padding: 10px 12px; border-bottom: 1px solid #d7e0e8; text-align: left; vertical-align: top; }}
    th {{ width: 240px; background: #f3f6f8; }}
    a {{ color: #64248c; font-weight: 900; }}
  </style>
</head>
<body>
  <main>
    <h1>Status do Banco</h1>
    <div class="badge">{"OK" if ok else "VERIFICAR"}</div>
    <table>{rows}</table>
    <p><a href="/editar">Voltar para Base editavel</a></p>
  </main>
</body>
</html>"""
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
            self.send_home()
            return
        if parsed.path == "/dashboard":
            if not self.require_permission("dashboard"):
                return
            self.send_dashboard()
            return
        if parsed.path == "/editar":
            if not self.require_permission("editar"):
                return
            self.send_edit()
            return
        if parsed.path == "/capacidades":
            if not self.require_permission("capacidades"):
                return
            self.send_capacities()
            return
        if parsed.path == "/relatorio-diario":
            if not self.require_permission("relatorio_diario"):
                return
            self.send_daily_report()
            return
        if parsed.path == "/relatorio-entrada-notas":
            if not self.require_permission("entrada_notas"):
                return
            self.send_note_entry_report()
            return
        if parsed.path == "/controle-medicao":
            if not self.require_permission("controle_medicao"):
                return
            self.send_measurement_control()
            return
        if parsed.path == "/controle-ct":
            if not self.require_permission("controle_ct"):
                return
            self.send_ct_control()
            return
        if parsed.path == "/controle-ct/exportar":
            if not self.require_permission("exportar_ct"):
                return
            self.send_ct_control_export()
            return
        if parsed.path == "/usuarios":
            if not self.require_master():
                return
            self.send_users()
            return
        if parsed.path == "/db-status":
            if not self.require_login():
                return
            self.send_db_status()
            return
        if parsed.path == "/template.csv":
            if not self.require_login():
                return
            self.send_download(template_csv(), "template_dashboard.csv", "text/csv; charset=utf-8")
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/login":
            params = self.body_params()
            user = params.get("user", [""])[0].strip()
            password = params.get("password", [""])[0].strip()
            if authenticate_user(user, password):
                self.audit("login_sucesso", "auth", {"usuario": clean_username(user)}, username=clean_username(user))
                self.send_response(HTTPStatus.SEE_OTHER)
                self.send_header("Location", "/home")
                self.set_session_cookie(clean_username(user))
                self.end_headers()
                return
            self.audit("login_falha", "auth", {"usuario": clean_username(user)}, ok=False, username=clean_username(user))
            self.send_login('<div class="error">Usuario ou senha invalidos.</div>')
            return

        if parsed.path == "/importar":
            if not self.require_permission("editar"):
                return
            try:
                length = int(self.headers.get("Content-Length", "0"))
                if length > MAX_UPLOAD_BYTES:
                    raise ValueError("Arquivo muito grande")
                body = self.rfile.read(length)
                files = parse_multipart(self.headers.get("Content-Type", ""), body)
                filename, content = files.get("base_file", ("", b""))
                if not filename or not content:
                    raise ValueError("Selecione um arquivo CSV ou XLSX")
                imported_rows = parse_import_file(filename, content)
                if not imported_rows:
                    raise ValueError("Nenhuma linha encontrada no arquivo")
                save_editable_rows(merge_import_rows(editable_rows(), imported_rows))
                rebuild_dashboard()
            except Exception as exc:
                self.audit("importar_base_falha", "base_editavel", {"erro": str(exc)}, ok=False)
                self.redirect("/editar?erro=" + quote(str(exc)))
                return
            self.audit("importar_base", "base_editavel", {"arquivo": filename, "linhas_importadas": len(imported_rows)})
            self.redirect("/editar?ok=1")
            return

        if parsed.path == "/capacidades":
            if not self.require_permission("capacidades"):
                return
            try:
                params = self.body_params()
                rows = json.loads(params.get("rows_json", ["[]"])[0])
                conductors = json.loads(params.get("conductors_json", ["[]"])[0])
                if not isinstance(rows, list):
                    raise ValueError("Cadastro invalido")
                if not isinstance(conductors, list):
                    raise ValueError("Cadastro de condutores invalido")
                save_capacity_rows(rows)
                save_conductor_rows(conductors)
                rebuild_dashboard()
            except Exception as exc:
                self.audit("salvar_capacidades_falha", "capacidades", {"erro": str(exc)}, ok=False)
                self.redirect("/capacidades?erro=" + quote(str(exc)))
                return
            self.audit("salvar_capacidades", "capacidades", {"placas": len(rows), "motoristas": len(conductors)})
            self.redirect("/capacidades?ok=1")
            return

        if parsed.path == "/relatorio-diario/observacoes":
            if not self.require_permission("relatorio_diario"):
                return
            try:
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8", errors="ignore") or "{}")
                observations = payload.get("observations", {})
                if not isinstance(observations, dict):
                    raise ValueError("Observacoes invalidas")
                save_daily_report_observations(observations)
            except Exception as exc:
                self.audit("salvar_observacoes_falha", "relatorio_diario", {"erro": str(exc)}, ok=False)
                self.send_bytes(
                    json.dumps({"ok": False, "error": str(exc)}).encode("utf-8"),
                    "application/json; charset=utf-8",
                    HTTPStatus.BAD_REQUEST,
                )
                return
            self.audit("salvar_observacoes", "relatorio_diario", {"observacoes": len(observations)})
            self.send_bytes(json.dumps({"ok": True}).encode("utf-8"), "application/json; charset=utf-8")
            return

        if parsed.path == "/relatorio-entrada-notas/importar":
            if not self.require_permission("entrada_notas"):
                return
            try:
                length = int(self.headers.get("Content-Length", "0"))
                if length > MAX_UPLOAD_BYTES:
                    raise ValueError("Arquivo muito grande")
                body = self.rfile.read(length)
                files = parse_multipart(self.headers.get("Content-Type", ""), body)
                filename, content = files.get("note_file", ("", b""))
                if not filename or not content:
                    raise ValueError("Selecione um arquivo XLSX")
                if Path(filename).suffix.lower() != ".xlsx":
                    raise ValueError("Importe apenas arquivo XLSX")
                imported_rows = parse_note_entry_file(content)
                if not imported_rows:
                    raise ValueError("Nenhuma nota encontrada na planilha")
                save_note_entry_rows(imported_rows)
            except Exception as exc:
                self.audit("importar_notas_falha", "entrada_notas", {"erro": str(exc)}, ok=False)
                self.redirect("/relatorio-entrada-notas?erro=" + quote(str(exc)))
                return
            self.audit("importar_notas", "entrada_notas", {"arquivo": filename, "notas": len(imported_rows)})
            self.redirect("/relatorio-entrada-notas?ok=1")
            return

        if parsed.path == "/controle-medicao/importar":
            if not self.require_permission("controle_medicao"):
                return
            try:
                length = int(self.headers.get("Content-Length", "0"))
                if length > MAX_UPLOAD_BYTES:
                    raise ValueError("Arquivo muito grande")
                body = self.rfile.read(length)
                files = parse_multipart(self.headers.get("Content-Type", ""), body)
                filename, content = files.get("measurement_file", ("", b""))
                if not filename or not content:
                    raise ValueError("Selecione um arquivo XLSX")
                if Path(filename).suffix.lower() != ".xlsx":
                    raise ValueError("Importe apenas arquivo XLSX")
                imported_rows = parse_measurement_file(content)
                if not imported_rows:
                    raise ValueError("Nenhuma medicao encontrada na planilha")
                save_measurement_rows(imported_rows)
            except Exception as exc:
                self.audit("importar_medicao_falha", "controle_medicao", {"erro": str(exc)}, ok=False)
                self.redirect("/controle-medicao?erro=" + quote(str(exc)))
                return
            self.audit("importar_medicao", "controle_medicao", {"arquivo": filename, "medicoes": len(imported_rows)})
            self.redirect("/controle-medicao?ok=1")
            return

        if parsed.path == "/controle-ct":
            if not self.require_permission("controle_ct"):
                return
            try:
                params = self.body_params()
                rows = json.loads(params.get("rows_json", ["[]"])[0])
                if not isinstance(rows, list):
                    raise ValueError("Controle invalido")
                save_ct_control_rows(rows)
            except Exception as exc:
                self.audit("salvar_controle_ct_falha", "controle_ct", {"erro": str(exc)}, ok=False)
                self.redirect("/controle-ct?erro=" + quote(str(exc)))
                return
            self.audit("salvar_controle_ct", "controle_ct", {"linhas": len(rows)})
            self.redirect("/controle-ct?ok=1")
            return

        if parsed.path == "/usuarios":
            if not self.require_master():
                return
            username = ""
            original_username = ""
            permissions: list[str] = []
            password = ""
            try:
                params = self.body_params()
                action = params.get("action", ["save"])[0]
                records = load_user_records()
                username = clean_username(params.get("username", [""])[0])
                if action == "delete":
                    save_user_records([item for item in records if item.get("username") != username])
                    self.audit("excluir_usuario", "usuarios", {"usuario_alvo": username})
                    self.redirect("/usuarios?removido=1")
                    return
                if action == "toggle":
                    changed = False
                    new_active = False
                    for item in records:
                        if item.get("username") == username:
                            item["active"] = not bool(item.get("active", True))
                            new_active = bool(item["active"])
                            changed = True
                    if not changed:
                        raise ValueError("Usuario nao encontrado")
                    save_user_records(records)
                    self.audit("alterar_status_usuario", "usuarios", {"usuario_alvo": username, "ativo": new_active})
                    self.redirect("/usuarios?ok=1")
                    return
                original_username = clean_username(params.get("original_username", [""])[0])
                if not username:
                    raise ValueError("Informe o usuario")
                if username == clean_username(app_user()):
                    raise ValueError("O usuario master principal e controlado pelas variaveis do sistema")
                permissions = sorted({key for key in params.get("permissions", []) if key in ALL_PERMISSION_KEYS})
                password = params.get("password", [""])[0]
                existing = next((item for item in records if item.get("username") == original_username), None)
                if not existing and any(item.get("username") == username for item in records):
                    raise ValueError("Ja existe um usuario com esse login")
                if existing and username != original_username and any(item.get("username") == username for item in records):
                    raise ValueError("Ja existe um usuario com esse login")
                if not existing and not password:
                    raise ValueError("Informe a senha do usuario")
                user_record = {
                    "username": username,
                    "name": params.get("name", [""])[0].strip(),
                    "password_hash": password_hash(password) if password else str(existing.get("password_hash", "")) if existing else "",
                    "active": "active" in params,
                    "permissions": permissions,
                }
                if existing:
                    records = [item for item in records if item.get("username") != original_username]
                records.append(user_record)
                save_user_records(records)
            except Exception as exc:
                self.audit("salvar_usuario_falha", "usuarios", {"usuario_alvo": username, "erro": str(exc)}, ok=False)
                self.redirect("/usuarios?erro=" + quote(str(exc)))
                return
            self.audit(
                "salvar_usuario",
                "usuarios",
                {
                    "usuario_alvo": username,
                    "usuario_anterior": original_username,
                    "ativo": "active" in params,
                    "permissoes": permissions,
                    "senha_alterada": bool(password),
                },
            )
            self.redirect("/usuarios?ok=1")
            return

        if parsed.path != "/editar":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        if not self.require_permission("editar"):
            return

        try:
            params = self.body_params()
            rows = json.loads(params.get("rows_json", ["[]"])[0])
            if not isinstance(rows, list):
                raise ValueError("Base invalida")
            save_editable_rows(rows)
            rebuild_dashboard()
        except Exception as exc:
            self.audit("salvar_base_falha", "base_editavel", {"erro": str(exc)}, ok=False)
            self.redirect("/editar?erro=" + quote(str(exc)))
            return
        self.audit("salvar_base", "base_editavel", {"linhas": len(rows)})
        self.redirect("/editar?ok=1")


def main() -> None:
    os.environ.pop("BUILDING_DASHBOARD", None)
    DATA_DIR.mkdir(exist_ok=True)
    build_dashboard.ensure_database_storage()
    if build_dashboard.use_postgres():
        ensure_postgres_audit_table()
        ensure_postgres_user_table()
        ensure_postgres_conductor_table()
        ensure_postgres_daily_observation_table()
        ensure_postgres_note_entry_table()
        ensure_postgres_measurement_table()
    if build_dashboard.use_postgres() or not INDEX_PATH.exists():
        rebuild_dashboard()
    port = int(os.environ.get("PORT", "8000"))
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"Servidor rodando em http://0.0.0.0:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()

