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
import unicodedata
from http import HTTPStatus
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse
from zipfile import ZipFile
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


def app_user() -> str:
    return os.environ.get("DASH_USER", "admin").strip()


def app_password() -> str:
    return os.environ.get("DASH_PASSWORD", "admin").strip()


def app_secret() -> str:
    return os.environ.get("DASH_SECRET", "dashboard-log-secret")


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
        <a class="side-link" href="/controle-ct"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M4 7h16"></path><path d="M4 12h16"></path><path d="M4 17h10"></path><path d="M17 15l2 2 4-4"></path></svg>Controle de CT</a>
      </nav>
      <div class="sidebar-footer">Dislub Equador<br>Ambiente de acompanhamento logistico.</div>
    </aside>
    <div class="content">
      <header class="topbar">
        <div>
          <h1>Home operacional</h1>
          <p class="subtitle">Acesse os modulos de acompanhamento, cadastro e controle.</p>
        </div>
        <a class="logout" href="/logout">Sair</a>
      </header>
      <main>
        <section class="summary">
          <div class="metric"><span>Modulos ativos</span><strong>5</strong></div>
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
        </section>
      </main>
    </div>
  </div>
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
        <div class="hint">Colunas usadas pelo dashboard: data, placa, terminal, viagens, capacidade, nota fiscal, produto, cliente e quantidade. Terminal deve ser 10 para Equador ou 19 para Ipiranga.</div>
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
          <button type="button" id="editModeToggle" class="secondary edit-toggle" title="Alternar modo de edição" aria-label="Editar">
            <span class="edit-icon" aria-hidden="true"></span>
            <span class="edit-label">Editar</span>
          </button>
          <button type="button" id="markEntry" class="secondary">Marcar entrada</button>
          <button type="button" id="markExit" class="secondary">Marcar saida</button>
          <button type="button" id="deleteRows" class="secondary">Excluir</button>
          <button type="submit">Salvar controle</button>
        </div>
        <div class="filters">
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
      const keys = rows.map(tripKey);
      const totals = keys.reduce((map, key) => {
        if (key) map.set(key, (map.get(key) || 0) + 1);
        return map;
      }, new Map());
      rows = rows.map((row, index) => {
        const clean = cleanRow(row);
        clean.viagens = keys[index] ? String(totals.get(keys[index]) || 1) : "";
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
      button.title = editMode ? "Visualizar" : "Alternar modo de edição";
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
          && (!status || row.status === status)
          && (!freight || row.tipoFrete === freight);
      });
    }
    function renderCounters() {
      const active = rows.map(cleanRow);
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
      $("searchFilter").value = "";
      $("statusFilter").value = "";
      $("freightFilter").value = "";
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
      if (!remove.size) return;
      rows = rows.filter((_, index) => !remove.has(index));
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
    ["searchFilter", "statusFilter", "freightFilter"].forEach((id) => $(id).addEventListener("input", render));
    $("ctForm").addEventListener("submit", () => {
      syncFromTableIfReady();
      recalculateTrips();
      $("rowsJson").value = JSON.stringify(rows);
    });
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
    table { width: 100%; border-collapse: collapse; min-width: 980px; }
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
        <a class="top-link" href="/logout">Sair</a>
      </nav>
    </div>
    <div class="hero-grid">
      <div class="report-date" id="reportDate">-</div>
      <div class="filters">
        <label>Data
          <select id="dateSelect"></select>
        </label>
        <label>Terminal
          <select id="terminalSelect">
            <option value="">Todos</option>
            <option value="10">Equador</option>
            <option value="19">Ipiranga</option>
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
        <h2>Detalhamento das viagens</h2>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Placa</th>
                <th>Motorista</th>
                <th>Terminal</th>
                <th class="num">Viagens</th>
                <th class="num">Capacidade</th>
                <th class="num">Volume</th>
                <th class="num">Notas</th>
                <th>Produtos</th>
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
    const fmt = new Intl.NumberFormat("pt-BR");
    const volume = (value) => `${fmt.format(Math.round(value / 1000))} mil`;
    const $ = (id) => document.getElementById(id);
    const logoUrl = "{favicon_url}";

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

    function filteredRows() {
      const date = $("dateSelect").value;
      const terminal = $("terminalSelect").value;
      return rows.filter((row) => row.data === date && (!terminal || row.terminal === terminal));
    }

    function groupedRowsByPlate(data) {
      const terminalOrder = { Equador: 1, Ipiranga: 2 };
      const groups = new Map();
      data.forEach((row) => {
        const key = row.placa;
        const current = groups.get(key) || {
          ...row,
          terminal: "",
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
        return {
          ...row,
          motorista: motoristas.join(" / "),
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
        $("reportRows").innerHTML = `<tr><td colspan="8" class="empty">Sem dados para o filtro.</td></tr>`;
        return;
      }
      $("reportRows").innerHTML = detailData
        .slice()
        .sort((a, b) => b.viagens - a.viagens || b.quantidade - a.quantidade || a.placa.localeCompare(b.placa))
        .map((row) => `
          <tr>
            <td><span class="pill">${row.placa}</span></td>
            <td>${row.motorista || "-"}</td>
            <td>${row.terminalNome}</td>
            <td class="num">${fmt.format(row.viagens)}</td>
            <td class="num">${volume(row.capacidade)}</td>
            <td class="num">${volume(row.quantidade)}</td>
            <td class="num">${fmt.format(row.notas)}</td>
            <td>${row.mixProdutos}</td>
          </tr>
        `).join("");
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
      const productWidth = 214;
      const rowMetrics = report.detailData.map((row) => {
        const lines = wrapText(ctx, row.mixProdutos, productWidth).slice(0, 3);
        const driverLines = wrapText(ctx, row.motorista || "-", 170).slice(0, 2);
        return { row, lines, driverLines, height: Math.max(58, 34 + Math.max(lines.length, driverLines.length) * 20) };
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
      ctx.fillText("MOTORISTA", 218, y + 94);
      ctx.fillText("TERM.", 408, y + 94);
      ctx.fillText("VIAG.", 500, y + 94);
      ctx.fillText("CAP.", 588, y + 94);
      ctx.fillText("VOL.", 698, y + 94);
      ctx.fillText("PRODUTOS", 790, y + 94);
      ctx.strokeStyle = "#d7e0e8";
      ctx.beginPath();
      ctx.moveTo(82, y + 112);
      ctx.lineTo(998, y + 112);
      ctx.stroke();

      let tableY = y + 140;
      rowMetrics.forEach(({ row, lines, driverLines, height }, idx) => {
        if (idx % 2 === 0) {
          ctx.fillStyle = "#f8fafb";
          roundRect(ctx, 78, tableY, 916, height - 8, 8);
          ctx.fill();
        }
        ctx.fillStyle = "#16212d";
        ctx.font = "900 20px Arial";
        ctx.fillText(row.placa, 88, tableY + 29);
        ctx.font = "800 19px Arial";
        driverLines.forEach((line, lineIdx) => {
          ctx.fillText(line, 218, tableY + 20 + lineIdx * 20);
        });
        ctx.fillText(row.terminalShort || row.terminalNome.slice(0, 3), 414, tableY + 29);
        ctx.fillText(fmt.format(row.viagens), 516, tableY + 29);
        ctx.fillText(volume(row.capacidade).replace(" mil", "k"), 592, tableY + 29);
        ctx.fillText(volume(row.quantidade).replace(" mil", "k"), 700, tableY + 29);
        ctx.fillStyle = "#657282";
        ctx.font = "700 17px Arial";
        lines.forEach((line, lineIdx) => {
          ctx.fillText(line, 790, tableY + 20 + lineIdx * 20);
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
      link.download = `relatorio-diario-${$("dateSelect").value.replaceAll("/", "-")}.png`;
      link.href = canvas.toDataURL("image/png");
      link.click();
    }

    async function shareImage() {
      await drawShareImage();
      const blob = await canvasToBlob($("shareCanvas"));
      if (!blob) return downloadImage();
      const file = new File([blob], `relatorio-diario-${$("dateSelect").value.replaceAll("/", "-")}.png`, { type: "image/png" });
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
      $("reportDate").textContent = $("dateSelect").value ? dateLabel($("dateSelect").value) : "-";
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
    $("dateSelect").addEventListener("change", render);
    $("terminalSelect").addEventListener("change", render);
    $("generateImage").addEventListener("click", (event) => withButtonLoading(event.currentTarget, "Gerando...", drawShareImage));
    $("downloadImage").addEventListener("click", (event) => withButtonLoading(event.currentTarget, "Baixando...", downloadImage));
    $("shareImage").addEventListener("click", (event) => withButtonLoading(event.currentTarget, "Compartilhando...", shareImage));
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


def merge_import_rows(current_rows: list[dict[str, object]], imported_rows: list[dict[str, object]]) -> list[dict[str, object]]:
    merged: list[dict[str, object]] = []
    index: dict[tuple[str, ...], int] = {}
    for row in [*current_rows, *imported_rows]:
        if is_return_row(row):
            continue
        key = import_key(row)
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
            "quantidade",
        ]
    )
    writer.writerow(["16/03/2026", "ABC1D23", "10", "1", "30000", "MOTORISTA EXEMPLO", "123456", "DIESEL S10", "CLIENTE EXEMPLO", "5000"])
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

    def send_bytes(self, content: bytes, content_type: str, status: int = 200) -> None:
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
        page = LOGIN_HTML.replace("{message}", message).replace("{favicon_url}", FAVICON_URL)
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
    <a class="top-link" href="/capacidades">Capacidades</a>
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
        page = (
            DAILY_REPORT_HTML.replace("{favicon_url}", FAVICON_URL)
            .replace("__DATA__", json_for_script(data))
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
            page = HOME_HTML.replace("{favicon_url}", FAVICON_URL)
            self.send_bytes(page.encode("utf-8"), "text/html; charset=utf-8")
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
        if parsed.path == "/capacidades":
            if not self.require_login():
                return
            self.send_capacities()
            return
        if parsed.path == "/relatorio-diario":
            if not self.require_login():
                return
            self.send_daily_report()
            return
        if parsed.path == "/controle-ct":
            if not self.require_login():
                return
            self.send_ct_control()
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
            if user == app_user() and password == app_password():
                self.send_response(HTTPStatus.SEE_OTHER)
                self.send_header("Location", "/home")
                self.set_session_cookie(user)
                self.end_headers()
                return
            self.send_login('<div class="error">Usuario ou senha invalidos.</div>')
            return

        if parsed.path == "/importar":
            if not self.require_login():
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
                self.redirect("/editar?erro=" + quote(str(exc)))
                return
            self.redirect("/editar?ok=1")
            return

        if parsed.path == "/capacidades":
            if not self.require_login():
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
                self.redirect("/capacidades?erro=" + quote(str(exc)))
                return
            self.redirect("/capacidades?ok=1")
            return

        if parsed.path == "/controle-ct":
            if not self.require_login():
                return
            try:
                params = self.body_params()
                rows = json.loads(params.get("rows_json", ["[]"])[0])
                if not isinstance(rows, list):
                    raise ValueError("Controle invalido")
                save_ct_control_rows(rows)
            except Exception as exc:
                self.redirect("/controle-ct?erro=" + quote(str(exc)))
                return
            self.redirect("/controle-ct?ok=1")
            return

        if parsed.path != "/editar":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        if not self.require_login():
            return

        try:
            params = self.body_params()
            rows = json.loads(params.get("rows_json", ["[]"])[0])
            if not isinstance(rows, list):
                raise ValueError("Base invalida")
            save_editable_rows(rows)
            rebuild_dashboard()
        except Exception as exc:
            self.redirect("/editar?erro=" + quote(str(exc)))
            return
        self.redirect("/editar?ok=1")


def main() -> None:
    os.environ.pop("BUILDING_DASHBOARD", None)
    DATA_DIR.mkdir(exist_ok=True)
    build_dashboard.ensure_database_storage()
    if build_dashboard.use_postgres():
        ensure_postgres_conductor_table()
    if build_dashboard.use_postgres() or not INDEX_PATH.exists():
        rebuild_dashboard()
    port = int(os.environ.get("PORT", "8000"))
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"Servidor rodando em http://0.0.0.0:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
