from __future__ import annotations

import hashlib
import hmac
import html
import json
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
FAVICON_URL = "https://pages.greatpages.com.br/www.dislubequador.com.br/1777495651/imagens/mobile/3562683_1_177616861364933621_m.svg"


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
  <link rel="icon" href="{favicon_url}" type="image/svg+xml">
  <title>Login - Dashboard Log</title>
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
      background:
        linear-gradient(135deg, rgba(255,255,255,.03) 0 1px, transparent 1px 48px),
        radial-gradient(720px circle at var(--mx, 76%) var(--my, 24%), rgba(43, 132, 203, .28), transparent 58%),
        radial-gradient(520px circle at 18% 78%, rgba(226, 38, 60, .18), transparent 62%),
        linear-gradient(135deg, #34104f, #4c176d 52%, #1b255f),
        var(--bg);
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
      width: 150px;
      height: 58px;
      display: flex;
      align-items: center;
      gap: 10px;
      color: #fff;
      font-weight: 900;
      text-transform: uppercase;
    }
    .brand-mark img {
      width: 46px;
      height: 46px;
      object-fit: contain;
      filter: drop-shadow(0 10px 18px rgba(0, 0, 0, .22));
    }
    .brand-mark span {
      line-height: 1.05;
      letter-spacing: .02em;
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
    .scene {
      position: absolute;
      right: 30px;
      bottom: 30px;
      width: clamp(118px, 12vw, 160px);
      aspect-ratio: 1;
      pointer-events: none;
      opacity: .90;
      transition: filter .2s ease;
      z-index: 1;
    }
    .orbit {
      position: absolute;
      inset: 1%;
      border: 1px solid rgba(255, 255, 255, .16);
      border-radius: 50%;
    }
    .orbit:nth-child(1) { transform: rotate(18deg) scaleX(1.16); }
    .orbit:nth-child(2) { transform: rotate(-26deg) scaleY(.72); }
    .bot {
      position: absolute;
      left: 50%;
      top: 50%;
      width: 62%;
      aspect-ratio: .78;
      transform: translate(-50%, -50%);
      border-radius: 42% 42% 30% 30%;
      background:
        radial-gradient(circle at 34% 18%, rgba(255,255,255,.98), transparent 25%),
        linear-gradient(155deg, #ffffff 0%, #e6f0f3 48%, #bed0d7 100%);
      border: 1px solid rgba(255, 255, 255, .9);
      box-shadow: 0 30px 72px rgba(0, 0, 0, .26), inset 0 1px 0 rgba(255, 255, 255, .9);
    }
    .bot::before {
      content: "";
      position: absolute;
      left: 22%;
      right: 22%;
      top: -9%;
      height: 11%;
      border-radius: 999px;
      background: linear-gradient(90deg, #2b84cb, #e2263c);
      box-shadow: 0 0 22px rgba(43, 132, 203, .52);
    }
    .bot::after {
      content: "";
      position: absolute;
      left: 28%;
      right: 28%;
      bottom: 13%;
      height: 12%;
      border-radius: 999px;
      background: linear-gradient(180deg, rgba(100, 36, 140, .16), rgba(43, 132, 203, .08));
      border: 1px solid rgba(100, 36, 140, .22);
    }
    .face {
      position: absolute;
      left: 11%;
      right: 11%;
      top: 22%;
      height: 39%;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 14%;
      border-radius: 30px;
      background: linear-gradient(180deg, #34104f, #1b255f);
      box-shadow: inset 0 0 22px rgba(43, 132, 203, .22);
    }
    .arm, .foot {
      position: absolute;
      background: #2b84cb;
      box-shadow: 0 5px 12px rgba(0, 0, 0, .12);
    }
    .arm {
      top: 49%;
      width: 11%;
      height: 27%;
      border-radius: 999px;
      transform-origin: top center;
    }
    .arm.left { left: -5%; transform: rotate(12deg); }
    .arm.right { right: -5%; transform: rotate(-12deg); }
    .foot {
      bottom: -7%;
      width: 24%;
      height: 9%;
      border-radius: 999px;
    }
    .foot.left { left: 18%; }
    .foot.right { right: 18%; }
    .eye {
      width: 30%;
      aspect-ratio: 1;
      border-radius: 50%;
      background: radial-gradient(circle at 35% 32%, #ffffff 0 20%, #eaf8f7 21% 100%);
      position: relative;
      overflow: hidden;
    }
    .pupil {
      position: absolute;
      left: 50%;
      top: 50%;
      width: 42%;
      aspect-ratio: 1;
      border-radius: 50%;
      background: linear-gradient(135deg, #2b84cb, #e2263c);
      transform: translate(calc(-50% + var(--look-x, 0px)), calc(-50% + var(--look-y, 0px)));
      transition: transform .08s ease-out, height .16s ease, border-radius .16s ease;
    }
    .bot.shy .pupil {
      height: 6px;
      border-radius: 999px;
      transform: translate(-50%, -50%);
    }
    .smile {
      position: absolute;
      left: 35%;
      right: 35%;
      bottom: 27%;
      height: 8%;
      border-bottom: 3px solid #64248c;
      border-radius: 0 0 999px 999px;
    }
    .shadow {
      position: absolute;
      left: 25%;
      right: 25%;
      bottom: 3%;
      height: 8%;
      border-radius: 50%;
      background: rgba(0, 0, 0, .22);
      filter: blur(8px);
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
      .scene { display: none; }
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
      <div class="scene" aria-hidden="true">
        <div class="orbit"></div>
        <div class="orbit"></div>
        <div class="shadow"></div>
        <div class="bot">
          <div class="arm left"></div>
          <div class="arm right"></div>
          <div class="face">
            <div class="eye"><div class="pupil"></div></div>
            <div class="eye"><div class="pupil"></div></div>
          </div>
          <div class="smile"></div>
          <div class="foot left"></div>
          <div class="foot right"></div>
        </div>
      </div>
      <div class="stats">
        <div class="stat"><span>Area</span><strong>CT LOG</strong></div>
        <div class="stat"><span>Acesso</span><strong>Seguro</strong></div>
        <div class="stat"><span>Dados</span><strong>Online</strong></div>
      </div>
    </section>
    <section class="login">
      <div class="eyebrow">Acesso ao sistema</div>
      <h2>Dashboard Log</h2>
      <p>Informe suas credenciais para entrar no painel.</p>
      {message}
      <form method="post" action="/login">
        <label>Usuario
          <input name="user" autocomplete="username" required>
        </label>
        <label>Senha
          <input name="password" type="password" autocomplete="current-password" required>
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
    const bot = document.querySelector(".bot");
    const pupils = document.querySelectorAll(".pupil");
    document.addEventListener("pointermove", (event) => {
      document.body.style.setProperty("--mx", `${event.clientX}px`);
      document.body.style.setProperty("--my", `${event.clientY}px`);
      pupils.forEach((pupil) => {
        const box = pupil.parentElement.getBoundingClientRect();
        const dx = event.clientX - (box.left + box.width / 2);
        const dy = event.clientY - (box.top + box.height / 2);
        const angle = Math.atan2(dy, dx);
        const distance = Math.min(box.width * .18, Math.hypot(dx, dy) / 18);
        pupil.style.setProperty("--look-x", `${Math.cos(angle) * distance}px`);
        pupil.style.setProperty("--look-y", `${Math.sin(angle) * distance}px`);
      });
    });
    document.querySelector('input[type="password"]').addEventListener("focus", () => bot.classList.add("shy"));
    document.querySelector('input[type="password"]').addEventListener("blur", () => bot.classList.remove("shy"));
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
  <title>Home - Dashboard Log</title>
  <style>
    :root {
      --bg: #34104f;
      --panel: #ffffff;
      --ink: #16212d;
      --muted: #657282;
      --line: #d7e0e8;
      --teal: #64248c;
      --green: #2b84cb;
      --gold: #e2263c;
      --shadow: 0 18px 42px rgba(0, 0, 0, .18);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      overflow-x: hidden;
      background:
        radial-gradient(620px circle at var(--mx, 80%) var(--my, 24%), rgba(43, 132, 203, .24), transparent 58%),
        radial-gradient(520px circle at 18% 84%, rgba(226, 38, 60, .16), transparent 64%),
        var(--bg);
      color: var(--ink);
      font-family: Inter, Segoe UI, Roboto, Arial, sans-serif;
    }
    .assistant { display: none; }
    .assistant .ring {
      position: absolute;
      inset: 1%;
      border: 1px solid rgba(255, 255, 255, .15);
      border-radius: 50%;
      transform: rotate(-18deg) scaleX(1.18);
    }
    .assistant .head {
      position: absolute;
      left: 50%;
      top: 50%;
      width: 62%;
      aspect-ratio: .78;
      transform: translate(-50%, -50%);
      border-radius: 42% 42% 30% 30%;
      background:
        radial-gradient(circle at 34% 20%, rgba(255,255,255,.95), transparent 24%),
        linear-gradient(160deg, #ffffff, #d7e5ea 58%, #b8cbd3);
      border: 1px solid rgba(255, 255, 255, .76);
      box-shadow: 0 26px 66px rgba(0, 0, 0, .22);
    }
    .assistant .head::before {
      content: "";
      position: absolute;
      left: 22%;
      right: 22%;
      top: -9%;
      height: 11%;
      border-radius: 999px;
      background: linear-gradient(90deg, #2b84cb, #e2263c);
      box-shadow: 0 0 30px rgba(43, 132, 203, .58);
    }
    .assistant .visor {
      position: absolute;
      left: 11%;
      right: 11%;
      top: 22%;
      height: 39%;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 14%;
      border-radius: 26px;
      background: linear-gradient(180deg, #34104f, #1b255f);
    }
    .assistant .eye {
      width: 29%;
      aspect-ratio: 1;
      border-radius: 50%;
      background: radial-gradient(circle at 35% 32%, #ffffff 0 20%, #eaf8f7 21% 100%);
      position: relative;
      overflow: hidden;
    }
    .assistant .pupil {
      position: absolute;
      left: 50%;
      top: 50%;
      width: 42%;
      aspect-ratio: 1;
      border-radius: 50%;
      background: linear-gradient(135deg, #2b84cb, #e2263c);
      transform: translate(calc(-50% + var(--look-x, 0px)), calc(-50% + var(--look-y, 0px)));
      transition: transform .08s ease-out;
    }
    .assistant .mouth {
      position: absolute;
      left: 35%;
      right: 35%;
      bottom: 27%;
      height: 8%;
      border-bottom: 3px solid #64248c;
      border-radius: 0 0 999px 999px;
    }
    .assistant .arm, .assistant .foot {
      position: absolute;
      background: #2b84cb;
      box-shadow: 0 8px 16px rgba(0, 0, 0, .12);
    }
    .assistant .arm {
      top: 49%;
      width: 12%;
      height: 27%;
      border-radius: 999px;
    }
    .assistant .arm.left { left: -5%; transform: rotate(12deg); }
    .assistant .arm.right { right: -5%; transform: rotate(-12deg); }
    .assistant .foot {
      bottom: -7%;
      width: 24%;
      height: 9%;
      border-radius: 999px;
    }
    .assistant .foot.left { left: 18%; }
    .assistant .foot.right { right: 18%; }
    .assistant .shadow {
      position: absolute;
      left: 27%;
      right: 27%;
      bottom: 4%;
      height: 8%;
      border-radius: 50%;
      background: rgba(0, 0, 0, .22);
      filter: blur(8px);
    }
    header {
      position: relative;
      z-index: 2;
      padding: 30px clamp(16px, 4vw, 44px) 22px;
      color: #fff;
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 16px;
    }
    h1 { margin: 0; font-size: clamp(30px, 4vw, 48px); letter-spacing: 0; }
    .subtitle { margin: 8px 0 0; color: #c8d6dc; }
    .brand-title {
      display: flex;
      align-items: center;
      gap: 18px;
    }
    .brand-title img {
      width: 76px;
      height: 76px;
      object-fit: contain;
      filter: drop-shadow(0 10px 18px rgba(0, 0, 0, .26));
    }
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
      position: relative;
      z-index: 2;
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
      box-shadow: 0 22px 60px rgba(0, 0, 0, .20), inset 0 1px 0 rgba(255, 255, 255, .86);
      transition: transform .18s ease, box-shadow .18s ease;
    }
    .card:hover {
      transform: translateY(-3px);
      box-shadow: 0 28px 76px rgba(0, 0, 0, .24), inset 0 1px 0 rgba(255, 255, 255, .9);
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
      .assistant { display: none; }
    }
  </style>
</head>
<body>
  <header>
    <div>
      <div class="brand-title"><img src="{favicon_url}" alt=""><h1>Dashboard Log</h1></div>
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
  <script>
    const pupils = document.querySelectorAll(".pupil");
    document.addEventListener("pointermove", (event) => {
      document.body.style.setProperty("--mx", `${event.clientX}px`);
      document.body.style.setProperty("--my", `${event.clientY}px`);
      pupils.forEach((pupil) => {
        const box = pupil.parentElement.getBoundingClientRect();
        const dx = event.clientX - (box.left + box.width / 2);
        const dy = event.clientY - (box.top + box.height / 2);
        const angle = Math.atan2(dy, dx);
        const distance = Math.min(box.width * .18, Math.hypot(dx, dy) / 18);
        pupil.style.setProperty("--look-x", `${Math.cos(angle) * distance}px`);
        pupil.style.setProperty("--look-y", `${Math.sin(angle) * distance}px`);
      });
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
  <title>Base editavel - Dashboard Log</title>
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
      <a class="top-link" href="/logout">Sair</a>
    </nav>
  </header>
  <main>
    <section class="panel">
      {message}
      <form id="sheetForm" method="post" action="/editar">
        <input type="hidden" name="rows_json" id="rowsJson">
        <div class="toolbar">
          <div class="meta"><span id="rowCount">__ROW_COUNT__</span> linhas na base</div>
          <div class="actions">
            <button type="button" id="addRow">Adicionar linha</button>
            <button type="button" id="deleteRows" class="button">Excluir selecionadas</button>
            <a class="button" href="/editar?recarregar=1">Recarregar base</a>
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
      ["notaFiscal", "Nota fiscal"],
      ["produto", "Produto"],
      ["cliente", "Cliente"],
      ["quantidade", "Quantidade"]
    ];
    let rows = __ROWS__;
    const thead = document.querySelector("#sheet thead");
    const tbody = document.querySelector("#sheet tbody");
    const rowCount = document.querySelector("#rowCount");
    let selectionStart = null;
    let selectionEnd = null;
    let isSelecting = false;

    function cleanRow(row) {
      return Object.fromEntries(columns.map(([key]) => [key, row[key] ?? ""]));
    }

    function render() {
      thead.innerHTML = `<tr><th></th>${columns.map(([, label]) => `<th>${label}</th>`).join("")}</tr>`;
      tbody.innerHTML = rows.map((row, idx) => {
        const clean = cleanRow(row);
        return `<tr data-row="${idx}">
          <td><input type="checkbox" aria-label="Selecionar linha ${idx + 1}"><br>${idx + 1}</td>
          ${columns.map(([key], colIdx) => `<td contenteditable="true" data-key="${key}" data-col="${colIdx}">${String(clean[key]).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")}</td>`).join("")}
        </tr>`;
      }).join("");
      rowCount.textContent = rows.length.toLocaleString("pt-BR");
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
    });
    document.querySelector("#deleteRows").addEventListener("click", () => {
      syncFromTable();
      const selected = new Set([...tbody.querySelectorAll("tr")].filter((tr) => tr.querySelector("input").checked).map((tr) => Number(tr.dataset.row)));
      rows = rows.filter((_, idx) => !selected.has(idx));
      render();
    });
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
      if (!text.includes("\t") && !text.includes("\n")) return;
      event.preventDefault();
      syncFromTable();
      const start = cellPosition(cell);
      text.trimEnd().split(/\r?\n/).forEach((line, rowOffset) => {
        line.split("\t").forEach((value, colOffset) => {
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
          matrix.push(values.join("\t"));
        }
        navigator.clipboard?.writeText(matrix.join("\n"));
      }
    });
    document.querySelector("#sheetForm").addEventListener("submit", () => {
      syncFromTable();
      document.querySelector("#rowsJson").value = JSON.stringify(rows);
    });
    if (rows.length) render();
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
    build_dashboard.main()


def editable_rows() -> list[dict[str, object]]:
    return build_dashboard.ensure_editable_data()


def save_editable_rows(rows: list[dict[str, object]]) -> None:
    if not rows:
        rows = build_dashboard.editable_rows_from_sources()
    clean_rows = []
    for row in rows:
        clean_rows.append({key: row.get(key, "") for key in build_dashboard.EDITABLE_COLUMNS})
    build_dashboard.save_editable_data(clean_rows)


def html_escape(value: object) -> str:
    return html.escape(str(value if value is not None else ""), quote=False)


def render_sheet_parts(rows: list[dict[str, object]]) -> tuple[str, str]:
    columns = [
        ("data", "Data"),
        ("placa", "Placa"),
        ("terminal", "Terminal"),
        ("viagens", "Viagens"),
        ("capacidade", "Capacidade"),
        ("notaFiscal", "Nota fiscal"),
        ("produto", "Produto"),
        ("cliente", "Cliente"),
        ("quantidade", "Quantidade"),
    ]
    thead = "<tr><th></th>" + "".join(f"<th>{label}</th>" for _, label in columns) + "</tr>"
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
        page = LOGIN_HTML.replace("{message}", message).replace("{favicon_url}", FAVICON_URL)
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
        if "recarregar" in params:
            save_editable_rows(build_dashboard.editable_rows_from_sources())
            self.redirect("/editar")
            return
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
            .replace("__ROWS__", json.dumps(rows, ensure_ascii=False))
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
    DATA_DIR.mkdir(exist_ok=True)
    if not INDEX_PATH.exists():
        rebuild_dashboard()
    port = int(os.environ.get("PORT", "8000"))
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"Servidor rodando em http://0.0.0.0:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
