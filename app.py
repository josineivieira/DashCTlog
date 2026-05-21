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
      padding: 32px;
      overflow: hidden;
      background:
        linear-gradient(135deg, rgba(255,255,255,.03) 0 1px, transparent 1px 48px),
        radial-gradient(720px circle at var(--mx, 76%) var(--my, 24%), rgba(20, 153, 160, .30), transparent 58%),
        radial-gradient(520px circle at 18% 78%, rgba(60, 140, 79, .20), transparent 62%),
        linear-gradient(135deg, #0c1f27, #12363c 52%, #10232b),
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
      width: 50px;
      height: 50px;
      border-radius: 14px;
      display: grid;
      place-items: center;
      color: #fff;
      font-weight: 900;
      background: linear-gradient(135deg, #0c7c83, #18a7ad);
      box-shadow: 0 16px 34px rgba(12, 124, 131, .28);
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
      right: clamp(18px, 5vw, 54px);
      bottom: 124px;
      width: clamp(190px, 21vw, 280px);
      aspect-ratio: 1;
      pointer-events: none;
      opacity: .98;
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
      background: linear-gradient(90deg, #0c7c83, #15a0a8);
      box-shadow: 0 0 32px rgba(12, 124, 131, .72);
    }
    .bot::after {
      content: "";
      position: absolute;
      left: 28%;
      right: 28%;
      bottom: 13%;
      height: 12%;
      border-radius: 999px;
      background: linear-gradient(180deg, rgba(12, 124, 131, .16), rgba(12, 124, 131, .06));
      border: 1px solid rgba(12, 124, 131, .22);
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
      background: linear-gradient(180deg, #10232b, #0b1e25);
      box-shadow: inset 0 0 22px rgba(20, 153, 160, .22);
    }
    .arm, .foot {
      position: absolute;
      background: #0c7c83;
      box-shadow: 0 8px 16px rgba(0, 0, 0, .12);
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
      background: linear-gradient(135deg, #0c7c83, #18a7ad);
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
      border-bottom: 3px solid #0c7c83;
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
      color: #0c7c83;
      font-size: 12px;
      font-weight: 900;
      text-transform: uppercase;
    }
    .eyebrow::before {
      content: "";
      width: 9px;
      height: 9px;
      border-radius: 50%;
      background: #0c7c83;
      box-shadow: 0 0 0 6px rgba(12, 124, 131, .12);
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
      outline: 3px solid rgba(12, 124, 131, .18);
      border-color: var(--teal);
    }
    button {
      min-height: 50px;
      border: 0;
      border-radius: 12px;
      background: linear-gradient(135deg, #0c7c83, #0f9694);
      color: #fff;
      cursor: pointer;
      font: inherit;
      font-weight: 900;
      box-shadow: 0 16px 32px rgba(12, 124, 131, .28);
      transition: transform .16s ease, box-shadow .16s ease;
    }
    button:hover {
      transform: translateY(-1px);
      box-shadow: 0 20px 40px rgba(12, 124, 131, .34);
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
      background: #3c8c4f;
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
      <div class="brand-mark">DL</div>
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
      overflow-x: hidden;
      background:
        radial-gradient(620px circle at var(--mx, 80%) var(--my, 24%), rgba(20, 153, 160, .24), transparent 58%),
        radial-gradient(520px circle at 18% 84%, rgba(159, 122, 28, .14), transparent 64%),
        var(--bg);
      color: var(--ink);
      font-family: Inter, Segoe UI, Roboto, Arial, sans-serif;
    }
    .assistant {
      position: fixed;
      right: clamp(34px, 7vw, 110px);
      bottom: clamp(34px, 8vh, 84px);
      width: clamp(180px, 19vw, 280px);
      aspect-ratio: 1;
      pointer-events: none;
      opacity: .96;
      z-index: 0;
    }
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
      background: linear-gradient(90deg, #0c7c83, #15a0a8);
      box-shadow: 0 0 30px rgba(12, 124, 131, .68);
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
      background: linear-gradient(180deg, #10232b, #0b1e25);
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
      background: linear-gradient(135deg, #0c7c83, #18a7ad);
      transform: translate(calc(-50% + var(--look-x, 0px)), calc(-50% + var(--look-y, 0px)));
      transition: transform .08s ease-out;
    }
    .assistant .mouth {
      position: absolute;
      left: 35%;
      right: 35%;
      bottom: 27%;
      height: 8%;
      border-bottom: 3px solid #0c7c83;
      border-radius: 0 0 999px 999px;
    }
    .assistant .arm, .assistant .foot {
      position: absolute;
      background: #0c7c83;
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
  <div class="assistant" aria-hidden="true">
    <div class="ring"></div>
    <div class="shadow"></div>
    <div class="head">
      <div class="arm left"></div>
      <div class="arm right"></div>
      <div class="visor">
        <div class="eye"><div class="pupil"></div></div>
        <div class="eye"><div class="pupil"></div></div>
      </div>
      <div class="mouth"></div>
      <div class="foot left"></div>
      <div class="foot right"></div>
    </div>
  </div>
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
