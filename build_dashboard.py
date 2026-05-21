from __future__ import annotations

import json
import math
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
ORIGINAL_ORDERS_PATH = ROOT / "20260520T162822.501-ordens 1.xlsx"
ORIGINAL_DETAIL_PATH = ROOT / "20260520T202750.243-geral ct log.xlsx"
UPLOAD_ORDERS_PATH = DATA_DIR / "ordens.xlsx"
UPLOAD_DETAIL_PATH = DATA_DIR / "geral_ct_log.xlsx"
EDITABLE_DATA_PATH = DATA_DIR / "dashboard_base.json"
OUTPUT_PATH = ROOT / "index.html"

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
TERMINAL_NAMES = {
    "10": "Equador",
    "19": "Ipiranga",
}
EDITABLE_COLUMNS = [
    "data",
    "placa",
    "terminal",
    "viagens",
    "capacidade",
    "notaFiscal",
    "produto",
    "cliente",
    "quantidade",
]


def current_orders_path() -> Path:
    return UPLOAD_ORDERS_PATH if UPLOAD_ORDERS_PATH.exists() else ORIGINAL_ORDERS_PATH


def current_detail_path() -> Path:
    return UPLOAD_DETAIL_PATH if UPLOAD_DETAIL_PATH.exists() else ORIGINAL_DETAIL_PATH


def col_to_idx(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref)
    letters = match.group(1) if match else "A"
    total = 0
    for char in letters:
        total = total * 26 + ord(char) - 64
    return total - 1


def excel_datetime(value: str) -> datetime | None:
    text = str(value).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    try:
        return datetime(1899, 12, 30) + timedelta(days=float(text))
    except (TypeError, ValueError):
        return None


def read_xlsx(path: Path) -> tuple[list[dict[str, str]], list[str]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", NS):
                shared_strings.append(
                    "".join(text.text or "" for text in item.findall(".//a:t", NS))
                )

        sheet_root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        rows: list[list[str]] = []
        for row in sheet_root.findall(".//a:sheetData/a:row", NS):
            values: list[str] = []
            for cell in row.findall("a:c", NS):
                idx = col_to_idx(cell.attrib.get("r", "A1"))
                while len(values) <= idx:
                    values.append("")

                value_node = cell.find("a:v", NS)
                value = ""
                if value_node is not None:
                    if cell.attrib.get("t") == "s":
                        value = shared_strings[int(value_node.text or "0")]
                    else:
                        value = value_node.text or ""
                values[idx] = value
            rows.append(values)

    headers = rows[0]
    records = [
        dict(zip(headers, row + [""] * (len(headers) - len(row)))) for row in rows[1:]
    ]
    return records, headers


def num(value: str) -> float:
    try:
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return 0.0


def clean_product(description: str) -> str:
    product = re.split(r"\s+-\s+| \(?ONU| \(?Cod ANP| \(?ANP", description, maxsplit=1)[0]
    return " ".join(product.split()).strip() or "Sem produto"


def day(value: str) -> str:
    parsed = excel_datetime(value)
    return parsed.strftime("%d/%m/%Y") if parsed else ""


def infer_capacities(
    orders_by_key: Counter[tuple[str, str, str]],
    load_by_key: dict[tuple[str, str, str], float],
) -> dict[str, int]:
    samples: dict[str, list[int]] = defaultdict(list)
    for key, total in load_by_key.items():
        trips = orders_by_key.get(key, 0)
        if trips <= 0 or total <= 0:
            continue
        rounded = int(round((total / trips) / 5000) * 5000)
        if rounded >= 5000:
            samples[key[1]].append(rounded)

    capacities: dict[str, int] = {}
    for plate, values in samples.items():
        capacities[plate] = Counter(values).most_common(1)[0][0]
    return capacities


def editable_rows_from_sources() -> list[dict[str, object]]:
    orders_path = current_orders_path()
    detail_path = current_detail_path()
    order_rows, order_headers = read_xlsx(orders_path)
    detail_rows, detail_headers = read_xlsx(detail_path)

    oh = {
        "date": order_headers[1],
        "plate": order_headers[4],
        "terminal": order_headers[3],
    }
    dh = {
        "date": detail_headers[1],
        "nf": detail_headers[2],
        "product": detail_headers[6],
        "client": detail_headers[8],
        "plate": detail_headers[11],
        "qty": detail_headers[12],
        "terminal": detail_headers[13],
    }

    orders_by_key: Counter[tuple[str, str, str]] = Counter()
    for row in order_rows:
        terminal = row[oh["terminal"]].strip()
        plate = row[oh["plate"]].strip().upper()
        date = day(row[oh["date"]])
        if terminal in TERMINAL_NAMES and plate and date:
            orders_by_key[(date, plate, terminal)] += 1

    load_by_key: dict[tuple[str, str, str], float] = defaultdict(float)
    for row in detail_rows:
        terminal = row[dh["terminal"]].strip()
        plate = row[dh["plate"]].strip().upper()
        date = day(row[dh["date"]])
        quantity = num(row[dh["qty"]])
        if terminal in TERMINAL_NAMES and plate and date and 0 < quantity <= 50000:
            load_by_key[(date, plate, terminal)] += quantity

    capacities = infer_capacities(orders_by_key, load_by_key)
    editable_rows: list[dict[str, object]] = []
    keys_with_detail: set[tuple[str, str, str]] = set()
    for row in detail_rows:
        terminal = row[dh["terminal"]].strip()
        plate = row[dh["plate"]].strip().upper()
        date = day(row[dh["date"]])
        quantity = num(row[dh["qty"]])
        if terminal not in TERMINAL_NAMES or not plate or not date or quantity <= 0 or quantity > 50000:
            continue
        key = (date, plate, terminal)
        keys_with_detail.add(key)
        editable_rows.append(
            {
                "data": date,
                "placa": plate,
                "terminal": terminal,
                "viagens": orders_by_key.get(key, 0),
                "capacidade": capacities.get(plate) or 30000,
                "notaFiscal": row[dh["nf"]].strip(),
                "produto": clean_product(row[dh["product"]]),
                "cliente": row[dh["client"]].strip(),
                "quantidade": quantity,
            }
        )

    for date, plate, terminal in sorted(set(orders_by_key) - keys_with_detail):
        editable_rows.append(
            {
                "data": date,
                "placa": plate,
                "terminal": terminal,
                "viagens": orders_by_key[(date, plate, terminal)],
                "capacidade": capacities.get(plate) or 30000,
                "notaFiscal": "",
                "produto": "",
                "cliente": "",
                "quantidade": 0,
            }
        )
    return editable_rows


def ensure_editable_data() -> list[dict[str, object]]:
    DATA_DIR.mkdir(exist_ok=True)
    if EDITABLE_DATA_PATH.exists():
        return json.loads(EDITABLE_DATA_PATH.read_text(encoding="utf-8"))
    rows = editable_rows_from_sources()
    EDITABLE_DATA_PATH.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    return rows


def build_data_from_editable(rows: list[dict[str, object]]) -> dict[str, object]:
    orders_by_key: Counter[tuple[str, str, str]] = Counter()
    load_by_key: dict[tuple[str, str, str], float] = defaultdict(float)
    products_by_key: dict[tuple[str, str, str], Counter[str]] = defaultdict(Counter)
    notes_by_key: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    clients_by_key: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    capacities: dict[str, int] = {}
    detail_line_count = 0

    for row in rows:
        terminal = str(row.get("terminal", "")).strip()
        plate = str(row.get("placa", "")).strip().upper()
        date = day(str(row.get("data", "")))
        if terminal not in TERMINAL_NAMES or not plate or not date:
            continue
        key = (date, plate, terminal)
        trips = int(num(str(row.get("viagens", 0))) or 0)
        if trips > orders_by_key.get(key, 0):
            orders_by_key[key] = trips
        capacity = int(num(str(row.get("capacidade", 0))) or 0)
        if capacity > 0:
            capacities[plate] = capacity
        quantity = num(str(row.get("quantidade", 0)))
        if quantity > 0:
            load_by_key[key] += quantity
            product = clean_product(str(row.get("produto", ""))) or "SEM PRODUTO"
            products_by_key[key][product] += quantity
            detail_line_count += 1
        note = str(row.get("notaFiscal", "")).strip()
        client = str(row.get("cliente", "")).strip()
        if note:
            notes_by_key[key].add(note)
        if client:
            clients_by_key[key].add(client)

    all_keys = sorted(
        set(orders_by_key) | set(load_by_key),
        key=lambda item: (
            datetime.strptime(item[0], "%d/%m/%Y"),
            item[1],
            item[2],
        ),
    )
    daily_plate_rows: list[dict[str, object]] = []
    for date, plate, terminal in all_keys:
        loaded = load_by_key.get((date, plate, terminal), 0.0)
        capacity = capacities.get(plate) or 30000
        trips_from_orders = orders_by_key.get((date, plate, terminal), 0)
        trips_from_load = math.ceil(loaded / capacity) if loaded else 0
        trips = max(trips_from_orders, trips_from_load)
        products = [
            {"produto": product, "quantidade": qty}
            for product, qty in products_by_key.get((date, plate, terminal), Counter()).most_common()
        ]
        daily_plate_rows.append(
            {
                "data": date,
                "placa": plate,
                "terminal": terminal,
                "terminalNome": TERMINAL_NAMES[terminal],
                "viagens": trips,
                "viagensOrdens": trips_from_orders,
                "viagensCarga": trips_from_load,
                "capacidade": capacity,
                "quantidade": loaded,
                "notas": len(notes_by_key.get((date, plate, terminal), set())),
                "clientes": len(clients_by_key.get((date, plate, terminal), set())),
                "produtos": products,
                "mixProdutos": ", ".join(
                    f"{item['produto']} ({item['quantidade'] / 1000:.0f}k)"
                    for item in products[:3]
                )
                or "Sem produto detalhado",
            }
        )

    return {
        "dailyPlateRows": daily_plate_rows,
        "meta": {
            "ordersFile": "Base editável",
            "detailFile": EDITABLE_DATA_PATH.name,
            "orderRows": len(rows),
            "detailRows": detail_line_count,
        },
    }


def build_data() -> dict[str, object]:
    if EDITABLE_DATA_PATH.exists():
        return build_data_from_editable(ensure_editable_data())

    orders_path = current_orders_path()
    detail_path = current_detail_path()
    order_rows, order_headers = read_xlsx(orders_path)
    detail_rows, detail_headers = read_xlsx(detail_path)

    oh = {
        "date": order_headers[1],
        "plate": order_headers[4],
        "terminal": order_headers[3],
    }
    dh = {
        "date": detail_headers[1],
        "nf": detail_headers[2],
        "product": detail_headers[6],
        "client": detail_headers[8],
        "plate": detail_headers[11],
        "qty": detail_headers[12],
        "terminal": detail_headers[13],
    }

    orders_by_key: Counter[tuple[str, str, str]] = Counter()
    for row in order_rows:
        terminal = row[oh["terminal"]].strip()
        plate = row[oh["plate"]].strip().upper()
        date = day(row[oh["date"]])
        if terminal in TERMINAL_NAMES and plate and date:
            orders_by_key[(date, plate, terminal)] += 1

    load_by_key: dict[tuple[str, str, str], float] = defaultdict(float)
    products_by_key: dict[tuple[str, str, str], Counter[str]] = defaultdict(Counter)
    notes_by_key: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    clients_by_key: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    detail_line_count = 0

    for row in detail_rows:
        terminal = row[dh["terminal"]].strip()
        plate = row[dh["plate"]].strip().upper()
        date = day(row[dh["date"]])
        quantity = num(row[dh["qty"]])
        if terminal not in TERMINAL_NAMES or not plate or not date:
            continue
        if quantity <= 0 or quantity > 50000:
            continue

        key = (date, plate, terminal)
        product = clean_product(row[dh["product"]])
        load_by_key[key] += quantity
        products_by_key[key][product] += quantity
        if row[dh["nf"]].strip():
            notes_by_key[key].add(row[dh["nf"]].strip())
        if row[dh["client"]].strip():
            clients_by_key[key].add(row[dh["client"]].strip())
        detail_line_count += 1

    capacities = infer_capacities(orders_by_key, load_by_key)
    all_keys = sorted(
        set(orders_by_key) | set(load_by_key),
        key=lambda item: (
            datetime.strptime(item[0], "%d/%m/%Y"),
            item[1],
            item[2],
        ),
    )

    daily_plate_rows: list[dict[str, object]] = []
    for date, plate, terminal in all_keys:
        loaded = load_by_key.get((date, plate, terminal), 0.0)
        capacity = capacities.get(plate) or 30000
        trips_from_orders = orders_by_key.get((date, plate, terminal), 0)
        trips_from_load = math.ceil(loaded / capacity) if loaded else 0
        trips = max(trips_from_orders, trips_from_load)
        products = [
            {"produto": product, "quantidade": qty}
            for product, qty in products_by_key.get((date, plate, terminal), Counter()).most_common()
        ]
        daily_plate_rows.append(
            {
                "data": date,
                "placa": plate,
                "terminal": terminal,
                "terminalNome": TERMINAL_NAMES[terminal],
                "viagens": trips,
                "viagensOrdens": trips_from_orders,
                "viagensCarga": trips_from_load,
                "capacidade": capacity,
                "quantidade": loaded,
                "notas": len(notes_by_key.get((date, plate, terminal), set())),
                "clientes": len(clients_by_key.get((date, plate, terminal), set())),
                "produtos": products,
                "mixProdutos": ", ".join(
                    f"{item['produto']} ({item['quantidade'] / 1000:.0f}k)"
                    for item in products[:3]
                )
                or "Sem produto detalhado",
            }
        )

    return {
        "dailyPlateRows": daily_plate_rows,
        "meta": {
            "ordersFile": orders_path.name,
            "detailFile": detail_path.name,
            "orderRows": len(order_rows),
            "detailRows": detail_line_count,
        },
    }


HTML_TEMPLATE = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Dashboard Log</title>
  <style>
    :root {
      --bg: #eef2f5;
      --top: #10232b;
      --top-2: #173843;
      --panel: #ffffff;
      --panel-soft: #f8fafb;
      --ink: #16212d;
      --muted: #657282;
      --line: #d7e0e8;
      --teal: #0c7c83;
      --green: #2f855a;
      --rust: #c16b3f;
      --blue: #4c63b6;
      --gold: #b98218;
      --shadow: 0 18px 42px rgba(23, 32, 51, .10);
    }
    body[data-theme="forest"] {
      --bg: #eef4ef;
      --top: #123126;
      --top-2: #1c4a39;
      --teal: #1f7a5c;
      --green: #3c8c4f;
      --rust: #a5672a;
      --blue: #3f6f7d;
      --gold: #9f7a1c;
    }
    body[data-theme="graphite"] {
      --bg: #edf0f3;
      --top: #171b22;
      --top-2: #2c3440;
      --teal: #287980;
      --green: #45775f;
      --rust: #a85f44;
      --blue: #596ca8;
      --gold: #a47b29;
    }
    body[data-theme="marine"] {
      --bg: #edf4f7;
      --top: #0e2536;
      --top-2: #16415c;
      --teal: #087f8c;
      --green: #2e7d64;
      --rust: #b56a40;
      --blue: #315fb0;
      --gold: #b2811e;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      background: var(--top);
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
    h1 { margin: 0; font-size: clamp(28px, 3vw, 42px); letter-spacing: 0; font-weight: 800; }
    .subtitle { margin: 8px 0 0; color: #c8d6dc; }
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
      text-decoration: none;
      font-size: 13px;
      font-weight: 800;
      background: rgba(255, 255, 255, .08);
    }
    .top-link:hover { background: rgba(255, 255, 255, .15); }
    main { padding: 0 clamp(16px, 4vw, 44px) 42px; }
    .filters {
      display: grid;
      grid-template-columns: 1.4fr repeat(5, minmax(132px, 1fr));
      gap: 12px;
      margin-bottom: 16px;
      padding: 14px;
      background: rgba(255, 255, 255, .96);
      border: 1px solid rgba(255, 255, 255, .7);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }
    label { display: grid; gap: 6px; color: #506071; font-size: 13px; font-weight: 700; }
    select, input {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 10px 11px;
      background: var(--panel-soft);
      color: var(--ink);
      font: inherit;
      width: 100%;
    }
    select:focus, input:focus {
      outline: 3px solid rgba(12, 124, 131, .18);
      border-color: var(--teal);
      background: #fff;
    }
    .kpis {
      display: grid;
      grid-template-columns: repeat(4, minmax(150px, 1fr));
      gap: 12px;
      margin-bottom: 16px;
    }
    .kpi, .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }
    .kpi {
      padding: 15px;
      min-height: 100px;
      border-top: 4px solid var(--teal);
    }
    .kpi:nth-child(2) { border-top-color: var(--green); }
    .kpi:nth-child(3) { border-top-color: var(--gold); }
    .kpi:nth-child(4) { border-top-color: var(--blue); }
    .kpi span { display: block; color: var(--muted); font-size: 12px; font-weight: 800; text-transform: uppercase; }
    .kpi strong { display: block; margin-top: 11px; font-size: 29px; line-height: 1.05; color: #111b26; }
    .grid {
      display: grid;
      grid-template-columns: 1.3fr .7fr;
      gap: 15px;
      align-items: start;
    }
    .panel { padding: 17px; overflow: hidden; }
    .wide { grid-column: 1 / -1; }
    h2 { margin: 0 0 13px; font-size: 17px; color: #1f2f3d; }
    .bars { display: grid; gap: 10px; }
    .bar-row {
      display: grid;
      grid-template-columns: minmax(130px, 230px) 1fr 86px;
      gap: 10px;
      align-items: center;
      min-height: 30px;
    }
    .bar-label { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: #263645; font-weight: 650; }
    .track { height: 11px; border-radius: 999px; background: #e7edf2; overflow: hidden; box-shadow: inset 0 1px 2px rgba(22, 33, 45, .08); }
    .fill { height: 100%; border-radius: inherit; background: var(--teal); }
    .value { color: var(--muted); text-align: right; font-variant-numeric: tabular-nums; }
    .hero-board {
      display: grid;
      grid-template-columns: 1fr 330px;
      gap: 15px;
      margin-bottom: 15px;
    }
    .focus-card {
      min-height: 320px;
      background: #fff;
    }
    .focus-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 14px;
      padding-bottom: 13px;
      border-bottom: 1px solid var(--line);
    }
    .focus-label { color: var(--muted); font-weight: 800; text-transform: uppercase; font-size: 12px; }
    .focus-plate-pill {
      display: inline-flex;
      align-items: center;
      min-height: 42px;
      padding: 7px 14px;
      border-radius: 8px;
      background: #10232b;
      color: #fff;
      font-size: clamp(22px, 3vw, 34px);
      font-weight: 900;
      letter-spacing: 0;
    }
    .focus-metrics {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 10px;
      margin-bottom: 16px;
    }
    .focus-metric {
      padding: 11px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel-soft);
    }
    .focus-metric span { display: block; color: var(--muted); font-size: 11px; font-weight: 800; text-transform: uppercase; }
    .focus-metric strong { display: block; margin-top: 6px; font-size: 20px; color: var(--ink); }
    .legend {
      display: flex;
      flex-wrap: wrap;
      gap: 10px 16px;
      margin-top: 10px;
      color: var(--muted);
      font-size: 12px;
    }
    .legend span { display: inline-flex; align-items: center; gap: 6px; }
    .swatch { width: 10px; height: 10px; border-radius: 3px; display: inline-block; }
    .mini-grid {
      display: grid;
      gap: 12px;
    }
    .mini-card {
      padding: 15px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
    }
    .mini-card span {
      display: block;
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
      text-transform: uppercase;
    }
    .mini-card strong {
      display: block;
      margin-top: 7px;
      font-size: 28px;
    }
    .product-list {
      display: grid;
      gap: 12px;
    }
    .product-row {
      display: grid;
      grid-template-columns: minmax(0, 1fr) 110px;
      gap: 12px;
      align-items: center;
      padding-bottom: 10px;
      border-bottom: 1px solid var(--line);
    }
    .product-name { font-weight: 800; color: #253545; }
    .product-sub { color: var(--muted); font-size: 12px; margin-top: 3px; }
    .product-value { color: var(--ink); text-align: right; font-weight: 800; font-variant-numeric: tabular-nums; }
    .line-chart { min-height: 320px; }
    .chart-toolbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      margin-bottom: 8px;
    }
    .segment {
      display: inline-flex;
      padding: 3px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #f3f6f8;
    }
    .segment button {
      border: 0;
      border-radius: 6px;
      background: transparent;
      color: var(--muted);
      cursor: pointer;
      font: inherit;
      font-size: 12px;
      font-weight: 800;
      padding: 7px 10px;
    }
    .segment button.active {
      background: #fff;
      color: var(--ink);
      box-shadow: 0 1px 4px rgba(23, 32, 51, .12);
    }
    .chart-total {
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
      text-transform: uppercase;
    }
    .line-chart svg { width: 100%; height: 300px; display: block; }
    .chart-grid { stroke: #dde6ed; stroke-width: 1; }
    .chart-axis { stroke: #b8c5d0; stroke-width: 1.2; }
    .chart-label { fill: var(--muted); font-size: 11px; }
    .chart-value { fill: #1a2a36; font-size: 12px; font-weight: 800; }
    .chart-value-volume { fill: #8d4c31; font-size: 11px; font-weight: 800; }
    .chart-line-trips { fill: none; stroke: var(--teal); stroke-width: 3; stroke-linecap: round; stroke-linejoin: round; }
    .chart-line-volume { fill: none; stroke: var(--rust); stroke-width: 3; stroke-linecap: round; stroke-linejoin: round; }
    .chart-area { fill: rgba(12, 124, 131, .10); }
    .chart-dot-trips { fill: #fff; stroke: var(--teal); stroke-width: 2.5; }
    .chart-dot-volume { fill: #fff; stroke: var(--rust); stroke-width: 2.5; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { padding: 12px 10px; border-bottom: 1px solid var(--line); text-align: left; vertical-align: middle; }
    tbody tr:nth-child(even) { background: #fafcfd; }
    tbody tr:hover { background: #eef7f7; }
    th { color: #506071; background: #f3f6f8; position: sticky; top: 0; z-index: 1; font-size: 12px; text-transform: uppercase; }
    .table-wrap { max-height: 520px; overflow: auto; border: 1px solid var(--line); border-radius: 8px; background: #fff; }
    .num-cell { text-align: right; font-variant-numeric: tabular-nums; }
    .terminal-cell { color: #344457; line-height: 1.3; }
    .load-cell {
      min-width: 130px;
    }
    .load-value {
      display: flex;
      justify-content: space-between;
      gap: 8px;
      margin-bottom: 5px;
      font-weight: 800;
      font-variant-numeric: tabular-nums;
    }
    .mini-track {
      height: 7px;
      border-radius: 999px;
      background: #e8eef5;
      overflow: hidden;
    }
    .mini-fill {
      height: 100%;
      border-radius: inherit;
      background: var(--green);
    }
    .pill {
      display: inline-flex;
      align-items: center;
      min-height: 24px;
      padding: 3px 8px;
      border-radius: 999px;
      background: #e5f3f2;
      color: #0d6268;
      font-weight: 700;
      font-size: 12px;
    }
    .note { color: var(--muted); font-size: 12px; margin-top: 10px; }
    .empty { color: var(--muted); padding: 22px; text-align: center; }
    @media (max-width: 980px) {
      .filters, .kpis, .grid, .hero-board { grid-template-columns: 1fr 1fr; }
      .wide { grid-column: 1 / -1; }
    }
    @media (max-width: 650px) {
      header { flex-direction: column; }
      .nav { justify-content: flex-start; }
      .filters, .kpis, .grid, .hero-board { grid-template-columns: 1fr; }
      .bar-row { grid-template-columns: 1fr; gap: 5px; }
      .focus-head { align-items: flex-start; flex-direction: column; }
      .focus-metrics { grid-template-columns: 1fr; }
      .product-row { grid-template-columns: 1fr; }
      .product-value { text-align: left; }
      .value { text-align: left; }
    }
  </style>
</head>
<body>
  <header>
    <div>
      <h1>Dashboard Log</h1>
      <p class="subtitle">Viagens por placa, produtos carregados e terminais Equador/Ipiranga.</p>
    </div>
    <a class="top-link" href="/editar">Atualizar dados</a>
  </header>
  <main>
    <section class="filters">
      <label>Buscar placa ou produto
        <input id="search" type="search" placeholder="NPB-4686, diesel, gasolina...">
      </label>
      <label>Data inicial
        <input id="dateStart" type="date">
      </label>
      <label>Data final
        <input id="dateEnd" type="date">
      </label>
      <label>Terminal
        <select id="terminal"></select>
      </label>
      <label>Placa
        <select id="plate"></select>
      </label>
      <label>Tema
        <select id="theme">
          <option value="marine">Marine</option>
          <option value="forest">Forest</option>
          <option value="graphite">Graphite</option>
        </select>
      </label>
    </section>

    <section class="kpis">
      <div class="kpi"><span>Viagens</span><strong id="kTrips">0</strong></div>
      <div class="kpi"><span>Carregado</span><strong id="kQty">0</strong></div>
      <div class="kpi"><span>Placas</span><strong id="kPlates">0</strong></div>
      <div class="kpi"><span>Notas</span><strong id="kNotes">0</strong></div>
    </section>

    <section class="hero-board">
      <div class="panel focus-card">
        <div class="focus-head">
          <div>
            <div class="focus-label">Placa destaque</div>
            <div id="focusPlate" class="focus-plate-pill">-</div>
          </div>
          <div class="focus-label">Maior volume/viagens no filtro</div>
        </div>
        <div class="focus-metrics">
          <div class="focus-metric"><span>Viagens</span><strong id="focusTrips">0</strong></div>
          <div class="focus-metric"><span>Volume</span><strong id="focusVolume">0</strong></div>
          <div class="focus-metric"><span>Capacidade</span><strong id="focusCapacity">0</strong></div>
        </div>
        <div id="focusProducts" class="product-list"></div>
      </div>
      <div class="mini-grid">
        <div class="mini-card"><span>Top produto</span><strong id="kTopProduct">-</strong></div>
        <div class="mini-card"><span>Top terminal</span><strong id="kTopTerminal">-</strong></div>
        <div class="mini-card"><span>Maior dia</span><strong id="kBestDay">-</strong></div>
      </div>
    </section>

    <section class="grid">
      <div class="panel wide">
        <div class="chart-toolbar">
          <h2>Evolução por Dia</h2>
          <div class="segment" aria-label="Segmentação da evolução diária">
            <button type="button" class="active" data-chart-mode="both">Ambos</button>
            <button type="button" data-chart-mode="trips">Viagens</button>
            <button type="button" data-chart-mode="volume">Volume</button>
          </div>
        </div>
        <div id="chartTotal" class="chart-total"></div>
        <div id="lineChart" class="line-chart"></div>
      </div>
      <div class="panel">
        <h2>Top Placas por Viagens</h2>
        <div id="plateTripBars" class="bars"></div>
      </div>
      <div class="panel">
        <h2>Top Placas por Volume</h2>
        <div id="plateQtyBars" class="bars"></div>
      </div>
      <div class="panel">
        <h2>Produtos</h2>
        <div id="productList" class="product-list"></div>
      </div>
      <div class="panel">
        <h2>Terminais</h2>
        <div id="terminalBars" class="bars"></div>
      </div>
      <div class="panel wide">
        <h2>Viagens no Dia por Placa</h2>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Data</th>
                <th>Placa</th>
                <th>Terminal</th>
                <th>Viagens</th>
                <th>Capacidade usada</th>
                <th>Carregado</th>
                <th>Notas</th>
              </tr>
            </thead>
            <tbody id="dailyTable"></tbody>
          </table>
        </div>
        <p class="note">Terminal 19 tratado como Ipiranga. Viagens usam a planilha de ordens e, quando necessário, completam por capacidade inferida pela placa.</p>
      </div>
    </section>
  </main>
  <script>
    const dataset = __DATA__;
    const rows = dataset.dailyPlateRows;
    const $ = (id) => document.getElementById(id);
    const fmt = new Intl.NumberFormat("pt-BR");
    const volume = (value) => `${fmt.format(Math.round(value / 1000))} mil`;
    const palette = ["#0c7c83", "#2f855a", "#c16b3f", "#4c63b6", "#b98218", "#6a5aa8", "#9a5a72", "#52796f"];
    let chartMode = "both";

    function unique(key) {
      return [...new Set(rows.map((row) => row[key]).filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), "pt-BR"));
    }

    function fillSelect(id, values, label) {
      $(id).innerHTML = `<option value="">${label}</option>` + values.map((value) => `<option>${value}</option>`).join("");
    }

    function filteredRows() {
      const query = $("search").value.trim().toLowerCase();
      const dateStart = $("dateStart").value;
      const dateEnd = $("dateEnd").value;
      const terminal = $("terminal").value;
      const plate = $("plate").value;
      return rows.filter((row) => {
        const haystack = [row.placa, row.terminalNome, row.mixProdutos, ...row.produtos.map((item) => item.produto)].join(" ").toLowerCase();
        const rowDate = brToIso(row.data);
        return (!query || haystack.includes(query))
          && (!dateStart || rowDate >= dateStart)
          && (!dateEnd || rowDate <= dateEnd)
          && (!terminal || row.terminal === terminal)
          && (!plate || row.placa === plate);
      });
    }

    function sumBy(data, key, valueKey) {
      const map = new Map();
      data.forEach((row) => map.set(row[key], (map.get(row[key]) || 0) + row[valueKey]));
      return [...map.entries()].sort((a, b) => b[1] - a[1]);
    }

    function productTotals(data) {
      const map = new Map();
      data.forEach((row) => row.produtos.forEach((item) => {
        map.set(item.produto, (map.get(item.produto) || 0) + item.quantidade);
      }));
      return [...map.entries()].sort((a, b) => b[1] - a[1]);
    }

    function dateValue(label) {
      const [d, m, y] = label.split("/").map(Number);
      return new Date(y, m - 1, d).getTime();
    }

    function brToIso(label) {
      const [d, m, y] = label.split("/");
      return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
    }

    function toInputDate(date) {
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, "0");
      const d = String(date.getDate()).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }

    function setCurrentMonthRange() {
      const now = new Date();
      $("dateStart").value = toInputDate(new Date(now.getFullYear(), now.getMonth(), 1));
      $("dateEnd").value = toInputDate(now);
    }

    function dailyTotals(data) {
      const map = new Map();
      data.forEach((row) => {
        const item = map.get(row.data) || { data: row.data, viagens: 0, quantidade: 0 };
        item.viagens += row.viagens;
        item.quantidade += row.quantidade;
        map.set(row.data, item);
      });
      return [...map.values()].sort((a, b) => dateValue(a.data) - dateValue(b.data));
    }

    function productShort(name) {
      return name
        .replace("GASOLINA C ECO ADITIVADA DURAMAIS", "Gas. Aditivada")
        .replace("GASOLINA C COMUM", "Gas. Comum")
        .replace("OLEO DIESEL B S10-COMUM", "Diesel S10")
        .replace("OLEO DIESEL B S500 COMUM", "Diesel S500")
        .replace("ETANOL HIDRATADO ECO ADITIVADO DURAMAIS", "Etanol Adit.")
        .replace("ETANOL HIDRATADO CARBURANTE", "Etanol")
        .replace("OLEO DIESEL B S10 ECO ADITIVADO DURAMAIS", "Diesel S10 Adit.");
    }

    function terminalTotals(data) {
      const map = new Map();
      data.forEach((row) => {
        const label = `${row.terminal} - ${row.terminalNome}`;
        map.set(label, (map.get(label) || 0) + row.viagens);
      });
      return [...map.entries()].sort((a, b) => b[1] - a[1]);
    }

    function bars(id, pairs, limit, color, formatter = fmt.format) {
      const max = Math.max(1, ...pairs.map((pair) => pair[1]));
      const visible = pairs.slice(0, limit);
      $(id).innerHTML = visible.map(([label, value]) => `
        <div class="bar-row" title="${label}: ${formatter(value)}">
          <div class="bar-label">${label}</div>
          <div class="track"><div class="fill" style="width:${(value / max) * 100}%; background:${color}"></div></div>
          <div class="value">${formatter(value)}</div>
        </div>
      `).join("") || `<div class="empty">Sem dados</div>`;
    }

    function uniqueFrom(data, key) {
      return [...new Set(data.map((row) => row[key]).filter(Boolean))];
    }

    function productList(id, products, limit = 7) {
      const total = products.reduce((sum, item) => sum + item[1], 0);
      $(id).innerHTML = products.slice(0, limit).map(([label, value], idx) => {
        const pct = total ? Math.round((value / total) * 100) : 0;
        return `
          <div class="product-row">
            <div>
              <div class="product-name">${productShort(label)}</div>
              <div class="product-sub">${pct}% do volume filtrado</div>
            </div>
            <div class="product-value">${volume(value)}</div>
          </div>`;
      }).join("") || `<div class="empty">Sem dados</div>`;
    }

    function lineChart(data) {
      const days = dailyTotals(data);
      if (!days.length) {
        $("chartTotal").textContent = "";
        $("lineChart").innerHTML = `<div class="empty">Sem dados</div>`;
        return;
      }
      const totalTrips = days.reduce((sum, item) => sum + item.viagens, 0);
      const totalQty = days.reduce((sum, item) => sum + item.quantidade, 0);
      $("chartTotal").textContent = `${fmt.format(totalTrips)} viagens | ${volume(totalQty)} carregados no período`;
      const width = 1120;
      const height = 300;
      const pad = { left: 48, right: 28, top: 34, bottom: 42 };
      const chartW = width - pad.left - pad.right;
      const chartH = height - pad.top - pad.bottom;
      const maxTrips = Math.max(1, ...days.map((item) => item.viagens));
      const maxQty = Math.max(1, ...days.map((item) => item.quantidade));
      const x = (idx) => pad.left + (idx / Math.max(1, days.length - 1)) * chartW;
      const yTrips = (value) => pad.top + chartH - (value / maxTrips) * chartH;
      const yQty = (value) => pad.top + chartH - (value / maxQty) * chartH;
      const tripPoints = days.map((item, idx) => [x(idx), yTrips(item.viagens)]);
      const qtyPoints = days.map((item, idx) => [x(idx), yQty(item.quantidade)]);
      const path = (points) => points.map((point, idx) => `${idx ? "L" : "M"}${point[0].toFixed(1)},${point[1].toFixed(1)}`).join(" ");
      const area = `${path(tripPoints)} L${x(days.length - 1).toFixed(1)},${(height - pad.bottom).toFixed(1)} L${pad.left},${height - pad.bottom} Z`;
      const tickEvery = Math.max(1, Math.ceil(days.length / 10));
      const grid = [0, .25, .5, .75, 1].map((ratio) => {
        const y = pad.top + chartH * ratio;
        return `<line class="chart-grid" x1="${pad.left}" y1="${y}" x2="${width - pad.right}" y2="${y}"></line>`;
      }).join("");
      const labels = days.map((item, idx) => idx % tickEvery ? "" : `<text class="chart-label" x="${x(idx)}" y="${height - 12}" text-anchor="middle">${item.data.slice(0, 5)}</text>`).join("");
      const tripDots = days.map((item, idx) => `<circle class="chart-dot-trips" cx="${x(idx)}" cy="${yTrips(item.viagens)}" r="4"><title>${item.data}: ${fmt.format(item.viagens)} viagens</title></circle>`).join("");
      const qtyDots = days.map((item, idx) => `<circle class="chart-dot-volume" cx="${x(idx)}" cy="${yQty(item.quantidade)}" r="3.5"><title>${item.data}: ${volume(item.quantidade)}</title></circle>`).join("");
      const tripLabels = days.map((item, idx) => `<text class="chart-value" x="${x(idx)}" y="${yTrips(item.viagens) - 10}" text-anchor="middle">${fmt.format(item.viagens)}</text>`).join("");
      const qtyLabels = days.map((item, idx) => `<text class="chart-value-volume" x="${x(idx)}" y="${yQty(item.quantidade) + 18}" text-anchor="middle">${Math.round(item.quantidade / 1000)}k</text>`).join("");
      const showTrips = chartMode === "both" || chartMode === "trips";
      const showVolume = chartMode === "both" || chartMode === "volume";
      $("lineChart").innerHTML = `
        <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Evolução diária de viagens e volume">
          ${grid}
          <line class="chart-axis" x1="${pad.left}" y1="${height - pad.bottom}" x2="${width - pad.right}" y2="${height - pad.bottom}"></line>
          ${showTrips ? `<path class="chart-area" d="${area}"></path>` : ""}
          ${showVolume ? `<path class="chart-line-volume" d="${path(qtyPoints)}"></path>${qtyDots}${qtyLabels}` : ""}
          ${showTrips ? `<path class="chart-line-trips" d="${path(tripPoints)}"></path>${tripDots}${tripLabels}` : ""}
          ${labels}
          ${showTrips ? `<text class="chart-label" x="${pad.left}" y="15">Viagens</text>` : ""}
          ${showVolume ? `<text class="chart-label" x="${width - 96}" y="15">Volume</text>` : ""}
        </svg>
        <div class="legend">
          ${showTrips ? `<span><i class="swatch" style="background:var(--teal)"></i> Viagens</span>` : ""}
          ${showVolume ? `<span><i class="swatch" style="background:var(--rust)"></i> Volume carregado</span>` : ""}
        </div>`;
    }

    function focusPanel(data, plateTrips, plateQty) {
      const topPlate = plateTrips[0]?.[0];
      if (!topPlate) {
        $("focusPlate").textContent = "-";
        $("focusTrips").textContent = "0";
        $("focusVolume").textContent = "0";
        $("focusCapacity").textContent = "0";
        $("focusProducts").innerHTML = `<div class="empty">Sem dados</div>`;
        return;
      }
      const plateRows = data.filter((row) => row.placa === topPlate);
      $("focusPlate").textContent = topPlate;
      $("focusTrips").textContent = fmt.format(plateTrips[0][1]);
      $("focusVolume").textContent = volume(plateQty.find((item) => item[0] === topPlate)?.[1] || 0);
      $("focusCapacity").textContent = volume(Math.max(...plateRows.map((row) => row.capacidade || 0), 0));
      productList("focusProducts", productTotals(plateRows), 5);
    }

    function render() {
      const data = filteredRows();
      const trips = data.reduce((total, row) => total + row.viagens, 0);
      const qty = data.reduce((total, row) => total + row.quantidade, 0);
      const notes = data.reduce((total, row) => total + row.notas, 0);
      const plateTrips = sumBy(data, "placa", "viagens");
      const plateQty = sumBy(data, "placa", "quantidade");
      const products = productTotals(data);

      $("kTrips").textContent = fmt.format(trips);
      $("kQty").textContent = volume(qty);
      $("kPlates").textContent = fmt.format(new Set(data.map((row) => row.placa)).size);
      $("kNotes").textContent = fmt.format(notes);
      $("kTopProduct").textContent = products[0]?.[0]?.split(" ").slice(0, 2).join(" ") || "-";
      $("kTopTerminal").textContent = terminalTotals(data)[0]?.[0] || "-";
      const bestDay = dailyTotals(data).sort((a, b) => b.viagens - a.viagens)[0];
      $("kBestDay").textContent = bestDay ? bestDay.data.slice(0, 5) : "-";

      bars("plateTripBars", plateTrips, 12, "var(--teal)");
      bars("plateQtyBars", plateQty, 12, "var(--green)", volume);
      bars("terminalBars", terminalTotals(data), 4, "var(--blue)");
      productList("productList", products, 8);
      lineChart(data);
      focusPanel(data, plateTrips, plateQty);

      $("dailyTable").innerHTML = data
        .slice()
        .sort((a, b) => b.viagens - a.viagens || b.quantidade - a.quantidade)
        .map((row) => `
          <tr>
            <td>${row.data}</td>
            <td><span class="pill">${row.placa}</span></td>
            <td class="terminal-cell">${row.terminal} - ${row.terminalNome}</td>
            <td class="num-cell">${fmt.format(row.viagens)}</td>
            <td class="num-cell">${volume(row.capacidade)}</td>
            <td class="load-cell">
              <div class="load-value"><span>${volume(row.quantidade)}</span><span>${Math.round((row.quantidade / Math.max(1, row.capacidade * Math.max(1, row.viagens))) * 100)}%</span></div>
              <div class="mini-track"><div class="mini-fill" style="width:${Math.min(100, (row.quantidade / Math.max(1, row.capacidade * Math.max(1, row.viagens))) * 100)}%"></div></div>
            </td>
            <td class="num-cell">${fmt.format(row.notas)}</td>
          </tr>
        `).join("") || `<tr><td colspan="7" class="empty">Nenhum registro encontrado</td></tr>`;
    }

    fillSelect("terminal", [["10", "10 - Equador"], ["19", "19 - Ipiranga"]].map((item) => item.join(" - ").replace(" - ", " - ")), "Todos os terminais");
    $("terminal").innerHTML = `<option value="">Todos os terminais</option><option value="10">10 - Equador</option><option value="19">19 - Ipiranga</option>`;
    fillSelect("plate", unique("placa"), "Todas as placas");
    $("theme").addEventListener("input", () => document.body.dataset.theme = $("theme").value);
    document.body.dataset.theme = $("theme").value;
    setCurrentMonthRange();
    document.querySelectorAll("[data-chart-mode]").forEach((button) => {
      button.addEventListener("click", () => {
        chartMode = button.dataset.chartMode;
        document.querySelectorAll("[data-chart-mode]").forEach((item) => item.classList.toggle("active", item === button));
        render();
      });
    });
    ["search", "dateStart", "dateEnd", "terminal", "plate"].forEach((id) => $(id).addEventListener("input", render));
    render();
  </script>
</body>
</html>
"""


def main() -> None:
    data = build_data()
    html = HTML_TEMPLATE.replace("__DATA__", json.dumps(data, ensure_ascii=False))
    OUTPUT_PATH.write_text(html, encoding="utf-8")
    meta = data["meta"]
    print(f"Dashboard criado: {OUTPUT_PATH}")
    print(f"Ordens: {meta['orderRows']} | Linhas detalhadas usadas: {meta['detailRows']}")
    print(f"Linhas dia/placa/terminal: {len(data['dailyPlateRows'])}")


if __name__ == "__main__":
    main()
