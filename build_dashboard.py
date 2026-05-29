from __future__ import annotations

import json
import math
import os
import re
import sys
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
CAPACITY_DATA_PATH = DATA_DIR / "capacidades_carretas.json"
OUTPUT_PATH = ROOT / "index.html"
FAVICON_URL = "https://pages.greatpages.com.br/www.dislubequador.com.br/1777495651/imagens/mobile/3562683_1_177616861364933621_m.svg"

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
    "motorista1",
    "notaFiscal",
    "produto",
    "cliente",
    "municipioDestino",
    "quantidade",
    "cfopDescricao",
]
CAPACITY_COLUMNS = [
    "id",
    "tipo",
    "capacidade",
    "placaCavalo",
    "tanques",
    "carreta",
    "observacao",
]
DEFAULT_CAPACITY_ROWS = [
    {"id": "1", "tipo": "CARRETA", "capacidade": "30", "placaCavalo": "OAH-4329", "tanques": "", "carreta": "JXA-1558", "observacao": ""},
    {"id": "2", "tipo": "CARRETA", "capacidade": "30", "placaCavalo": "NPB-5A17", "tanques": "", "carreta": "JXA-4749", "observacao": "5 TQ"},
    {"id": "3", "tipo": "CARRETA", "capacidade": "30", "placaCavalo": "OAO-2G70", "tanques": "", "carreta": "JXA-5453", "observacao": ""},
    {"id": "4", "tipo": "CARRETA", "capacidade": "30", "placaCavalo": "QZX-0E87", "tanques": "", "carreta": "JXA-5463", "observacao": "6 TQ"},
    {"id": "5", "tipo": "CARRETA", "capacidade": "30", "placaCavalo": "QZR-4D17", "tanques": "", "carreta": "JXA-5473", "observacao": "6 TQ"},
    {"id": "6", "tipo": "CARRETA", "capacidade": "30", "placaCavalo": "NPB-8487", "tanques": "", "carreta": "JXA-5531", "observacao": "5 TQ"},
    {"id": "7", "tipo": "CARRETA", "capacidade": "30", "placaCavalo": "OAH-4129", "tanques": "", "carreta": "JXA-5543", "observacao": "6 TQ"},
    {"id": "8", "tipo": "MANOBRA", "capacidade": "30", "placaCavalo": "JXA-2216", "tanques": "MANOBRA", "carreta": "JXA-7910", "observacao": ""},
    {"id": "9", "tipo": "CAMINHAO", "capacidade": "20", "placaCavalo": "", "tanques": "4X5", "carreta": "NPB-4686", "observacao": ""},
    {"id": "10", "tipo": "CAMINHAO", "capacidade": "25", "placaCavalo": "", "tanques": "3X5+2X3+1X4", "carreta": "OAK-8G51", "observacao": ""},
    {"id": "11", "tipo": "CAMINHAO", "capacidade": "25", "placaCavalo": "", "tanques": "5X5", "carreta": "PHO-4D02", "observacao": ""},
    {"id": "12", "tipo": "CAMINHAO", "capacidade": "25", "placaCavalo": "", "tanques": "5X5", "carreta": "PHO-4D32", "observacao": ""},
    {"id": "13", "tipo": "CAMINHAO", "capacidade": "25", "placaCavalo": "", "tanques": "", "carreta": "PHZ-8J44", "observacao": ""},
    {"id": "14", "tipo": "CAMINHAO", "capacidade": "25", "placaCavalo": "", "tanques": "BVB", "carreta": "PHZ-8J64", "observacao": ""},
    {"id": "15", "tipo": "CAMINHAO", "capacidade": "20", "placaCavalo": "", "tanques": "", "carreta": "QZN-8E65", "observacao": ""},
    {"id": "16", "tipo": "CAMINHAO", "capacidade": "25", "placaCavalo": "", "tanques": "", "carreta": "QZW-6A95", "observacao": ""},
    {"id": "17", "tipo": "CAMINHAO", "capacidade": "20", "placaCavalo": "", "tanques": "", "carreta": "TAB-2G94", "observacao": ""},
]


def database_url() -> str:
    return os.environ.get("DATABASE_URL", "").strip()


def postgres_driver_available() -> bool:
    return bool(postgres_driver_name())


def postgres_driver_name() -> str:
    try:
        import psycopg2  # noqa: F401
    except ModuleNotFoundError:
        try:
            import psycopg  # noqa: F401
        except ModuleNotFoundError:
            return ""
        return "psycopg"
    return "psycopg2"


def postgres_driver_error() -> str:
    errors = []
    try:
        import psycopg2  # noqa: F401
        return ""
    except Exception as exc:
        errors.append(f"psycopg2: {type(exc).__name__}: {exc}")
    try:
        import psycopg  # noqa: F401
        return ""
    except Exception as exc:
        errors.append(f"psycopg: {type(exc).__name__}: {exc}")
    return " | ".join(errors)


def python_executable() -> str:
    return sys.executable


def use_postgres() -> bool:
    return (
        bool(database_url())
        and postgres_driver_available()
    )


def database_required() -> bool:
    return any(
        os.environ.get(name)
        for name in (
            "DATABASE_REQUIRED",
            "RENDER",
            "RENDER_SERVICE_ID",
            "RENDER_EXTERNAL_URL",
        )
    )


def ensure_database_storage() -> None:
    if use_postgres():
        return
    if not database_required():
        return
    if not database_url():
        raise RuntimeError("DATABASE_URL nao esta configurada; abortando para nao usar JSON no deploy.")
    raise RuntimeError(
        "Banco configurado, mas driver Postgres indisponivel; abortando para nao usar JSON no deploy. "
        + postgres_driver_error()
    )


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


def normalize_plate(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


def capacity_liters(value: object) -> int:
    amount = int(num(str(value)) or 0)
    if 0 < amount < 1000:
        return amount * 1000
    return amount


def clean_product(description: str) -> str:
    product = re.split(r"\s+-\s+| \(?ONU| \(?Cod ANP| \(?ANP", description, maxsplit=1)[0]
    return " ".join(product.split()).strip() or "Sem produto"


def invoice_numbers(value: object) -> list[str]:
    text = str(value or "").strip()
    if not text:
        return []
    return [item.strip() for item in re.split(r"\s+/\s+", text) if item.strip()]


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
    capacities.update(capacity_registry())
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
                "capacidade": capacities.get(normalize_plate(plate)) or capacities.get(plate) or 30000,
                "motorista1": "",
                "notaFiscal": row[dh["nf"]].strip(),
                "produto": clean_product(row[dh["product"]]),
                "cliente": row[dh["client"]].strip(),
                "municipioDestino": "",
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
                "capacidade": capacities.get(normalize_plate(plate)) or capacities.get(plate) or 30000,
                "motorista1": "",
                "notaFiscal": "",
                "produto": "",
                "cliente": "",
                "municipioDestino": "",
                "quantidade": 0,
            }
        )
    return editable_rows


def postgres_connection():
    try:
        import psycopg2
    except ModuleNotFoundError:
        import psycopg

        return psycopg.connect(database_url())

    return psycopg2.connect(database_url())


def clean_capacity_row(row: dict[str, object], idx: int = 0) -> dict[str, str]:
    cleaned = {key: str(row.get(key, "") or "").strip().upper() for key in CAPACITY_COLUMNS}
    cleaned["id"] = str(row.get("id", "") or idx or "").strip()
    cleaned["capacidade"] = str(row.get("capacidade", "") or "").strip()
    cleaned["observacao"] = str(row.get("observacao", "") or "").strip().upper()
    return cleaned


def ensure_postgres_capacity_table() -> None:
    with postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS capacidade_carretas (
                    id TEXT PRIMARY KEY,
                    row_order INTEGER NOT NULL DEFAULT 0,
                    tipo TEXT NOT NULL DEFAULT '',
                    capacidade TEXT NOT NULL DEFAULT '',
                    placa_cavalo TEXT NOT NULL DEFAULT '',
                    tanques TEXT NOT NULL DEFAULT '',
                    carreta TEXT NOT NULL DEFAULT '',
                    observacao TEXT NOT NULL DEFAULT '',
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )


def postgres_capacity_rows() -> list[dict[str, object]]:
    ensure_postgres_capacity_table()
    with postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT id, tipo, capacidade, placa_cavalo, tanques, carreta, observacao
                FROM capacidade_carretas
                ORDER BY row_order, id
                """
            )
            rows = [
                {
                    "id": item[0],
                    "tipo": item[1],
                    "capacidade": item[2],
                    "placaCavalo": item[3],
                    "tanques": item[4],
                    "carreta": item[5],
                    "observacao": item[6],
                }
                for item in cur.fetchall()
            ]
            if rows:
                return rows
    save_postgres_capacity_rows(DEFAULT_CAPACITY_ROWS)
    return list(DEFAULT_CAPACITY_ROWS)


def save_postgres_capacity_rows(rows: list[dict[str, object]]) -> None:
    ensure_postgres_capacity_table()
    with postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("TRUNCATE TABLE capacidade_carretas")
            for idx, row in enumerate(rows, start=1):
                item = clean_capacity_row(row, idx)
                cur.execute(
                    """
                    INSERT INTO capacidade_carretas (
                        id, row_order, tipo, capacidade, placa_cavalo, tanques, carreta, observacao
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        item["id"] or str(idx),
                        idx,
                        item["tipo"],
                        item["capacidade"],
                        item["placaCavalo"],
                        item["tanques"],
                        item["carreta"],
                        item["observacao"],
                    ),
                )


def save_capacity_rows(rows: list[dict[str, object]]) -> None:
    clean_rows = [clean_capacity_row(row, idx) for idx, row in enumerate(rows, start=1)]
    if use_postgres():
        save_postgres_capacity_rows(clean_rows)
        return
    DATA_DIR.mkdir(exist_ok=True)
    CAPACITY_DATA_PATH.write_text(json.dumps(clean_rows, ensure_ascii=False, indent=2), encoding="utf-8")


def ensure_capacity_rows() -> list[dict[str, object]]:
    if use_postgres():
        return postgres_capacity_rows()
    DATA_DIR.mkdir(exist_ok=True)
    if CAPACITY_DATA_PATH.exists():
        rows = json.loads(CAPACITY_DATA_PATH.read_text(encoding="utf-8"))
        if isinstance(rows, list):
            return [clean_capacity_row(row, idx) for idx, row in enumerate(rows, start=1)]
    save_capacity_rows(DEFAULT_CAPACITY_ROWS)
    return list(DEFAULT_CAPACITY_ROWS)


def capacity_registry() -> dict[str, int]:
    capacities: dict[str, int] = {}
    for row in ensure_capacity_rows():
        capacity = capacity_liters(row.get("capacidade", ""))
        if capacity <= 0:
            continue
        for key in ("placaCavalo", "carreta"):
            plate = normalize_plate(str(row.get(key, "")))
            if plate:
                capacities[plate] = capacity
    return capacities


def apply_capacity_registry_to_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    registry = capacity_registry()
    updated = []
    for row in rows:
        item = dict(row)
        capacity = registry.get(normalize_plate(str(item.get("placa", ""))))
        if capacity:
            item["capacidade"] = capacity
        updated.append(item)
    return updated


def ensure_postgres_table() -> None:
    with postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS dashboard_base (
                    id SERIAL PRIMARY KEY,
                    row_order INTEGER NOT NULL DEFAULT 0,
                    data TEXT NOT NULL DEFAULT '',
                    placa TEXT NOT NULL DEFAULT '',
                    terminal TEXT NOT NULL DEFAULT '',
                    viagens TEXT NOT NULL DEFAULT '',
                    capacidade TEXT NOT NULL DEFAULT '',
                    motorista_1 TEXT NOT NULL DEFAULT '',
                    nota_fiscal TEXT NOT NULL DEFAULT '',
                    produto TEXT NOT NULL DEFAULT '',
                    cliente TEXT NOT NULL DEFAULT '',
                    municipio_destino TEXT NOT NULL DEFAULT '',
                    quantidade TEXT NOT NULL DEFAULT '',
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )
            cur.execute("ALTER TABLE dashboard_base ADD COLUMN IF NOT EXISTS motorista_1 TEXT NOT NULL DEFAULT ''")
            cur.execute("ALTER TABLE dashboard_base ADD COLUMN IF NOT EXISTS municipio_destino TEXT NOT NULL DEFAULT ''")


def postgres_rows() -> list[dict[str, object]]:
    ensure_postgres_table()
    with postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT data, placa, terminal, viagens, capacidade,
                       motorista_1, nota_fiscal, produto, cliente, municipio_destino, quantidade
                FROM dashboard_base
                ORDER BY row_order, id
                """
            )
            rows = []
            for item in cur.fetchall():
                rows.append(
                    {
                        "data": item[0],
                        "placa": item[1],
                        "terminal": item[2],
                        "viagens": item[3],
                        "capacidade": item[4],
                        "motorista1": item[5],
                        "notaFiscal": item[6],
                        "produto": item[7],
                        "cliente": item[8],
                        "municipioDestino": item[9],
                        "quantidade": item[10],
                    }
                )
            return rows


def save_postgres_rows(rows: list[dict[str, object]]) -> None:
    ensure_postgres_table()
    with postgres_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("TRUNCATE TABLE dashboard_base RESTART IDENTITY")
            for idx, row in enumerate(rows):
                cur.execute(
                    """
                    INSERT INTO dashboard_base (
                        row_order, data, placa, terminal, viagens, capacidade,
                        motorista_1, nota_fiscal, produto, cliente, municipio_destino, quantidade
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        idx,
                        str(row.get("data", "")),
                        str(row.get("placa", "")),
                        str(row.get("terminal", "")),
                        str(row.get("viagens", "")),
                        str(row.get("capacidade", "")),
                        str(row.get("motorista1", "")),
                        str(row.get("notaFiscal", "")),
                        str(row.get("produto", "")),
                        str(row.get("cliente", "")),
                        str(row.get("municipioDestino", "")),
                        str(row.get("quantidade", "")),
                    ),
                )


def save_editable_data(rows: list[dict[str, object]]) -> None:
    if use_postgres():
        if not rows:
            raise ValueError("Base vazia: salvamento no banco cancelado para evitar perda de dados.")
        rows = apply_capacity_registry_to_rows(rows)
        save_postgres_rows(rows)
        return
    rows = apply_capacity_registry_to_rows(rows)
    if not rows:
        rows = editable_rows_from_sources()
    DATA_DIR.mkdir(exist_ok=True)
    EDITABLE_DATA_PATH.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")


def ensure_editable_data() -> list[dict[str, object]]:
    if use_postgres():
        return postgres_rows()
    DATA_DIR.mkdir(exist_ok=True)
    if EDITABLE_DATA_PATH.exists():
        rows = json.loads(EDITABLE_DATA_PATH.read_text(encoding="utf-8"))
        if isinstance(rows, list) and rows:
            return apply_capacity_registry_to_rows(rows)
    rows = editable_rows_from_sources()
    save_editable_data(rows)
    return rows


def build_data_from_editable(rows: list[dict[str, object]]) -> dict[str, object]:
    report_key = tuple[str, str, str, str]
    orders_by_key: Counter[report_key] = Counter()
    load_by_key: dict[report_key, float] = defaultdict(float)
    products_by_key: dict[report_key, Counter[str]] = defaultdict(Counter)
    notes_by_key: dict[report_key, set[str]] = defaultdict(set)
    clients_by_key: dict[report_key, set[str]] = defaultdict(set)
    drivers_by_key: dict[report_key, set[str]] = defaultdict(set)
    capacities: dict[str, int] = {}
    detail_line_count = 0

    for row in rows:
        terminal = str(row.get("terminal", "")).strip()
        plate = str(row.get("placa", "")).strip().upper()
        date = day(str(row.get("data", "")))
        driver = str(row.get("motorista1", "")).strip().upper()
        if terminal not in TERMINAL_NAMES or not plate or not date:
            continue
        key = (date, plate, terminal, driver)
        trips = int(num(str(row.get("viagens", 0))) or 0)
        if trips > orders_by_key.get(key, 0):
            orders_by_key[key] = trips
        capacity = capacity_liters(row.get("capacidade", 0))
        if capacity > 0 and normalize_plate(plate) not in capacities:
            capacities[plate] = capacity
        quantity = num(str(row.get("quantidade", 0)))
        if quantity > 0:
            load_by_key[key] += quantity
            product = clean_product(str(row.get("produto", ""))) or "SEM PRODUTO"
            products_by_key[key][product] += quantity
            detail_line_count += 1
        client = str(row.get("cliente", "")).strip()
        for note in invoice_numbers(row.get("notaFiscal", "")):
            notes_by_key[key].add(note)
        if client:
            clients_by_key[key].add(client)
        if driver:
            drivers_by_key[key].add(driver)

    all_keys = sorted(
        set(orders_by_key) | set(load_by_key),
        key=lambda item: (
            datetime.strptime(item[0], "%d/%m/%Y"),
            item[1],
            item[2],
            item[3],
        ),
    )
    daily_plate_rows: list[dict[str, object]] = []
    for date, plate, terminal, driver in all_keys:
        key = (date, plate, terminal, driver)
        loaded = load_by_key.get(key, 0.0)
        capacity = capacities.get(normalize_plate(plate)) or capacities.get(plate) or 30000
        trips_from_orders = orders_by_key.get(key, 0)
        trips_from_load = math.ceil(loaded / capacity) if loaded else 0
        trips = max(trips_from_orders, trips_from_load)
        products = [
            {"produto": product, "quantidade": qty}
            for product, qty in products_by_key.get(key, Counter()).most_common()
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
                "notas": len(notes_by_key.get(key, set())),
                "clientes": len(clients_by_key.get(key, set())),
                "motorista": " / ".join(sorted(drivers_by_key.get(key, set()))),
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
            "detailFile": "Postgres" if use_postgres() else EDITABLE_DATA_PATH.name,
            "orderRows": len(rows),
            "detailRows": detail_line_count,
        },
    }


def build_data() -> dict[str, object]:
    if use_postgres() or EDITABLE_DATA_PATH.exists():
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
        for note in invoice_numbers(row[dh["nf"]]):
            notes_by_key[key].add(note)
        if row[dh["client"]].strip():
            clients_by_key[key].add(row[dh["client"]].strip())
        detail_line_count += 1

    capacities = infer_capacities(orders_by_key, load_by_key)
    capacities.update(capacity_registry())
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
        capacity = capacities.get(normalize_plate(plate)) or capacities.get(plate) or 30000
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
                "motorista": "",
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
  <link rel="icon" href="__FAVICON__" type="image/svg+xml">
  <title>Dashboard</title>
  <style>
    :root {
      --bg: #f4f6fb;
      --top: #34104f;
      --top-2: #4c176d;
      --panel: #ffffff;
      --panel-soft: #f8fafb;
      --ink: #16212d;
      --muted: #657282;
      --line: #d7e0e8;
      --teal: #64248c;
      --green: #20a86b;
      --rust: #f43f6e;
      --blue: #1268d9;
      --gold: #f59e0b;
      --orange: #fb8c00;
      --shadow: 0 14px 34px rgba(23, 32, 51, .08);
    }
    body[data-theme="forest"] {
      --bg: #f0f5ee;
      --top: #34104f;
      --top-2: #2b84cb;
      --teal: #64248c;
      --green: #2b84cb;
      --rust: #e2263c;
      --blue: #1b255f;
      --gold: #d72d51;
    }
    body[data-theme="graphite"] {
      --bg: #f0f0f6;
      --top: #34104f;
      --top-2: #1b255f;
      --teal: #64248c;
      --green: #2b84cb;
      --rust: #e2263c;
      --blue: #1b255f;
      --gold: #d72d51;
    }
    body[data-theme="marine"] {
      --bg: #edf4f7;
      --top: #1b255f;
      --top-2: #34104f;
      --teal: #64248c;
      --green: #2b84cb;
      --rust: #e2263c;
      --blue: #1b255f;
      --gold: #d72d51;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      background: var(--bg);
      color: var(--ink);
      font-family: Inter, Segoe UI, Roboto, Arial, sans-serif;
      overflow-x: hidden;
    }
    header {
      position: relative;
      overflow: hidden;
      padding: 22px clamp(16px, 4vw, 42px) 28px;
      background: radial-gradient(720px circle at 76% 35%, rgba(43,132,203,.34), transparent 62%), linear-gradient(135deg,#34104f,#4c176d 58%,#1b255f);
      color: #fff;
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 16px;
    }
    header::after {
      content:"";
      position:absolute;
      right:clamp(18px, 5vw, 72px);
      bottom:-86px;
      width:min(46vw,520px);
      aspect-ratio:1.8;
      background:url("__FAVICON__") center / contain no-repeat;
      opacity:.16;
      pointer-events:none;
    }
    header > * {
      position: relative;
      z-index: 2;
    }
    h1 { margin: 0; font-size: clamp(28px, 3vw, 38px); letter-spacing: 0; font-weight: 950; }
    .subtitle { margin: 8px 0 0; color: rgba(255,255,255,.9); }
    .brand-pill {
      display: inline-flex;
      align-items: center;
      gap: 12px;
      margin-bottom: 10px;
      color: #fff;
      font-size: 13px;
      font-weight: 900;
      text-transform: uppercase;
    }
    .brand-pill img {
      width: 74px;
      height: auto;
      object-fit: contain;
      filter: drop-shadow(0 6px 10px rgba(0, 0, 0, .24));
    }
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
      border: 1px solid var(--line);
      border-radius: 8px;
      color: #fff;
      text-decoration: none;
      font-size: 13px;
      font-weight: 800;
      background: rgba(255,255,255,.10);
      border-color: rgba(255,255,255,.30);
      box-shadow: inset 0 0 0 1px rgba(255,255,255,.06);
    }
    main { padding: 14px clamp(16px, 4vw, 44px) 42px; }
    .filters {
      display: grid;
      grid-template-columns: minmax(240px, 1.4fr) repeat(6, minmax(138px, 1fr));
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
      outline: 3px solid rgba(100, 36, 140, .18);
      border-color: var(--teal);
      background: #fff;
    }
    .kpis {
      display: grid;
      grid-template-columns: repeat(5, minmax(160px, 1fr));
      gap: 14px;
      margin-bottom: 16px;
    }
    .kpi, .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 10px;
      box-shadow: var(--shadow);
    }
    .kpi {
      min-height: 118px;
      padding: 20px;
      display: grid;
      grid-template-columns: 56px 1fr;
      gap: 14px;
      align-items: center;
    }
    .kpi-icon {
      width: 52px;
      height: 52px;
      border-radius: 12px;
      display: grid;
      place-items: center;
      color: #fff;
      background: linear-gradient(135deg, #64248c, #3f48cc);
      box-shadow: inset 0 0 0 1px rgba(255,255,255,.22);
    }
    .kpi-icon svg { width: 25px; height: 25px; stroke: currentColor; fill: none; stroke-width: 2.2; stroke-linecap: round; stroke-linejoin: round; }
    .kpi:nth-child(2) .kpi-icon { background: linear-gradient(135deg, #0b66d8, #2b84cb); }
    .kpi:nth-child(3) .kpi-icon { background: linear-gradient(135deg, #fb8c00, #f59e0b); }
    .kpi:nth-child(4) .kpi-icon { background: linear-gradient(135deg, #20a86b, #2dbb7f); }
    .kpi:nth-child(5) .kpi-icon { background: linear-gradient(135deg, #64248c, #7c3aed); }
    .kpi span { display: block; color: var(--muted); font-size: 12px; font-weight: 800; text-transform: uppercase; }
    .kpi strong { display: block; margin-top: 6px; font-size: 30px; line-height: 1.05; color: #111b26; }
    .kpi small { display:block; margin-top:8px; color:var(--muted); font-size:12px; font-weight:800; }
    .kpi small .up { color:#16a34a; }
    .kpi small .down { color:#e2263c; }
    .dashboard-grid {
      display: grid;
      grid-template-columns: minmax(0, 1.45fr) minmax(340px, .85fr);
      gap: 15px;
      align-items: start;
    }
    .lower-grid {
      display:grid;
      grid-template-columns:minmax(0, 1.45fr) minmax(340px, .85fr);
      gap:15px;
      align-items:start;
      margin-top:15px;
    }
    .side-stack { display:grid; gap:15px; }
    .panel { padding: 18px; overflow: hidden; }
    .hero-board, .grid { display:none !important; }
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
      grid-template-columns: minmax(0, 1fr) minmax(280px, 360px);
      gap: 15px;
      margin-bottom: 15px;
    }
    .focus-card {
      min-height: 250px;
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
      background: #34104f;
      color: #fff;
      font-size: clamp(22px, 3vw, 34px);
      font-weight: 900;
      letter-spacing: 0;
    }
    .focus-driver {
      margin-top: 8px;
      color: #344457;
      font-size: 14px;
      font-weight: 850;
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
      font-size: 24px;
      line-height: 1.1;
    }
    .mini-card.terminal-card {
      display: grid;
      align-content: start;
      gap: 11px;
    }
    .mini-card .bars {
      gap: 9px;
    }
    .mini-card .bar-row {
      grid-template-columns: minmax(94px, 130px) 1fr 44px;
      gap: 8px;
      min-height: 24px;
      font-size: 13px;
    }
    .mini-card .track { height: 9px; }
    .product-list {
      display: grid;
      gap: 12px;
    }
    .product-card {
      min-height: 410px;
    }
    .product-viz {
      display:grid;
      grid-template-columns:minmax(180px, 240px) 1fr;
      gap:18px;
      align-items:center;
    }
    .donut-wrap {
      position:relative;
      min-height:220px;
      display:grid;
      place-items:center;
    }
    .donut-total {
      position:absolute;
      inset:auto;
      display:grid;
      gap:4px;
      text-align:center;
      color:var(--muted);
      font-size:12px;
      font-weight:800;
    }
    .donut-total strong {
      color:var(--ink);
      font-size:22px;
    }
    .product-card .product-list { gap:9px; }
    .product-card .product-row {
      grid-template-columns:14px minmax(96px,1fr) minmax(78px,auto) 40px;
      gap:10px;
      padding-bottom:8px;
    }
    .product-dot { width:10px; height:10px; border-radius:50%; }
    .product-pct { color:var(--muted); text-align:right; font-weight:800; }
    .product-row {
      display: grid;
      grid-template-columns: minmax(0, 1fr) 110px;
      gap: 12px;
      align-items: center;
      padding-bottom: 10px;
      border-bottom: 1px solid var(--line);
    }
    .product-name { font-weight: 800; color: #253545; line-height: 1.18; overflow-wrap: anywhere; }
    .product-sub { color: var(--muted); font-size: 12px; margin-top: 3px; }
    .product-value { color: var(--ink); text-align: right; font-weight: 800; font-variant-numeric: tabular-nums; white-space: nowrap; }
    .line-chart { min-height: clamp(300px, 38vh, 440px); }
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
    .line-chart svg { width: 100%; height: auto; display: block; }
    .chart-grid { stroke: #dde6ed; stroke-width: 1; }
    .chart-axis { stroke: #b8c5d0; stroke-width: 1.2; }
    .chart-label { fill: var(--muted); font-size: 11px; }
    .chart-value { fill: #1a2a36; font-size: 12px; font-weight: 800; }
    .chart-value-volume { fill: #8d4c31; font-size: 11px; font-weight: 800; }
    .chart-line-trips { fill: none; stroke: var(--teal); stroke-width: 3; stroke-linecap: round; stroke-linejoin: round; }
    .chart-line-volume { fill: none; stroke: var(--rust); stroke-width: 3; stroke-linecap: round; stroke-linejoin: round; }
    .chart-area { fill: rgba(100, 36, 140, .10); }
    .chart-dot-trips { fill: #fff; stroke: var(--teal); stroke-width: 2.5; }
    .chart-dot-volume { fill: #fff; stroke: var(--rust); stroke-width: 2.5; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { padding: 12px 10px; border-bottom: 1px solid var(--line); text-align: left; vertical-align: middle; }
    tbody tr:nth-child(even) { background: #fafcfd; }
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
    .plate-pill { background:#dcfce7; color:#047857; font-weight:900; }
    .side-metric {
      display:grid;
      grid-template-columns:56px 1fr;
      gap:14px;
      align-items:center;
      min-height:118px;
    }
    .side-icon {
      width:52px;
      height:52px;
      border-radius:14px;
      display:grid;
      place-items:center;
      color:#fff;
      background:linear-gradient(135deg,#7c3aed,#9b5cf6);
    }
    .side-icon svg { width:25px; height:25px; stroke:currentColor; fill:none; stroke-width:2.2; stroke-linecap:round; stroke-linejoin:round; }
    .side-metric span { color:var(--muted); font-size:12px; font-weight:900; text-transform:uppercase; }
    .side-metric strong { display:block; margin-top:7px; font-size:30px; }
    .side-metric small { display:block; margin-top:8px; color:var(--muted); font-size:12px; font-weight:850; }
    .performance-card { min-height:200px; }
    .performance-body { display:grid; grid-template-columns:150px 1fr; gap:18px; align-items:center; }
    .perf-ring { position:relative; width:138px; height:138px; border-radius:50%; display:grid; place-items:center; background:conic-gradient(#20a86b 0deg, #20a86b var(--perfDeg), #edf1f5 var(--perfDeg), #edf1f5 360deg); }
    .perf-ring::before { content:""; width:92px; height:92px; border-radius:50%; background:#fff; box-shadow:inset 0 0 0 1px var(--line); }
    .perf-ring-label { position:absolute; text-align:center; font-weight:950; }
    .perf-ring-label strong { display:block; font-size:30px; }
    .perf-ring-label span { color:var(--muted); font-size:10px; }
    .perf-list { display:grid; gap:12px; font-weight:850; font-size:13px; }
    .perf-row { display:grid; grid-template-columns:10px 1fr auto; gap:8px; align-items:center; color:#344457; }
    .perf-row::before { content:""; width:8px; height:8px; border-radius:50%; background:#20a86b; }
    .perf-row:nth-child(3)::before { background:#f59e0b; }
    .note { color: var(--muted); font-size: 12px; margin-top: 10px; }
    .empty { color: var(--muted); padding: 22px; text-align: center; }
    @media (max-width: 1440px) {
      main {
        padding-inline: 14px;
        zoom: .86;
        width: calc(100% / .86);
      }
      header { padding: 18px 24px 24px; }
      h1 { font-size: 34px; }
      .subtitle { margin-top: 6px; }
      .kpis { grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 12px; }
      .kpi { min-height: 104px; padding: 16px; grid-template-columns: 48px 1fr; gap: 12px; }
      .kpi-icon { width: 44px; height: 44px; border-radius: 10px; }
      .kpi-icon svg { width: 22px; height: 22px; }
      .kpi strong { font-size: 28px; }
      .dashboard-grid, .lower-grid { grid-template-columns: minmax(0, 1.4fr) minmax(300px, .6fr); }
      .product-viz { grid-template-columns: minmax(170px, 210px) minmax(0, 1fr); gap: 14px; }
      .donut-wrap { min-height: 190px; }
      .donut-wrap svg { width: 190px; height: 190px; }
      .product-card .product-row { grid-template-columns:12px minmax(86px,1fr) minmax(72px,auto) 36px; gap:8px; }
    }
    @media (max-width: 1180px) {
      .kpis { grid-template-columns: repeat(3, minmax(0, 1fr)); }
      .dashboard-grid, .lower-grid { grid-template-columns: 1fr; }
      .side-stack { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .performance-card { grid-column: 1 / -1; }
    }
    @media (max-width: 980px) {
      main { zoom: 1; width: auto; }
      .filters, .kpis { grid-template-columns: 1fr 1fr; }
      .dashboard-grid, .lower-grid { grid-template-columns: 1fr; }
      .wide { grid-column: 1 / -1; }
      .product-viz { grid-template-columns: 1fr; }
      .side-stack { grid-template-columns: 1fr; }
    }
    @media (max-width: 650px) {
      header { flex-direction: column; }
      .nav { justify-content: flex-start; }
      .filters, .kpis, .dashboard-grid, .lower-grid, .performance-body { grid-template-columns: 1fr; }
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
      <div class="brand-pill"><img src="__FAVICON__" alt="">Grupo Dislub Equador</div>
      <h1>Dashboard</h1>
      <p class="subtitle">Viagens por placa, produtos carregados e terminais Equador/Ipiranga.</p>
    </div>
    <a class="top-link" href="/editar">Atualizar dados</a>
  </header>
  <main>
    <section class="kpis">
      <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24"><path d="M3 16V8h11v8"></path><path d="M14 11h4l3 3v2h-7"></path><circle cx="7" cy="18" r="2"></circle><circle cx="17" cy="18" r="2"></circle></svg></div><div><span>Viagens</span><strong id="kTrips">0</strong><small>Total no filtro</small></div></div>
      <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24"><path d="M4 19h16"></path><path d="M7 16V9"></path><path d="M12 16V5"></path><path d="M17 16v-7"></path><path d="M9 9h6"></path></svg></div><div><span>Volume carregado</span><strong id="kQty">0</strong><small>Litros carregados no filtro</small></div></div>
      <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24"><path d="M3 11h18"></path><path d="M5 7h14v10H5z"></path><circle cx="8" cy="17" r="2"></circle><circle cx="16" cy="17" r="2"></circle></svg></div><div><span>Placas ativas</span><strong id="kPlates">0</strong><small>Placas unicas no filtro</small></div></div>
      <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24"><path d="M7 3h8l4 4v14H7z"></path><path d="M15 3v5h5"></path><path d="M10 13h6"></path><path d="M10 17h4"></path></svg></div><div><span>Notas fiscais</span><strong id="kNotes">0</strong><small>Notas informadas no filtro</small></div></div>
      <div class="kpi"><div class="kpi-icon"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"></circle><path d="M12 7v5l3 2"></path><path d="m8 16 2-2"></path></svg></div><div><span>Capacidade media</span><strong id="kUtilization">0%</strong><small>Carregado / capacidade prevista</small></div></div>
    </section>

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
      <label>Motorista
        <select id="driver"></select>
      </label>
      <label>Tema
        <select id="theme">
          <option value="marine">Marine</option>
          <option value="forest">Forest</option>
          <option value="graphite">Graphite</option>
        </select>
      </label>
    </section>

    <section class="dashboard-grid">
      <div class="panel">
        <div class="chart-toolbar">
          <h2>Evolucao por Dia</h2>
          <div class="segment" aria-label="Segmentacao da evolucao diaria">
            <button type="button" class="active" data-chart-mode="both">Diario</button>
            <button type="button" data-chart-mode="trips">Viagens</button>
            <button type="button" data-chart-mode="volume">Volume</button>
          </div>
        </div>
        <div id="chartTotal" class="chart-total"></div>
        <div id="lineChart" class="line-chart"></div>
      </div>
      <div class="panel product-card">
        <h2>Top produtos</h2>
        <div id="productDonut" class="product-viz"></div>
      </div>
    </section>

    <section class="lower-grid">
      <div class="panel">
        <h2>Viagens no Dia por Placa</h2>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Data</th>
                <th>Placa</th>
                <th>Motorista</th>
                <th>Terminal</th>
                <th>Viagens</th>
                <th>Capacidade</th>
                <th>Carregado</th>
                <th>Notas</th>
              </tr>
            </thead>
            <tbody id="dailyTable"></tbody>
          </table>
        </div>
      </div>
      <div class="side-stack">
        <div class="panel mini-card terminal-card"><span>Terminais</span><div id="terminalSummaryTop" class="bars"></div></div>
        <div class="panel side-metric">
          <div class="side-icon"><svg viewBox="0 0 24 24"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M22 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg></div>
          <div><span>Motoristas ativos</span><strong id="kDrivers">0</strong><small>Motoristas com viagens no filtro</small></div>
        </div>
        <div class="panel performance-card">
          <h2>Uso de capacidade</h2>
          <div class="performance-body">
            <div id="capacityRing" class="perf-ring" style="--perfDeg:0deg"><div class="perf-ring-label"><strong id="kCapacityUsed">0%</strong><span>Carregado / previsto</span></div></div>
            <div class="perf-list">
              <div class="perf-row"><span>Volume carregado</span><strong id="capacityVolume">0</strong></div>
              <div class="perf-row"><span>Capacidade prevista</span><strong id="capacityTotal">0</strong></div>
              <div class="perf-row"><span>Viagens consideradas</span><strong id="capacityTrips">0</strong></div>
              <div class="perf-row"><span>Placas ativas</span><strong id="capacityPlates">0</strong></div>
            </div>
          </div>
        </div>
      </div>
    </section>

    <section class="hero-board">
      <div class="panel focus-card">
        <div class="focus-head">
          <div>
            <div class="focus-label">Placa destaque</div>
            <div id="focusPlate" class="focus-plate-pill">-</div>
            <div id="focusDriver" class="focus-driver">-</div>
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
        <div class="mini-card terminal-card"><span>Terminais</span><div id="terminalSummaryTop" class="bars"></div></div>
        <div class="mini-card"><span>Motoristas</span><strong id="kDrivers">0</strong></div>
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
        <h2>Produtos</h2>
        <div id="productList" class="product-list"></div>
      </div>
      <div class="panel wide">
        <h2>Viagens no Dia por Placa</h2>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Data</th>
                <th>Placa</th>
                <th>Motorista</th>
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
    const palette = ["#64248c", "#2b84cb", "#e2263c", "#1b255f", "#d72d51", "#8b3fb5", "#4c176d", "#56616b"];
    let chartMode = "both";

    function unique(key) {
      return [...new Set(rows.map((row) => row[key]).filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), "pt-BR"));
    }

    function uniqueDrivers() {
      return [...new Set(rows.flatMap((row) => String(row.motorista || "").split("/").map((item) => item.trim()).filter(Boolean)))]
        .sort((a, b) => a.localeCompare(b, "pt-BR"));
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
      const driver = $("driver").value;
      return rows.filter((row) => {
        const haystack = [row.placa, row.motorista, row.terminalNome, row.mixProdutos, ...row.produtos.map((item) => item.produto)].join(" ").toLowerCase();
        const rowDrivers = String(row.motorista || "").split("/").map((item) => item.trim());
        const rowDate = brToIso(row.data);
        return (!query || haystack.includes(query))
          && (!dateStart || rowDate >= dateStart)
          && (!dateEnd || rowDate <= dateEnd)
          && (!terminal || row.terminal === terminal)
          && (!plate || row.placa === plate)
          && (!driver || rowDrivers.includes(driver));
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

    function productDonut(products, totalQty) {
      const visible = products.slice(0, 7);
      if (!visible.length) {
        $("productDonut").innerHTML = `<div class="empty">Sem dados</div>`;
        return;
      }
      const total = visible.reduce((sum, item) => sum + item[1], 0) || 1;
      let offset = 25;
      const circles = visible.map(([label, value], idx) => {
        const segment = value / total * 100;
        const item = `<circle r="42" cx="60" cy="60" pathLength="100" fill="none" stroke="${palette[idx % palette.length]}" stroke-width="22" stroke-dasharray="${segment} ${100 - segment}" stroke-dashoffset="${offset}"><title>${productShort(label)}: ${volume(value)}</title></circle>`;
        offset -= segment;
        return item;
      }).join("");
      $("productDonut").innerHTML = `
        <div class="donut-wrap">
          <svg viewBox="0 0 120 120" width="220" height="220" aria-label="Top produtos">
            <circle r="42" cx="60" cy="60" fill="none" stroke="#edf1f5" stroke-width="22"></circle>
            <g transform="rotate(-90 60 60)">${circles}</g>
          </svg>
          <div class="donut-total"><span>Total carregado</span><strong>${volume(totalQty)}</strong></div>
        </div>
        <div class="product-list">
          ${visible.map(([label, value], idx) => {
            const pct = totalQty ? Math.round(value / totalQty * 100) : 0;
            return `<div class="product-row"><span class="product-dot" style="background:${palette[idx % palette.length]}"></span><div class="product-name">${productShort(label)}</div><div class="product-value">${volume(value)}</div><div class="product-pct">${pct}%</div></div>`;
          }).join("")}
        </div>
      `;
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
      const boxWidth = $("lineChart").clientWidth || 1120;
      const width = Math.max(720, Math.round(boxWidth));
      const viewportHeight = window.innerHeight || 900;
      const height = Math.round(Math.max(250, Math.min(440, viewportHeight * .34, width * .26)));
      const pad = {
        left: width > 1100 ? 64 : 48,
        right: width > 1100 ? 52 : 36,
        top: 36,
        bottom: 44
      };
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
        $("focusDriver").textContent = "-";
        $("focusTrips").textContent = "0";
        $("focusVolume").textContent = "0";
        $("focusCapacity").textContent = "0";
        $("focusProducts").innerHTML = `<div class="empty">Sem dados</div>`;
        return;
      }
      const plateRows = data.filter((row) => row.placa === topPlate);
      const drivers = [...new Set(plateRows.flatMap((row) => String(row.motorista || "").split("/").map((item) => item.trim()).filter(Boolean)))];
      $("focusPlate").textContent = topPlate;
      $("focusDriver").textContent = drivers.join(" / ") || "Sem motorista informado";
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
      const activePlates = new Set(data.map((row) => row.placa)).size;
      const plateTrips = sumBy(data, "placa", "viagens");
      const plateQty = sumBy(data, "placa", "quantidade");
      const products = productTotals(data);
      const capacityTotal = data.reduce((total, row) => total + ((row.capacidade || 0) * Math.max(1, row.viagens || 0)), 0);
      const utilization = capacityTotal ? Math.round(qty / capacityTotal * 100) : 0;

      $("kTrips").textContent = fmt.format(trips);
      $("kQty").textContent = volume(qty);
      $("kPlates").textContent = fmt.format(activePlates);
      $("kNotes").textContent = fmt.format(notes);
      $("kUtilization").textContent = `${fmt.format(utilization)}%`;
      $("kTopProduct").textContent = products[0]?.[0]?.split(" ").slice(0, 2).join(" ") || "-";
      $("kDrivers").textContent = fmt.format(new Set(data.flatMap((row) => String(row.motorista || "").split("/").map((item) => item.trim()).filter(Boolean))).size);
      const bestDay = dailyTotals(data).sort((a, b) => b.viagens - a.viagens)[0];
      $("kBestDay").textContent = bestDay ? bestDay.data.slice(0, 5) : "-";
      $("kCapacityUsed").textContent = `${fmt.format(utilization)}%`;
      $("capacityRing").style.setProperty("--perfDeg", `${Math.min(100, utilization) / 100 * 360}deg`);
      $("capacityVolume").textContent = volume(qty);
      $("capacityTotal").textContent = volume(capacityTotal);
      $("capacityTrips").textContent = fmt.format(trips);
      $("capacityPlates").textContent = fmt.format(activePlates);

      bars("terminalSummaryTop", terminalTotals(data), 2, "var(--blue)");
      productList("productList", products, 8);
      productDonut(products, qty);
      lineChart(data);
      focusPanel(data, plateTrips, plateQty);

      $("dailyTable").innerHTML = data
        .slice()
        .sort((a, b) => b.viagens - a.viagens || b.quantidade - a.quantidade)
        .map((row) => `
          <tr>
            <td>${row.data}</td>
            <td><span class="pill plate-pill">${row.placa}</span></td>
            <td>${row.motorista || "-"}</td>
            <td class="terminal-cell">${row.terminal} - ${row.terminalNome}</td>
            <td class="num-cell">${fmt.format(row.viagens)}</td>
            <td class="num-cell">${volume(row.capacidade)}</td>
            <td class="load-cell">
              <div class="load-value"><span>${volume(row.quantidade)}</span><span>${Math.round((row.quantidade / Math.max(1, row.capacidade * Math.max(1, row.viagens))) * 100)}%</span></div>
              <div class="mini-track"><div class="mini-fill" style="width:${Math.min(100, (row.quantidade / Math.max(1, row.capacidade * Math.max(1, row.viagens))) * 100)}%"></div></div>
            </td>
            <td class="num-cell">${fmt.format(row.notas)}</td>
          </tr>
        `).join("") || `<tr><td colspan="8" class="empty">Nenhum registro encontrado</td></tr>`;
    }

    fillSelect("terminal", [["10", "10 - Equador"], ["19", "19 - Ipiranga"]].map((item) => item.join(" - ").replace(" - ", " - ")), "Todos os terminais");
    $("terminal").innerHTML = `<option value="">Todos os terminais</option><option value="10">10 - Equador</option><option value="19">19 - Ipiranga</option>`;
    fillSelect("plate", unique("placa"), "Todas as placas");
    fillSelect("driver", uniqueDrivers(), "Todos os motoristas");
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
    let resizeTimer;
    window.addEventListener("resize", () => {
      window.clearTimeout(resizeTimer);
      resizeTimer = window.setTimeout(() => lineChart(filteredRows()), 120);
    });
    ["search", "dateStart", "dateEnd", "terminal", "plate", "driver"].forEach((id) => $(id).addEventListener("input", render));
    render();
  </script>
</body>
</html>
"""


def main() -> None:
    data = build_data()
    html = (
        HTML_TEMPLATE.replace("__DATA__", json.dumps(data, ensure_ascii=False))
        .replace("__FAVICON__", FAVICON_URL)
    )
    OUTPUT_PATH.write_text(html, encoding="utf-8")
    meta = data["meta"]
    print(f"Dashboard criado: {OUTPUT_PATH}")
    print(f"Ordens: {meta['orderRows']} | Linhas detalhadas usadas: {meta['detailRows']}")
    print(f"Linhas dia/placa/terminal: {len(data['dailyPlateRows'])}")


if __name__ == "__main__":
    os.environ["BUILDING_DASHBOARD"] = "1"
    main()
