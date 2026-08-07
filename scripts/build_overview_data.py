#!/usr/bin/env python3
import csv
import json
import re
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data" / "processed"
OLD_CSV = Path("/Users/veronikalagutkina/Downloads/Кефир. Вероника. Основная таблица. Условия работы - июль.csv")
KPI_SOURCES = [
    {"path": OLD_CSV, "label": "Основная таблица, июль", "kind": "csv"},
    {"path": Path("/Users/veronikalagutkina/Downloads/Вероника 22.xlsx"), "label": "Вероника 22", "kind": "xlsx"},
    {"path": Path("/Users/veronikalagutkina/Downloads/Вероника 7.08.xlsx"), "label": "Вероника 7.08", "kind": "xlsx"},
]

TARGET_DIRECTIONS = ["Сантехника", "Электрика", "Кофемашины", "Кондиционеры", "Телевизоры", "Холодильники"]
FRIDGE_CITIES = ["Москва", "Санкт-Петербург", "Ростов-на-Дону", "Воронеж", "Краснодар", "Новосибирск"]
DEFAULT_CITIES = ["Москва", "Санкт-Петербург"]

RU_MONTHS = {
    "января": 1,
    "февраля": 2,
    "марта": 3,
    "апреля": 4,
    "мая": 5,
    "июня": 6,
    "июля": 7,
    "августа": 8,
    "сентября": 9,
    "октября": 10,
    "ноября": 11,
    "декабря": 12,
}


def clean_header(value):
    return str(value or "").replace("\n", " ").strip()


def clean_text(value):
    text = str(value or "").replace(",,,,", "").strip()
    text = re.sub(r"\s+", " ", text)
    if text == "Кинякин Денис":
        return "Киякин Денис"
    text = re.sub(r"\s+\(БАН\)$", "", text)
    return text


def numeric(value):
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
    text = re.sub(r"[^0-9.\-]", "", text)
    return float(text) if text else 0.0


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    match = re.match(r"(\d{1,2})\s+([а-яё]+)\s+(\d{4})", text, re.I)
    if match:
        return date(int(match.group(3)), RU_MONTHS[match.group(2).lower()], int(match.group(1)))
    return datetime.fromisoformat(text).date()


def normalize_direction(value):
    text = clean_text(value)
    aliases = {
        "Холодильники, морозильные камеры": "Холодильники",
        "Кондиционеры, вентиляция": "Кондиционеры",
        "": "Без направления",
    }
    return aliases.get(text, text)


def row_from_values(row, idx):
    direction = normalize_direction(row[idx["Группы ТО в Отчёте KPI - Основные.Название"]])
    day = parse_date(row[idx["День.Название"]])
    return {
        "date": day.isoformat(),
        "profile": clean_text(row[idx["Группа линии.Название"]]),
        "direction": direction,
        "city": clean_text(row[idx["Город.Название"]]),
        "promo": numeric(row[idx.get("Итог2.Продвижение", -1)]) if idx.get("Итог2.Продвижение", -1) >= 0 else 0.0,
        "tariff": numeric(row[idx.get("Итог2.Тариф", -1)]) if idx.get("Итог2.Тариф", -1) >= 0 else 0.0,
        "spend": numeric(row[idx.get("Итог.Затраты Сумм", -1)]) if idx.get("Итог.Затраты Сумм", -1) >= 0 else 0.0,
        "orders": numeric(row[idx.get("Итог.Принято заказов", -1)]) if idx.get("Итог.Принято заказов", -1) >= 0 else 0.0,
        "calls": numeric(row[idx.get("КЦ.Звонки всего", -1)]) if idx.get("КЦ.Звонки всего", -1) >= 0 else 0.0,
        "target_calls": numeric(row[idx.get("КЦ.Целевые звонки", -1)]) if idx.get("КЦ.Целевые звонки", -1) >= 0 else 0.0,
        "views": numeric(row[idx.get("Реклама.Просмотры объявления", -1)]) if idx.get("Реклама.Просмотры объявления", -1) >= 0 else 0.0,
        "contacts": numeric(row[idx.get("Реклама. Контакты", -1)]) if idx.get("Реклама. Контакты", -1) >= 0 else 0.0,
        "rating": numeric(row[idx.get("Реклама.Рейтинг продавца", -1)]) if idx.get("Реклама.Рейтинг продавца", -1) >= 0 else 0.0,
    }


def read_csv_rows(path):
    text = path.read_text(encoding="utf-8-sig")
    dialect = csv.Sniffer().sniff(text[:4096])
    reader = csv.reader(text.splitlines(), dialect)
    headers = [clean_header(x) for x in next(reader)]
    idx = {h: i for i, h in enumerate(headers)}
    return [row_from_values(row, idx) for row in reader if any(row)]


def read_xlsx_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    ws.reset_dimensions()
    rows = ws.iter_rows(values_only=True)
    headers = [clean_header(x) for x in next(rows)]
    idx = {h: i for i, h in enumerate(headers)}
    return [row_from_values(row, idx) for row in rows if any(v is not None for v in row)]


def add_metrics(row):
    row["cpo"] = round(row["spend"] / row["orders"], 4) if row["orders"] else None
    row["cp_call"] = round(row["spend"] / row["calls"], 4) if row["calls"] else None
    row["cp_contact"] = round(row["spend"] / row["contacts"], 4) if row["contacts"] else None
    row["call_to_order"] = round(row["orders"] / row["calls"], 4) if row["calls"] else None
    row["contact_to_order"] = round(row["orders"] / row["contacts"], 4) if row["contacts"] else None
    row["view_to_contact"] = round(row["contacts"] / row["views"], 4) if row["views"] else None
    row["active"] = bool(row["spend"] or row["orders"] or row["calls"] or row["contacts"] or row["views"])
    row["threshold"] = 1200 if row.get("direction") in ["Сантехника", "Электрика"] else 1800
    if row["orders"] == 0 and row["spend"] > 0:
        row["status"] = "no_orders"
    elif row["orders"] > 0 and row["cpo"] and row["cpo"] > row["threshold"]:
        row["status"] = "high_cpo"
    elif row["orders"] > 0:
        row["status"] = "ok"
    else:
        row["status"] = "inactive"
    return row


def aggregate(rows, keys):
    result = {}
    for row in rows:
        key = tuple(row[k] for k in keys)
        if key not in result:
            result[key] = {k: row[k] for k in keys}
            for metric in ["spend", "orders", "calls", "target_calls", "views", "contacts", "promo", "tariff"]:
                result[key][metric] = 0.0
        for metric in ["spend", "orders", "calls", "target_calls", "views", "contacts", "promo", "tariff"]:
            result[key][metric] += row.get(metric, 0.0)
    return [add_metrics(v) for v in result.values()]


def expected_cities(direction):
    return FRIDGE_CITIES if direction == "Холодильники" else DEFAULT_CITIES


def format_period(dates):
    first, last = min(dates), max(dates)
    if first.year == last.year and first.month == last.month:
        month_name = [k for k, v in RU_MONTHS.items() if v == first.month][0]
        return f"{first.day}-{last.day} {month_name} {last.year}"
    first_month = [k for k, v in RU_MONTHS.items() if v == first.month][0]
    last_month = [k for k, v in RU_MONTHS.items() if v == last.month][0]
    return f"{first.day} {first_month} - {last.day} {last_month} {last.year}"


def fmt_int(value):
    return f"{round(value):,}".replace(",", " ")


def build_insights(campaigns, by_direction, by_profile):
    insights = []
    good_dirs = [
        x for x in by_direction
        if x["direction"] in TARGET_DIRECTIONS and x["orders"] >= 3 and x["spend"] >= 1000 and x["cpo"]
    ]
    if good_dirs:
        best = min(good_dirs, key=lambda x: x["cpo"])
        insights.append(
            f"{best['direction']} сейчас дает самый низкий CPO среди направлений с существенным объемом: {fmt_int(best['cpo'])} ₽ при {int(best['orders'])} заказах. Это направление стоит использовать как эталон по бюджету, гео и обработке лидов."
        )
    no_orders = [x for x in campaigns if x["orders"] == 0 and x["spend"] >= 1000]
    if no_orders:
        worst = max(no_orders, key=lambda x: x["spend"])
        insights.append(
            f"Связка {worst['profile']} · {worst['direction']} · {worst['city']} потратила {fmt_int(worst['spend'])} ₽ без заказов. Сначала проверяем качество трафика и обработку контактов, затем решаем: наблюдать или заменить."
        )
    high_cpo = [x for x in campaigns if x["orders"] > 0 and x["cpo"] and x["cpo"] > x["threshold"]]
    if high_cpo:
        item = max(high_cpo, key=lambda x: x["spend"])
        insights.append(
            f"У {item['profile']} в связке {item['direction']} · {item['city']} CPO выше ориентира: {fmt_int(item['cpo'])} ₽. Тут нужна точечная оптимизация ставок, объявления и качества контактов."
        )
    contact_candidates = [x for x in campaigns if x["contacts"] >= 10 and (not x["orders"] or x["contact_to_order"] is not None and x["contact_to_order"] < 0.08)]
    if contact_candidates:
        item = max(contact_candidates, key=lambda x: x["contacts"])
        insights.append(
            f"{item['direction']} · {item['city']} у {item['profile']} набрала {int(item['contacts'])} контактов, но слабо довела их до заказов. Это зона проверки скрипта, дозвона, цены после консультации и причин отказа."
        )
    small_profiles = [x for x in by_profile if x["spend"] > 0 and x["orders"] <= 1]
    if small_profiles:
        names = ", ".join(x["profile"] for x in sorted(small_profiles, key=lambda x: x["spend"], reverse=True)[:3])
        insights.append(f"Новые или малые профили пока лучше не сравнивать с основными напрямую: по ним мало заказов. Для них полезнее смотреть запуск направлений и первые сигналы по контактам: {names}.")
    insights.append("Для детального блока объявлений нужны свежие Avito Pro и XML-выгрузки по тем же профилям и периоду: только там есть показы, CTR из показов в просмотр, первое фото, заголовки и тексты.")
    return insights[:6]


def read_source(source):
    path = source["path"]
    if not path.exists():
        return []
    if source["kind"] == "csv":
        return read_csv_rows(path)
    return read_xlsx_rows(path)


def summarize_source(source, rows):
    if not rows:
        return {
            "label": source["label"],
            "file": str(source["path"]),
            "rows": 0,
            "period": "нет данных",
            "orders": 0,
            "spend": 0,
            "contacts": 0,
            "views": 0,
        }
    totals = aggregate([add_metrics(row.copy()) for row in rows], [])[0]
    dates = [datetime.fromisoformat(row["date"]).date() for row in rows]
    return {
        "label": source["label"],
        "file": str(source["path"]),
        "rows": len(rows),
        "period": format_period(dates),
        "orders": round(totals["orders"], 2),
        "spend": round(totals["spend"], 2),
        "contacts": round(totals["contacts"], 2),
        "views": round(totals["views"], 2),
    }


def build_payload(daily_rows, source_summary):
    daily_rows = [add_metrics(row.copy()) for row in daily_rows]
    daily_rows.sort(key=lambda x: (x["date"], x["profile"], x["direction"], x["city"]))
    by_key = {}
    for row in daily_rows:
        by_key[(row["profile"], row["direction"], row["city"], row["date"])] = row

    campaigns = aggregate(daily_rows, ["profile", "direction", "city"])
    campaigns.sort(key=lambda x: (x["orders"] == 0 and x["spend"] > 0, x["cpo"] or 999999, -x["orders"]))
    by_profile = sorted(aggregate(daily_rows, ["profile"]), key=lambda x: x["spend"], reverse=True)
    by_direction = sorted(aggregate(daily_rows, ["direction"]), key=lambda x: x["spend"], reverse=True)
    by_city = sorted(aggregate(daily_rows, ["city"]), key=lambda x: x["city"])
    daily = sorted(aggregate(daily_rows, ["date"]), key=lambda x: x["date"])

    lookup = {(x["profile"], x["direction"], x["city"]): x for x in campaigns}
    profiles = sorted({row["profile"] for row in daily_rows if row["profile"]})
    coverage = []
    for profile in profiles:
        for direction in TARGET_DIRECTIONS:
            for city in expected_cities(direction):
                row = lookup.get((profile, direction, city), {})
                spend, orders = row.get("spend", 0.0), row.get("orders", 0.0)
                status = "есть заказы" if orders > 0 else ("есть траты, нет заказов" if spend > 0 else "нет запуска")
                coverage.append(add_metrics({
                    "profile": profile,
                    "direction": direction,
                    "city": city,
                    "spend": spend,
                    "orders": orders,
                    "calls": row.get("calls", 0.0),
                    "target_calls": row.get("target_calls", 0.0),
                    "views": row.get("views", 0.0),
                    "contacts": row.get("contacts", 0.0),
                    "promo": row.get("promo", 0.0),
                    "tariff": row.get("tariff", 0.0),
                    "status": status,
                }) | {"status": status})

    totals = aggregate(daily_rows, [])[0]
    dates = [datetime.fromisoformat(row["date"]).date() for row in daily_rows]
    payload = {
        "generated": date.today().isoformat(),
        "period": format_period(dates),
        "source_files": [item["file"] for item in source_summary if item["rows"]],
        "sourceSummary": source_summary,
        "totals": {k: round(totals[k], 2) for k in ["spend", "orders", "calls", "target_calls", "views", "contacts"]},
        "byProfile": by_profile,
        "byDirection": by_direction,
        "byCity": by_city,
        "campaigns": campaigns,
        "coverage": coverage,
        "daily": daily,
        "insights": build_insights(campaigns, by_direction, by_profile),
        "profiles": profiles,
        "directions": sorted({row["direction"] for row in daily_rows if row["direction"]}),
        "cities": sorted({row["city"] for row in daily_rows if row["city"]}),
        "targetDirections": TARGET_DIRECTIONS,
        "dailyRows": daily_rows,
    }
    return payload


def main():
    source_summary = []
    by_key = {}
    for source in KPI_SOURCES:
        rows = read_source(source)
        source_summary.append(summarize_source(source, rows))
        for row in rows:
            by_key[(row["profile"], row["direction"], row["city"], row["date"])] = row

    daily_rows = list(by_key.values())
    payload = build_payload(daily_rows, source_summary)

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    (DATA_DIR / "overview_data.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    (DATA_DIR / "overview_data.js").write_text("window.OVERVIEW_DATA = " + json.dumps(payload, ensure_ascii=False) + ";\n", encoding="utf-8")
    (DATA_DIR / "kpi_table_data.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    (DATA_DIR / "kpi_table_data.js").write_text("window.KPI_TABLE_DATA = " + json.dumps(payload, ensure_ascii=False) + ";\n", encoding="utf-8")
    print(json.dumps({
        "period": payload["period"],
        "rows": len(daily_rows),
        "profiles": len(payload["profiles"]),
        "campaigns": len(payload["campaigns"]),
        "totals": payload["totals"],
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
