from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import os
import re
import tempfile
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


ACTIVE_TECHS = {"Innocent Mhora", "Talent Mutanda"}
MONTH_LABELS = {
    1: "Jan",
    2: "Feb",
    3: "Mar",
    4: "Apr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Aug",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dec",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build dashboard_data.json from the Liseo workbook.")
    parser.add_argument("--workbook", help="Local workbook path (.xlsx/.xlsm).")
    parser.add_argument("--workbook-url", help="Public share or direct download URL for the workbook.")
    parser.add_argument("--output", default="dashboard_data.json", help="Where to write the dashboard JSON.")
    return parser.parse_args()


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def canonical_sku(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return clean_text(value)


def to_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = clean_text(value).replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> int | None:
    num = to_float(value)
    if num is None:
        return None
    return int(round(num))


def parse_date_value(value: Any) -> dt.date | None:
    if value is None or value == "":
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = clean_text(value)
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def normalize_month_label(value: Any, month_num: int | None) -> str:
    text = clean_text(value)
    if text:
        short = text[:3].title()
        if short in MONTH_LABELS.values():
            return short
    if month_num in MONTH_LABELS:
        return MONTH_LABELS[month_num]
    return ""


def add_download_variant(url: str) -> str:
    parsed = urllib.parse.urlparse(url)
    query = dict(urllib.parse.parse_qsl(parsed.query, keep_blank_values=True))
    query["download"] = "1"
    return urllib.parse.urlunparse(parsed._replace(query=urllib.parse.urlencode(query)))


def extract_download_url_from_html(page_html: str) -> str | None:
    patterns = [
        r'"downloadUrl":"([^"]+)"',
        r'"@content\.downloadUrl":"([^"]+)"',
        r'"downloadUrl\\":\\"([^"]+)\\"',
    ]
    for pattern in patterns:
        match = re.search(pattern, page_html)
        if not match:
            continue
        candidate = match.group(1)
        candidate = candidate.replace("\\u0026", "&").replace("\\/", "/").replace('\\"', '"')
        return html.unescape(candidate)
    return None


def write_temp_workbook(payload: bytes) -> Path:
    fd, temp_path = tempfile.mkstemp(suffix=".xlsm")
    os.close(fd)
    path = Path(temp_path)
    path.write_bytes(payload)
    return path


def download_workbook(url: str) -> Path:
    headers = {"User-Agent": "Mozilla/5.0 Codex Dashboard Refresher"}
    candidates = [url]
    if "download=1" not in url:
        candidates.append(add_download_variant(url))

    last_error: Exception | None = None
    for candidate in candidates:
        try:
            request = urllib.request.Request(candidate, headers=headers)
            with urllib.request.urlopen(request, timeout=60) as response:
                payload = response.read()
                content_type = response.headers.get("Content-Type", "")
                if payload.startswith(b"PK\x03\x04"):
                    return write_temp_workbook(payload)
                if "html" in content_type.lower():
                    nested_url = extract_download_url_from_html(payload.decode("utf-8", errors="ignore"))
                    if nested_url:
                        nested_request = urllib.request.Request(nested_url, headers=headers)
                        with urllib.request.urlopen(nested_request, timeout=60) as nested_response:
                            nested_payload = nested_response.read()
                            if nested_payload.startswith(b"PK\x03\x04"):
                                return write_temp_workbook(nested_payload)
                raise RuntimeError(f"Workbook URL did not return an Excel file: {candidate}")
        except Exception as exc:  # noqa: BLE001
            last_error = exc
    raise RuntimeError(f"Unable to download workbook from {url}") from last_error


def resolve_workbook_source(workbook: str | None, workbook_url: str | None) -> tuple[Path, bool]:
    if workbook:
        path = Path(workbook).expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(f"Workbook not found: {path}")
        return path, False

    env_path = os.getenv("WORKBOOK_PATH")
    if env_path:
        path = Path(env_path).expanduser().resolve()
        if path.exists():
            return path, False

    url = workbook_url or os.getenv("WORKBOOK_URL")
    if url:
        return download_workbook(url), True

    raise ValueError("Provide --workbook, WORKBOOK_PATH, --workbook-url, or WORKBOOK_URL.")


def load_reference_map(workbook_path: Path) -> dict[str, dict[str, Any]]:
    wb = load_workbook(workbook_path, data_only=True, read_only=True)
    ws = wb["Reference"]
    ref_map: dict[str, dict[str, Any]] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        sku = canonical_sku(row[0])
        if not sku:
            continue
        ref_map[sku] = {
            "model": clean_text(row[1]),
            "difficulty": to_float(row[2]) or 0.0,
            "rate": to_float(row[3]) or 0.0,
        }
    wb.close()
    return ref_map


def build_rows(workbook_path: Path, ref_map: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    wb = load_workbook(workbook_path, data_only=True, read_only=True)
    ws = wb["Production Data"]
    rows: list[dict[str, Any]] = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        qc_date = parse_date_value(row[1])
        qty = to_int(row[5])
        sku = canonical_sku(row[6])
        tech = clean_text(row[8])

        if not qc_date and qty is None and not sku and not tech:
            continue
        if qty is None or qty <= 0 or not tech:
            continue

        year = to_int(row[2]) or (qc_date.year if qc_date else None)
        month_num = to_int(row[4]) or (qc_date.month if qc_date else None)
        if year is None or month_num is None:
            continue

        month_label = normalize_month_label(row[3], month_num)
        ref = ref_map.get(sku, {})
        difficulty = to_float(row[10])
        if difficulty in (None, 0.0):
            difficulty = float(ref.get("difficulty", 0.0))

        total_points = to_float(row[11])
        if total_points is None:
            total_points = round(qty * difficulty, 2)

        status = clean_text(row[9]) or ("Active" if tech in ACTIVE_TECHS else "Inactive")
        rows.append(
            {
                "tech": tech,
                "yr": year,
                "mo": month_label,
                "moNum": month_num,
                "u": qty,
                "p": round(total_points, 1),
                "s": status,
            }
        )

    wb.close()
    return rows


def build_payload(rows: list[dict[str, Any]], workbook_name: str) -> dict[str, Any]:
    generated_at = dt.datetime.now(dt.timezone.utc)
    total_units = sum(int(row["u"]) for row in rows)
    total_points = round(sum(float(row["p"]) for row in rows), 1)
    techs = sorted({row["tech"] for row in rows})
    active_techs = sorted({row["tech"] for row in rows if row["s"] == "Active"})

    return {
        "meta": {
            "workbookName": workbook_name,
            "generatedAt": generated_at.isoformat(),
            "generatedAtLabel": generated_at.strftime("%d %b %Y %H:%M UTC"),
            "sourceSheet": "Production Data",
        },
        "stats": {
            "rows": len(rows),
            "totalUnits": total_units,
            "totalPoints": total_points,
            "techCount": len(techs),
            "activeTechCount": len(active_techs),
        },
        "rows": rows,
    }


def refresh_dashboard_data(workbook: str | None = None, workbook_url: str | None = None, output: str | Path = "dashboard_data.json") -> Path:
    workbook_path, is_temp = resolve_workbook_source(workbook, workbook_url)
    try:
        ref_map = load_reference_map(workbook_path)
        rows = build_rows(workbook_path, ref_map)
        payload = build_payload(rows, workbook_path.name)
        output_path = Path(output).expanduser().resolve()
        output_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        return output_path
    finally:
        if is_temp and workbook_path.exists():
            workbook_path.unlink()


def main() -> None:
    args = parse_args()
    output_path = refresh_dashboard_data(
        workbook=args.workbook,
        workbook_url=args.workbook_url,
        output=args.output,
    )
    print(output_path)


if __name__ == "__main__":
    main()
