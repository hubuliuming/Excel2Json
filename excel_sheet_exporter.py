#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import keyword
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

try:
    from openpyxl import load_workbook
except ImportError as exc:  # pragma: no cover
    print("Missing dependency: openpyxl. Install with: pip install openpyxl", file=sys.stderr)
    raise

INVALID_FILE_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1F]')
INVALID_IDENTIFIER_CHARS = re.compile(r'[^0-9a-zA-Z_]')

SUPPORTED_TYPES = {
    "int": "int",
    "long": "long",
    "float": "float",
    "double": "double",
    "bool": "bool",
    "string": "string",
    "json": "string",
}

CLASS_NAME_KEYWORD = "classname"


@dataclass
class ColumnDef:
    source_name: str
    field_name: str
    cs_type: str
    raw_type: str
    index: int


@dataclass
class SheetExportResult:
    workbook_name: str
    sheet_name: str
    export_name: str
    json_path: Path
    cs_path: Path | None
    row_count: int


class ExportError(Exception):
    pass



def sanitize_filename(name: str) -> str:
    value = INVALID_FILE_CHARS.sub("_", str(name or "")).strip().rstrip(".")
    return value or "Unnamed"



def sanitize_identifier(name: str, pascal: bool = False) -> str:
    text = str(name or "Field").strip()
    if not text:
        text = "Field"
    text = text.replace(" ", "_").replace("-", "_")
    text = INVALID_IDENTIFIER_CHARS.sub("_", text)
    text = re.sub(r"_+", "_", text).strip("_") or "Field"
    if text and text[0].isdigit():
        text = f"_{text}"
    if keyword.iskeyword(text):
        text += "_"
    if pascal:
        parts = [p for p in re.split(r"_+", text) if p]
        text = "".join(p[:1].upper() + p[1:] for p in parts) or "Field"
        if text[0].isdigit():
            text = f"_{text}"
    return text



def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")



def normalize_cell_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()



def parse_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "y", "on"}:
        return True
    if text in {"0", "false", "no", "n", "off", ""}:
        return False
    raise ValueError(f"Cannot parse bool from '{value}'")



def convert_scalar(value: Any, raw_type: str) -> Any:
    if is_blank(value):
        return None
    base_type = raw_type.lower()
    if base_type == "string":
        return str(value)
    if base_type == "json":
        return json.dumps(value, ensure_ascii=False) if not isinstance(value, str) else value
    if base_type == "int":
        return int(value)
    if base_type == "long":
        return int(value)
    if base_type == "float":
        return float(value)
    if base_type == "double":
        return float(value)
    if base_type == "bool":
        return parse_bool(value)
    return value



def split_array_text(text: str, delimiter: str) -> list[str]:
    return [item.strip() for item in text.split(delimiter)] if text else []



def convert_value(value: Any, raw_type: str, array_delimiter: str) -> Any:
    raw_type = (raw_type or "string").strip().lower()
    if raw_type.endswith("[]"):
        element_type = raw_type[:-2]
        if is_blank(value):
            return []
        if isinstance(value, str):
            items = split_array_text(value, array_delimiter)
        elif isinstance(value, Iterable) and not isinstance(value, (str, bytes, dict)):
            items = list(value)
        else:
            items = [value]
        return [convert_scalar(item, element_type) for item in items if not is_blank(item)]
    return convert_scalar(value, raw_type)



def infer_type(values: list[Any]) -> str:
    samples = [v for v in values if not is_blank(v)]
    if not samples:
        return "string"

    def is_bool_like(v: Any) -> bool:
        try:
            parse_bool(v)
            return True
        except Exception:
            return False

    if all(is_bool_like(v) for v in samples):
        return "bool"

    def is_int_like(v: Any) -> bool:
        try:
            f = float(v)
            return f.is_integer()
        except Exception:
            return False

    if all(is_int_like(v) for v in samples):
        return "int"

    def is_float_like(v: Any) -> bool:
        try:
            float(v)
            return True
        except Exception:
            return False

    if all(is_float_like(v) for v in samples):
        return "float"

    return "string"



def find_class_name_column(ws, header_row: int, type_row: int) -> tuple[int | None, str | None]:
    header_values = [normalize_cell_text(cell.value) for cell in ws[header_row]]
    for idx, value in enumerate(header_values):
        if CLASS_NAME_KEYWORD in value.lower():
            output_name = normalize_cell_text(ws.cell(row=type_row, column=idx + 1).value)
            if not output_name:
                raise ExportError(
                    f"Sheet '{ws.title}' has a '{value}' marker in row {type_row}, but row {type_row + 1} is empty in that column."
                )
            return idx, output_name
    return None, None



def build_columns(ws, type_row: int, header_row: int, data_start_row: int, sample_rows: int) -> tuple[list[ColumnDef], str] | None:
    class_name_col, export_name = find_class_name_column(ws, header_row, type_row)
    if class_name_col is None:
        return None

    headers = [normalize_cell_text(cell.value) for cell in ws[header_row]]
    if not any(headers):
        raise ExportError(f"Sheet '{ws.title}' header row {header_row} is empty")

    type_values = [normalize_cell_text(cell.value).lower() for cell in ws[type_row]]
    value_samples_by_col: list[list[Any]] = [[] for _ in headers]
    max_sample_row = min(ws.max_row, data_start_row + max(sample_rows, 1) - 1)
    for row_idx in range(data_start_row, max_sample_row + 1):
        row = ws[row_idx]
        for col_idx, cell in enumerate(row):
            if col_idx < len(value_samples_by_col):
                value_samples_by_col[col_idx].append(cell.value)

    columns: list[ColumnDef] = []
    seen_names: set[str] = set()
    for idx, header in enumerate(headers):
        if idx == class_name_col:
            continue
        if not header:
            continue

        field_name = sanitize_identifier(header)
        original_field_name = field_name
        suffix = 2
        while field_name.lower() in seen_names:
            field_name = f"{original_field_name}_{suffix}"
            suffix += 1
        seen_names.add(field_name.lower())

        raw_type = type_values[idx] if idx < len(type_values) else ""
        if CLASS_NAME_KEYWORD in raw_type:
            continue
        if raw_type and raw_type not in SUPPORTED_TYPES and not raw_type.endswith("[]"):
            raise ExportError(
                f"Sheet '{ws.title}' column '{header}' has unsupported type '{raw_type}'. "
                "Supported: int, long, float, double, bool, string, json, and [] arrays."
            )
        if not raw_type:
            raw_type = infer_type(value_samples_by_col[idx])
        base_cs_type = SUPPORTED_TYPES.get(raw_type[:-2], raw_type[:-2]) if raw_type.endswith("[]") else SUPPORTED_TYPES.get(raw_type, raw_type)
        cs_type = f"List<{base_cs_type}>" if raw_type.endswith("[]") else base_cs_type
        columns.append(ColumnDef(header, field_name, cs_type, raw_type, idx))

    if not columns:
        raise ExportError(f"Sheet '{ws.title}' has no valid export columns")
    return columns, export_name



def row_is_effectively_empty(row_values: list[Any], columns: list[ColumnDef]) -> bool:
    return all(is_blank(row_values[col.index] if col.index < len(row_values) else None) for col in columns)



def parse_sheet_rows(ws, columns: list[ColumnDef], data_start_row: int, array_delimiter: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for row in ws.iter_rows(min_row=data_start_row, values_only=True):
        values = list(row)
        if row_is_effectively_empty(values, columns):
            continue
        item: dict[str, Any] = {}
        for col in columns:
            value = values[col.index] if col.index < len(values) else None
            item[col.field_name] = convert_value(value, col.raw_type, array_delimiter)
        rows.append(item)
    return rows



def generate_cs_source(namespace: str, class_name: str, columns: list[ColumnDef]) -> str:
    using_lines = ["using System;"]
    if any(col.cs_type.startswith("List<") for col in columns):
        using_lines.append("using System.Collections.Generic;")
    using_block = "\n".join(using_lines)

    field_lines = [f"        public {col.cs_type} {sanitize_identifier(col.field_name, pascal=True)};" for col in columns]
    fields = "\n".join(field_lines)
    return (
        f"{using_block}\n\n"
        f"namespace {namespace}\n"
        "{\n"
        "    [Serializable]\n"
        f"    public class {class_name}\n"
        "    {\n"
        f"{fields}\n"
        "    }\n"
        "}\n"
    )



def export_workbook(
    excel_path: Path,
    json_folder: Path,
    cs_folder: Path | None,
    type_row: int,
    header_row: int,
    data_start_row: int,
    namespace: str,
    array_delimiter: str,
    sample_rows: int,
    include_hidden_sheets: bool,
) -> list[SheetExportResult]:
    wb = load_workbook(excel_path, data_only=True)
    results: list[SheetExportResult] = []
    workbook_name = excel_path.stem
    try:
        for ws in wb.worksheets:
            if not include_hidden_sheets and ws.sheet_state != "visible":
                continue
            if ws.title.startswith("#"):
                continue

            built = build_columns(ws, type_row, header_row, data_start_row, sample_rows)
            if built is None:
                continue
            columns, export_name_raw = built
            export_name = sanitize_filename(export_name_raw)

            rows = parse_sheet_rows(ws, columns, data_start_row, array_delimiter)
            json_path = json_folder / f"{export_name}.json"
            json_path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")

            cs_path: Path | None = None
            if cs_folder is not None:
                class_name = sanitize_identifier(export_name_raw, pascal=True)
                cs_source = generate_cs_source(namespace, class_name, columns)
                cs_path = cs_folder / f"{export_name}.cs"
                cs_path.write_text(cs_source, encoding="utf-8")

            results.append(SheetExportResult(workbook_name, ws.title, export_name, json_path, cs_path, len(rows)))
    finally:
        wb.close()
    return results



def copy_json_outputs(source_folder: Path, target_folder: Path) -> None:
    target_folder.mkdir(parents=True, exist_ok=True)
    for src in source_folder.glob("*.json"):
        dst = target_folder / src.name
        dst.write_bytes(src.read_bytes())



def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export Excel sheets that contain a ClassName marker into JSON/C# files.")
    parser.add_argument("--excel-folder", required=True, help="Folder containing .xlsx files")
    parser.add_argument("--json-folder", required=True, help="Output folder for JSON")
    parser.add_argument("--cs-folder", help="Output folder for C#; omit to skip C# generation")
    parser.add_argument("--copy-json-to", help="Optional folder to copy JSON outputs into")
    parser.add_argument("--type-row", type=int, default=1, help="Type row index (1-based). Supports legacy layout. Default: 1")
    parser.add_argument("--header-row", type=int, default=2, help="Field-name row index (1-based). Supports legacy layout. Default: 2")
    parser.add_argument("--data-start-row", type=int, default=3, help="First data row (1-based). Default: 3")
    parser.add_argument("--namespace", default="Game.Config", help="C# namespace")
    parser.add_argument("--array-delimiter", default="|", help="Delimiter for [] array columns. Default: |")
    parser.add_argument("--sample-rows", type=int, default=30, help="Rows used for type inference when type row is empty")
    parser.add_argument("--include-hidden-sheets", action="store_true", help="Also export hidden sheets")
    return parser.parse_args()



def validate_args(args: argparse.Namespace) -> None:
    if args.type_row <= 0:
        raise ExportError("--type-row must be >= 1")
    if args.header_row == args.type_row:
        raise ExportError("--header-row and --type-row cannot be the same")
    if args.data_start_row <= max(args.header_row, args.type_row):
        raise ExportError("--data-start-row must be greater than both --header-row and --type-row")



def main() -> int:
    args = parse_args()
    try:
        validate_args(args)
        excel_folder = Path(args.excel_folder).resolve()
        json_folder = Path(args.json_folder).resolve()
        cs_folder = Path(args.cs_folder).resolve() if args.cs_folder else None
        copy_json_to = Path(args.copy_json_to).resolve() if args.copy_json_to else None

        if not excel_folder.exists():
            raise ExportError(f"Excel folder does not exist: {excel_folder}")

        json_folder.mkdir(parents=True, exist_ok=True)
        if cs_folder is not None:
            cs_folder.mkdir(parents=True, exist_ok=True)

        excel_files = sorted([p for p in excel_folder.rglob("*.xlsx") if not p.name.startswith("~$")])
        if not excel_files:
            raise ExportError(f"No .xlsx files found under: {excel_folder}")

        all_results: list[SheetExportResult] = []
        for excel_path in excel_files:
            workbook_results = export_workbook(
                excel_path=excel_path,
                json_folder=json_folder,
                cs_folder=cs_folder,
                type_row=args.type_row,
                header_row=args.header_row,
                data_start_row=args.data_start_row,
                namespace=args.namespace,
                array_delimiter=args.array_delimiter,
                sample_rows=args.sample_rows,
                include_hidden_sheets=args.include_hidden_sheets,
            )
            all_results.extend(workbook_results)
            for item in workbook_results:
                msg = f"[OK] {item.workbook_name} / {item.sheet_name} -> {item.json_path.name}"
                if item.cs_path:
                    msg += f" , {item.cs_path.name}"
                print(msg)

        if copy_json_to is not None:
            copy_json_outputs(json_folder, copy_json_to)
            print(f"[COPY] JSON copied to: {copy_json_to}")

        print(f"[DONE] Exported {len(all_results)} sheet(s).")
        return 0
    except ExportError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # pragma: no cover
        print(f"[ERROR] Unexpected failure: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
