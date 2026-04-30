from __future__ import annotations

import argparse
from pathlib import Path

from analyzer import InputOptions, SUPPORTED_SUFFIXES, analyze_files
from report_exporter import save_reports_to_excel


def parse_manager_hints(raw: str) -> tuple[str, ...]:
    return tuple(part.strip() for part in raw.split(",") if part.strip())


def main() -> None:
    parser = argparse.ArgumentParser(description="Create one combined XLSX report from Telegram/chat exports")
    parser.add_argument("--input_dir", required=True, help="Directory with JSON/CSV/XLSX/TXT chat exports")
    parser.add_argument("--output", required=True, help="Path to one combined XLSX file")
    parser.add_argument(
        "--input_format",
        default="auto",
        choices=["auto", "telegram_json", "csv_table", "excel_table", "plain_text"],
        help="Input format. Use auto for mixed folders.",
    )
    parser.add_argument("--manager_hints", default="", help="Comma-separated manager names/keywords, e.g. Dana,Ruslan,Manager")
    parser.add_argument("--window_size", type=int, default=6, help="Messages per mood-analysis window")
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory not found: {input_dir}")

    files = sorted([p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_SUFFIXES])
    if not files:
        raise FileNotFoundError(f"No supported files found in {input_dir}")

    options = InputOptions(
        input_format=args.input_format,
        manager_hints=parse_manager_hints(args.manager_hints),
        window_size=max(1, args.window_size),
    )
    reports = analyze_files(files, options)
    output_path = save_reports_to_excel(reports, args.output)
    print(f"Done! Combined report created: {output_path.resolve()}")


if __name__ == "__main__":
    main()
