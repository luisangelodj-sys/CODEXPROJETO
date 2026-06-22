from __future__ import annotations

import csv
import shutil
from datetime import datetime
from pathlib import Path


SOURCE_CSV = Path("music_analysis/all_music_files.csv")
DEST_ROOT = Path(r"D:\Biblioteca DJ Consolidada")
REPORT_CSV = Path("music_analysis/consolidation_report.csv")


def safe_part(text: str) -> str:
    for ch in '<>:"/\\|?*':
        text = text.replace(ch, "_")
    return text.strip().rstrip(".") or "_"


def destination_for(source: Path) -> Path:
    drive = safe_part(source.drive.replace(":", "")) or "sem_drive"
    parts = [safe_part(part) for part in source.parts[1:]]
    return DEST_ROOT / "_origens" / f"Drive_{drive}" / Path(*parts)


def main() -> int:
    if not SOURCE_CSV.exists():
        raise SystemExit("Arquivo music_analysis/all_music_files.csv nao encontrado. Rode o raio-x primeiro.")

    DEST_ROOT.mkdir(parents=True, exist_ok=True)
    report_rows = []
    copied = skipped = errors = 0

    with SOURCE_CSV.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))

    for index, row in enumerate(rows, start=1):
        src = Path(row["path"])
        dest = destination_for(src)
        status = "pending"
        message = ""

        try:
            if not src.exists():
                status = "missing_source"
                errors += 1
            elif DEST_ROOT in src.parents:
                status = "already_inside_destination"
                skipped += 1
            elif dest.exists() and dest.stat().st_size == src.stat().st_size:
                status = "already_copied_same_size"
                skipped += 1
            else:
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dest)
                status = "copied"
                copied += 1
        except Exception as exc:
            status = "error"
            message = str(exc)
            errors += 1

        report_rows.append(
            {
                "index": index,
                "status": status,
                "source": str(src),
                "destination": str(dest),
                "message": message,
            }
        )

        if index % 500 == 0:
            print(f"{datetime.now().isoformat(timespec='seconds')} - processados {index}/{len(rows)}")

    REPORT_CSV.parent.mkdir(exist_ok=True)
    with REPORT_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["index", "status", "source", "destination", "message"])
        writer.writeheader()
        writer.writerows(report_rows)

    print(f"Destino: {DEST_ROOT}")
    print(f"Copiados: {copied}")
    print(f"Pulados: {skipped}")
    print(f"Erros: {errors}")
    print(f"Relatorio: {REPORT_CSV}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
