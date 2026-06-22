from __future__ import annotations

import csv
import hashlib
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path


PATH_LIST = Path("music_analysis_paths.txt")
OUT_DIR = Path("music_analysis")


def norm_name(path: Path) -> str:
    stem = path.stem.lower()
    stem = re.sub(r"\s*-\s*(youtube|topic|official|audio|video|lyrics|vevo)\b.*$", "", stem)
    stem = re.sub(r"\[(.*?)\]|\((.*?)\)", " ", stem)
    stem = re.sub(r"\b(official|audio|video|lyrics|youtube|remaster(ed)?|hd|hq)\b", " ", stem)
    stem = re.sub(r"[^a-z0-9]+", " ", stem)
    return re.sub(r"\s+", " ", stem).strip()


def file_hash(path: Path) -> str:
    h = hashlib.sha1()
    with path.open("rb") as f:
        h.update(f.read(1024 * 1024))
        size = path.stat().st_size
        if size > 3 * 1024 * 1024:
            f.seek(size // 2)
            h.update(f.read(1024 * 1024))
            f.seek(max(0, size - 1024 * 1024))
            h.update(f.read(1024 * 1024))
    return h.hexdigest()


def main() -> None:
    skip_hash = "--skip-hash" in sys.argv
    OUT_DIR.mkdir(exist_ok=True)
    rows = []
    ext_counter = Counter()
    folder_sizes = Counter()
    by_size = defaultdict(list)
    by_norm = defaultdict(list)
    missing_or_error = []

    for raw in PATH_LIST.read_text(encoding="utf-8", errors="ignore").splitlines():
        raw = raw.strip()
        if not raw:
            continue
        p = Path(raw)
        try:
            st = p.stat()
        except OSError as exc:
            missing_or_error.append({"path": raw, "error": str(exc)})
            continue
        row = {
            "path": str(p),
            "name": p.name,
            "ext": p.suffix.lower(),
            "size": st.st_size,
            "size_mb": round(st.st_size / (1024 * 1024), 2),
            "modified": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
            "folder": str(p.parent),
            "normalized_name": norm_name(p),
        }
        rows.append(row)
        ext_counter[row["ext"]] += 1
        folder_sizes[row["folder"]] += st.st_size
        by_size[st.st_size].append(row)
        by_norm[row["normalized_name"]].append(row)

    fields = ["path", "name", "ext", "size", "size_mb", "modified", "folder", "normalized_name"]
    with (OUT_DIR / "all_music_files.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)

    exact = []
    if not skip_hash:
        candidate_groups = [g for g in by_size.values() if len(g) > 1 and g[0]["size"] > 100 * 1024]
        for group in candidate_groups:
            hashes = defaultdict(list)
            for row in group:
                try:
                    hashes[file_hash(Path(row["path"]))].append(row)
                except OSError:
                    continue
            for digest, hgroup in hashes.items():
                if len(hgroup) > 1:
                    for row in hgroup:
                        item = dict(row)
                        item["sha1"] = digest
                        item["duplicate_group_count"] = len(hgroup)
                        exact.append(item)

    exact_fields = ["sha1", "duplicate_group_count"] + fields
    with (OUT_DIR / "exact_duplicates_same_audio.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=exact_fields)
        writer.writeheader()
        writer.writerows(sorted(exact, key=lambda r: (r["sha1"], r["path"])))

    same_name = []
    for name, group in by_norm.items():
        if name and len(group) > 1:
            for row in group:
                item = dict(row)
                item["name_group_count"] = len(group)
                same_name.append(item)
    same_fields = ["name_group_count"] + fields
    with (OUT_DIR / "possible_duplicates_same_name.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=same_fields)
        writer.writeheader()
        writer.writerows(sorted(same_name, key=lambda r: (r["normalized_name"], r["size"], r["path"])))

    tiny = [r for r in rows if r["size"] < 100 * 1024]
    with (OUT_DIR / "suspicious_tiny_audio_files.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(sorted(tiny, key=lambda r: r["size"]))

    with (OUT_DIR / "scan_errors.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["path", "error"])
        writer.writeheader()
        writer.writerows(missing_or_error)

    total_size = sum(r["size"] for r in rows)
    exact_waste = sum(r["size"] for r in exact) - sum(max([r["size"] for r in g], default=0) for g in _groups_by(exact, "sha1"))
    lines = [
        "# Raio-x das musicas do PC",
        "",
        f"Arquivos de audio encontrados: {len(rows)}",
        f"Espaco total em audio: {total_size / (1024 ** 3):.2f} GB",
        f"Duplicatas exatas encontradas: {len(exact)} arquivos em {len(_groups_by(exact, 'sha1'))} grupos",
        f"Espaco potencial recuperavel mantendo 1 por duplicata exata: {exact_waste / (1024 ** 3):.2f} GB",
        f"Possiveis duplicatas por nome: {len(same_name)} arquivos",
        f"Arquivos muito pequenos/suspeitos: {len(tiny)}",
        "",
        "## Formatos",
    ]
    for ext, count in ext_counter.most_common():
        lines.append(f"- {ext}: {count}")
    lines.append("")
    lines.append("## Pastas com mais audio")
    for folder, size in folder_sizes.most_common(40):
        lines.append(f"- {size / (1024 ** 3):.2f} GB - {folder}")
    lines += [
        "",
        "## Arquivos de analise",
        "- all_music_files.csv: lista completa",
        "- exact_duplicates_same_audio.csv: duplicatas confirmadas por hash",
        "- possible_duplicates_same_name.csv: possiveis duplicatas por nome",
        "- suspicious_tiny_audio_files.csv: audios muito pequenos",
        "- scan_errors.csv: arquivos que nao puderam ser lidos",
    ]
    (OUT_DIR / "report.md").write_text("\n".join(lines), encoding="utf-8")
    print(OUT_DIR.resolve())


def _groups_by(rows, key):
    groups = defaultdict(list)
    for row in rows:
        groups[row[key]].append(row)
    return groups


if __name__ == "__main__":
    main()
