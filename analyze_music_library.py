from __future__ import annotations

import csv
import hashlib
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path


AUDIO_EXTS = {".mp3", ".wav", ".flac", ".m4a", ".aac", ".ogg", ".wma", ".aiff", ".aif", ".opus"}
SKIP_DIRS = {
    "$recycle.bin",
    "system volume information",
    "windows",
    "program files",
    "program files (x86)",
    "appdata",
    "node_modules",
}


def norm_title(path: Path) -> str:
    stem = path.stem.lower()
    stem = re.sub(r"\s*-\s*(youtube|topic|official|audio|video|lyrics|vevo)\b.*$", "", stem)
    stem = re.sub(r"\[(.*?)\]|\((.*?)\)", " ", stem)
    stem = re.sub(r"\b(official|audio|video|lyrics|youtube|remaster(ed)?|hd|hq)\b", " ", stem)
    stem = re.sub(r"[^a-z0-9]+", " ", stem)
    return re.sub(r"\s+", " ", stem).strip()


def quick_hash(path: Path) -> str:
    h = hashlib.sha1()
    with path.open("rb") as f:
        h.update(f.read(1024 * 1024))
        if path.stat().st_size > 2 * 1024 * 1024:
            f.seek(max(0, path.stat().st_size - 1024 * 1024))
            h.update(f.read(1024 * 1024))
    return h.hexdigest()


def walk_audio(roots: list[Path]):
    for root in roots:
        if not root.exists():
            continue
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if d.lower() not in SKIP_DIRS]
            for name in filenames:
                p = Path(dirpath) / name
                if p.suffix.lower() in AUDIO_EXTS:
                    try:
                        st = p.stat()
                    except OSError:
                        continue
                    yield p, st


def main() -> int:
    roots = [Path(arg) for arg in sys.argv[1:]]
    out_dir = Path("music_analysis")
    out_dir.mkdir(exist_ok=True)
    rows = []
    ext_counter = Counter()
    folder_sizes = Counter()
    by_size = defaultdict(list)
    by_norm_name = defaultdict(list)

    for p, st in walk_audio(roots):
        ext = p.suffix.lower()
        row = {
            "path": str(p),
            "name": p.name,
            "ext": ext,
            "size": st.st_size,
            "modified": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
            "folder": str(p.parent),
            "norm_name": norm_title(p),
        }
        rows.append(row)
        ext_counter[ext] += 1
        folder_sizes[str(p.parent)] += st.st_size
        by_size[st.st_size].append(row)
        by_norm_name[row["norm_name"]].append(row)

    with (out_dir / "all_music_files.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["path", "name", "ext", "size", "modified", "folder", "norm_name"])
        writer.writeheader()
        writer.writerows(rows)

    exact_candidates = []
    for size, group in by_size.items():
        if len(group) < 2 or size < 1024 * 100:
            continue
        hashes = defaultdict(list)
        for row in group:
            try:
                hashes[quick_hash(Path(row["path"]))].append(row)
            except OSError:
                pass
        for h, hgroup in hashes.items():
            if len(hgroup) > 1:
                for row in hgroup:
                    r = dict(row)
                    r["quick_hash"] = h
                    exact_candidates.append(r)

    with (out_dir / "exact_duplicate_candidates.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["quick_hash", "path", "name", "ext", "size", "modified", "folder", "norm_name"])
        writer.writeheader()
        writer.writerows(exact_candidates)

    name_candidates = [row for group in by_norm_name.values() if len(group) > 1 for row in sorted(group, key=lambda r: (r["norm_name"], r["size"], r["path"]))]
    with (out_dir / "same_name_candidates.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["path", "name", "ext", "size", "modified", "folder", "norm_name"])
        writer.writeheader()
        writer.writerows(name_candidates)

    small = [r for r in rows if r["size"] < 100 * 1024]
    with (out_dir / "suspicious_tiny_audio_files.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["path", "name", "ext", "size", "modified", "folder", "norm_name"])
        writer.writeheader()
        writer.writerows(small)

    total_size = sum(r["size"] for r in rows)
    top_folders = folder_sizes.most_common(30)
    report = [
        "# Relatorio da biblioteca de musicas",
        "",
        f"Arquivos de audio encontrados: {len(rows):,}".replace(",", "."),
        f"Espaco usado por audio: {total_size / (1024 ** 3):.2f} GB",
        f"Possiveis duplicadas exatas: {len(exact_candidates):,} arquivos".replace(",", "."),
        f"Possiveis duplicadas por nome: {len(name_candidates):,} arquivos".replace(",", "."),
        f"Arquivos de audio muito pequenos/suspeitos: {len(small):,}".replace(",", "."),
        "",
        "## Formatos",
    ]
    for ext, count in ext_counter.most_common():
        report.append(f"- {ext}: {count}")
    report += ["", "## Pastas com mais audio"]
    for folder, size in top_folders:
        report.append(f"- {size / (1024 ** 3):.2f} GB - {folder}")
    report += [
        "",
        "## Arquivos gerados",
        "- all_music_files.csv",
        "- exact_duplicate_candidates.csv",
        "- same_name_candidates.csv",
        "- suspicious_tiny_audio_files.csv",
    ]
    (out_dir / "report.md").write_text("\n".join(report), encoding="utf-8")
    print(out_dir.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
