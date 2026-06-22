from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path


IN_FILE = Path("music_analysis/all_music_files.csv")
OUT_FILE = Path("music_analysis/duplicates_same_name_and_size.csv")
SUMMARY_FILE = Path("music_analysis/duplicate_folder_pairs.csv")


def read_rows():
    with IN_FILE.open("r", encoding="utf-8-sig", newline="") as f:
        yield from csv.DictReader(f)


def main() -> None:
    groups = defaultdict(list)
    for row in read_rows():
        key = (row["normalized_name"], row["size"], row["ext"])
        groups[key].append(row)

    dup_rows = []
    folder_pairs = defaultdict(int)
    group_id = 0
    for (_name, _size, _ext), rows in groups.items():
        if len(rows) < 2:
            continue
        folders = sorted({r["folder"] for r in rows})
        if len(folders) > 1:
            for i, a in enumerate(folders):
                for b in folders[i + 1 :]:
                    folder_pairs[(a, b)] += 1
        group_id += 1
        sorted_rows = sorted(rows, key=lambda r: (r["folder"], r["name"]))
        for idx, row in enumerate(sorted_rows, start=1):
            item = dict(row)
            item["duplicate_group"] = group_id
            item["copies_in_group"] = len(rows)
            item["suggestion"] = "KEEP_ONE_REVIEW" if idx == 1 else "REMOVE_FROM_REKORDBOX_REVIEW"
            dup_rows.append(item)

    fields = ["duplicate_group", "copies_in_group", "suggestion", "path", "name", "ext", "size", "size_mb", "modified", "folder", "normalized_name"]
    with OUT_FILE.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(dup_rows)

    with SUMMARY_FILE.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["matches", "folder_a", "folder_b"])
        writer.writeheader()
        for (a, b), count in sorted(folder_pairs.items(), key=lambda x: x[1], reverse=True):
            writer.writerow({"matches": count, "folder_a": a, "folder_b": b})

    print(f"duplicate groups: {group_id}")
    print(f"duplicate files listed: {len(dup_rows)}")


if __name__ == "__main__":
    main()
