import argparse
from pathlib import Path
import shutil
import pandas as pd
from typing import Union
import unicodedata
import re

def load_filenames_from_excel(excel_path: Path, col: Union[int,str] = 0) -> set:
    
    df = pd.read_excel(excel_path, dtype=str)
    names = set()
    try:
        if isinstance(col, int):
            if col < 0 or col >= df.shape[1]:
                return names
            series = df.iloc[:, col]
        else:
            if col not in df.columns:
                return names
            series = df[col]
    except Exception:
        return names
    for v in series.dropna().astype(str).str.strip():
        if v:
            names.add(v)
    return names

def norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKC", s).strip()
    s = re.sub(r"[\u200B-\u200D\uFEFF]", "", s)
    return s.lower()

def find_files(root: Path, candidates: set):
    
    cand_norm = {norm(c): c for c in candidates}
    found = {c: [] for c in candidates}
    for p in root.rglob('*.pdf'):
        try:
            n = norm(p.name)
            r = norm(str(p.relative_to(root)))
        except Exception:
            
            continue
        for kn, orig in cand_norm.items():
            if n.startswith(kn) or r.startswith(kn) or kn in n or kn in r:
                found[orig].append(p)
    return found

def ensure_dir(path: Path):
    path.mkdir(parents=True, exist_ok=True)

def parse_col_arg(val: str):
    try:
        return int(val)
    except ValueError:
        return val

def main():
    parser = argparse.ArgumentParser(description="Move/copy PDFs listed in an Excel file (single column).")
    parser.add_argument("excel", type=Path, help="Path to Excel file (xlsx/xls).")
    parser.add_argument("source_root", type=Path, help="Directory to search recursively for PDFs.")
    parser.add_argument("dest_dir", type=Path, help="Directory where matched PDFs will be placed.")
    parser.add_argument("--copy", action="store_true", help="Copy instead of move (default: move).")
    parser.add_argument("--all-matches", action="store_true", help="If multiple matches per name, keep all instead of only latest.")
    parser.add_argument("--col", metavar='COL', help="Column index (0-based) or name to read filenames from. Defaults to 0.")
    args = parser.parse_args()

    excel = args.excel.resolve()
    root = args.source_root.resolve()
    dest = args.dest_dir.resolve()
    ensure_dir(dest)

    if args.col:
        col = parse_col_arg(args.col)
        candidates = load_filenames_from_excel(excel, col=col)
    else:
        candidates = load_filenames_from_excel(excel, col=0)

    if not candidates:
        print("No filenames found in the Excel file (in the specified column).")
        return

    matches = find_files(root, candidates)

    moved = 0
    missing = []
    processed_srcs = set()

    for name, paths in matches.items():
        if not paths:
            missing.append(name)
            continue

       
        if args.all_matches:
            targets = paths
        else:
            
            targets = [max(paths, key=lambda p: p.stat().st_mtime)]

        for src in targets:
           
            if str(src.resolve()) in processed_srcs:
                continue

            if not src.exists():
                print(f"Skipping (not found): {src}")
                continue

            dst = dest / src.name
           
            if dst.exists():
                stem = dst.stem
                suf = dst.suffix
                i = 1
                while True:
                    candidate = dest / f"{stem}_{i}{suf}"
                    if not candidate.exists():
                        dst = candidate
                        break
                    i += 1

            try:
                if args.copy:
                    shutil.copy2(src, dst)
                else:
                    shutil.move(str(src), str(dst))
            except FileNotFoundError:
                print(f"Failed to move/copy (not found during operation): {src}")
                continue
            except Exception as e:
                print(f"Failed to move/copy {src} -> {dst}: {e}")
                continue

            moved += 1
            processed_srcs.add(str(src.resolve()))
            print(f"{'Copied' if args.copy else 'Moved'}: {src} -> {dst}")

    print(f"\nDone. {moved} files processed.")

    if missing:
        missing_file = dest / "missing_files.txt"
        with missing_file.open("w", encoding="utf-8") as f:
            for name in missing:
                f.write(f"{name}\n")
        print(f"Wrote {len(missing)} missing names to: {missing_file}")

if __name__ == "__main__":
    main()