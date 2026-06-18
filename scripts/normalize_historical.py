"""
Normalize historical CV folder:
  1. Flatten the duplicate nested '06 - Versiones Históricas (2010-2022)/' inside our 06_versiones_historicas_2010_2022/.
  2. Delete 0-byte placeholders.
  3. Rename subdirs to snake_case.
  4. Rename files to ASCII-safe snake_case.
  5. Report what was kept vs removed.
"""
from __future__ import annotations
import re
import shutil
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HISTORIC = ROOT / "data" / "raw" / "06_versiones_historicas_2010_2022"


def ascii_slug(name: str) -> str:
    stem = Path(name).stem
    ext = Path(name).suffix.lower()
    # normalize unicode
    stem = unicodedata.normalize("NFKD", stem)
    stem = "".join(c for c in stem if not unicodedata.combining(c))
    # replace problem chars
    stem = stem.replace("[", "").replace("]", "")
    stem = re.sub(r"[\s\-\(\)\.]+", "_", stem)
    stem = re.sub(r"[^a-zA-Z0-9_]", "", stem)
    stem = re.sub(r"_+", "_", stem).strip("_")
    return f"{stem.lower()}{ext}"


def main():
    print(f"📂 Normalizing: {HISTORIC}")

    # Step 1: Flatten duplicate nested folder if exists
    nested = list(HISTORIC.glob("06 *"))
    if nested and nested[0].is_dir():
        nested_dir = nested[0]
        print(f"📤 Flattening: moving contents of '{nested_dir.name}/' up one level")
        for item in nested_dir.iterdir():
            target = HISTORIC / item.name
            shutil.move(str(item), str(target))
        nested_dir.rmdir()

    # Step 2: Delete 0-byte files
    deleted_zero = 0
    for f in HISTORIC.rglob("*"):
        if f.is_file() and f.stat().st_size == 0:
            print(f"  🗑 0-byte: {f.relative_to(HISTORIC)}")
            f.unlink()
            deleted_zero += 1

    # Step 3 + 4: Rename dirs and files to snake_case
    # Rename dirs first (depth-first, bottom-up to avoid path issues)
    for d in sorted([p for p in HISTORIC.rglob("*") if p.is_dir()],
                    key=lambda p: -len(str(p))):
        new_name = ascii_slug(d.name).rstrip("_")
        if d.name != new_name:
            new_path = d.parent / new_name
            print(f"📁 rename dir: {d.name}/ → {new_name}/")
            d.rename(new_path)

    # Then rename files
    renamed = 0
    for f in sorted(HISTORIC.rglob("*")):
        if f.is_file():
            new_name = ascii_slug(f.name)
            if f.name != new_name:
                new_path = f.parent / new_name
                # avoid collision
                suffix = 1
                while new_path.exists():
                    base = new_path.stem
                    ext = new_path.suffix
                    new_path = f.parent / f"{base}_{suffix}{ext}"
                    suffix += 1
                f.rename(new_path)
                renamed += 1

    # Final report
    print(f"\n✅ Normalization complete:")
    print(f"   Deleted {deleted_zero} 0-byte placeholders")
    print(f"   Renamed {renamed} files to ASCII snake_case")

    remaining = [f for f in HISTORIC.rglob("*") if f.is_file()]
    print(f"\n📊 Final inventory ({len(remaining)} files):")
    for f in sorted(remaining):
        rel = f.relative_to(HISTORIC)
        size_kb = f.stat().st_size / 1024
        print(f"  {size_kb:>7.1f} KB  {rel}")


if __name__ == "__main__":
    main()
