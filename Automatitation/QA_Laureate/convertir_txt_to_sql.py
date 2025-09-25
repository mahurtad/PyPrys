# rename_txt_to_sql.py
import os
from pathlib import Path

# <<< EDITA AQUÍ SI QUIERES OTRA CARPETA >>>
TARGET_DIR = Path(
    r"C:\Users\manue\Downloads\New folder\New folder\change_request_a876b895fb676290aaf4f6d87befdc7c_attachments (1)\scripts_ins_retencion_deuda_22082025"
)

def safe_rename(src: Path, dst: Path) -> Path:
    """
    Renombra sin sobrescribir: si dst existe, genera nombre único.
    """
    if not dst.exists():
        src.rename(dst)
        return dst

    base = dst.with_suffix("")  # sin extensión
    ext = dst.suffix            # ".sql"
    i = 1
    candidate = Path(f"{base} ({i}){ext}")
    while candidate.exists():
        i += 1
        candidate = Path(f"{base} ({i}){ext}")
    src.rename(candidate)
    return candidate

def main():
    if not TARGET_DIR.exists():
        print(f"Carpeta no encontrada:\n{TARGET_DIR}")
        return

    total, renamed, skipped = 0, 0, 0

    for entry in TARGET_DIR.iterdir():  # solo en la carpeta (no recursivo)
        if entry.is_file():
            total += 1
            if entry.suffix.lower() == ".txt":
                new_path = entry.with_suffix(".sql")
                final_path = safe_rename(entry, new_path)
                print(f"[OK] {entry.name}  ->  {final_path.name}")
                renamed += 1
            else:
                skipped += 1

    print("\n--- Resumen ---")
    print(f"Total de archivos: {total}")
    print(f"Renombrados (.txt → .sql): {renamed}")
    print(f"Omitidos (no .txt): {skipped}")

if __name__ == "__main__":
    main()
