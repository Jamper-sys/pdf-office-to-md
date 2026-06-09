"""
pdf_office_to_md.py
Convierte PDF y archivos de Office (.docx, .xlsx, .pptx) — y varios más
soportados por MarkItDown — a Markdown (.md).

Requisitos (una sola vez):
    pip install "markitdown[all]"

Por defecto, los .md se guardan en la carpeta Descargas del usuario
(C:\\Users\\<usuario>\\Downloads). Puedes cambiar el destino con -o,
o forzar que se guarden junto al original con --beside.

Uso:
    # Un archivo (genera contrato.md en Descargas)
    python pdf_office_to_md.py "C:\\ruta\\contrato.pdf"

    # Varios archivos
    python pdf_office_to_md.py archivo1.pdf archivo2.docx informe.xlsx

    # Toda una carpeta (no recursivo)
    python pdf_office_to_md.py "C:\\ruta\\carpeta"

    # Toda una carpeta de forma recursiva
    python pdf_office_to_md.py "C:\\ruta\\carpeta" --recursive

    # Mandar los .md a otra carpeta
    python pdf_office_to_md.py archivo.pdf -o "C:\\salida"

    # Guardar el .md junto al archivo original
    python pdf_office_to_md.py archivo.pdf --beside
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    from markitdown import MarkItDown
except ImportError:
    sys.stderr.write(
        'Falta MarkItDown. Instálalo con:\n    pip install "markitdown[all]"\n'
    )
    sys.exit(1)


SUPPORTED_EXTS = {
    ".pdf",
    ".docx", ".doc",
    ".xlsx", ".xls", ".csv",
    ".pptx", ".ppt",
    ".html", ".htm",
    ".txt", ".rtf",
    ".epub",
    ".png", ".jpg", ".jpeg",
    ".mp3", ".wav", ".m4a",
    ".json", ".xml",
    ".zip",
}


def collect_files(inputs: list[Path], recursive: bool) -> list[Path]:
    files: list[Path] = []
    for p in inputs:
        if p.is_file():
            files.append(p)
        elif p.is_dir():
            it = p.rglob("*") if recursive else p.iterdir()
            for child in it:
                if child.is_file() and child.suffix.lower() in SUPPORTED_EXTS:
                    files.append(child)
        else:
            sys.stderr.write(f"[aviso] No existe: {p}\n")
    return files


def default_downloads_dir() -> Path:
    return Path.home() / "Downloads"


def convert_one(md: MarkItDown, src: Path, out_dir: Path | None, beside: bool) -> Path:
    if beside:
        target_dir = src.parent
    elif out_dir is not None:
        target_dir = out_dir
    else:
        target_dir = default_downloads_dir()
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / (src.stem + ".md")
    result = md.convert(str(src))
    target.write_text(result.text_content or "", encoding="utf-8")
    return target


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convierte PDF y archivos de Office a Markdown usando MarkItDown."
    )
    parser.add_argument("inputs", nargs="+", help="Archivos o carpetas a convertir.")
    parser.add_argument(
        "-o", "--output-dir",
        help="Carpeta de salida. Por defecto: la carpeta Descargas del usuario.",
    )
    parser.add_argument(
        "--beside", action="store_true",
        help="Guardar cada .md junto al archivo original (ignora -o y el default).",
    )
    parser.add_argument(
        "-r", "--recursive", action="store_true",
        help="Si se pasa una carpeta, recorrerla recursivamente.",
    )
    args = parser.parse_args()

    inputs = [Path(p).expanduser() for p in args.inputs]
    out_dir = Path(args.output_dir).expanduser() if args.output_dir else None

    files = collect_files(inputs, args.recursive)
    if not files:
        sys.stderr.write("No se encontraron archivos soportados para convertir.\n")
        return 1

    md = MarkItDown()
    ok = 0
    fail = 0
    for src in files:
        try:
            target = convert_one(md, src, out_dir, args.beside)
            print(f"[ok]   {src.name}  ->  {target}")
            ok += 1
        except Exception as e:
            print(f"[err]  {src.name}: {e}", file=sys.stderr)
            fail += 1

    print(f"\nTerminado. Convertidos: {ok}. Fallidos: {fail}.")
    return 0 if fail == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
