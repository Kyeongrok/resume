from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "output"

RESUMES = {
    "en": {
        "md": PROJECT_ROOT / "CV_Kyeongrok_Kim_Eng.md",
        "stem": "CV_Kyeongrok_Kim_Eng",
    },
    "ko": {
        "md": PROJECT_ROOT / "CV_Kyeongrok_Kim_Kor.md",
        "stem": "CV_Kyeongrok_Kim_Kor",
    },
}


def find_pandoc() -> str:
    pandoc = shutil.which("pandoc")
    if pandoc:
        return pandoc
    if sys.platform == "win32":
        for base in [
            Path.home() / "AppData/Local/Microsoft/WinGet/Packages",
            Path.home() / "AppData/Local/Pandoc",
            Path("C:/Program Files/Pandoc"),
            Path.home() / "scoop/apps/pandoc/current",
        ]:
            if base.exists():
                for match in base.rglob("pandoc.exe"):
                    return str(match)
    raise RuntimeError(
        "Pandoc not found.\n"
        "  winget install JohnMacFarlane.Pandoc\n"
        "  or: https://pandoc.org/installing.html"
    )


def find_libreoffice() -> str:
    for name in ("libreoffice", "soffice"):
        path = shutil.which(name)
        if path:
            return path
    if sys.platform == "win32":
        for candidate in [
            Path("C:/Program Files/LibreOffice/program/soffice.exe"),
            Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
        ]:
            if candidate.exists():
                return str(candidate)
    raise RuntimeError(
        "LibreOffice not found.\n"
        "  sudo apt-get install libreoffice  (Linux)\n"
        "  or: https://www.libreoffice.org/download/"
    )


def set_page_layout(docx_path: Path) -> None:
    # A4: 11906 x 16838 twips  |  margins: ~20mm top/bottom, ~25mm left/right
    page_w, page_h = 11906, 16838
    margin_tb = 1134
    margin_lr = 1417

    with zipfile.ZipFile(docx_path, "r") as src:
        entries = {name: src.read(name) for name in src.namelist()}

    doc = entries["word/document.xml"].decode("utf-8")

    pg_size = f'<w:pgSz w:w="{page_w}" w:h="{page_h}"/>'
    if re.search(r"<w:pgSz[^/]*/>", doc):
        doc = re.sub(r"<w:pgSz[^/]*/>", pg_size, doc)
    else:
        doc = doc.replace("</w:sectPr>", f"{pg_size}</w:sectPr>")

    pg_margin = (
        f'<w:pgMar w:top="{margin_tb}" w:right="{margin_lr}" '
        f'w:bottom="{margin_tb}" w:left="{margin_lr}" '
        'w:header="709" w:footer="709" w:gutter="0"/>'
    )
    if re.search(r"<w:pgMar[^/]*/>", doc):
        doc = re.sub(r"<w:pgMar[^/]*/>", pg_margin, doc)
    else:
        doc = doc.replace("</w:sectPr>", f"{pg_margin}</w:sectPr>")

    entries["word/document.xml"] = doc.encode("utf-8")

    temp_path = docx_path.with_suffix(".tmp.docx")
    try:
        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as dst:
            for name, data in entries.items():
                dst.writestr(name, data)
        temp_path.replace(docx_path)
    finally:
        if temp_path.exists():
            temp_path.unlink(missing_ok=True)


def build_docx(lang: str, pandoc: str) -> Path:
    info = RESUMES[lang]
    md_path: Path = info["md"]
    out_docx = OUTPUT_DIR / f"{info['stem']}.docx"

    if not md_path.exists():
        raise FileNotFoundError(f"Markdown not found: {md_path}")

    args = [
        pandoc,
        str(md_path),
        "--from=markdown",
        "--to=docx",
        f"--output={out_docx}",
        "--columns=80",
    ]

    reference = PROJECT_ROOT / "reference.docx"
    if reference.exists():
        args.append(f"--reference-doc={reference}")

    print(f"  {out_docx.name}")
    result = subprocess.run(args, check=False)
    if result.returncode != 0:
        raise RuntimeError(f"Pandoc failed (exit {result.returncode})")

    set_page_layout(out_docx)
    return out_docx


def build_pdf(docx_path: Path, soffice: str) -> Path:
    out_pdf = docx_path.with_suffix(".pdf")
    print(f"  {out_pdf.name}")
    result = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(OUTPUT_DIR),
            str(docx_path),
        ],
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice failed (exit {result.returncode})")
    return out_pdf


def build(langs: list[str], pdf: bool) -> int:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    try:
        pandoc = find_pandoc()
        soffice = find_libreoffice() if pdf else None
    except RuntimeError as e:
        print(e, file=sys.stderr)
        return 1

    print(f"Pandoc:      {pandoc}")
    if soffice:
        print(f"LibreOffice: {soffice}")
    print()

    errors: list[tuple[str, Exception]] = []
    for lang in langs:
        print(f"[{lang.upper()}]")
        try:
            docx_path = build_docx(lang, pandoc)
            if pdf and soffice:
                build_pdf(docx_path, soffice)
        except Exception as exc:
            print(f"  ERROR: {exc}", file=sys.stderr)
            errors.append((lang, exc))

    print()
    if errors:
        print(f"Build failed ({len(errors)} error(s)).", file=sys.stderr)
        return 1

    print("Done!")
    for f in sorted(OUTPUT_DIR.glob("*")):
        print(f"  {f.name}  ({f.stat().st_size // 1024} KB)")
    return 0


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build resume md → docx + pdf")
    parser.add_argument(
        "--lang", choices=["en", "ko", "all"], default="all",
        help="Language to build (default: all)",
    )
    parser.add_argument(
        "--no-pdf", action="store_true",
        help="Skip PDF generation",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    langs = list(RESUMES.keys()) if args.lang == "all" else [args.lang]
    return build(langs=langs, pdf=not args.no_pdf)


if __name__ == "__main__":
    raise SystemExit(main())
