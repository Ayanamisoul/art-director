"""Рендер настоящих превью слайдов: PPTX -> PDF (PowerPoint COM) -> PNG (fitz).

Зачем: CSS-превью в чате показывает нашу «дефолтную» стилизацию и врёт,
если презентация собрана по пользовательскому шаблону. PNG из самого файла
показывают ровно то, что увидит пользователь в PowerPoint.

Вызывается после каждой сборки презентации (MCP save_presentation и правки
в редакторе). При любой ошибке возвращает None — веб тогда рисует CSS-вариант.
"""
import subprocess
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SCRIPT = ROOT / "scripts" / "pptx2pdf.ps1"
PREVIEW_DIR = ROOT / "data" / "artifacts" / "previews"
PREVIEW_DIR.mkdir(parents=True, exist_ok=True)


def render_pptx_previews(pptx_path: Path, max_slides: int | None = None,
                         dpi: int = 72) -> int | None:
    """Сохранить превью-картинки слайдов. Возвращает число картинок или None."""
    try:
        import fitz  # pymupdf
    except ImportError:
        return None
    pptx_path = Path(pptx_path)
    if not pptx_path.exists() or not SCRIPT.exists():
        return None
    out_dir = pptx_path.parent / "previews" / pptx_path.stem
    try:
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = Path(tmp) / "out.pdf"
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPT),
                 "-in", str(pptx_path.resolve()), "-out", str(pdf_path.resolve())],
                capture_output=True, timeout=180,
            )
            if not pdf_path.exists():
                print(f"preview: pptx2pdf failed: "
                      f"{result.stderr.decode('utf-8', errors='replace')[:200]}")
                return None
            doc = fitz.open(str(pdf_path))
            out_dir.mkdir(parents=True, exist_ok=True)
            for old in out_dir.glob("p*.png"):
                old.unlink()
            count = len(doc) if max_slides is None else min(max_slides, len(doc))
            for i in range(count):
                doc[i].get_pixmap(dpi=dpi).save(str(out_dir / f"p{i + 1}.png"))
            doc.close()
            return count
    except Exception as exc:
        print(f"preview render failed: {exc}")
        return None
