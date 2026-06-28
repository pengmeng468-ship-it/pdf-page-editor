from __future__ import annotations

import shutil
import subprocess
import tempfile
import time
from pathlib import Path


SUPPORTED_OFFICE_EXTS = {".doc", ".docx", ".ppt", ".pptx"}
WORD_EXTS = {".doc", ".docx"}
POWERPOINT_EXTS = {".ppt", ".pptx"}


class OfficeConversionError(RuntimeError):
    pass


def convert_office_to_pdf(source: str | Path, output: str | Path) -> dict:
    src = Path(source).resolve()
    dst = Path(output).resolve()
    if src.suffix.lower() not in SUPPORTED_OFFICE_EXTS:
        raise OfficeConversionError(f"不支持的 Office 文件类型：{src.suffix}")
    if not src.exists():
        raise OfficeConversionError(f"文件不存在：{src}")

    dst.parent.mkdir(parents=True, exist_ok=True)
    errors: list[str] = []

    try:
        _convert_with_office(src, dst)
        return _result(src, dst, "Microsoft Office")
    except Exception as exc:
        errors.append(f"Office 导出失败：{exc}")

    try:
        _convert_with_libreoffice(src, dst)
        return _result(src, dst, "LibreOffice")
    except Exception as exc:
        errors.append(f"LibreOffice 导出失败：{exc}")

    raise OfficeConversionError("；".join(errors))


def app_temp_dir() -> Path:
    base = Path(tempfile.gettempdir()) / "PDF Editor-lite"
    base.mkdir(parents=True, exist_ok=True)
    return base


def unique_pdf_path(folder: str | Path, source_name: str) -> Path:
    folder_path = Path(folder)
    folder_path.mkdir(parents=True, exist_ok=True)
    stem = Path(source_name).stem
    candidate = folder_path / f"{stem}.pdf"
    if not candidate.exists():
        return candidate
    stamp = time.strftime("%Y%m%d-%H%M%S")
    candidate = folder_path / f"{stem}-{stamp}.pdf"
    if not candidate.exists():
        return candidate
    for index in range(2, 1000):
        candidate = folder_path / f"{stem}-{stamp}-{index}.pdf"
        if not candidate.exists():
            return candidate
    raise OfficeConversionError(f"无法生成不重名 PDF：{source_name}")


def _convert_with_office(src: Path, dst: Path) -> None:
    suffix = src.suffix.lower()
    if suffix in WORD_EXTS:
        _convert_word_with_office(src, dst)
    elif suffix in POWERPOINT_EXTS:
        _convert_powerpoint_with_office(src, dst)
    else:
        raise OfficeConversionError(f"不支持的 Office 文件类型：{suffix}")
    _ensure_pdf(dst)


def _convert_word_with_office(src: Path, dst: Path) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    doc = None
    try:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)
        doc.ExportAsFixedFormat(str(dst), 17)
    finally:
        if doc is not None:
            doc.Close(False)
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()


def _convert_powerpoint_with_office(src: Path, dst: Path) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    deck = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        deck = app.Presentations.Open(str(src), WithWindow=False, ReadOnly=True)
        deck.SaveAs(str(dst), 32)
    finally:
        if deck is not None:
            deck.Close()
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()


def _convert_with_libreoffice(src: Path, dst: Path) -> None:
    soffice = _find_soffice()
    out_dir = dst.parent
    profile_dir = app_temp_dir() / "lo-profile"
    profile_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        str(soffice),
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        f"-env:UserInstallation=file:///{profile_dir.as_posix()}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(src),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    produced = out_dir / f"{src.stem}.pdf"
    if proc.returncode != 0:
        raise OfficeConversionError((proc.stderr or proc.stdout or "LibreOffice 返回错误").strip())
    if produced.exists() and produced.resolve() != dst.resolve():
        if dst.exists():
            dst.unlink()
        produced.replace(dst)
    _ensure_pdf(dst)


def _find_soffice() -> Path:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return Path(candidate)
    raise OfficeConversionError("未找到 LibreOffice soffice.exe")


def _ensure_pdf(path: Path) -> None:
    if not path.exists() or path.stat().st_size == 0:
        raise OfficeConversionError(f"没有生成有效 PDF：{path}")


def _result(src: Path, dst: Path, engine: str) -> dict:
    return {
        "source": str(src),
        "path": str(dst),
        "name": dst.name,
        "engine": engine,
        "bytes": dst.stat().st_size,
    }
