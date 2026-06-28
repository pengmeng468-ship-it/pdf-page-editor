from pathlib import Path

from office_convert import unique_pdf_path


def test_unique_pdf_path_uses_source_stem_and_avoids_overwrite(tmp_path):
    folder = tmp_path / "中文 路径"
    folder.mkdir(parents=True)
    existing = folder / "测试文档.pdf"
    existing.write_bytes(b"%PDF-1.4\n")

    result = unique_pdf_path(folder, "测试文档.docx")

    assert result.parent == folder
    assert result.name != existing.name
    assert result.suffix == ".pdf"
    assert result.stem.startswith("测试文档")
