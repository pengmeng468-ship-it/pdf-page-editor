from __future__ import annotations

import tempfile
from pathlib import Path

from tests.test_pdf_ops import test_export_and_split_with_chinese_paths


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="pdf工具_") as tmp:
        test_export_and_split_with_chinese_paths(Path(tmp))
    print("SELF_CHECK_OK")


if __name__ == "__main__":
    main()
