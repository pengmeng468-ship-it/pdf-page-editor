from __future__ import annotations

import shutil
from pathlib import Path


def main() -> None:
    here = Path(__file__).resolve().parent
    out_dir = here.parents[1] / "outputs" / "PDF轻工具"
    out_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(here / "dist" / "PDFTool.exe", out_dir / "PDF工具.exe")
    shutil.copy2(here / "README.txt", out_dir / "README.txt")
    print(out_dir / "PDF工具.exe")


if __name__ == "__main__":
    main()
