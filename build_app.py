"""Build Document Translator for the operating system running this script."""

from __future__ import annotations

import os
import platform
import sys
from pathlib import Path


APP_NAME = "DocumentTranslator"
PROJECT_DIR = Path(__file__).resolve().parent
ENTRY_POINT = PROJECT_DIR / "pptxtranslator.py"


def build_arguments() -> list[str]:
    system = platform.system()
    if system not in {"Windows", "Darwin"}:
        raise SystemExit("Packaging is currently configured for Windows and macOS.")

    icon = PROJECT_DIR / ("pptx_icon.ico" if system == "Windows" else "pptx_icon.png")
    if not icon.exists():
        raise SystemExit(f"Missing application icon: {icon}")

    args = [
        str(ENTRY_POINT),
        "--name", APP_NAME,
        "--noconfirm",
        "--clean",
        "--windowed",
        "--onedir",
        "--collect-all", "customtkinter",
        "--collect-all", "tkinterdnd2",
        "--add-data", f"{PROJECT_DIR / 'assets'}{os.pathsep}assets",
        "--add-data", f"{PROJECT_DIR / 'pptx_icon.png'}{os.pathsep}.",
        "--add-data", f"{PROJECT_DIR / 'pptx_icon.ico'}{os.pathsep}.",
        "--icon", str(icon),
        "--distpath", str(PROJECT_DIR / "dist"),
        "--workpath", str(PROJECT_DIR / "build"),
        "--specpath", str(PROJECT_DIR),
    ]

    if system == "Darwin":
        args.extend(["--osx-bundle-identifier", "com.documenttranslator.desktop"])

    return args


def main() -> None:
    try:
        import PyInstaller.__main__
    except ImportError as exc:
        raise SystemExit(
            "PyInstaller is not installed. Run: "
            f"{sys.executable} -m pip install -r requirements-build.txt"
        ) from exc

    print(f"Building {APP_NAME} for {platform.system()} ({platform.machine()})...")
    PyInstaller.__main__.run(build_arguments())


if __name__ == "__main__":
    main()
