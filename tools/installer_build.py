from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import textwrap
import urllib.request
import zipfile
from pathlib import Path


APP_ID = "{{8D87D822-A7AF-4A2D-93EF-33E724C05F38}"
APP_NAME = "期刊论文引用通知自动化管理系统"
APP_VERSION = "1.1.1"
APP_PUBLISHER = "Citation Notifier"
PYTHON_VERSION = "3.10.11"
PYTHON_EMBED_URL = (
    "https://www.python.org/ftp/python/"
    f"{PYTHON_VERSION}/python-{PYTHON_VERSION}-embed-amd64.zip"
)

PROJECT_FILES = [
    ".gitignore",
    "CHANGELOG.md",
    "LICENSE",
    "README.md",
    "app.py",
    "config.example.yaml",
    "db_lookup.py",
    "extract_references.py",
    "lookup_authors.py",
    "main.py",
    "requirements.txt",
    "send_emails.py",
    "templates",
    "启动系统.bat",
]

SKIP_PARTS = {
    ".git",
    ".pytest_cache",
    "__pycache__",
    "build",
    "dist",
    "data",
    "数据",
    "整理归档",
    "软件著作权申请资料",
    "邮箱检索",
}

COMMON_ISCC_DIRS = [
    Path(r"C:\Program Files (x86)\Inno Setup 6"),
    Path(r"C:\Program Files\Inno Setup 6"),
    Path(r"C:\Program Files (x86)\Inno Setup 5"),
    Path(r"C:\Program Files\Inno Setup 5"),
    Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Inno Setup 6",
    Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Inno Setup 5",
]


def should_skip_source_path(path: Path) -> bool:
    parts = {part.lower() for part in path.parts}
    return bool(parts.intersection({part.lower() for part in SKIP_PARTS}))


def find_iscc(extra_dirs: list[Path] | None = None) -> Path | None:
    found = shutil.which("ISCC.exe") or shutil.which("ISCC")
    if found:
        return Path(found)

    for directory in [*(extra_dirs or []), *COMMON_ISCC_DIRS]:
        candidate = directory / "ISCC.exe"
        if candidate.exists():
            return candidate
    return None


def make_installed_launcher() -> str:
    return textwrap.dedent(
        r"""
        @echo off
        setlocal

        set "APP_DIR=%~dp0"
        set "CITATION_NOTIFIER_HOME=%LOCALAPPDATA%\CitationNotifier"
        set "PATH=%APP_DIR%runtime\python;%APP_DIR%runtime\python\Scripts;%PATH%"

        if not exist "%CITATION_NOTIFIER_HOME%" mkdir "%CITATION_NOTIFIER_HOME%"
        if not exist "%CITATION_NOTIFIER_HOME%\data" mkdir "%CITATION_NOTIFIER_HOME%\data"
        if not exist "%CITATION_NOTIFIER_HOME%\data\input" mkdir "%CITATION_NOTIFIER_HOME%\data\input"
        if not exist "%CITATION_NOTIFIER_HOME%\config.yaml" (
            copy "%APP_DIR%app\config.example.yaml" "%CITATION_NOTIFIER_HOME%\config.yaml" >nul
        )

        cd /d "%APP_DIR%app"
        title Citation Notifier
        echo Starting Citation Notifier...
        echo URL: http://127.0.0.1:5000
        echo User data: %CITATION_NOTIFIER_HOME%
        "%APP_DIR%runtime\python\python.exe" main.py web
        pause
        """
    ).strip() + "\n"


def make_inno_script() -> str:
    return textwrap.dedent(
        rf"""
        #define MyAppName "{APP_NAME}"
        #define MyAppVersion "{APP_VERSION}"
        #define MyAppPublisher "{APP_PUBLISHER}"
        #define MyAppExeName "启动 Citation Notifier.bat"

        [Setup]
        AppId={APP_ID}
        AppName={{#MyAppName}}
        AppVersion={{#MyAppVersion}}
        AppPublisher={{#MyAppPublisher}}
        DefaultDirName={{autopf}}\CitationNotifier
        DefaultGroupName={{#MyAppName}}
        DisableProgramGroupPage=yes
        OutputDir=dist\installer
        OutputBaseFilename=CitationNotifier_Setup_v{{#MyAppVersion}}
        Compression=lzma2/ultra64
        SolidCompression=yes
        WizardStyle=modern
        ArchitecturesAllowed=x64compatible
        ArchitecturesInstallIn64BitMode=x64compatible
        UninstallDisplayName={{#MyAppName}}

        [Tasks]
        Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加图标："; Flags: unchecked

        [Files]
        Source: "build\installer\payload\*"; DestDir: "{{app}}"; Flags: ignoreversion recursesubdirs createallsubdirs

        [Icons]
        Name: "{{group}}\{{#MyAppName}}"; Filename: "{{app}}\{{#MyAppExeName}}"; WorkingDir: "{{app}}"
        Name: "{{autodesktop}}\{{#MyAppName}}"; Filename: "{{app}}\{{#MyAppExeName}}"; WorkingDir: "{{app}}"; Tasks: desktopicon

        [Run]
        Filename: "{{app}}\{{#MyAppExeName}}"; Description: "启动 {{#MyAppName}}"; Flags: postinstall shellexec skipifsilent nowait

        [UninstallDelete]
        Type: filesandordirs; Name: "{{app}}\runtime"
        Type: filesandordirs; Name: "{{app}}\app"
        """
    ).strip() + "\n"


def clean_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)


def copy_project_files(project_root: Path, app_dir: Path) -> None:
    app_dir.mkdir(parents=True, exist_ok=True)
    for name in PROJECT_FILES:
        source = project_root / name
        if not source.exists():
            continue
        if should_skip_source_path(Path(name)):
            continue
        destination = app_dir / name
        if source.is_dir():
            shutil.copytree(
                source,
                destination,
                ignore=lambda _dir, names: [
                    item for item in names if should_skip_source_path(Path(item))
                ],
            )
        else:
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)


def download_file(url: str, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists() and destination.stat().st_size > 0:
        return
    urllib.request.urlretrieve(url, destination)


def enable_embed_site(runtime_dir: Path) -> None:
    pth_files = sorted(runtime_dir.glob("python*._pth"))
    if not pth_files:
        return
    pth = pth_files[0]
    lines = pth.read_text(encoding="utf-8").splitlines()
    normalized = ["import site" if line.strip() == "#import site" else line for line in lines]
    pth.write_text("\n".join(normalized) + "\n", encoding="utf-8")


def build_python_runtime(project_root: Path, runtime_dir: Path, cache_dir: Path) -> None:
    clean_dir(runtime_dir)
    embed_zip = cache_dir / f"python-{PYTHON_VERSION}-embed-amd64.zip"
    download_file(PYTHON_EMBED_URL, embed_zip)
    with zipfile.ZipFile(embed_zip) as zf:
        zf.extractall(runtime_dir)
    enable_embed_site(runtime_dir)

    get_pip = cache_dir / "get-pip.py"
    download_file("https://bootstrap.pypa.io/get-pip.py", get_pip)
    python_exe = runtime_dir / "python.exe"
    subprocess.run([str(python_exe), str(get_pip)], check=True)
    subprocess.run(
        [
            str(python_exe),
            "-m",
            "pip",
            "install",
            "--no-warn-script-location",
            "-r",
            str(project_root / "requirements.txt"),
        ],
        check=True,
    )


def write_installer_files(project_root: Path) -> Path:
    payload_dir = project_root / "build" / "installer" / "payload"
    clean_dir(payload_dir)
    copy_project_files(project_root, payload_dir / "app")
    (payload_dir / "启动 Citation Notifier.bat").write_text(
        make_installed_launcher(), encoding="utf-8"
    )
    iss_path = project_root / "installer.iss"
    iss_path.write_text(make_inno_script(), encoding="utf-8")
    return payload_dir


def build(args: argparse.Namespace) -> int:
    project_root = Path(args.project_root).resolve()
    payload_dir = write_installer_files(project_root)
    if not args.skip_runtime:
        build_python_runtime(
            project_root,
            payload_dir / "runtime" / "python",
            project_root / "build" / "installer" / "cache",
        )

    iscc = find_iscc()
    if not iscc:
        print("Inno Setup compiler ISCC.exe was not found.")
        print("Install Inno Setup 6, then rerun: python tools/installer_build.py")
        print(f"Prepared payload: {payload_dir}")
        return 2

    subprocess.run([str(iscc), str(project_root / "installer.iss")], cwd=project_root, check=True)
    print(f"Installer output: {project_root / 'dist' / 'installer'}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build the Windows installer.")
    parser.add_argument("--project-root", default=Path(__file__).resolve().parents[1])
    parser.add_argument(
        "--skip-runtime",
        action="store_true",
        help="Only prepare payload and installer.iss; do not download/build Python runtime.",
    )
    args = parser.parse_args(argv)
    return build(args)


if __name__ == "__main__":
    raise SystemExit(main())
