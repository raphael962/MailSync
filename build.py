#!/usr/bin/env python3
"""
Script de build — MailSync
===========================
Lance ce script sur la machine cible pour générer l'installeur correspondant.

  Mac     ->  dist/MailSync.dmg
  Windows ->  dist/MailSync Setup.exe
  Linux   ->  dist/MailSync.AppImage

Prérequis communs :
  pip install pyinstaller tqdm pillow

Prérequis supplémentaires selon l'OS :
  Mac     : brew install create-dmg
  Windows : installer NSIS depuis https://nsis.sourceforge.io
  Linux   : rien (appimagetool est téléchargé automatiquement)

Usage :
  python3 build.py
"""

import os
import sys
import shutil
import platform
import subprocess
import textwrap

APP_NAME    = "MailSync"
APP_VERSION = "1.0.0"
SCRIPT      = "mailsync.py"
BUILD_DIR   = "build_tmp"
DIST_DIR    = "dist"

OS = platform.system()

# ─── Utilitaires ─────────────────────────────────────────────────────────────

def run(cmd, **kwargs):
    print(f"\n  $ {cmd}")
    result = subprocess.run(cmd, shell=True, **kwargs)
    if result.returncode != 0:
        print(f"\n  ERREUR (code {result.returncode})")
        sys.exit(result.returncode)

def require(tool, install_hint):
    if not shutil.which(tool):
        print(f"\n  Outil manquant : '{tool}'")
        print(f"  Installation : {install_hint}")
        sys.exit(1)

def step(msg):
    print(f"\n{'='*60}\n  {msg}\n{'='*60}")

# ─── Build PyInstaller (sans icône) ──────────────────────────────────────────

def pyinstaller_build():
    step("Build PyInstaller")
    require("pyinstaller", "pip install pyinstaller")

    windowed = "--windowed" if OS in ("Darwin", "Windows") else ""
    manifest = f'--manifest "mailsync.manifest"' if OS == "Windows" and os.path.exists("mailsync.manifest") else ""

    cmd = (
        f'pyinstaller --clean {windowed} --onedir '
        f'--name "{APP_NAME}" '
        f'--distpath "{BUILD_DIR}/pyinstaller" '
        f'--workpath "{BUILD_DIR}/work" '
        f'--specpath "{BUILD_DIR}" '
        f'{manifest} '
        f'"{SCRIPT}"'
    )
    run(cmd)

# ─── Mac ──────────────────────────────────────────────────────────────────────

def build_mac():
    pyinstaller_build()

    step("Création du DMG")
    require("create-dmg", "brew install create-dmg")

    app_path = f'{BUILD_DIR}/pyinstaller/{APP_NAME}.app'
    dmg_out  = f'{DIST_DIR}/{APP_NAME}.dmg'
    os.makedirs(DIST_DIR, exist_ok=True)

    cmd = (
        f'create-dmg '
        f'--volname "{APP_NAME}" '
        f'--window-pos 200 120 '
        f'--window-size 600 400 '
        f'--icon-size 100 '
        f'--icon "{APP_NAME}.app" 175 190 '
        f'--hide-extension "{APP_NAME}.app" '
        f'--app-drop-link 425 190 '
        f'"{dmg_out}" '
        f'"{app_path}"'
    )
    run(cmd)
    print(f"\n  Installeur Mac créé : {dmg_out}")

# ─── Windows ─────────────────────────────────────────────────────────────────

def build_windows():
    pyinstaller_build()

    step("Création de l'installeur NSIS")

    nsis_candidates = [
        shutil.which("makensis"),
        r"C:\Program Files (x86)\NSIS\makensis.exe",
        r"C:\Program Files\NSIS\makensis.exe",
    ]
    makensis = next((p for p in nsis_candidates if p and os.path.exists(p)), None)
    if not makensis:
        print("\n  NSIS introuvable.")
        print("  Télécharger : https://nsis.sourceforge.io/Download")
        sys.exit(1)

    os.makedirs(DIST_DIR, exist_ok=True)
    app_dir  = os.path.abspath(f"{BUILD_DIR}\\pyinstaller\\{APP_NAME}")
    dist_dir = os.path.abspath(DIST_DIR)

    nsi_content = textwrap.dedent(f"""
        !include "MUI2.nsh"

        Name "{APP_NAME}"
        OutFile "{dist_dir}\\{APP_NAME} Setup.exe"
        InstallDir "$PROGRAMFILES64\\{APP_NAME}"
        RequestExecutionLevel admin
        SetCompressor /SOLID lzma

        !define MUI_ABORTWARNING
        !insertmacro MUI_PAGE_WELCOME
        !insertmacro MUI_PAGE_DIRECTORY
        !insertmacro MUI_PAGE_INSTFILES
        !insertmacro MUI_PAGE_FINISH
        !insertmacro MUI_UNPAGE_CONFIRM
        !insertmacro MUI_UNPAGE_INSTFILES
        !insertmacro MUI_LANGUAGE "French"

        Section "Installation" SecMain
            SetOutPath "$INSTDIR"
            File /r "{app_dir}\\*.*"
            WriteUninstaller "$INSTDIR\\Uninstall.exe"
            CreateShortcut "$DESKTOP\\{APP_NAME}.lnk" "$INSTDIR\\{APP_NAME}.exe"
            CreateDirectory "$SMPROGRAMS\\{APP_NAME}"
            CreateShortcut "$SMPROGRAMS\\{APP_NAME}\\{APP_NAME}.lnk" "$INSTDIR\\{APP_NAME}.exe"
            CreateShortcut "$SMPROGRAMS\\{APP_NAME}\\Désinstaller.lnk" "$INSTDIR\\Uninstall.exe"
            WriteRegStr HKCU "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{APP_NAME}" "DisplayName" "{APP_NAME}"
            WriteRegStr HKCU "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{APP_NAME}" "UninstallString" "$INSTDIR\\Uninstall.exe"
            WriteRegStr HKCU "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{APP_NAME}" "DisplayVersion" "{APP_VERSION}"
        SectionEnd

        Section "Uninstall"
            RMDir /r "$INSTDIR"
            Delete "$DESKTOP\\{APP_NAME}.lnk"
            RMDir /r "$SMPROGRAMS\\{APP_NAME}"
            DeleteRegKey HKCU "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{APP_NAME}"
        SectionEnd
    """).strip()

    nsi_path = f"{BUILD_DIR}\\installer.nsi"
    with open(nsi_path, "w", encoding="utf-8") as f:
        f.write(nsi_content)

    run(f'"{makensis}" "{nsi_path}"')
    print(f"\n  Installeur Windows créé : {DIST_DIR}\\{APP_NAME} Setup.exe")

# ─── Linux ────────────────────────────────────────────────────────────────────

def build_linux():
    pyinstaller_build()

    step("Création de l'AppImage")

    appimage_tool = shutil.which("appimagetool")
    if not appimage_tool:
        print("  appimagetool non trouvé, téléchargement...")
        run("wget -q https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-x86_64.AppImage -O appimagetool")
        run("chmod +x appimagetool")
        appimage_tool = "./appimagetool"

    app_dir = f"{BUILD_DIR}/pyinstaller/{APP_NAME}"
    appdir  = f"{BUILD_DIR}/AppDir"
    icon_dir = f"{appdir}/usr/share/icons/hicolor/512x512/apps"
    os.makedirs(f"{appdir}/usr/bin", exist_ok=True)
    os.makedirs(icon_dir, exist_ok=True)

    run(f'cp -r "{app_dir}/." "{appdir}/usr/bin/"')

    # Icône de substitution minimale (carré bleu 512x512)
    try:
        from PIL import Image, ImageDraw
        img  = Image.new("RGBA", (512, 512), "#4f7cff")
        draw = ImageDraw.Draw(img)
        draw.ellipse([60, 60, 452, 452], fill="white")
        draw.text((180, 200), "MS", fill="#4f7cff")
        img.save(f"{icon_dir}/{APP_NAME}.png")
        img.save(f"{appdir}/{APP_NAME}.png")
    except Exception:
        # Sans Pillow : icône vide (l'AppImage fonctionnera quand même)
        open(f"{icon_dir}/{APP_NAME}.png", "wb").close()
        open(f"{appdir}/{APP_NAME}.png", "wb").close()

    desktop = textwrap.dedent(f"""
        [Desktop Entry]
        Name={APP_NAME}
        Exec={APP_NAME}
        Icon={APP_NAME}
        Type=Application
        Categories=Utility;Network;
    """).strip()
    with open(f"{appdir}/{APP_NAME}.desktop", "w") as f:
        f.write(desktop)

    apprun = textwrap.dedent(f"""
        #!/bin/bash
        HERE="$(dirname "$(readlink -f "$0")")"
        exec "$HERE/usr/bin/{APP_NAME}" "$@"
    """).strip()
    apprun_path = f"{appdir}/AppRun"
    with open(apprun_path, "w") as f:
        f.write(apprun)
    run(f'chmod +x "{apprun_path}"')

    os.makedirs(DIST_DIR, exist_ok=True)
    out = f'{DIST_DIR}/{APP_NAME}.AppImage'
    run(f'ARCH=x86_64 "{appimage_tool}" "{appdir}" "{out}"')
    run(f'chmod +x "{out}"')
    print(f"\n  AppImage Linux créée : {out}")

# ─── Nettoyage ────────────────────────────────────────────────────────────────

def cleanup():
    shutil.rmtree(BUILD_DIR, ignore_errors=True)

# ─── Point d'entrée ───────────────────────────────────────────────────────────

def main():
    print(f"\n{'='*60}")
    print(f"  Build — {APP_NAME} v{APP_VERSION}")
    print(f"  OS détecté : {OS}")
    print(f"{'='*60}")

    if not os.path.exists(SCRIPT):
        print(f"\n  Script source introuvable : {SCRIPT}")
        sys.exit(1)

    require("pyinstaller", "pip install pyinstaller")

    if OS == "Darwin":
        build_mac()
    elif OS == "Windows":
        build_windows()
    elif OS == "Linux":
        build_linux()
    else:
        print(f"\n  OS non supporté : {OS}")
        sys.exit(1)

    cleanup()

    print(f"\n{'='*60}")
    print(f"  Build terminé. Résultat dans : {DIST_DIR}/")
    print(f"{'='*60}\n")

if __name__ == "__main__":
    main()