#!/usr/bin/env python3
"""Build script for the PyInstaller artifacts (CLI + GUI)."""

import subprocess
import sys


def run_command(cmd):
    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, check=False)
    if result.returncode != 0:
        print(f"Error: Command failed with exit code {result.returncode}")
        sys.exit(result.returncode)


def main():
    print("Building CrispTranslator apps...")

    print("\n--- Building CLI App ---")
    run_command(["pyinstaller", "format-transplant-cli.spec", "--clean", "--noconfirm"])

    print("\n--- Building GUI App ---")
    run_command(["pyinstaller", "transplant-app-gui.spec", "--clean", "--noconfirm"])

    print("\nBuild process complete. Check 'dist/' folder for artifacts.")


if __name__ == "__main__":
    main()
