#!/usr/bin/env python3
import subprocess
import sys
import os

def run_command(cmd):
    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd)
    if result.returncode != 0:
        print(f"Error: Command failed with exit code {result.returncode}")
        sys.exit(result.returncode)

def main():
    print("Building CrispTranslator apps...")
    
    # 1. CLI App
    print("
--- Building CLI App ---")
    run_command(["pyinstaller", "format-transplant-cli.spec", "--clean", "--noconfirm"])
    
    # 2. GUI App
    print("
--- Building GUI App ---")
    run_command(["pyinstaller", "transplant-app-gui.spec", "--clean", "--noconfirm"])
    
    print("
Build process complete. Check 'dist/' folder for artifacts.")

if __name__ == "__main__":
    main()
