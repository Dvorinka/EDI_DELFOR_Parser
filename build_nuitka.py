# build_nuitka.py
import os
import shutil
import subprocess
import sys
from pathlib import Path

def run_command(command):
    """Helper function to run shell commands"""
    print(f"Running: {' '.join(command)}")
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode != 0:
        print("Error:", result.stderr)
        sys.exit(1)
    return result

def main():
    # Configuration
    script_name = "edi_parser_main.py"
    app_name = "EDI_Parser"
    icon_path = None  # Set to path of your .ico file if you have one
    output_dir = "dist"
    temp_dir = "build"
    
    # Clean up previous builds
    for dir_path in [output_dir, temp_dir]:
        if os.path.exists(dir_path):
            print(f"Cleaning up {dir_path}...")
            shutil.rmtree(dir_path)
    
    # Create output directories
    os.makedirs(output_dir, exist_ok=True)
    
    # Base Nuitka command
    cmd = [
        "python", "-m", "nuitka",
        "--standalone",
        "--onefile",
        "--windows-disable-console",  # For GUI apps
        f"--output-dir={output_dir}",
        f"--windows-icon-from-ico={icon_path}" if icon_path else "",
        f"--include-package=tkinter",
        f"--include-package=openpyxl",
        "--enable-plugin=tk-inter",
        "--remove-output",
        "--assume-yes-for-downloads",
        "--follow-imports",
        "--follow-import-to=*",
        "--nofollow-import-to=*.test",
        "--nofollow-import-to=*.tests",
        "--nofollow-import-to=*.unittest",
        "--nofollow-import-to=*.test_*",
        "--nofollow-import-to=*conftest*",
        "--nofollow-import-to=*pytest*",
        "--nofollow-import-to=*setuptools*",
        "--nofollow-import-to=*pip*",
        "--nofollow-import-to=*distutils*",
        "--nofollow-import-to=*numpy*",  # Exclude if not needed
        "--nofollow-import-to=*matplotlib*",  # Exclude if not needed
        "--windows-company-name=YourCompany",
        f"--windows-file-version=1.0",
        f"--windows-product-version=1.0",
        f"--windows-file-description={app_name}",
        f"--windows-product-name={app_name}",
        script_name
    ]
    
    # Remove empty strings from command list
    cmd = [x for x in cmd if x]
    
    # Run Nuitka
    print("Starting Nuitka compilation...")
    result = run_command(cmd)
    
    # Move the final executable if needed
    if os.name == 'nt':  # Windows
        exe_name = f"{app_name}.exe"
        src = os.path.join(output_dir, script_name.replace('.py', '.exe'))
        dst = os.path.join(".", exe_name)
        if os.path.exists(dst):
            os.remove(dst)
        shutil.move(src, dst)
        print(f"\nCompilation complete! Executable created: {exe_name}")
    else:
        print("\nCompilation complete! Check the dist/ directory for the output.")

if __name__ == "__main__":
    main()