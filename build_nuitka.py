# build_nuitka.py
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from tqdm import tqdm

def run_command(command, description="Running command"):
    """Helper function to run shell commands with progress"""
    print(f"\n{description}:")
    print("[" + " " * 50 + "] 0%", end="\r")
    
    # Start the process
    process = subprocess.Popen(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        universal_newlines=True,
        bufsize=1
    )
    
    # Track progress through known Nuitka build stages
    stages = [
        'Nuitka:INFO: Starting Python compilation.',
        'Nuitka:INFO: Completed Python level compilation and optimization.',
        'Nuitka:INFO: Generating source code for C backend.',
        'Nuitka:INFO: Running data composer tool for optimal constant value handling.',
        'Nuitka:INFO: Running C compilation via gcc.',
        'Nuitka:INFO: Successfully created'  # Final stage
    ]
    current_stage = 0
    
    # Process output in real-time
    while True:
        output = process.stdout.readline()
        if output == '' and process.poll() is not None:
            break
        if output:
            output = output.strip()
            print(output)  # Show the actual build output
            
            # Check for stage completion
            for i, stage in enumerate(stages):
                if stage in output and i >= current_stage:
                    current_stage = i + 1
                    progress = int((current_stage / len(stages)) * 100)
                    filled = int(50 * progress / 100)
                    bar = '=' * filled + ' ' * (50 - filled)
                    print(f"\r[{bar}] {progress}%", end="\r")
                    break
    
    # Get any remaining output and errors
    stdout, stderr = process.communicate()
    
    # Show completion
    print("\r[" + "=" * 50 + "] 100%")
    
    if process.returncode != 0:
        print("\n❌ Error during compilation:")
        print(stderr.strip())
        sys.exit(1)
    
    print("\n✅ Build stage completed successfully!")

def main():
    # Configuration
    script_name = "edi_parser_main.py"
    app_name = "EDI_Parser"
    icon_path = "icon.ico"  # Converted from icon.png
    output_dir = "dist"
    temp_dir = "build"
    
    # Convert PNG to ICO if needed
    if os.path.exists("icon.png") and not os.path.exists("icon.ico"):
        try:
            from PIL import Image
            img = Image.open("icon.png")
            img.save("icon.ico")
            print("✅ Converted icon.png to icon.ico")
        except ImportError:
            print("⚠️  Install Pillow with 'pip install pillow' for icon conversion")
            icon_path = None
    
    print(" Starting build process...")
    
    # Clean up previous builds
    print(" Cleaning up previous builds...")
    for dir_path in [output_dir, temp_dir]:
        if os.path.exists(dir_path):
            shutil.rmtree(dir_path)
    
    # Create output directories
    os.makedirs(output_dir, exist_ok=True)
    
    # Base Nuitka command
    cmd = [
        "python", "-m", "nuitka",
        "--standalone",
        "--onefile",
        "--windows-disable-console",  # Hide console window
        f"--output-dir={output_dir}",
        f"--windows-icon-from-ico={icon_path}" if icon_path and os.path.exists(icon_path) else "",
        f"--include-package=tkinter",
        f"--include-package=openpyxl",
        "--enable-plugin=tk-inter",
        "--remove-output",
        "--assume-yes-for-downloads",
        "--follow-imports",
        script_name
    ]
    
    # Remove empty strings from command list
    cmd = [x for x in cmd if x]
    
    try:
        # Run Nuitka with progress
        run_command(cmd, " Compiling with Nuitka")
        
        # Move the final executable if needed
        if os.name == 'nt':  # Windows
            exe_name = f"{app_name}.exe"
            src = os.path.join(output_dir, script_name.replace('.py', '.exe'))
            dst = os.path.join(".", exe_name)
            if os.path.exists(dst):
                os.remove(dst)
            shutil.move(src, dst)
            print(f"\n Compilation complete! Executable created: {exe_name}")
        else:
            print("\n Compilation complete! Check the dist/ directory for the output.")
            
    except Exception as e:
        print(f"\n Error during build: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()