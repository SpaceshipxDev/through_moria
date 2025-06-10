import subprocess
import os
import sys
from pathlib import Path

def convert_ppt_to_images(ppt_file, output_dir=None):
    """
    Convert PowerPoint slides to PNG images using LibreOffice headless mode
    
    Args:
        ppt_file: Path to PowerPoint file (.ppt, .pptx)
        output_dir: Directory to save images (optional, defaults to same dir as ppt)
    """
    
    ppt_path = Path(ppt_file)
    
    # Check if file exists
    if not ppt_path.exists():
        print(f"Error: File {ppt_file} not found")
        return False
    
    # Set output directory
    if output_dir is None:
        output_dir = ppt_path.parent
    else:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    
    # LibreOffice command
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "png",
        "--outdir", str(output_dir),
        str(ppt_path)
    ]
    
    try:
        print(f"Converting {ppt_file}...")
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        print(f"Success! Images saved to: {output_dir}")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"Error converting file: {e}")
        print(f"Command output: {e.stdout}")
        print(f"Command error: {e.stderr}")
        return False
    except FileNotFoundError:
        print("Error: LibreOffice not found. Make sure it's installed and in your PATH")
        return False

def main():
    if len(sys.argv) < 2:
        print("Usage: python script.py <powerpoint_file> [output_directory]")
        print("Example: python script.py presentation.pptx screenshots/")
        return
    
    ppt_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    convert_ppt_to_images(ppt_file, output_dir)

if __name__ == "__main__":
    main()