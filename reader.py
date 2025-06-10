import subprocess
import sys
from pathlib import Path

def convert_ppt_to_images(ppt_file, output_dir=None):
    """
    Convert PowerPoint slides to individual PNG images using LibreOffice
    """
    
    ppt_path = Path(ppt_file)
    
    if not ppt_path.exists():
        print(f"Error: File {ppt_file} not found")
        return False
    
    # Set output directory
    if output_dir is None:
        output_dir = Path.cwd()
    else:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    
    # LibreOffice command to export all slides as images
    cmd = [
        "libreoffice",
        "--headless",
        "--invisible",
        "--convert-to", "png",
        "--outdir", str(output_dir),
        str(ppt_path)
    ]
    
    try:
        print(f"Converting slides from {ppt_file}...")
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        print(f"Done! Images saved to: {output_dir}")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        print(f"Output: {e.stdout}")
        print(f"Error: {e.stderr}")
        return False
    except FileNotFoundError:
        print("LibreOffice not found. Try the full path:")
        print('"/Applications/LibreOffice.app/Contents/MacOS/soffice"')
        return False

def main():
    if len(sys.argv) < 2:
        print("Usage: python script.py <powerpoint_file> [output_directory]")
        return
    
    ppt_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    convert_ppt_to_images(ppt_file, output_dir)

if __name__ == "__main__":
    main()