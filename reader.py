import subprocess
import os
import sys
from pathlib import Path

def convert_ppt_to_images(ppt_file, output_dir=None):
    """
    Convert PowerPoint slides to PNG images via PDF intermediate
    
    Args:
        ppt_file: Path to PowerPoint file (.ppt, .pptx)
        output_dir: Directory to save images (optional, defaults to current directory)
    """
    
    ppt_path = Path(ppt_file)
    
    # Check if file exists
    if not ppt_path.exists():
        print(f"Error: File {ppt_file} not found")
        return False
    
    # Set output directory - force current working directory if none specified
    if output_dir is None:
        output_dir = Path.cwd()
    else:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    
    # Step 1: Convert PowerPoint to PDF
    pdf_path = output_dir / f"{ppt_path.stem}.pdf"
    
    pdf_cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(ppt_path)
    ]
    
    try:
        print(f"Converting {ppt_file} to PDF...")
        result = subprocess.run(pdf_cmd, capture_output=True, text=True, check=True)
        
        if not pdf_path.exists():
            print("Error: PDF conversion failed")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"Error converting to PDF: {e}")
        print(f"Command error: {e.stderr}")
        return False
    except FileNotFoundError:
        print("Error: LibreOffice not found. Make sure it's installed and in your PATH")
        return False
    
    # Step 2: Convert PDF pages to PNG images using ImageMagick
    try:
        print("Converting PDF pages to PNG images...")
        
        # ImageMagick command to convert all PDF pages to PNG
        convert_cmd = [
            "convert",
            "-density", "300",  # High resolution
            "-quality", "100",
            str(pdf_path),
            str(output_dir / f"{ppt_path.stem}-slide-%02d.png")
        ]
        
        result = subprocess.run(convert_cmd, capture_output=True, text=True, check=True)
        
        # Clean up the temporary PDF
        pdf_path.unlink()
        
        print(f"Success! Slide images saved to: {output_dir}")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"Error converting PDF to images: {e}")
        print(f"Command error: {e.stderr}")
        print("Make sure ImageMagick is installed (brew install imagemagick)")
        return False
    except FileNotFoundError:
        print("Error: ImageMagick 'convert' command not found")
        print("Install it with: brew install imagemagick")
        return False

def main():
    if len(sys.argv) < 2:
        print("Usage: python script.py <powerpoint_file> [output_directory]")
        print("Example: python script.py presentation.pptx screenshots/")
        print("\nRequires: LibreOffice and ImageMagick")
        print("Install ImageMagick: brew install imagemagick")
        return
    
    ppt_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    convert_ppt_to_images(ppt_file, output_dir)

if __name__ == "__main__":
    main()