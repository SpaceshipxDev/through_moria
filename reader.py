#!/usr/bin/env python3
import subprocess
import os
import sys
import glob
from pathlib import Path

def convert_pptx_to_images(pptx_file, output_dir="./read"):
    """
    Convert PPTX to individual slide images using LibreOffice + pdftoppm
    """
    # Create output directory if it doesn't exist
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    
    # Get the base filename without extension
    base_name = Path(pptx_file).stem
    temp_pdf = f"{base_name}.pdf"
    
    try:
        # Step 1: Convert PPTX to PDF using LibreOffice
        print(f"Converting {pptx_file} to PDF...")
        subprocess.run([
            "libreoffice", "--headless", "--convert-to", "pdf", pptx_file
        ], check=True)
        
        # Step 2: Check if PDF was created
        if not os.path.exists(temp_pdf):
            raise FileNotFoundError(f"PDF conversion failed - {temp_pdf} not found")
        
        # Step 3: Convert PDF to PNG images using pdftoppm
        print(f"Converting PDF to images...")
        output_prefix = os.path.join(output_dir, f"{base_name}")
        subprocess.run([
            "pdftoppm", "-png", temp_pdf, output_prefix
        ], check=True)
        
        # Step 4: Clean up temp PDF
        os.remove(temp_pdf)
        
        # Step 5: Count and report results
        png_files = glob.glob(os.path.join(output_dir, f"{base_name}-*.png"))
        print(f"✅ Successfully created {len(png_files)} slide images in {output_dir}")
        
        return png_files
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Command failed: {e}")
        return None
    except FileNotFoundError as e:
        print(f"❌ File error: {e}")
        return None
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return None

def main():
    if len(sys.argv) != 2:
        print("Usage: python convert_slides.py <pptx_file>")
        sys.exit(1)
    
    pptx_file = sys.argv[1]
    
    if not os.path.exists(pptx_file):
        print(f"❌ File not found: {pptx_file}")
        sys.exit(1)
    
    if not pptx_file.lower().endswith('.pptx'):
        print("❌ File must be a .pptx file")
        sys.exit(1)
    
    convert_pptx_to_images(pptx_file)

if __name__ == "__main__":
    main()