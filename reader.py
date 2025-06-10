import subprocess
import os
import sys
from pathlib import Path

def convert_ppt_to_images(ppt_file, output_dir=None):
    """
    Convert PowerPoint slides to individual PNG images using LibreOffice macro
    
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
    
    # Create a temporary macro file for LibreOffice
    macro_content = f'''
Sub ExportSlides
    Dim oDoc As Object
    Dim oSlides As Object
    Dim oSlide As Object
    Dim i As Integer
    Dim sURL As String
    Dim sOutputDir As String
    Dim sBaseName As String
    Dim sFileName As String
    Dim oExportFilter As Object
    Dim aArgs(2) As New com.sun.star.beans.PropertyValue
    
    oDoc = ThisComponent
    oSlides = oDoc.getDrawPages()
    
    sOutputDir = "{str(output_dir).replace('\\', '/')}"
    sBaseName = "{ppt_path.stem}"
    
    For i = 0 To oSlides.getCount() - 1
        oSlide = oSlides.getByIndex(i)
        sFileName = sOutputDir + "/" + sBaseName + "-slide-" + Format(i+1, "00") + ".png"
        sURL = ConvertToURL(sFileName)
        
        aArgs(0).Name = "URL"
        aArgs(0).Value = sURL
        aArgs(1).Name = "FilterName"
        aArgs(1).Value = "draw_png_Export"
        aArgs(2).Name = "PageRange"
        aArgs(2).Value = CStr(i+1)
        
        oDoc.storeToURL(sURL, aArgs())
    Next i
End Sub
'''
    
    macro_path = output_dir / "export_macro.bas"
    with open(macro_path, 'w') as f:
        f.write(macro_content)
    
    # Alternative approach: Use LibreOffice with page range export
    base_name = ppt_path.stem
    
    # First, let's try converting to PDF then splitting
    pdf_path = output_dir / f"{base_name}.pdf"
    
    pdf_cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(ppt_path)
    ]
    
    try:
        print(f"Converting {ppt_file} to PDF...")
        subprocess.run(pdf_cmd, capture_output=True, text=True, check=True)
        
        if not pdf_path.exists():
            print("Error: PDF conversion failed")
            return False
            
        print("PDF created, now splitting into individual images...")
        
        # Use Python's pdf2image if available, otherwise try sips (Mac native)
        try:
            from pdf2image import convert_from_path
            images = convert_from_path(str(pdf_path), dpi=300)
            
            for i, image in enumerate(images):
                image_path = output_dir / f"{base_name}-slide-{i+1:02d}.png"
                image.save(str(image_path), 'PNG')
                print(f"Saved: {image_path}")
            
            # Clean up PDF
            pdf_path.unlink()
            macro_path.unlink(missing_ok=True)
            
            print(f"Success! {len(images)} slide images saved to: {output_dir}")
            return True
            
        except ImportError:
            print("pdf2image not found, trying sips (Mac native)...")
            
            # Use Mac's sips command
            for page_num in range(1, 100):  # Assume max 100 slides
                output_file = output_dir / f"{base_name}-slide-{page_num:02d}.png"
                
                sips_cmd = [
                    "sips", "-s", "format", "png",
                    str(pdf_path) + f"[{page_num-1}]",  # 0-indexed for sips
                    "--out", str(output_file)
                ]
                
                result = subprocess.run(sips_cmd, capture_output=True, text=True)
                if result.returncode != 0:
                    break  # No more pages
                    
                print(f"Saved: {output_file}")
            
            # Clean up
            pdf_path.unlink()
            macro_path.unlink(missing_ok=True)
            
            print(f"Success! Slide images saved to: {output_dir}")
            return True
            
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False

def main():
    if len(sys.argv) < 2:
        print("Usage: python script.py <powerpoint_file> [output_directory]")
        print("Example: python script.py presentation.pptx")
        print("\nOptional: pip install pdf2image for better conversion")
        return
    
    ppt_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    convert_ppt_to_images(ppt_file, output_dir)

if __name__ == "__main__":
    main()