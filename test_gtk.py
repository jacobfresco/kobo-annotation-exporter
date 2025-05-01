from weasyprint import HTML
from cairosvg import svg2png
import os
import sys

def test_gtk_installation():
    print("Starting GTK3 installation test...")
    print(f"Python version: {sys.version}")
    print(f"Current working directory: {os.getcwd()}")
    
    try:
        # Test WeasyPrint
        print("\nTesting WeasyPrint...")
        html = HTML(string='<h1>Test Page</h1><p>If you can see this, WeasyPrint is working!</p>')
        pdf = html.write_pdf()
        print("✓ WeasyPrint test passed!")
        
        # Test Cairo
        print("\nTesting Cairo...")
        svg_content = '''<?xml version="1.0" encoding="UTF-8" standalone="no"?>
        <svg width="100" height="100" xmlns="http://www.w3.org/2000/svg">
            <rect width="100" height="100" fill="blue"/>
        </svg>'''
        png_data = svg2png(bytestring=svg_content)
        print("✓ Cairo test passed!")
        
        print("\nAll tests passed! GTK3 is installed correctly.")
        return True
        
    except ImportError as e:
        print(f"\nImport Error: {str(e)}")
        print("This might indicate that some Python packages are not installed correctly.")
        return False
    except Exception as e:
        print(f"\nError during testing: {str(e)}")
        print("\nGTK3 installation might be incomplete or incorrect.")
        print("Please make sure you have:")
        print("1. Installed GTK3 runtime")
        print("2. Restarted your computer after installation")
        print("3. Installed all required Python packages")
        return False

if __name__ == "__main__":
    print("Starting test...")
    result = test_gtk_installation()
    print(f"\nTest completed with result: {'Success' if result else 'Failed'}") 