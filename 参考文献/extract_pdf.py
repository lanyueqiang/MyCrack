import pdfplumber
import sys

# Set UTF-8 encoding for stdout
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

pdf_path = r'd:\work\research\论文\Crack\Network for robust and high-accuracy pavement crack segmentation.pdf'
with pdfplumber.open(pdf_path) as pdf:
    # Extract all pages
    full_text = ''
    for i, page in enumerate(pdf.pages):
        text = page.extract_text()
        if text:
            full_text += f"\n--- Page {i+1} ---\n{text}"
    
# Write to file with UTF-8 encoding
with open('pdf_content.txt', 'w', encoding='utf-8') as f:
    f.write(full_text)

print("PDF content extracted successfully!")
