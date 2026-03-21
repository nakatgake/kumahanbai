import fitz
import sys

def convert_pdf_to_images(pdf_path, output_prefix):
    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=150)
        output_file = f"{output_prefix}_page_{page_num + 1}.png"
        pix.save(output_file)
        print(f"Saved {output_file}")
    doc.close()

if __name__ == "__main__":
    pdf_path = sys.argv[1]
    output_prefix = sys.argv[2]
    convert_pdf_to_images(pdf_path, output_prefix)
