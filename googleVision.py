import os
from pdf2image import convert_from_path
from google.cloud import vision
from PIL import Image
from tqdm import tqdm
import pandas as pd  


# === Folder paths ===
input_folder = "/Users/fahad/Desktop/car_stolen/PDFs"   # Folder containing your scanned PDFs
output_folder = "/Users/fahad/Desktop/car_stolen/textsVision"   # Folder to store OCR text output
os.makedirs(output_folder, exist_ok=True)

# === Initialize Vision API client ===
client = vision.ImageAnnotatorClient()

failed_files = []

def extract_text_from_image(image_path):
    """Send image to Google Vision OCR and return extracted text"""
    with open(image_path, "rb") as image_file:
        content = image_file.read()
    image = vision.Image(content=content)

    # Add simple retry logic for API failures
    for attempt in range(3):
        try:
            response = client.document_text_detection(image=image)
            return response.full_text_annotation.text
        except Exception as e:
            if attempt < 2:
                time.sleep(2)
            else:
                raise e

for pdf_file in tqdm(os.listdir(input_folder), desc="Processing PDFs"):
    if not pdf_file.lower().endswith(".pdf"):
        continue

    pdf_path = os.path.join(input_folder, pdf_file)
    output_path = os.path.join(output_folder, pdf_file.replace(".pdf", ".txt"))

    try:
        # Convert only the FIRST PAGE of the PDF
        pages = convert_from_path(pdf_path, first_page=1, last_page=1, dpi=200)
        if not pages:
            raise ValueError("No pages found")

        page = pages[0]
        width, height = page.size
        cropped = page.crop((0, 0, width, int(height * 0.25)))

        temp_image = "temp_image.jpg"
        cropped.save(temp_image)

        text = extract_text_from_image(temp_image)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(text.strip())

    except Exception as e:
        print(f"❌ Error processing {pdf_file}: {e}")
        failed_files.append({"filename": pdf_file, "error": str(e)})
        continue

# === Save failures to Excel ===
if failed_files:
    df = pd.DataFrame(failed_files)
    df.to_excel("ocr_failures.xlsx", index=False)
    print(f"\n⚠️  Some files failed. Logged in ocr_failures.xlsx ({len(failed_files)} total).")
else:
    print("\n✅ All PDFs processed successfully!")
