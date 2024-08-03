import argparse
import os
from dataclasses import dataclass
from typing import List, Dict, Tuple

import pdf2image
import pytesseract
import tabula
from PIL import Image
# from PyPDF2 import PdfWriter, PdfReader
from PyPDF2 import PdfFileWriter, PdfFileReader
from collections import defaultdict
import re
import pandas as pd
from tqdm import tqdm
from PIL import ImageEnhance, ImageFilter

# Arbitrarily large integer for sorting rank
MAX_LABEL_NUMBER = 1000000
DEFAULT_CONVERSION_FILE = 'Conversion File.xlsx'

# poppler_path = "C:/Users/Administrator/Downloads/Release-24.07.0-0/poppler-24.07.0/Library/bin/"
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

@dataclass
class ShippingLabel:
    pdf_index: int
    pick_list_rank: int
    upc_ref: str

# Character replacement for pseudo fuzzy matching to counteract OCR failures
fuzzy_replacements = {
    r'[OQ]': '0',
    r'[S]': '5',
    r'[-]': '',
    # r'[1]': 'I',
    r'[1|]': 'I',
    r'[\s]': ''
}


def fuzz(text):
    # Remove consecutive 'I's
    for k, v in fuzzy_replacements.items():
        text = re.sub(k, v, text)
    text = re.sub(r'II+', 'I', text)
    return text

def preprocess_image(image: Image) -> Image:
    image = image.convert('L')
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(3)
    image = image.filter(ImageFilter.MedianFilter())
    return image

def get_packing_rank(upc_ref, packing_order):
    return packing_order[upc_ref]

def sort_slips(pick_list_path, shipping_label_path, conversion_file_path) -> List[ShippingLabel]:
    pick_list = tabula.read_pdf(pick_list_path, pages='all', area=(0, 0, 100000, 100000),
                                pandas_options={"header": None})
    
    packing_order = defaultdict(lambda: MAX_LABEL_NUMBER)
    upc_lookup = read_conversion(conversion_file_path)

    i = 0
    for table in pick_list:
        entries = [row for row in table.values if isinstance(row[0], str)]
        for entry in entries:
            fuzzed_entry = fuzz(entry[0])
            packing_order[fuzzed_entry] = i
            i += 1

    slips = parse_label_pdf(shipping_label_path)

    total_labels = len(slips)
    print(f"Total labels parsed: {total_labels}")

    unmatched_labels = []
    for label in tqdm(slips, desc="Processing labels"):
        fuzzed_ref = fuzz(label.upc_ref)
        if fuzzed_ref in upc_lookup:
            label.upc_ref = upc_lookup[fuzzed_ref]
        else:
            unmatched_labels.append(label)
        
        label.pick_list_rank = get_packing_rank(label.upc_ref, packing_order)
        if label.pick_list_rank == MAX_LABEL_NUMBER:
            print(f"Unmatched label: {label.upc_ref}, Original: {label.upc_ref}, PDF Index: {label.pdf_index}")

    if unmatched_labels:
        print(f"Unmatched labels: {len(unmatched_labels)}")
        for label in unmatched_labels:
            print(f"Unmatched label: {label}")

    # Sort all slips, unmatched ones will get the MAX_LABEL_NUMBER rank and go to the end
    all_slips = sorted(slips, key=lambda label: (label.pick_list_rank, label.pdf_index))
    print(f"Total sorted labels: {len(all_slips)}")
    return all_slips

# Read in the UPC conversion file
def read_conversion(conversion_file_path) -> Dict[str, str]:
    lookup = {}
    conversions = pd.read_excel(conversion_file_path)
    for conversion in conversions.values:
        lookup[fuzz(conversion[1].upper().strip())] = fuzz(conversion[0].upper().strip())
    return lookup

# Parse the entire label pdf into a list of labels
def parse_label_pdf(label_file_name: str) -> List[ShippingLabel]:
    refs = []
    print(f"Loading {label_file_name}...")
    page_images = pdf2image.convert_from_path(label_file_name, dpi=500, grayscale=True, thread_count=10)

    print(f"Total pages parsed: {len(page_images)}")
    for i, page in tqdm(enumerate(page_images), "Reading reference numbers...", total=len(page_images)):
        refs.append(ShippingLabel(i, MAX_LABEL_NUMBER, read_reference_number_usps(page)))
    return refs

def read_reference_number(image: Image, coords: Tuple[int, int, int, int]) -> str:
    cropped_image = image.crop(coords)
    padded_image = Image.new(cropped_image.mode, (cropped_image.width, cropped_image.height + 200), 'white')
    padded_image.paste(cropped_image, (0, 100))

    # Preprocess the image to enhance OCR accuracy
    padded_image = preprocess_image(padded_image)

    text = str(pytesseract.image_to_string(padded_image,
                                           config='''-c tessedit_char_whitelist="Trx Ref No.: 1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ-" --dpi 500 --psm 6'''))
    text = re.sub(r'[^A-Z0-9]+$', '', text)
    text = re.split(r'-\s*\d+[^xX]*[xX]\s*', text)[-1]
    text = re.sub(r'\s', '', text)
    return fuzz(text.upper())

def read_reference_number_ups(image: Image) -> str:
    coords = (0, 2820, 1500, 2970)
    return read_reference_number(image, coords)

def read_reference_number_usps(image: Image) -> str:
    coords = (0, 2033, 1437, 2100)
    return read_reference_number(image, coords)

def write_pdf(slips: List[ShippingLabel], labels_pdf_path: str, output_path: str) -> None:
    output_writer = PdfFileWriter()
    input_reader = PdfFileReader(labels_pdf_path)
    for i, slip in enumerate(slips):
        output_writer.addPage(input_reader.getPage(slip.pdf_index))
        if slip.pick_list_rank >= MAX_LABEL_NUMBER:
            print(f"Slip {slip.pdf_index + 1} cannot be matched. Appended as page {slips.index(slip) + 1}")

    written = False
    while not written:
        try:
            with open(output_path, 'wb') as of:
                output_writer.write(of)
            written = True
        except Exception as e:
            if input(f"ERROR: Cannot open {output_path}. Please make sure it is not open elsewhere\n"
                     f"To retry, press ENTER. To exit, enter 'e'\n") == 'e':
                break

def Main():
    argParseDescription = (
        'Pick List and Label sorting tool. Takes a PDF for a pick list and a PDF for labels, '
        'then outputs a new PDF for shipping labels that is in the order of the pick list.')

    parser = argparse.ArgumentParser(description=argParseDescription)
    parser.add_argument('-p', required=False, dest='pickList', metavar='picklist', help='The PDF for the packing slips')
    parser.add_argument('-l', required=False, dest='shippingLabels', metavar='labels', help='The PDF for the pick list')
    parser.add_argument('-o', dest='outputFile', help='The path to the desired output file')
    parser.add_argument('-c', default=DEFAULT_CONVERSION_FILE, dest='conversionFile', help='The path to the UPC conversion file')
    args = parser.parse_args()

    if args.pickList is None:
        args.pickList = input("Enter path to pick list: ").strip()

    if args.shippingLabels is None:
        args.shippingLabels = input("Enter path to shipping labels: ").strip()

    if args.outputFile is None:
        args.outputFile = input("Enter path to output file (leave empty for default): ").strip()
        if args.outputFile == '':
            args.outputFile = os.path.abspath(args.shippingLabels.replace('.pdf', '_reordered.pdf'))

    if args.conversionFile == DEFAULT_CONVERSION_FILE:
        args.conversionFile = os.path.join(os.path.abspath(os.path.dirname(__file__)), args.conversionFile)

    sorted_slips = sort_slips(args.pickList, args.shippingLabels, args.conversionFile)
    print(f"Final sorted slips count: {len(sorted_slips)}")
    write_pdf(sorted_slips, args.shippingLabels, args.outputFile)
    print(f"Ordered list written at {args.outputFile}")

if __name__ == "__main__":
    Main()