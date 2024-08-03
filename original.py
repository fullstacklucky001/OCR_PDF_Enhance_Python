import argparse
import os
from dataclasses import dataclass
from typing import List, Dict, Tuple

import pdf2image
import pytesseract
import tabula
from PIL import Image
from PyPDF2 import PdfWriter, PdfReader
from collections import defaultdict
import re
import pandas as pd
from tqdm import tqdm

# arbitrarily large integer for sorting rank
MAX_LABEL_NUMBER = 1000000
DEFAULT_CONVERSION_FILE = 'Conversion File.xlsx'

poppler_path = "C:/Users/Administrator/Downloads/Release-24.07.0-0/poppler-24.07.0/Library/bin/"
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

@dataclass
class ShippingLabel:
    pdf_index: int
    pick_list_rank: int
    upc_ref: str


# character replacement for pseudo fuzzy matching to counteract OCR failures
fuzzy_replacements = {
    r'[OQ]': '0',
    r'[S]': '5',
    r'[-]': '',
    r'[1|]': 'I',
    r'[\s]': ''
}


def fuzz(text):
    for k, v in fuzzy_replacements.items():
        text = re.sub(k, v, text)
    return text





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
    for i, label in tqdm(enumerate(slips)):
        fuzzed_ref = fuzz(label.upc_ref)
        if fuzzed_ref in upc_lookup:
            label.upc_ref = upc_lookup[fuzzed_ref]
        else:
            # If no match is found, use the original upc_ref or skip
            label.upc_ref = label.upc_ref  # or use a placeholder like 'UNMATCHED'
            # Optionally, you could skip processing this label with `continue`
            # continue
        
        label.pick_list_rank = get_packing_rank(label.upc_ref, packing_order)

    # Sort only those slips which have a valid packing order (not MAX_LABEL_NUMBER)
    slips = sorted([slip for slip in slips if slip.pick_list_rank != MAX_LABEL_NUMBER], key=lambda label: label.pick_list_rank)
    return slips

# read in the UPC conversion file
def read_conversion(conversion_file_path) -> Dict[str, str]:
    lookup = {}
    conversions = pd.read_excel(conversion_file_path)
    for conversion in conversions.values:
        lookup[fuzz(conversion[1].upper().strip())] = fuzz(conversion[0].upper().strip())
    return lookup


# parse the entire label pdf into a list of labels
def parse_label_pdf(label_file_name: str) -> List[ShippingLabel]:
    refs = []
    print(f"Loading {label_file_name}...")
    page_images = pdf2image.convert_from_path(label_file_name, dpi=500, grayscale=True, thread_count=10, poppler_path=poppler_path)
    for i, page in tqdm(enumerate(page_images), "Reading reference numbers...", len(page_images)):
        refs.append(ShippingLabel(i, MAX_LABEL_NUMBER, read_reference_number_usps(page)))
    return refs


##
# excerpt from pdf_combo_new.py
def read_reference_number(image: Image, coords: Tuple[int, int, int, int]) -> str:
    # expected coords for reference number
    cropped_image = image.crop(coords)
    padded_image = Image.new(cropped_image.mode, (cropped_image.width, cropped_image.height + 200), 'white')
    padded_image.paste(cropped_image, (0, 100))

    # Only include relevant characters; this keeps OCR on course
    # DPI 300 is just an approximation
    # PSM level 6 means "Assume a single uniform block of text."
    text = str(pytesseract.image_to_string(padded_image,
                                           config='''
                          -c tessedit_char_whitelist="Trx Ref No.: 1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ-"
                          --dpi 500
                          --psm 6'''))
    # remove any weird chars and spaces at end
    text = re.sub(r'[^A-Z0-9]+$', '', text)
    # split on "- 1X" or similar string
    text = re.split(r'-\s*\d+[^xX]*[xX]\s*', text)[-1]
    # remove any spaces
    text = re.sub(r'\s', '', text)

    return fuzz(text.upper())


def read_reference_number_ups(image: Image) -> str:
    coords = (0, 2820, 1500, 2970)
    return read_reference_number(image, coords)


def read_reference_number_usps(image: Image) -> str:
    # old coords = (0, 1930, 1437, 2003)
    coords = (0, 2033, 1437, 2100)
    return read_reference_number(image, coords)

def write_pdf(slips: List[ShippingLabel], labels_pdf_path: str, output_path: str) -> None:
    output_writer = PdfWriter()
    input_reader = PdfReader(labels_pdf_path)
    for i, slip in enumerate(slips):
        print(slip)
        output_writer.add_page(input_reader.pages[slip.pdf_index])
        if (slip.pick_list_rank >= MAX_LABEL_NUMBER):
            print(f"Slip {slip.pdf_index + 1} cannot be matched. Appended as page {i + 1}")
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
    # Argparse section
    argParseDescription = (
        'Pick List and Label sorting tool. Takes a PDF for a pick list and a PDF for labels, '
        'then outputs a new PDF for shipping labels that is in the order of the pick list.')

    parser = argparse.ArgumentParser(description=argParseDescription)
    parser.add_argument('-p', required=False, dest='pickList',
                        metavar='picklist', help='The PDF for the packing slips')
    parser.add_argument('-l', required=False, dest='shippingLabels',
                        metavar='labels', help='The PDF for the pick list')
    parser.add_argument('-o', dest='outputFile', help='The path to the desired output file')
    parser.add_argument('-c', default=DEFAULT_CONVERSION_FILE, dest='conversionFile', help='The path to the UPC '
                                                                                           'conversion file')
    args = parser.parse_args()

    if args.pickList == None:
        args.pickList = input("Enter path to pick list: ").strip()

    if args.shippingLabels == None:
        args.shippingLabels = input("Enter path to shipping labels: ").strip()

    if args.outputFile == None:
        args.outputFile = input("Enter path to output file (leave empty for default): ").strip()
        if args.outputFile == '':
            args.outputFile = os.path.abspath(args.shippingLabels.replace('.pdf', '_reordered.pdf'))

    if args.conversionFile == DEFAULT_CONVERSION_FILE:
        args.conversionFile = os.path.join(
            os.path.abspath(os.path.dirname(__file__)),
            args.conversionFile
        )

    # Main section

    sorted_slips = sort_slips(args.pickList, args.shippingLabels, args.conversionFile)
    write_pdf(sorted_slips, args.shippingLabels, args.outputFile)
    print(f"Ordered list written at {args.outputFile}")


if __name__ == "__main__":
    Main()
