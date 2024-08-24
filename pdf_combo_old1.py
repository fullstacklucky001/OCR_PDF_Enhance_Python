# \package pdfCombo
#
#     \brief   This script takes in packing slips and labels and sorts the packing slips to be in the same order as the labels.
#
#


import argparse
import re
from enum import Enum
from typing import List, Tuple
from dataclasses import dataclass
import dataclasses
import json

from PIL import Image
import pdf2image
import pytesseract
from tqdm import tqdm
import tabula
import pandas as pd
from PyPDF2 import PdfFileWriter, PdfFileReader

@dataclass
class Mode:
    name: str
    sort_key: str
    scan_area: (int, int, int, int)
    ship_scan: (int, int, int, int)
    order_scan: (int, int, int, int)
    index_val: str
    shipping_svc: str
    slips_path: str
    labels_path: str
    scan_area_2: (int, int, int, int) = (0, 0, 0, 0)


class Store(Enum):
    Target = "target"
    Belk = "belk"
    GSI = "gsi"
    HSN = "hsn"
    Hibbett = "hibbett"
    BedBath = "bedbath"


def get_mode(slips_path: str, labels_path: str, custom_store: Store = None) -> Mode:
    store_string: str = slips_path.lower()

    # Only check the path if the custom store is none.
    if custom_store != None:
        store_string = custom_store.value

    # Check if the store string contains the name of the store (defined in the enum above)
    if Store.Target.value in store_string:
        mode = Mode(Store.Target.name, "MFG ID",
                    (210, 10, 400, 575), (100, 250, 225, 575), (93, 472, 107, 545),
                    "SEND TO:", "", slips_path, labels_path)
    elif Store.Belk.value in store_string:
        mode = Mode(Store.Belk.name, "Item Number",
                    (125, 10, 300, 585), (5, 125, 50, 250), (65, 65, 100, 250),
                    "Ship To:", "", slips_path, labels_path)
    elif Store.GSI.value in store_string:
        mode = Mode(Store.GSI.name, "Item Number",
                    (220, 10, 240, 560), (100, 305,
                                          180, 500), (220, 450, 240, 560),
                    "", "", slips_path, labels_path)
    elif Store.HSN.value in store_string:
        mode = Mode(Store.HSN.name, "Item Number",
                    (140, 10, 190, 600), (240, 10, 340, 220), (495, 10, 550, 155),
                    "", "", slips_path, labels_path)
    elif Store.Hibbett.value in store_string:
        mode = Mode(Store.Hibbett.name, "Item Number",
                    (210, 10, 240, 575), (100, 250, 225, 575), (222,487,236,538),
                    "Ship To:", "", slips_path, labels_path, scan_area_2=(240, 10, 280, 575))
    elif Store.BedBath.value in store_string:
        mode = Mode(Store.BedBath.name, "Vendor Part #",
                    (160, 20, 241, 601), (600, 305, 750, 600), (10, 200, 75, 600),
                    "Shipped To:", "", slips_path, labels_path)
    else:
        print("can't detect retailer name. Please select: ")
        print(f"[1] {Store.Target.name}")
        print(f"[2] {Store.Belk.name}")
        print(f"[3] {Store.GSI.name}")
        print(f"[4] {Store.HSN.name}")
        print(f"[5] {Store.Hibbett.name}")
        print(f"[6] {Store.BedBath.name}")

        store = input("[1,2,3,4,5,6] > ")
        if "1" in store:
            custom_store = Store.Target
        elif "2" in store:
            custom_store = Store.Belk
        elif "3" in store:
            custom_store = Store.GSI
        elif "4" in store:
            custom_store = Store.HSN
        elif "5" in store:
            custom_store = Store.Hibbett
        elif "6" in store:
            custom_store = Store.BedBath

        mode = get_mode(slips_path, labels_path, custom_store)
            

    return mode


@dataclass
class ShippingLabel:
    page_num: int
    full_name: str
    addr_line1: str
    addr_line2: str
    addr_line3: str
    addr_line4: str
    reference_num: str = ""


#region Shipping Labels
def _parseSingleShippingLabel_NotHSN(label_image, crop_coordinates: List, store_name: str) -> ShippingLabel:
    # TODO: refactor read_reference_number_ups to take in "is_GSI_OR_HSN instead of store_name"
    ref_num = read_reference_number_ups(label_image, store_name)
    if ref_num == "N/A":
        ref_num = read_reference_number_fedex(label_image)

    # the last label that we looped through. This will either be a valid label or the last attempt.
    last_parsed_label = None
    for coords in crop_coordinates:
        cropped_label = label_image.crop(coords).convert("L")
        text = str(pytesseract.image_to_string(cropped_label, config='--psm 6'))

        last_parsed_label = get_details_list_from_shipping_label(text)
        last_parsed_label.reference_num = ref_num
        if last_parsed_label.full_name != "Label_Error":
            break

    return last_parsed_label


def _parseShippingLabels_NotHSN(label_images: List, crop_coordinates: List[Tuple[int, int, int, int]], store_name: str) -> Tuple[List[dict], List[int]]:

    output: List[dict] = []
    errors: List[int] = []

    for i, label in tqdm(enumerate(label_images), total=len(label_images)):
        last_parsed_label = _parseSingleShippingLabel_NotHSN(
            label, crop_coordinates, store_name)
        last_parsed_label.page_num = i

        output.append(last_parsed_label)

    for i, label in enumerate(output):
        if label.full_name == "Label_Error":
            # errors should be a list of indices from output
            errors.append(i)

    return output, errors


def _parseShippingLabels_HSN(label_images, crop_coordinates) -> Tuple[List[ShippingLabel], List[int]]:
    output: List[ShippingLabel] = []
    errors: List[int] = []


    for i, label in tqdm(enumerate(label_images), total=len(label_images)):
        last_parsed_label = None
        for coords in crop_coordinates:
            cropped_label = label.crop(coords).convert("L")
            text = str(pytesseract.image_to_string(cropped_label))

            last_parsed_label = ShippingLabel(
                page_num=i,
                full_name="Label_Error",
                reference_num="",
                addr_line1="",
                addr_line2="",
                addr_line3="",
                addr_line4=""
            )

            lines = text.split("\n")
            for line in lines:
                if line.find(" - ") != -1:
                    line = line.split(" - ")
                    
                    line[0] = line[0].replace("#", "")
                    last_parsed_label.full_name = line[0]

                    line[1] = line[1].replace(":", "")
                    line[1] = line[1].strip()
                    last_parsed_label.reference_num = line[1]

                elif line.find("Trx Ref No") != -1:
                    name_coordinates = (0, 300, 1215, 475)
                    name_image = label.crop(name_coordinates).convert("L")
                    name_text = str(pytesseract.image_to_string(name_image))
                    name_text = name_text.split('\n')[1]

                    line = line.split(".:")
                    last_parsed_label.full_name = name_text
                    line[1] = line[1].replace(":", "")
                    line[1] = line[1].strip()
                    last_parsed_label.reference_num = line[1].strip()

            if last_parsed_label.full_name != "Label_Error":
                break

        output.append(last_parsed_label)

    for i, label in enumerate(output):
        if label.full_name == "Label_Error":
            errors.append(i)

    return output, errors


# This returns (parsed labels, indices of errored labels)
def parseShippingLabel(mode: Mode) -> Tuple[List[ShippingLabel], List[int]]:
    page_images: List = pdf2image.convert_from_path(
        mode.labels_path, dpi=500, grayscale=True)
    crop_coordinates = []

    # HSN is special
    if mode.name == Store.HSN.name:
        crop_coordinates = [(0, 1875, 1450, 2100), (0, 2850, 1000, 2950)]
        return _parseShippingLabels_HSN(page_images, crop_coordinates)

    if mode.name == Store.Target.name:
        # there are multiple possible locations for the information on the label.
        target1 = (70, 350, 1700, 820)  # Fedex Home Delivery
        target2 = (70, 400, 1700, 820)  # Fedex Home Delivery
        target3 = (144, 407, 1950, 730)
        crop_coordinates = [target1, target2, target3]
    elif mode.name == Store.Belk.name:
        crop_coordinates = [(200, 850, 1300, 1210)]
    elif mode.name == Store.BedBath.name:
        crop_coordinates = [(64, 350, 1700, 810)]
    elif mode.name == Store.GSI.name or mode.name == Store.Hibbett.name:
        crop_coordinates = [
            (70, 350, 1700, 820),
            (70, 400, 1700, 820),
            (144, 407, 1950, 730)]

    return _parseShippingLabels_NotHSN(page_images, crop_coordinates, mode.name)


def checkShippingLabels(labels: List[ShippingLabel], errors: List[int]) -> List[ShippingLabel]:
    def correctShippingLabel(label: ShippingLabel, label_number: int) -> ShippingLabel:
        new_name = input(f"Enter name for shipping label #{label_number}: ").upper()
        label.full_name = new_name
        return label

    
    # no errors, just return the list
    if len(errors) == 0:
        return labels
    
    manually_check = input("Some labels are missing their names. Enter manually? [y/N] ")
    # anything other than an explicit yes = loop through and replace errors with blanks
    if manually_check.capitalize() != "Y":
        for i, label in enumerate(labels):
            if label.full_name == "Label_Error":
                labels[i].full_name = ""
        return labels
    
    # if we got here it means they want to manually update the broken ones
    for i, label in enumerate(labels):
        if i in errors:
            labels[i] = correctShippingLabel(label, i+1)

    return labels


# \fn     labels_Ripper
#   \brief  This function take the text extracted from a shipping label and identifies
#           the important information: the name, shipping address, and the city/state
#           and zip code. These three lines are return as a list.
#   \return <list>
#
def get_details_list_from_shipping_label(text) -> ShippingLabel:
    # Define some variables used
    label_keys = ["full_name", "addr_line1",
                  "addr_line2", "addr_line3", "addr_line4"]
    label_vals = dict.fromkeys(label_keys, "")
    addr = text.splitlines()

    # Setup our regex patterns
    name = re.compile(r"[^/]([A-Za-z].?\s?)+$")
    address = re.compile(
        r"^([C][/][O])\s+[A-Za-z]+\s+[A-Za-z]+\s+[A-Za-z]+|^[A-Za-z]?(\d+[A-Za-z]?)+\s([A-Za-z0-9]\s?)+")
    cityStZip = re.compile(r"([A-Za-z]\s?)+,?\s[A-Z]{2}\s\d{5}")

    try:
        # First replace any incorrect characters, these are errors of OCR
        addr = [w.replace("$", "S") for w in addr]
        addr = [w.replace("ยง", "S") for w in addr]
        # Next search for the name
        newAddr = list(filter(name.match, addr))
        # filter out any bad lines we don't want, since name pattern is less restrictive
        newAddr = list(filter(lambda x: "APTS" not in x, newAddr))
        newAddr = list(filter(lambda x: "APT" not in x, newAddr))
        newAddr = list(filter(lambda x: "TO:" not in x, newAddr))


        # Finally, search for address and city/state/zip
        newAddr.append(list(filter(address.match, addr))[0])
        newAddr.append(list(filter(cityStZip.match, addr))[0])

        if len(newAddr) != 3:
            print(newAddr)
            raise ValueError(len(newAddr))

        # If the lines were successfully found, save them and return
        label_vals["full_name"] = AddressUtil().format_address(newAddr[0])
        for i in range(1, len(newAddr)):
            temp = newAddr[i].rstrip()
            temp = temp.replace("-  ", "-")
            temp = temp.replace("- ", "-")
            label_vals[label_keys[i]] = temp
    except Exception:
        label_vals["full_name"] = "Label_Error"

    return ShippingLabel(page_num=0,
                         full_name=label_vals["full_name"],
                         addr_line1=label_vals["addr_line1"],
                         addr_line2=label_vals["addr_line2"],
                         addr_line3=label_vals["addr_line3"],
                         addr_line4=label_vals["addr_line4"])

#endregion


class AddressUtil:
    def is_address(self, potential_address: str) -> bool:
        return re.compile(r"^[A-Za-z]?(\d+[A-Za-z]?)+\s([A-Za-z0-9]\s?)+").match(potential_address)

    def format_address(self, s: str) -> str:
        return s.upper().replace(",", "")

    def city_state(self, data) -> Tuple[str, str, str]:
        comma = data.find(",")
        city = self.format_address(data[0:comma])
        t = data[comma:].strip(",").split()
        state = self.format_address(t[0])
        zip_code = t[1]
        return city, state, zip_code

    def strip_address(self, q) -> None:
        addr_keys = ["full_name", "first_name", "last_name", "company", "addr_line1",
                     "addr_line2", "addr_line3", "city", "state", "zip_code", "csz"]
        addr_book = dict.fromkeys(addr_keys, "")
        addr_index_1, addr_index_2, addr_index_3 = 1, 2, 3

        for i, j in enumerate(q):
            if i == 0:
                try:
                    t = j.split()
                    addr_book["first_name"] = self.format_address(t[0])
                    addr_book["last_name"] = self.format_address(
                        t[1 if len(t) == 2 else 2])
                    addr_book["full_name"] = self.format_address(j)
                except:
                    addr_book["full_name"] = self.format_address(j)
            if i == 1 and not(self.is_address(j)):
                addr_book["company"] = self.format_address(j)
                addr_index_1 += 1
                addr_index_2 += 1
                addr_index_3 += 1
            elif i == addr_index_1:
                addr_book["addr_line1"] = self.format_address(j)
            elif i == addr_index_2 and i != len(q) - 1:
                addr_book["addr_line2"] = self.format_address(j)
            elif i == addr_index_3 and i != len(q) - 1:
                addr_book["addr_line3"] = self.format_address(j)
            if i == len(q) - 1:
                csz = self.city_state(j)
                addr_book["city"], addr_book["state"], addr_book["zip_code"] = csz[0], csz[1], csz[2]
                addr_book["csz"] = " ".join(csz)
        return addr_book

#region Packing Slips

@dataclass
class PackingSlip:
    name: str
    addr: str
    city_state_zip: str
    reference_num: str
    page: int

def processBedBathPackingSlips(mode: Mode) -> List[PackingSlip]:
    # collect shipping info: this is for BedBath
    order_info_list = tabula.read_pdf(mode.slips_path, area = mode.order_scan, pages = 'all', pandas_options={'header': None})
    ship_to_list = tabula.read_pdf(mode.slips_path, area = mode.ship_scan, stream = False, pages = 'all', pandas_options={'header': None})

    result = []
    for (i, (order,ship)) in enumerate(zip(order_info_list, ship_to_list)):
        ship = ship[0]
        # process recipient inf

        name = ship[1]
        addr = ""
        city_state_zip = ""
        # includes secondary address (lot or apartment)
        if len(ship) == 5:
            addr = ship[2] + " " + ship[3]
            city_state_zip = ship[4]
        else:
            addr = ship[2]
            city_state_zip = ship[3]

        reference_num = order[1][0]
        result.append(PackingSlip(name=name, addr=addr, city_state_zip=city_state_zip, reference_num=reference_num, page=i))
    return result


def processHsnPackingSlips(mode: Mode) -> List[PackingSlip]:
    order_info_list = tabula.read_pdf(mode.slips_path, area = mode.order_scan, pages = 'all', pandas_options={'header': None})
    ship_to_list = tabula.read_pdf(mode.slips_path, area = mode.ship_scan, guess=False, pages = 'all', pandas_options={'header': None}, multiple_tables=True)
    
    slips = []
    for (i,(order, ship)) in enumerate(zip(order_info_list, ship_to_list)):
    
        # This will be in the form of "Package ID:<the number>"
        reference_number = order[0][0]
        # Getting a negative index of a list goes from the end. We want the last element
        reference_number = reference_number.split(":")[-1]
        
        ship=ship[0]

        slips.append(PackingSlip(name=ship[0], addr=ship[1], city_state_zip=[2], reference_num=reference_number, page=i))

    return slips    
    
def processTargetPackingSlips(mode: Mode) -> List[PackingSlip]:
    
    order_info_list = tabula.read_pdf(mode.slips_path, area = mode.order_scan, pages = 'all', pandas_options={'header': None})
    ship_to_list = tabula.read_pdf(mode.slips_path, area = mode.ship_scan, pages = 'all', pandas_options={'header': None})

    ret_list = []

    for (i, (order, ship)) in enumerate(zip(order_info_list, ship_to_list)):
        order = order[0]
        ship = ship[1]


        reference_num = str(order[0])
        reference_num = reference_num[4:]
        print(reference_num)
        name = ship[1]
        address = ship[2]
        city_state_zip = ship[4]

        ret_list.append(PackingSlip(name, address, city_state_zip, reference_num, page=i))



    return ret_list

def processBelkPackingSlips(mode: Mode) -> List[PackingSlip]:
    order_info_list = tabula.read_pdf(mode.slips_path, area=mode.order_scan, pages="all", pandas_options={"header": None})
    ship_to_list = tabula.read_pdf(mode.slips_path, area=mode.ship_scan, pages="all", pandas_options={"header": None})

    ret_list = []

    for (i, (order, ship)) in  enumerate(zip(order_info_list, ship_to_list)):
        ship = ship[0]

        name = ship[1]
        addr = ship[2]
        city_state_zip = ship[3]
        
        reference_num = ""

        ret_list.append(PackingSlip(name, addr, city_state_zip, reference_num, page=i))
    return ret_list

def processHibbettPackingSlips(mode: Mode) -> List[PackingSlip]:
    order_info_list = tabula.read_pdf(mode.slips_path, area = mode.order_scan, pages = 'all', pandas_options={'header': None})
    ship_to_list = tabula.read_pdf(mode.slips_path, area = mode.ship_scan, pages = 'all', pandas_options={'header': None})

    ret_list = []

    for (i, (order, ship)) in  enumerate(zip(order_info_list, ship_to_list)):
        reference_num = order[0][0]
        
        name = ship[2][0]
        addr = ship[2][1]
        city_state_zip = ship[2][2]
        
        ret_list.append(PackingSlip(name, addr, city_state_zip, reference_num, page=i))

    return ret_list

#endregion

def processAndSortPackingSlips(mode: Mode) -> List[PackingSlip]:
    label_output, label_errors = parseShippingLabel(mode)
    labels: List[PackingSlip] = checkShippingLabels(label_output, label_errors)

    slips: List[PackingSlip] = []

    if mode.name == Store.Target.name:
        slips = processTargetPackingSlips(mode)
    elif mode.name == Store.Belk.name:
        slips = processBelkPackingSlips(mode)
    elif mode.name == Store.GSI.name:
        pass
    elif mode.name == Store.HSN.name:
        slips = processHsnPackingSlips(mode)
    elif mode.name == Store.Hibbett.name:
        slips = processHibbettPackingSlips(mode)
    # otherwise bedbath
    else:
        slips = processBedBathPackingSlips(mode)

    # There's a bug where sometimes tabula will see multiple spaces and skip them,
    # Pushing the first and last name together. Because of this, we're going to 
    # Remove all spaces before comparing. For consistency, make everything uppercase.
    for (label, slip) in zip(labels,slips):
        label.full_name = label.full_name.replace(" ", "")
        slip.name = slip.name.replace(" ", "").upper()


    def get_slip_key(slip: PackingSlip) -> int:
        names = [label.full_name for label in labels]
        order_numbers = [label.reference_num for label in labels]

        if mode.name == "Target":
            # Get the proper subsection of the target number that will match the number on the label
            shortened_reference_num = slip.reference_num[-6:-1]
            if shortened_reference_num in order_numbers:
                return order_numbers.index(shortened_reference_num)
        
        else:
            if slip.reference_num in order_numbers:
                return order_numbers.index(slip.reference_num)

        if slip.name.upper() in names:
            return names.index(slip.name)
        
        print(f"SLIP WITH NAME {slip.name} AND NUMBER {slip.reference_num} NOT FOUND IN LABELS")
        # TODO: Fail gracefully
        return -1
        

    ordered = sorted(slips, key=get_slip_key)

    return ordered

def exportPackingSlips(mode: Mode, slips: List[PackingSlip]):
    unordered_file = open(mode.slips_path, "rb")
    unordered_pdf = PdfFileReader(unordered_file)

    writer = PdfFileWriter()
    with open(mode.name + "_reordered.pdf", "wb") as file:
        for slip in slips:
            page = unordered_pdf.getPage(slip.page)
            writer.addPage(page)
        writer.write(file)
    
    unordered_file.close()
    print(f"saved sorted packing slips to {mode.name}_reordered.pdf")

def read_reference_number_ups(image: Image, retailer_name: str) -> str:
    # expected coords for reference number
    coords = (0, 2820, 700, 2970)
    cropped_image = image.crop(coords).convert("L")

    text = str(pytesseract.image_to_string(cropped_image))

    fullRefNoSplit = text.split("Trx Ref No.: ")
    partialRefNoSplit = text.split("No")

    if (retailer_name in [Store.HSN.name, Store.GSI.name]) and (len(fullRefNoSplit) > 1):
        text = fullRefNoSplit[1]
        text = text.split("\n")[0].strip()
        text = text.split(" ")[0]
        return text
    elif len(fullRefNoSplit) > 1:
        text = fullRefNoSplit[1]
        text = text.split("\n")[0]
        text = text.replace(" ", "")
        return text[-6:-1]
    elif len(partialRefNoSplit) > 1:
        text = partialRefNoSplit[1]
        text = text.split("\n")[0]
        text = text.replace(" ", "")
        return text[-6:-1]
    else:
        return "N/A"


def read_reference_number_fedex(image: Image) -> str:
    # expected coords for reference number
    coords = (792, 842, 1122, 912)
    cropped_image = image.crop(coords).convert("L")

    text = str(pytesseract.image_to_string(cropped_image))

    fullRefNoSplit = text.split("REF:")

    if len(fullRefNoSplit) > 1:
        text = fullRefNoSplit[1]
        text = text.split("\n")[0]
        text = text.replace(" ", "")
        text = text[-6:-1]
    else:
        text = "N/A"
    return text


def Main():
    # Argparse section

    argParseDescription = ('Packing Slip and Label combination tool. Takes a PDF for packing slips and a PDF for labels, '
                           'then outputs a new PDF for packing slips that is in the order of the labels.')

    parser = argparse.ArgumentParser(description=argParseDescription)
    parser.add_argument('-s', required=False, dest='packingSlips',
                        metavar='slips', help='The PDF for the packing slips')
    parser.add_argument('-l', required=False, dest='shippingLabels',
                        metavar='labels', help='The PDF for the shipping labels')
    parser.add_argument('-o', default='slips_reordered.pdf')
    # TODO: add option for selecting store



    args = parser.parse_args()

    if args.packingSlips == None:
        args.packingSlips = input("Enter path to packing slips: ")
    
    if args.shippingLabels == None:
        args.shippingLabels = input("Enter path to shipping labels: ")

    # Main section

    mode = get_mode(args.packingSlips, args.shippingLabels)

    sorted_slips = processAndSortPackingSlips(mode)
    exportPackingSlips(mode, sorted_slips)

    


if __name__ == "__main__":
    Main()
