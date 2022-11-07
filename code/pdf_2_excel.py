# This project involves creating an excel/csv file from a pdf file given to us by a client Daniel Borg
# The location of the pdf file is here: pdf_2_excel\data\BMF-Handbook-2022-Fullbook (1).pdf
# We have made a word version of the document here: pdf_2_excel\data\BMF-Handbook-2022-Fullbook (1).docx
# This is the version we will be using to extract the text data


from docx import Document
from itertools import chain
from docx.shared import Pt
import pandas as pd
import re


def make_vendors_indicies(vendors_list: list, full_text: list) -> list:
    """
    Makes a list of vendor's indicies to indicate the starting and stopping points
    in the overall full_text list.

    """
    indicies = []
    for vendor in vendors_list:
        try:
            indicies.append(full_text.index(vendor))
        except Exception as e:
            print(e)

    return indicies


def make_nested_list(full_text: list, indicies: list) -> list:
    """
    Makes a nested list to indicate the start and stop of each vendor and
    their relevant information.

    """
    nested_list = []
    for index, value in enumerate(indicies):
        if index + 1 < len(indicies) and index - 1 >= 0:
            previous = int(indicies[index - 1])
            current = int(value)

            nested_list.append(full_text[previous:current])

    return nested_list


def remove_email_adddress(_str: str) -> str:
    regex = r"\S*@\S*\s?"
    subst = ""

    return re.sub(regex, subst, _str, 0)


def find_branches(vendor_list: list) -> list:
    """
    Evaluates to see if an individual vendor list contains branches.

    """
    branches_list = []
    try:
        vendor_list.index("Branches")
        branches_list = vendor_list[
            vendor_list.index("Branches") + 1 : len(vendor_list)
        ]
        branches_list = list(filter(None, branches_list))

        telephone_numbers = [i for i in branches_list if i.startswith("T ")]

        for number in telephone_numbers:
            branches_list[branches_list.index(number) - 1] = (
                str(branches_list[branches_list.index(number) - 1])
                + " "
                + str(branches_list[branches_list.index(number)])
            )
            del branches_list[branches_list.index(number)]

    except ValueError:
        print("Branches not in list")

    return branches_list


def get_postcode(address_string: str) -> str:
    postcode = " ".join(address_string.split(" ")[-2:])

    if any(str.isdigit(c) for c in postcode):
        return postcode
    else:
        return ""


def make_vendor_dict(vendor_list: list) -> dict:
    """
    Makes a dictionary of vendors and their information, for example:
    {
        "merchant",
        "address",
        "telephone",
        "fax"
        "email",
        "website,"
        "core_activity"
    }

    """

    vendor_dict = {}

    # There should always be a merchant name
    vendor_dict["merchant"] = vendor_list[0]
    vendor_list = vendor_list[1:]

    # There should always be an address
    address_string = vendor_list[0]

    # calculate the postcode
    postcode = get_postcode(address_string)
    address_string = address_string.replace(postcode, "").strip()

    vendor_dict["address"] = address_string
    vendor_dict["postcode"] = postcode

    vendor_list = vendor_list[1:]

    # other fields
    others = {
        "T ": "telephone",
        "F ": "fax",
        "@": "email",
        "W www": "website",
        "Core Activity": "core_activity",
    }

    others_list = list(others.keys())

    for other in others_list:
        _index = [i for i, s in enumerate(vendor_list) if other in s]

        if len(_index) != 0:
            vendor_dict[others.get(other)] = vendor_list[_index[0]]
        elif other in vendor_dict.values():
            continue
        else:
            vendor_dict[others.get(other)] = None

    # Clean the email and website keys
    if vendor_dict["email"]:
        vendor_dict["email"] = vendor_dict["email"].split(" ")[0]

    if vendor_dict["website"]:
        vendor_dict["website"] = vendor_dict["website"].split(" ")[-1]

    if vendor_dict["core_activity"]:
        _str = vendor_dict["core_activity"]
        result = remove_email_adddress(_str)

        if result:
            vendor_dict["core_activity"] = result

    branches_list = find_branches(vendor_list)
    if len(branches_list) != 0:
        vendor_dict["branches"] = True
        vendor_dict["number_of_branches"] = len(branches_list)
    else:
        vendor_dict["branches"] = False
        vendor_dict["number_of_branches"] = 0

    return vendor_dict


def dataframe_per_vendor(vendor_dict: dict) -> pd.DataFrame:
    df = pd.DataFrame.from_dict(data=vendor_dict, orient="index").T

    # clean dataframe
    df["telephone"] = df["telephone"].str.replace("T ", "")
    df["fax"] = df["fax"].str.replace("F ", "")

    df["core_activity"] = df["core_activity"].str.replace("Core Activity ", "")

    return df


def dataframe_branches(vendor_list: list) -> pd.DataFrame:
    """
    Creates a dataframe for the merchant and its branches.

    """
    df_branches = pd.DataFrame(columns=["merchant", "branch", "telephone"])

    branches_list = find_branches(vendor_list)

    if len(branches_list) != 0:
        df_branches["branch"] = branches_list
        df_branches["merchant"] = vendor_list[0]

        df_branches[["branch", "telephone"]] = df_branches["branch"].str.split(
            " T ", expand=True
        )

        # Retrieve the address and calculate the postcode
        address_string = vendor_list[1]
        postcode = get_postcode(address_string)

        address_string = address_string.replace(postcode, "").strip()

        df_branches["address"] = address_string
        df_branches["postcode"] = postcode

    return df_branches


def all_vendors_dataframe(nested_vendors_list: list) -> pd.DataFrame:
    df_list = []
    for vendor_list in nested_vendors_list:
        try:
            df = dataframe_per_vendor(make_vendor_dict(vendor_list))
            df_list.append(df)
        except Exception as e:
            print(e)

    return pd.concat(df_list, axis=0)


def all_branches_dataframe(nested_vendors_list: list) -> pd.DataFrame:
    df_list = []
    for vendor_list in nested_vendors_list:
        try:
            df_list.append(dataframe_branches(vendor_list))
        except Exception as e:
            print(e)

    return pd.concat(df_list, axis=0)


def main():
    # Assume that the document is in Google Drive
    doc = Document("/content/drive/MyDrive/Upwork/pdf_2_excel/BMF-Handbook-2022-Fullbook (1).docx")

    runs = list(chain.from_iterable(list(p.runs) for p in doc.paragraphs))

    full_text = [r.text for r in runs]

    # The list of vendors will have a bold Arial font with a size of 9
    bold_text = [
        r.text for r in runs if r.font.name == "Arial" and r.font.size == Pt(9)
    ]

    starting_vendor = "A F Akehurst & Sons Ltd"
    ending_vendor = "Yorkshire Timber and Builders Merchant"

    # Let us first get a list of all the vendors
    vendors_list = bold_text[
        bold_text.index(starting_vendor) : bold_text.index(ending_vendor)
    ]
    indicies = make_vendors_indicies(vendors_list, full_text)

    nested_vendors_list = make_nested_list(full_text, indicies)

    # Small correction in nested_list - RS Industrial Services (Tyne & Wear)
    correction_list = nested_vendors_list[329] + nested_vendors_list[330]
    correction_list[0] = "RS Industrial Services (Tyne & Wear) Ltd"
    del correction_list[1]

    nested_vendors_list.append(correction_list)

    df_all = all_vendors_dataframe(nested_vendors_list)
    df_branches_all = all_branches_dataframe(nested_vendors_list)

    writer = pd.ExcelWriter(
        "/content/drive/MyDrive/Upwork/pdf_2_excel/BMF-Handbook-2022-Fullbook.xlsx",
        engine="xlsxwriter",
    )

    df_all.to_excel(writer, sheet_name="merchant_data")
    df_branches_all.to_excel(writer, sheet_name="branches_data")

    writer.close()


if __name__ == "__main__":
    main()
