import unittest
from docx import Document
from itertools import chain
from docx.shared import Pt
import pandas as pd


def get_vendors_list() -> list:
        doc = Document("/content/drive/MyDrive/Upwork/pdf_2_excel/BMF-Handbook-2022-Fullbook (1).docx")
        runs = list(chain.from_iterable(list(p.runs) for p in doc.paragraphs))

        bold_text = [r.text for r in runs if r.font.name == "Arial" and r.font.size == Pt(9)]

        starting_vendor = "A F Akehurst & Sons Ltd"
        ending_vendor = "Yorkshire Timber and Builders Merchant"

        return bold_text[bold_text.index(starting_vendor):bold_text.index(ending_vendor)]


def make_merchants_branches_dict() -> dict:
    filepath = "/content/drive/MyDrive/Upwork/pdf_2_excel/BMF-Handbook-2022-Fullbook.xlsx"
    df_merchants = pd.read_excel(filepath, sheet_name="merchant_data")

    merchants = df_merchants["merchant"]
    num_branches = df_merchants["number_of_branches"]

    return dict(zip(merchants, num_branches))


class TestOutput(unittest.TestCase):
    def test_total_vendors(self):
        filepath = "/content/drive/MyDrive/Upwork/pdf_2_excel/BMF-Handbook-2022-Fullbook.xlsx"
        df_merchants = pd.read_excel(filepath, sheet_name="merchant_data")

        vendors_list = get_vendors_list()

        percentage = len(df_merchants["merchant"].unique().tolist()) / len(vendors_list)

        self.assertEqual(percentage, 0.9929906542056075)
    
    def test_merchant_num_branches(self):
        all_merchants_branches_dict = make_merchants_branches_dict()

        merchants_branches_dict = {
            "AE Spink Ltd": 6,
            "Leekes Ltd": 6,
            "WJ Lewis BM Ltd": 2,
            "Turnbull & Co Ltd": 8,
            "Longwater Construction Supplies Ltd": 3
        }

        branches_output = []
        for merchant in list(merchants_branches_dict.keys()):
            branches_output.append(all_merchants_branches_dict.get(merchant))
        
        self.assertEqual(branches_output, list(merchants_branches_dict.values()))

if __name__ == "__main__":
    unittest.main()