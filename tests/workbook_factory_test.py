import unittest
import sys
import os

sys.path.append(os.path.join(os.getcwd(), "src"))

from rfexcel.factory.workbook_factory import WorkbookFactory

class MyTestCase(unittest.TestCase):
    def test_factory_check(self):
        print("=== Starting Factory Test ===\n")
        factory = WorkbookFactory()

        # Test 1: XLSX in Edit Mode (Read/Write)
        print("--- Test 1: Creating Edit Workbook (.xlsx, read_only=False) ---")
        wb_edit = factory.create_workbook(path="data.xlsx", read_only=False)
        wb_edit.print()
        print("")  # Newline

        # Test 2: XLSX in Stream Mode (Read Only)
        print("--- Test 2: Creating Stream Workbook (.xlsx, read_only=True) ---")
        wb_stream = factory.create_workbook(path="data.xlsx", read_only=True)
        wb_stream.print()
        print("")

        # Test 3: Invalid/Unknown File
        print("--- Test 3: Creating Invalid Workbook (.txt) ---")
        wb_invalid = factory.create_workbook(path="notes.txt")
        wb_invalid.print()
        print("\n=== Test Complete ===")


if __name__ == '__main__':
    unittest.main()