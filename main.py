import re

from openpyxl import Workbook, load_workbook, worksheet


def select_number_from_str(string: str) -> list[str]:
    all_phones = []
    for i in range(0, len(string)):
        if re.match(r"[0-9]", string[i]):
            print("YAY")


def filetype1(wb: Workbook, *args, **kwargs) -> Workbook:
    """
        Makes new worksheet for filetype 1 and returns 1 if created successfully

        Args:
               wb (Workbook): original workbook

        Returns:
            Workbook: new workbook with formatted data
    """

    output_data: list[str, list] = []
    ws: worksheet = wb.active
    for row in ws.values:
        for m in re.finditer(r"(?:(?:8|\+7)[\- ]?)?(?:\(?\d{3}\)?[\- ]?)?[\d\- ]{7,10}", str(row[1])):
            select_number_from_str(m.string)
    print(output_data)
    return wb


def filetype2(wb: Workbook, *args, **kwargs) -> Workbook:
    """
        Makes new worksheet for filetype 2 and returns new workbook if created successfully

    Args:
        wb (Workbook): original workbook

    Returns:
        Workbook: new workbook with formatted data
    """
    pass


def main(filename: str, *args, **kwargs):
    wb = load_workbook(filename)
    ws = wb.active
    row_len = len(list(ws.iter_rows())[0])
    if row_len == 3:
        new_workbook = filetype1(wb)
    else:
        new_workbook = filetype2(wb)
    new_workbook.save("output.xlsx")


if __name__ == '__main__':
    main("test1.xlsx")


