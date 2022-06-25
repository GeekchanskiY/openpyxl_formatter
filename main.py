from openpyxl import Workbook, load_workbook, worksheet


def filetype1(wb: Workbook, *args, **kwargs) -> Workbook:
    """
        Makes new worksheet for filetype 1 and returns 1 if created successfully

        Args:
               wb (Workbook): original workbook

        Returns:
            Workbook: new workbook with formatted data
    """
    pass


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
    main("test2.xlsx")


