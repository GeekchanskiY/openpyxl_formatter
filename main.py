import re
import phonenumbers
from phonenumbers.phonenumberutil import NumberParseException
from openpyxl import Workbook, load_workbook, worksheet


def alternative_select_number_from_str(string: str) -> str:

    # deleting all unnecessary symbols

    string = re.sub("[а-я]", "", string)
    string = re.sub("[А-Я]", "", string)
    string = re.sub("[a-z]", "", string)
    string = re.sub("[A-Z]", "", string)

    # deleting spaces
    string = string.strip()

    # for (xxx) case
    if string[0] == "(" and string[-1] == ")":
        string = string[1:-1]

    # flag for getting code
    code_starts = False
    code = ""
    this_phone = ""
    for i in range(0, len(string)):

        # getting code
        if string[i] == "(":
            code_starts = True
        elif string[i] == ")":
            code_starts = False

        # if numeric value in string pos, writing this number to code or phone
        if re.match(r"\d", string[i]):
            if code_starts is True:
                code += string[i]
            else:
                this_phone += string[i]

    # concat code with phone
    this_phone = code + this_phone

    return this_phone


def filetype1(wb: Workbook) -> Workbook:
    """
        Makes new worksheet for filetype 1 and returns 1 if created successfully

        Args:
               wb (Workbook): original workbook

        Returns:
            Workbook: new Workbook with formatted data
    """

    output_data: list[str, list] = []
    ws: worksheet = wb.active
    for row in ws.values:
        numbers: list[str] = []
        raw_numbers: list[str] = []

        # First filter with general regex

        for m in re.finditer(r"(?:(?:8|\+7)[\- ]?)?(?:\(?\d{3}\)?[\- ]?)?[\d\- ]{7,10}", str(row[2])):

            # Second filter with phonenumbers

            for match in phonenumbers.PhoneNumberMatcher(m.string, "RU"):
                raw_numbers.append(match.raw_string)
                numbers.append(str(match.number.national_number))

            # deleting everything phonenumbers found

            temp_str: str = m.string

            for number in raw_numbers:
                temp_str = temp_str.replace(number, "")

            # Third custom filter

            if len(str(re.sub(r"\D", "", temp_str))) >= 7:
                for t_s in temp_str.split(","):
                    if t_s != "" and len(re.sub("\D", "", t_s)) >= 7:
                        numbers.append(alternative_select_number_from_str(t_s))

        # deleting duplicates if exists

        numbers = list(set(numbers))

        print(numbers)
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
        new_workbook = filetype1(wb)
    new_workbook.save("output.xlsx")


if __name__ == '__main__':
    main("test1.xlsx")
