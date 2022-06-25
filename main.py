import re
import phonenumbers
from phonenumbers.phonenumberutil import NumberParseException
from openpyxl import Workbook, load_workbook, worksheet


def check_number(string: str) -> int:
    """
        Checks if number matches one of these cases:
            Case 1: code starts with 9
            Case 2: starts with 6 & 9
        and returns if matches

    Args:
        string(str): formatted number str (length will be checked dynamically)
    Returns:
        bool: 0 - does not match, 1 - case 1, 2 - case 2
    """
    if len(string) == 10:
        if string[0] == "9":
            return 1
        elif string[3] == "6":
            return 2
        elif string[3] == "9":
            return 2
    elif len(string) == 7:
        if string[0] == "6":
            return 2
        elif string[0] == "9":
            return 2
    return 0


def alternative_select_number_from_str(string: str) -> str:
    """

    Alternative string formatting for all specific cases

    Args:
         string(str): raw number str
    Returns:
        str: formatted string
    """
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

    output_data: list[dict] = []
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

        for m in re.finditer(r"(?:(?:8|\+7)[\- ]?)?(?:\(?\d{3}\)?[\- ]?)?[\d\- ]{7,10}", str(row[1])):

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
        for n in numbers:
            res = check_number(n)
            if res is 1:
                output_data.append({"name": str(row[0]), "phone": n, "comment": str(row[2])})
            elif res is 2:
                if len(n) == 7:
                    n = "7812" + n
                    output_data.append({"name": str(row[0]), "phone": n, "comment": str(row[2])})

    return wb


def filetype2(wb: Workbook) -> Workbook:
    """
        Makes new worksheet for filetype 2 and returns new workbook if created successfully

    Args:
        wb (Workbook): original workbook

    Returns:
        Workbook: new workbook with formatted data
    """
    pass


def main(filename: str):
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
