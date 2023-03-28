import re, sys, os 
from bs4 import BeautifulSoup
sys.path.append(os.path.join(os.path.dirname(__file__), ".."))
from office_helper.utils.excel_op import ExcelOperator



def read_data_from_html_and_excel(start, end, col_name, excel_path, html_path):
    pattern_with_dash = r"^\w{4}\s*-\s*"
    pattern_with_slash = r"^\w{4}\s*/\s*\w{4}"
    ecs_number_to_sheet = {}
    cell_range = f"{col_name}{start}:{col_name}{end}"
    with ExcelOperator(excel_path) as excel:
        ecs_number_data = excel.read_cell_value(0, cell_range)
    for ecs in ecs_number_data:
        start += 1
        if ecs in ecs_number_to_sheet:
            continue
        if match := re.search(pattern_with_dash, ecs):
            ecs_number_to_sheet[match.group()[:4]] = f"{chr(ord(col_name) + 2)}{start-2}"
        elif match := re.search(pattern_with_slash, ecs):
            ecs_number_to_sheet[match.group()] = f"{chr(ord(col_name) + 2)}{start-2}"
        else:
            print(f"No pattern matching with '{ecs}'")
    
    with open(html_path, encoding='utf-8') as d:
        soup = BeautifulSoup(d, features="html.parser")

    all_ecus = soup.find_all("p", "default_style_ecu")
    return (ecs_number_to_sheet, all_ecus)

if __name__ == '__main__':
    # Here goes the config
    start = 160
    end = 225
    col_name = "E"
    XL_PATH = r'C:\Users\Xac\Documents\WeChat Files\wxid_7itvgu9kd4tk22\FileStorage\File\2023-03\test_write.xlsm'
    HTML_PATH = r"C:\Users\Xac\Documents\WeChat Files\wxid_7itvgu9kd4tk22\FileStorage\File\2023-03\test.html"

    (ecs_number_to_sheet, all_ecu_elements_in_html) = read_data_from_html_and_excel(start, end, col_name, XL_PATH, HTML_PATH)

    with ExcelOperator(XL_PATH) as excel:
        for nr in ecs_number_to_sheet.keys():
            if any(nr in (ecu := i).contents[0] for i in all_ecu_elements_in_html):
                cur_ecu = ecu
            else:
                continue
            if any((parent := i).name == 'table' for i in cur_ecu.parents):
                cur_ecu_table = parent
            if any((sibling := i).name == 'span' for i in cur_ecu_table.next_siblings):
                cur_ecu_span = sibling
            templist = []
            for index, tr in enumerate(cur_ecu_span.find_all("tr")):
                if index == 4:
                    fill_value = [[templist[0] + templist[1]], [templist[2] + templist[3]]]
                    print(ecs_number_to_sheet[nr], fill_value)
                    excel.write_cell_value(0, ecs_number_to_sheet[nr], fill_value)
                    templist = []
                    break
                infolist = [td for td in tr.find_all("td")]
                if len(infolist) >= 2:
                    templist.append(infolist[1].contents[0])
                else:
                    templist.append("")