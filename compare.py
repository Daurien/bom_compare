import os
import pprint
from deepdiff import DeepDiff
import pandas as pd
import numpy as np
from datetime import datetime
import warnings
import re
import openpyxl
import argparse
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.comments import Comment
from openpyxl.packaging.custom import (
    BoolProperty,
    DateTimeProperty,
    StringProperty,
    CustomPropertyList,
    IntProperty
)


def bom_excel_to_dictionary(filePath: str):

    # read the excel BOM file ignoring the first row and convert it to a numpy array
    with warnings.catch_warnings(action="ignore"):
        file_type = get_file_type(filePath)
        columns_to_find = ['Level', 'Item', 'Description', 'Quantity', 'Supply Type', 'Supplier_Type']
        if file_type == 'ORACLE':
            skip = find_table_origin_line_number(filePath)-1
            BOM = pd.read_excel(filePath, dtype={'Level': int, 'Item': str, 'Description': str,
                                'Quantity': int, 'Supply Type': str}, skiprows=skip)
            columns_to_keep = [element for element in columns_to_find if element in BOM.keys()]
            BOM = BOM[columns_to_keep].to_numpy()
        elif file_type == 'CREO':
            skip = find_table_origin_line_number(filePath, 'Import_Creo')-1
            BOM = pd.read_excel(filePath, sheet_name='Import_Creo', dtype={'Level': int, 'Item': str, 'Description': str,
                                'Quantity': int, 'Supply Type': str}, skiprows=skip)
            columns_to_keep = [element for element in columns_to_find if element in BOM.keys()]
            BOM = BOM[columns_to_keep].to_numpy()

    # # unnify dat format to string
    # for row in BOM:
    #     if isinstance(row[10], datetime):
    #         row[10] = row[10].strftime('%m/%d/%Y')
    max_depth = max(BOM[:, 0])
    bom_length = len(BOM)
    last_row = bom_length - 1

    # # delete all useless columns
    # BOM = np.delete(BOM, [max_depth+6], 1)
    # BOM = np.delete(BOM, np.s_[max_depth+7:], 1)

    contents = []
    prev_contents = []
    current_content = []
    reading = False
    final_bom = {}

    for col in range(max_depth, 0, -1):
        # for deepest level
        if col == max_depth and col != 1:
            for row in range(bom_length):
                if BOM[row, 0] == col:
                    reading = True
                    current_content.append(append_row(BOM[row]))

                if row == last_row or reading:
                    reading = False
                    contents.append(current_content)
                    current_content = []

            prev_contents = contents
            contents = []

        # for intermediate levels
        elif col > 1:
            for row in range(bom_length):
                if BOM[row, 0] == col:
                    reading = True
                    try:
                        if BOM[row+1, 0] == col+1:  # if current row contains sub-level items
                            current_content.append(append_row(BOM[row], prev_contents.pop(0)))
                        else:  # if current does not contains sub-level items
                            current_content.append(append_row(BOM[row]))
                    except IndexError:  # if we are reading last row
                        current_content.append(append_row(BOM[row]))

                if row == last_row or (reading and BOM[row, 0] < col):
                    reading = False
                    contents.append(current_content)
                    current_content = []

            prev_contents = contents
            contents = []

        # for top level
        else:
            for row in range(bom_length):
                if BOM[row, 0] == col and row != last_row:
                    try:
                        if BOM[row+1, 0] == col+1:  # if current row contains sub-level items
                            current_content.append(append_row(BOM[row], prev_contents.pop(0)))
                        else:  # if current does not contains sub-level items
                            current_content.append(append_row(BOM[row]))
                    except IndexError:  # if we are reading last row
                        current_content.append(append_row(BOM[row]))

                elif row == last_row:
                    final_bom = array_to_dict(current_content)
                    # final_bom = {'BOM''Item name': 'QSMA', 'content': array_to_dict(current_content)}

    return final_bom, max_depth


def get_file_type(filePath: str):
    workbook = openpyxl.load_workbook(filePath)
    sheet_names = workbook.sheetnames

    if 'Bom' in sheet_names and 'Import_Creo' in sheet_names:
        return 'CREO'
    else:
        return 'ORACLE'


def array_to_dict(array):
    res = {}

    for row in array:
        res[row['Item name']] = row

    return res


def find_table_origin_line_number(file_path, sheet_name=None):
    # Load the workbook and select the active worksheet
    workbook = openpyxl.load_workbook(file_path)
    if sheet_name:
        worksheet = workbook[sheet_name]
    else:
        worksheet = workbook.active

    # Iterate through the tables to find the origin
    for table in worksheet.tables.values():
        origin_cell = table.ref.split(':')[0]
        # Extract the row number from the cell reference
        row_number = int(''.join(filter(str.isdigit, origin_cell)))

        return row_number

    return None


def append_row(row, content=[]):

    # search pattern like '*_02' in name to extract revision
    item_name = str(row[2])
    pattern = r'_(\d{2})$'
    match = re.search(pattern, item_name)
    if match:
        revision = match.group(1)
        item_name = item_name[:-3]
    else:
        revision = 'xx'

    result = {
        'Item name': item_name,
        'Revision': revision,
        'Description': row[2],
        'Quantity': row[3],
        # 'FromDate': row[max_depth+5],
        'SupplyType': row[4],
        'Depth': row[0]
    }

    if content:
        result['content'] = array_to_dict(content)

    return result


def append_to_dict(keys: list, bom_content: dict, modify_type: str, initial_dict: dict = {}):
    """_summary_

    Args:
        keys (list): list of key string ['34410718','34410662','34411697']
        bom_content (dict): nom generated with bom_excel_to_dictionary
        modify_type (dict): {'type' : 'ADDED'}, {'type' : 'REMOVED'} or {'type' : 'CHANGED', value_changed:'Revision' 'new_value': '03', 'old_value': '02'}
        initial_dict (dict, optional): _description_. Defaults to {}.

    Returns:
        _type_: _description_
    """

    # Initialize a variable to hold the current level of the dictionary
    current_level = initial_dict

    # Iterate through the keys to create the nested structure
    for i, key in enumerate(keys):
        if key not in current_level:
            content = get_content(bom_content, keys[:i+1])
            if i == len(keys) - 1:
                mf = [modify_type]
            else:
                mf = [{'type': f'Item {modify_type['type']} inside'}]
            content = {'Description':  content['Description'], 'Revision': content['Revision'],
                       'Quantity':  content['Quantity'], 'SupplyType':  content['SupplyType'], 'ModifyType': mf}
            current_level[key] = {'content': content}
        else:
            if i == len(keys) - 1:
                if modify_type not in current_level[key]['content']['ModifyType']:
                    current_level[key]['content']['ModifyType'].append(modify_type)
            else:
                new_item = {'type': f'Item {modify_type["type"]} inside'}
                if new_item not in current_level[key]['content']['ModifyType']:
                    current_level[key]['content']['ModifyType'].append(new_item)

        current_level = current_level[key]

    return initial_dict


def dict_to_table(d, maxDepth, level=1, result=None):
    """Transforms a dictionary BOM into a table like extracted from Oracle DB tables

    Args:
        d (dict): dictionnary to convert

    Returns:
        result: dict converted as list wiht level on 1rst row
    """
    if result is None:
        result = []
    for key, value in d.items():
        if key != 'content':
            result.append((level, *["" if i != level else level for i in range(1, maxDepth+1)],
                          key, *(d[key]['content'].values())))
            if isinstance(value, dict):
                dict_to_table(value, maxDepth, level=level + 1, result=result)
    return result


def get_content(bom: dict, keys: list):
    """get content of a dict bom from a string key

    Args:
        bom (dict): bom to analyse
        keys (list): list of key string ['34410718','34410662','34411697']
    """
    current_level = bom
    for i, key in enumerate(keys):
        if i == 0:
            current_level = current_level[key]
        else:
            current_level = current_level['content'][key]

    return current_level


def save_df_to_excel(result: pd.DataFrame, max_depth: int):

    # Create a workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active

    # Convert DataFrame to rows and append to worksheet
    rows = dataframe_to_rows(result.drop("ModifyType", axis=1), index=False, header=True)
    header = next(rows)
    ws.append(header)

    for r in rows:
        # Convert numerical strings back to numbers and dict to strings
        r = [int(x) if isinstance(x, str) and x.isdigit() else (str(x) if isinstance(x, list) else x) for x in r]
        ws.append(r)

    # Left-align all cells in the worksheet
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='left')

    # Set column widths to match the first file
    row_widths = [5] + [3] * max_depth + [31, 41, 8, 10, 19, 20]
    nb_columns = len(result.columns) - 1
    for i, width in enumerate(row_widths):
        ws.column_dimensions[chr(ord('A')+i)].width = width
        ws.column_dimensions[chr(ord('A')+i)].alignment = Alignment(horizontal='left')

    # Define table to written data
    tab = Table(displayName="Table1", ref=f'A1:{chr(ord('A')+result.shape[1]-2)}{result.shape[0]+1}')

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tab.tableStyleInfo = style

    # Add the table to the worksheet
    ws.add_table(tab)

    # Define the fill style with the desired background color
    fill_1 = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    fill_2 = PatternFill(start_color="DBDBDB", end_color="DBDBDB", fill_type="solid")
    fill_3 = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    fill_4 = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    fill_5 = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")

    fill = [fill_1, fill_2, fill_3, fill_4, fill_5]

    # Apply the fill style to the specified range
    for row in range(1, result.shape[0]+1):
        for col in range(1, nb_columns):
            ws.cell(row=row+1, column=col+1).fill = fill[col-1 if col <=
                                                         result.loc[row-1].at["Level"] else result.loc[row-1].at["Level"]-1]

    # Outline the specified range with the specified brush
    def outilne_range(min_row, min_col, max_row, max_col, brush):
        for row in range(min_row, max_row+1):
            ws.cell(row=row, column=min_col).border = Border(left=brush, bottom=(
                brush if row == max_row else None), top=(brush if row == min_row else None))
            ws.cell(row=row, column=max_col).border = Border(right=brush, bottom=(
                brush if row == max_row else None), top=(brush if row == min_row else None))
        for col in range(min_col+1, max_col):
            if max_row != min_row:
                ws.cell(row=min_row, column=col).border = Border(top=brush)
                ws.cell(row=max_row, column=col).border = Border(bottom=brush)
            else:
                ws.cell(row=min_row, column=col).border = Border(top=brush, bottom=brush)

    green_brush = Side(border_style="medium", color="2EB82E")
    red_brush = Side(border_style="medium", color="FF0000")

    bom1 = "bom1"
    bom2 = "bom2"

    for index, row in result.iterrows():
        if {'type': 'ADDED'} in row['ModifyType']:
            # Green outline added items
            outilne_range(index+2, row['Level']+1, index+2, nb_columns, green_brush)
            ws[f'{chr(ord('A')+max_depth+1)}{index+2}'].comment = Comment(f'Item {result.iloc[index]
                                                                                  ["Item"]} was added in {bom2} (not present in {bom1})', "Automatically Generated")

        elif {'type': 'REMOVED'} in row['ModifyType']:
            # Red outline removed items
            outilne_range(index+2, row['Level']+1, index+2, nb_columns, red_brush)
            ws[f'{chr(ord('A')+max_depth+1)}{index+2}'].comment = Comment(f'Item {result.iloc[index]
                                                                                  ["Item"]} was removed in {bom2} (present in {bom1})', "Automatically Generated")
        else:
            for modif in row['ModifyType']:
                if modif['type'] == 'CHANGED':
                    ws[f'{chr(ord('A')+result.columns.get_loc(modif['changed_value']))}{index+2}'].comment = Comment(f'{modif['changed_value']
                                                                                                                        } of item {result.iloc[index]["Item"]} has changed : {modif['old_value']} -> {modif['new_value']}', "Automatically Generated")

        # else:
        #     for modif in row['ModifyType']:

    # SETTING CONFIDENTIALITY LABEL TO GENERAL
    props = CustomPropertyList()
    props.append(BoolProperty(name="MSIP_Label_57443d00-af18-408c-9335-47b5de3ec9b9_Enabled", value=True))
    props.append(DateTimeProperty(name="MSIP_Label_57443d00-af18-408c-9335-47b5de3ec9b9_SetDate", value="2024-12-23T12:50:17Z"))
    props.append(StringProperty(name="MSIP_Label_57443d00-af18-408c-9335-47b5de3ec9b9_Method", value="Privileged"))
    props.append(StringProperty(name="MSIP_Label_57443d00-af18-408c-9335-47b5de3ec9b9_Name", value="General v2"))
    props.append(StringProperty(name="MSIP_Label_57443d00-af18-408c-9335-47b5de3ec9b9_SiteId",
                 value="6e51e1ad-c54b-4b39-b598-0ffe9ae68fef"))
    props.append(StringProperty(name="MSIP_Label_57443d00-af18-408c-9335-47b5de3ec9b9_ActionId",
                 value="accb3366-0258-48d3-95a0-00a6be342c5b"))
    props.append(IntProperty(name="MSIP_Label_57443d00-af18-408c-9335-47b5de3ec9b9_ContentBits", value=2))

    # Assign the custom properties to the workbook
    wb.custom_doc_props = props

    # Save the workbook
    wb.save("output_colored.xlsx")
    os.startfile("output_colored.xlsx")


def compare_bom(path1: str, path2: str):
    bom1, m1 = bom_excel_to_dictionary(path1)
    bom2, m2 = bom_excel_to_dictionary(path2)

    diff = DeepDiff(bom1, bom2, threshold_to_diff_deeper=0)
    max_depth = max(m1, m2)

    item_added = [re.findall(r'\[\'(.*?)\'\]', element.replace('[\'content\']', "").replace("root", ""))
                  for element in diff.get('dictionary_item_added', [])]
    item_removed = [re.findall(r'\[\'(.*?)\'\]', element.replace('[\'content\']', "").replace("root", ""))
                    for element in diff.get('dictionary_item_removed', [])]
    item_changed = [(re.findall(r'\[\'(.*?)\'\]', element.replace('[\'content\']', "").replace("root", "")), diff['values_changed'][element])
                    for element in diff.get('values_changed', [])]

    output = {}

    for item in item_added:
        output = append_to_dict(item, bom2, {'type': 'ADDED'})

    for item in item_removed:
        output = append_to_dict(item, bom1, {'type': 'REMOVED'})

    for item in item_changed:
        output = append_to_dict(item[0][:-1], bom1, {'type': 'CHANGED', 'changed_value': item[0][-1], **item[1]})

    table_output = dict_to_table(output, max_depth)
    columns = ['Level', *[str(i) for i in range(1, max_depth+1)], 'Item', 'Description', 'Revision',
               'Quantity', 'SupplyType', 'ModifyType']

    output_df = pd.DataFrame(table_output, columns=columns)

    save_df_to_excel(output_df, max_depth)

# Function to check if a file exists


def check_file(file_path):
    if os.path.isfile(file_path):
        print(f"File exists: {file_path}")
    else:
        print(f"File does not exist: {file_path}")


##################################################### WORKING CODE #################################################

# Create the parser
parser = argparse.ArgumentParser(description="A script compare two BOM document with stardard format")

# Add arguments with help descriptions
parser.add_argument('--bom1', type=str, required=True, help='Path to the first bom')
parser.add_argument('--bom2', type=str, required=True, help='Path to the second bom')

# Parse the arguments
args = parser.parse_args()

# Check the files
check_file(args.bom1)
check_file(args.bom2)

compare_bom(args.bom1, args.bom2)


# filePath = 'C:/Users/SESA787052/Downloads/BOM QPBE44026-15 Design to Manufacturing1.xlsx'
# filePath2 = 'C:/Users/SESA787052/Downloads/BOM QPBE44026-15 Design to Manufacturing2.xlsx'
# compare_bom(filePath, filePath2)
# print(get_file_type(filePath))
# skip = find_table_origin_line_number(filePath, 'Import_Creo')-1
# print(skip)
# compare_bom('C:/Users/SESA787052/Documents/BOM_compare/BOM DESIGN QSMA11485_01.xlsx',
#             "C:/Users/SESA787052/Documents/BOM_compare/BOM DESIGN QSMA11485_02.xlsx")
# print(bom_excel_to_dictionary('C:/Users/SESA787052/Downloads/QBVE94603.xlsx'))
# bom, _ = bom_excel_to_dictionary('C:/Users/SESA787052/Downloads/0M-3402007000.xlsx')
# bom, _ = bom_excel_to_dictionary(filePath)
# pprint.pprint(bom)
# bom2, _ = bom_excel_to_dictionary('C:/Users/SESA787052/Downloads/QBVE94603.xlsx')

# print(bom2['34413004'].keys())
# print(bom2)
# compare_bom(bom, bom2)
