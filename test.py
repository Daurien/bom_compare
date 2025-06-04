from PIL import ImageFont
import ctypes
from pprint import pprint
import pandas as pd
import tkinter as tk
from tkinter import font


def light_bom_oracle_to_dictionary(filePath: str):
    """Take txt file BOM extraceted from Oracle and return a simplified BOM dictionary.

    Args:
        filePath (str): file to read (from Oracle only)

    Returns:
        dict: bom transformed as a dictionary with 'Item name' as keys and other attributes as values.
    """
    # Read the tab-separated file into a DataFrame
    df = pd.read_csv(filePath, sep='\t')

    # Create the simplified DataFrame
    simplified_df = df[['Level', 'Item', 'Revision', 'Quantity', 'Description']].rename(
        columns={
            'Item': 'Item name',
            'Revision': 'Revision',
            'Quantity': 'Quantity',
            'Description': 'Description'
        })
    # Clean whitespace from all string columns
    simplified_df = simplified_df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

    # Add 'Depth' column full of 1
    simplified_df['Depth'] = 1

    # Remove suffix pattern from 'Item name' column
    simplified_df['Item name'] = simplified_df['Item name'].str.replace(r'_\d{2}$', '', regex=True)

    # Find parent for each row
    parent_list = []
    for idx, row in simplified_df.iterrows():
        current_level = row['Level']
        if current_level == 1:
            parent = "main"
        else:
            parent = None
            # Look upwards for the most recent row with Level == current_level - 1
            for prev_idx in range(idx - 1, -1, -1):
                if simplified_df.loc[prev_idx, 'Level'] == current_level - 1:
                    parent = simplified_df.loc[prev_idx, 'Item name']
                    break
        parent_list.append(parent)

    simplified_df['parent'] = parent_list

    # Add category column based on the logic
    category_list = []
    for idx, row in simplified_df.iterrows():
        current_item = str(row['Item name'])
        parent_item = str(row['parent'])
        # Check next row's Item name if it exists
        if idx + 1 < len(simplified_df):
            next_item_parent = str(simplified_df.loc[idx + 1, 'parent'])
        else:
            next_item_parent = None
        if (next_item_parent == current_item) and (current_item != parent_item):
            category = "Assembly"
        else:
            category = "Part"
        category_list.append(category)

    simplified_df['category'] = category_list

    # Filter out rows where category is 'Assembly'
    simplified_df = simplified_df[simplified_df['category'] != 'Assembly'].drop(
        ['parent', 'category', 'Level'], axis=1)

    light_bom = {}

    for _, row in simplified_df.iterrows():
        item = row['Item name']
        light_bom[item] = row.to_dict()

    return light_bom

# print(get_matieres_from_creo_db())


# result = light_bom_oracle_to_dictionary('fnd_gfm_204718568.txt')

# if '34413553' in result:
#     print("34413553_01 found in dictionary")
#     pprint(result['34413553'])
# else:
#     print("34413553_01 not found in dictionary")


# Create a persistent Tkinter root and font object at module level
_tk_root = tk.Tk()
_tk_root.withdraw()
_default_font = font.Font(root=_tk_root, family="Calibri", size=11)


def get_text_pixel_width(text, font_name="Calibri", font_size=11):
    # Use the persistent font object if default params, else create a new one
    if font_name == "Calibri" and font_size == 11:
        return _default_font.measure(text)
    else:
        tk_font = font.Font(root=_tk_root, family=font_name, size=font_size)
        return tk_font.measure(text)


def get_text_pixel_width2(text, font_name="Calibri", font_size=11, dpi=96):
    # Pillow default DPI is 72, so scale font size to match 96 DPI
    scaled_font_size = int(round(30))
    try:
        font = ImageFont.truetype(f"{font_name}.ttf", scaled_font_size)
    except OSError:
        font = ImageFont.load_default()
    bbox = font.getbbox(text)
    print(bbox)
    width = bbox[2] - bbox[0]
    return width


print(get_text_pixel_width("Hello world"))
print(get_text_pixel_width2("Hello world"))
print(get_text_pixel_width(" Old : FRONT UPPER/LOWER GVEAC7 CABINED PROTECTION SHIELD"))
print(get_text_pixel_width2(" Old : FRONT UPPER/LOWER GVEAC7 CABINED PROTECTION SHIELD"))

dpi_x = _tk_root.winfo_fpixels('1i')
dpi_y = _tk_root.winfo_fpixels('1i')  # Usually the same as dpi_x

print(f"Tkinter DPI: {dpi_x} x {dpi_y}")
