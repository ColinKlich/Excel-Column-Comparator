import pandas as pd
from difflib import SequenceMatcher
import jellyfish
from tqdm import tqdm
from pathlib import Path
from beaupy import select_multiple, select
import openpyxl

def __main__():
    print("Input Filepath for the first file (.xlsx)")
    file1_loc = Path(input()).as_posix().replace("\"", "")

    print("Which sheet will you be using?")
    sheets = list_excel_sheets(file1_loc)
    sheet_name1 = select(sheets)

    print("Input Filepath for the second file (.xlsx)")
    file2_loc = Path(input()).as_posix().replace("\"", "")

    print("Which sheet will you be using?")
    sheets = list_excel_sheets(file2_loc)
    sheet_name2 = select(sheets)

    print("Which column will you be comparing? (Should be identical)")
    col_comparison = input()
    print("\n")

    try:
        # df_first_file, df_second_file = readFiles(file1_loc,sheet_name1, file2_loc, sheet_name2)
        
        print("Reading in Files...")
        if file1_loc.__contains__(".csv"):
            df_first_file = pd.read_csv(file1_loc)
        else:
            df_first_file = pd.read_excel(file1_loc,sheet_name=sheet_name1)

        if file2_loc.__contains__(".csv"):
            df_second_file = pd.read_csv(file2_loc)
        else:
            df_second_file = pd.read_excel(file2_loc,sheet_name=sheet_name2)

        print("Files Read!")

    except:
        print("Input is incorrectly formatted! Try Again.")
        __main__()

    try:
        list1 = list(df_first_file[col_comparison])
        list2 = list(df_second_file[col_comparison])
        similar_items, dissimilar_items = compare_dataframes(list1, list2)
    except:
        print("The selected column cannot be found. Enter a new column")
        col_comparison = input()

        df_first_file,df_second_file = readFiles(file1_loc,sheet_name1, file2_loc, sheet_name2)
        list1 = list(df_first_file[col_comparison])
        list2 = list(df_second_file[col_comparison])

        similar_items, dissimilar_items = compare_dataframes(list1, list2)

    df_out = pd.DataFrame(similar_items, columns=['Input '+ col_comparison, 'Master ' + col_comparison, 'Similarity'])
    # df_out.to_excel("output.xlsx", sheet_name='Sheet1', index=False)
    print("Found "+ str(len(similar_items)) + " similar items!")
    
    new_df = pd.merge(df_out, df_first_file, left_on='Input '+ col_comparison, right_on=col_comparison)
    new_new_df = pd.merge(new_df, df_second_file, left_on='Master '+ col_comparison, right_on=col_comparison)
    
    cols = select_multiple(list(new_new_df.columns), tick_character='*', ticked_indices=[0,1,2], pagination=True, page_size=10)
    
    new_new_df.to_excel("output.xlsx", sheet_name='Sheet1', columns=cols, index=False)

    dissimilar_df = pd.DataFrame(dissimilar_items, columns=[col_comparison])
    dissimilar_out = pd.merge(dissimilar_df, df_first_file, left_on=col_comparison, right_on=col_comparison)
    dissimilar_out.to_excel("not found.xlsx", sheet_name='Sheet1', index=False)
    
    print("Written to file\n")
    wait = input("Press (enter) to close")

    return 0

def similar(a, b):
    try:
        return jellyfish.jaro_similarity(a, b)
    #SequenceMatcher(None, a, b).ratio()
    except:
        return 0

def list_excel_sheets(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet_names = wb.sheetnames
    return sheet_names
    
def readFiles(file1_loc, sheet_name1, file2_loc, sheet_name2):
    print("Reading in Files...")

    df_first_file = pd.read_excel(file1_loc,sheet_name=sheet_name1)
    df_second_file = pd.read_excel(file2_loc,sheet_name=sheet_name2)
    
    return df_first_file, df_second_file

def compare_dataframes(list1, list2):
    similar_items = []
    dissimilar_items = []
    for item1 in tqdm(list1, desc="Loading…", ascii=False, ncols=75):
        str1 = str(item1).replace('R90','R9')
        foundVar = False
        for item2 in list2:
            str2 = str(item2).replace('R90','R9')
            if (str1.startswith(str2) or str1.endswith(str2) or str2.startswith(str1) or str2.endswith(str1)) and (len(str2) > 4) and (len(str1) > 4):
                foundVar = True
                if (len(str1) <= 6 or len(str2) <= 6):
                    similar_items.append((item1, item2, 'Possible'))
                else:
                    similar_items.append((item1, item2, 'Exact'))
                break
        if foundVar == False:
            dissimilar_items.append(item1)
            # elif (str3.startswith(str4) or str3.endswith(str4) or str4.startswith(str3) or str4.endswith(str3)) and (len(str3) > 4) and (len(str4) > 4):
            #     similar_items.append((item1, item2, 'Possible'))
            #     break
    print("Complete")
    return similar_items, dissimilar_items

def clean(string):
    string2 = ''
    for char in string:
        if(char != '0'):
            string2 += char
    return string2


if __name__ == "__main__": __main__()