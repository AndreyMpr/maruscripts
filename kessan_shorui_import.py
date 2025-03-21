import pandas as pd
from openpyxl import load_workbook

def transfer_data():

    # Assign each source file a corresponding column in the target file.
    try:
        all_source_files = {
            '515.xlsx': 3,   # Column D
            '5151.xlsx': 4,  # Column E
            '5152.xlsx': 5,  # Column F
            '5153.xlsx': 6,  # Column G
            '5154.xlsx': 7,  # Column H
            '545.xlsx': 8,   # Column I
            '5451.xlsx': 9,  # Column J
            '5452.xlsx': 10, # Column K
            '514.xlsx': 11,  # Column L
            '544.xlsx': 12   # Column M
        }
        
        source_dicts = {}

        # Print keys and values from each input excel to the screen.
        for filename, col_index in all_source_files.items():
            print(f"Reading source file: {filename}...")

            # Read source files with pandas
            source_df = pd.read_excel(filename, dtype=str)
            keys = source_df.iloc[:, 1].astype(str)
            values = source_df.iloc[:, -1].astype(str)
            source_dicts[col_index] = dict(zip(keys, values))

            print(f"Keys for column {col_index} from {filename}: {list(source_dicts[col_index].keys())[:5]}...")

    except FileNotFoundError:
        print("Error: One or more Excel files not found. Make sure all files are in the same directory as this script.\nエラー：1 つまたは複数の Excel ファイルが見つからない。すべてのファイルがこのスクリプトと同じフォルダーにあることを確認してください。")
    except Exception as e:
        print(f"An error occured: {str(e)}")

if __name__ == "__main__":
    transfer_data()
