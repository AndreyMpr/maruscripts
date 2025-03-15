import pandas as pd
from openpyxl import load_workbook

def transfer_data():
    try: ""
         


    except FileNotFoundError:
        print("Error: One or more Excel files not found. Make sure all files are in the same directory as this script.\nエラー：1 つまたは複数の Excel ファイルが見つからない。すべてのファイルがこのスクリプトと同じフォルダーにあることを確認してください。")
    except Exception as e:
        print(f"An error occured: {str(e)}")

if __name__ == "__main__":
    transfer_data()
