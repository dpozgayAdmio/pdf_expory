import os
from os.path import isdir

import pandas as pd
import matplotlib.pyplot as plt
import win32com.client
import warnings

from datetime import datetime
from colorama import init, Fore, Style

MONTHS = {
    1: {"leden", "january", "január", "januar"},
    2: {"únor", "unor", "february", "február", "februar"},
    3: {"březen", "brezen", "march", "marec"},
    4: {"duben", "april", "apríl"},
    5: {"květen", "kveten", "may", "máj"},
    6: {"červen", "cerven", "june", "jún", "jun"},
    7: {"červenec", "cervenec", "july", "júl", "jul"},
    8: {"srpen", "august"},
    9: {"září", "zari", "september"},
    10: {"říjen", "rijen", "october", "október"},
    11: {"listopad", "november"},
    12: {"prosinec", "december"}
}

ERROR = Fore.RED
WARNING = Fore.YELLOW
INFO = Fore.BLUE
SKIPP = Fore.CYAN
GOOD = Fore.GREEN

def my_print(status, text):
    with open("log.txt", "a") as file:
        file.write(text + "\n")
    if status is not None:
        print(status + text)
    else:
        print(text)
    return 0

def get_sheets(path, month, year, tryes=None):
    xls = None
    for _ in range(5):
        try:
            xls = pd.ExcelFile(path)
            break
        except PermissionError:
            print(Fore.YELLOW + f"Premision deny, first close file: {path}")
            continue

    if xls is None:
        return []

    sheets = xls.sheet_names
    to_process = []
    current_months = MONTHS[month]

    for sheet_name in sheets:
        print(Fore.BLUE + sheet_name, end=", ")
        for c_m in current_months:
            if c_m in sheet_name:
                to_process.append(sheet_name)

        # Predosli rok kvoli zavierkam
        if ("zav" in sheet_name or "záv" in sheet_name) and  str(year - 1) in sheet_name:
            to_process.append(sheet_name)
    print()
    return to_process


def read(df, month, sheet, celkom=1):
    date = None
    for index, row in df.iterrows():
        if index == 11 and row[4] == "DUZP":
            date = row[5]
            print(date)
            if isinstance(date, str):
                try:
                    date = datetime.strptime(date, "%d.%m.%Y")
                except ValueError:
                    print(Fore.RED + f"Value error in date {date}")
                    return 0, date

        if isinstance(row[0], str) and ("total" in row[0].lower() or "celkem" in row[0].lower()):
            celkom -= 1

        try:
            if index == 15 and (row[1].strip().lower() not in MONTHS[date.month] or date.month != month):
                if "záv" in sheet:
                    print(Fore.RED + "skipp", f"not this ZAVIERKA: {row[1].lower()}, {date.month} but is {month}")
                else:
                    print(Fore.RED + f"invalid mont: {row[1].lower()} {date.month}")
                return  0, date
        except AttributeError:
            print(Fore.RED + f"EXEPT error with: {row[1]}")
            if pd.isna(row[1]):
                # TODO: zvazil by som continue alebo to sem dodat
                print(Fore.RED + "cant find month text")
            return 0, date

        if celkom == 0:

            if df.iloc[index, 7] == 0:
                print(Fore.CYAN + f"Skipp: Celkem = 0")
                return 0, date

            for i in range(index + 1, len(df)):
                end = df.iloc[i, 0]
                if isinstance(end, str) and end.lower() in {"odeslano", "odesláno"}:
                    return index + 1, date
            print(Fore.RED + f"Cant finde end \"ODESLANO\": {index + 1 + len(df)}")
            return 0, date

    print(Fore.RED + F"Cant found \"celkem\": {index}")
    return 0, date


def save_ugly(df, row):
    df_cut = df.iloc[:row + 1, :9]

    # 3. Vykresli tabuľku ako obrázok
    fig, ax = plt.subplots(figsize=(12, 5))
    ax.axis('tight')
    ax.axis('off')
    table = ax.table(cellText=df_cut.values, colLabels=df_cut.columns, loc='center')

    # 4. Ulož ako PDF
    plt.savefig("export.pdf", bbox_inches='tight')
    print(Fore.GREEN + "Úspešne exportované do export.pdf")


def save(path, file_name, sheet, row, out_name, debug_mode=False):
    # DEBUG
    if debug_mode:
        out_path = path + '\\Dodací listy' + "\\" + out_name
        print("rows:", row)
        print(Fore.GREEN + f"savet to: {out_path}")
        return False

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    if not os.path.exists(path + '\\Dodací listy'):
        print(Fore.RED + "Cant find file \"Dodací listy\"")
        return False

    # Otvor
    wb = excel.Workbooks.Open(path + "\\" + file_name)
    ws = wb.Sheets(sheet)

    ws.PageSetup.PrintArea = f"A1:I{row}"

    # 1 A4
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False

    # Export do PDF
    out_path = path + '\\Dodací listy' + "\\" + out_name
    if os.path.exists(out_path + ".pdf"):
        print(Fore.RED + f"File allredy exist: {out_path}")
        wb.Close(SaveChanges=False)
        excel.Quit()
        return False

    ws.ExportAsFixedFormat(0, out_path)
    print(Fore.GREEN + f"savet to: {out_path}")

    # Zavri
    wb.Close(SaveChanges=False)
    excel.Quit()

    return True


def main():
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet._reader")
    folder = r"C:\Users\dominik.pozgay\OneDrive - ADMIO s.r.o\FVL"
    #folder = "C:\Users\dominik.pozgay\OneDrive - ADMIO s.r.o\ADMIO - FV_DL"
    month = 1
    year = 2025
    black_list = {"Adriaan Van Selm", "argoterra-cz", "Hájek Pavel_společnosti"}
    extra_list = {"Zle"} # set()  # ak chcem spracovat IBA konkretne firmy zadam ich do extra_list
                        # ak nie je prazdny ignoruje vsetko ostatne


    compliet_dl = 0
    compliet_z = 0

    init(autoreset=True)

    for sub_folder in os.listdir(folder):
        print(Fore.BLUE + f"Process company: {sub_folder}")
        find = False
        saving = False

        if extra_list and sub_folder not in extra_list:
            print(Fore.CYAN + "Skipp beacous extra")
            print("------------------------------")
            continue

        if not isdir(folder + '\\' + sub_folder):
            print(Fore.YELLOW + "Not file")
            print("------------------------------")
            continue

        if sub_folder in black_list:
            print(Fore.YELLOW + "Black list")
            print("------------------------------")
            continue

        for file in os.listdir(folder + '\\' + sub_folder):
            if "Dodací list 2025" in file:
                path = folder + '\\' + sub_folder + "\\" + file
                print(f"Find: {path}")
                sheets_to_process = get_sheets(path, month, year)

                for sheet in sheets_to_process:
                    check = False
                    print(Fore.YELLOW + sheet, datetime.now().strftime("%H:%M:%S"))
                    # nacitaj sheet
                    df = pd.read_excel(path, sheet_name=sheet, header=None)

                    # vyber oblast a skontroluj datum
                    if "jic" in path.lower():
                        row_index, date = read(df, month, sheet, celkom=4)
                    else:
                        row_index, date = read(df, month, sheet)

                    if row_index == 0:
                        print()
                        continue

                    if "záv" not in sheet and date.month != month:
                        print(Fore.RED + "ERROR date")
                        print()
                        continue

                    # exportuj do pdf a uloz
                    if 'záv' in sheet or 'zav' in sheet:
                        name = "závěrka_" + str(year) + " Dodací list " + sub_folder.title()
                    else:
                        name = str(month) + "_" + str(year) + " Dodací list " + sub_folder.title()
                        check = True

                    saving = save(folder + '\\' + sub_folder, file, sheet, row_index, name)

                    if saving and check:
                        compliet_dl += 1
                    if saving and not check:
                        compliet_z += 1

                    print()

                find = True

        if not find:
            print(Fore.RED + f"Nenajdeny subor Dodací list 2025!!!")

        print("------------------------------")

    print(compliet_dl, compliet_z)


if __name__ == "__main__":
    main()