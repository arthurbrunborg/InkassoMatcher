COMPANY_NAME = "Demand"
API_KEY = "aa501ec42b97de3e1345abac9c18c5250b55d87a1141f3410826626e3ab84454"

import requests
import pandas as pd
import tkinter as tk
from tkinter.filedialog import askopenfilename
from pathlib import Path
from datetime import datetime


def get_pnrs():
    # make request and get data
    try:
        url = "https://api.inkassoregisteret.com/v1/debt/collector/requestEntitiesQuery"
        r = requests.get(url, headers={"X-Api-Key": API_KEY})
        # extract pnr's
        pnrs = []
        for entry in r.json()["entities"]:
            entry_id = entry["entityId"]
            id = entry_id["id"]
            type = entry_id["type"]
            if type != "PNR":
                # TODO: Raise some error
                tk.messagebox.showerror(
                    "Error",
                    "Found non-PNR when getting data from inkassoregisteret API",
                )
                return [], False
            pnrs.append(id)
        return pnrs, True
    except Exception as e:
        tk.messagebox.showerror(
            "Error",
            "Was not able to get data from inkassoregisteret API",
        )
        return None, False


def write_excel(df, pnrs, out_path):
    df = df.astype({"Debitor_Ident": str})

    rows = df.loc[df["Debitor_Ident"].isin(pnrs)]

    rows.to_excel(out_path, sheet_name="Sheet1", index=True)

    return True


def get_and_validate_excel(filename):
    try:
        df = pd.read_excel(filename, sheet_name="Sheet1", header=7, usecols="B:P")
        if not (
            df.columns
            == [
                "Sak_Nr",
                "Debitor_Nr",
                "Debitor_Ident",
                "Debitor_Fodselsdato",
                "Debitor_Navn1",
                "Unnamed: 6",
                "Debitor_Navn2",
                "Debitor_Adresse1",
                "Unnamed: 9",
                "Debitor_Adresse2",
                "Debitor_Postnr",
                "Debitor_Sted",
                "Debitor_EPost",
                "Unnamed: 14",
                "Telefon_Nummer",
            ]
        ).all():
            tk.messagebox.showerror(
                "Error",
                "The file does not look as expected. Was not able to find the expected columns",
            )
            return None, False
        return df, True
    except Exception as e:
        tk.messagebox.showerror(
            "Error",
            f"Was not able to read file '{filename}' as excel file.",
        )
        return None, False


def select_file():
    global filename
    filename = askopenfilename()
    if filename == "":
        return
    global df, df_status
    df, df_status = get_and_validate_excel(filename)
    if not df_status:
        return
    label1["text"] = filename.split("/")[-1]
    label2["text"] = f"Rows in loaded file: {len(df)}"
    extract_button["bg"] = "blue"


def extract_pressed():
    if not pnrs_status:
        tk.messagebox.showerror("Error", "No PNR's loaded")
        return
    if not df_status:
        tk.messagebox.showerror("Error", "No file selected")
        return

    out_filename = (
        Path(filename).parent
        / f"{datetime.today().strftime('%d-%b-%Y-%H-%M-%S')}-out.xlsx"
    )
    status = write_excel(df, pnrs, out_filename)
    if not status:
        return

    tk.messagebox.showinfo("Success", f"Extracted to {out_filename}")


def get_row():
    global row_counter
    row_counter += 1
    return row_counter


def update_pnr_label(pnrs, pnrs_status):
    if not pnrs_status:
        pnr_label["text"] = f"Error when getting PNR's from inkassoregisteret API"
    else:
        pnr_label[
            "text"
        ] = f"{len(pnrs)} number of records loaded from inkassoregisteret API"


def pnrs_button_pressed(brodcase_success=True):
    pnr_label["text"] = "No data from inkassoregisteret loaded.."
    global pnrs, pnrs_status
    pnrs, pnrs_status = get_pnrs()
    update_pnr_label(pnrs, pnrs_status)
    refresh_pnrs_button["bg"] = "blue"
    if pnrs_status and brodcase_success:
        tk.messagebox.showinfo("Success", "Got updated data from inkassoregisteret API")


row_counter = -1
pnrs, pnrs_status = None, False
df_status = False
filename = ""
APPLICATION_NAME = f"Inkassoregister-henter {COMPANY_NAME}"

root = tk.Tk()
root.title(APPLICATION_NAME)

label0 = tk.Label(text=APPLICATION_NAME, font=("Arial", 20))
label0.grid(row=get_row(), column=0, padx=(100, 100), pady=(30, 10))

refresh_pnrs_button = tk.Button(
    text="Refresh data from Inkassoregisteret",
    command=pnrs_button_pressed,
    bg="gray",
    fg="white",
)
refresh_pnrs_button.grid(row=get_row(), column=0, padx=(200, 200), pady=(10, 3))

pnr_label = tk.Label(text="Loading data from Inkassoregisteret...")
pnr_label.grid(row=get_row(), column=0, padx=(100, 100), pady=(3, 10))

select_file_button = tk.Button(
    text="Select file", command=select_file, bg="blue", fg="white"
)
select_file_button.grid(row=get_row(), column=0, padx=(200, 200), pady=(10, 3))

label1 = tk.Label(text="No file selected")
label1.grid(row=get_row(), column=0, padx=(50, 50))

label2 = tk.Label(text="")
label2.grid(row=get_row(), column=0, padx=(50, 50))

extract_button = tk.Button(
    text="Extract",
    command=extract_pressed,
    bg="gray",
    fg="white",
)
extract_button.grid(row=get_row(), column=0, padx=(200, 200), pady=(3, 30))

root.after(ms=200, func=lambda: pnrs_button_pressed(False))
root.mainloop()