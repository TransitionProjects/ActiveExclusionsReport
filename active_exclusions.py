import pandas as pd
import numpy as np
from datetime import datetime
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

def main():
    # import the exclusion report and read it into a pandas data frame
    df = pd.read_excel(askopenfilename(title="Open the Exclusion Report"))

    # create the relevant date variables
    today = datetime.today().date()

    # filter out exclusions that happened more than 6 months ago the exclusion
    # was not agency wide, and/or the exclusion was not created by the day center
    df_2 = df[
        (
            (
                df["Infraction Provider"].str.contains("Day") |
                df["Infraction Provider"].str.contains("Agency")
            )
        ) |
        (
            df["Infraction Banned End Date"].isna()
        )
    ]

    # write the final report to excel and exit
    intial_file_name = "Exclusion Report {}.xlsx".format(today)
    writer = pd.ExcelWriter(
        asksaveasfilename(
            title="Save the Recent Exclusion from the Resource Center Report",
            initialdir=".xlsx",
            initialfile=intial_file_name
        ),
        engine="xlsxwriter"
        )
    df_2[[
        "Client Uid",
        "Client First Name",
        "Client Last Name",
        "Infraction Provider",
        "Infraction Banned Start Date",
        "Infraction Banned End Date",
        "Infraction Banned Code",
        "Infraction Type"
    ]].to_excel(writer, sheet_name="Resource Center Exclusions", index=False)
    writer.save()

if __name__ == "__main__":
    main()
