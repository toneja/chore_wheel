#!/usr/bin/env python3

import os
import random
import numpy as np
import pandas as pd


def chore_wheel():
    # Change to the working directory
    os.chdir(os.path.dirname(__file__))
    # Definitions of chores
    chores = {
        1: "Bathrooms",
        2: "Hallways/Stairs",
        3: "Facilities",
        4: "Headhouses",
        5: "Growth Chamber",
    }
    # Get list of employee names from spreadsheet
    data_file = "ChoreAssignments.xlsx"
    if os.path.exists(data_file):
        data = pd.read_excel(data_file)
        df = pd.DataFrame(data)
        names = df["Employee"].tolist()
    else:
        print(f"No list of employee names [ChoreAssignments.xlsx] found.")
        input("Press ENTER to exit.\n")
        quit()
    print(f"There are {len(names)} employees to be split into {len(chores)} groups.\n")
    # Clear out old data if the spreadsheet is "full"
    if df.iloc[:, 7].notna().any():
        df.iloc[:, 4:8] = np.nan
    # Split employees into random groups
    group_num = 1
    # Try to prevent infinite loops
    duplicate = 0
    while len(names) > 0:
        if group_num > len(chores):
            group_num = 1
        while True:
            name = random.choice(names)
            name_index = df[df["Employee"] == name].index[0]
            name_data = df.loc[name_index].drop(labels="Employee", errors="ignore")
            if (name_data == chores.get(group_num)).any():
                print(
                    f"The chore {chores.get(group_num)} has already been assigned to {name}."
                )
                duplicate += 1
                if duplicate > 5:
                    # The loop can get stuck here; increment group number to continue
                    group_num += 1
            else:
                duplicate = 0
                # Skip the 1st 4 columns; insert assigned chore into first empty cell
                for col in df.columns[4:]:
                    if pd.isna(df.at[name_index, col]):
                        df.at[name_index, col] = chores.get(group_num)
                        names.remove(name)
                        group_num += 1
                        break
                break
    print(df)
    df.to_excel(data_file, index=False)


if __name__ == "__main__":
    chore_wheel()
