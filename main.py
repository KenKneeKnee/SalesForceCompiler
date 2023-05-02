
import sys
import pandas as pd
import numpy as np
import datetime
import os

# This file path is used when packaging script into an executable. Use normal file path for development
app_path = os.path.dirname(sys.executable)

# Function definition for swapping 2 columns in a dataframe
def swap(df, col1, col2):
    col_list = list(df.columns)
    x, y = col_list.index(col1), col_list.index(col2)
    col_list[y], col_list[x] = col_list[x], col_list[y]
    df = df[col_list]
    return df

if __name__ == '__main__':
    print("Reading files...")
    try:
        temp_path = os.path.join(app_path, "PrdnImport_YYYYMMDDHHMM_SPL_ALL_M.xlsx")
        temp = pd.read_excel(temp_path)
        temp = temp.set_index('RecordId') # Rows are matched based on RecordID
    except:
        print("Template File not found")
        input("Press Enter to exit")

    # Heat Treatment
    try:
        ht_path = os.path.join(app_path, "PrdnImport_HT.xlsx")
        ht = pd.read_excel(ht_path)
        ht = ht[ht['Unnamed: 0'] == "/SPL HT"] # Rows for each department are selected using this line
        ht = ht.set_index('RecordId')
        if ht.shape[1] is not temp.shape[1]:
            print("WARNING! HT Column numbers do not match")
        temp.update(ht, overwrite=False)
        print("...Heat Treatment data transferred")
    except:
        print("...Heat Treatment file not found")

    # Cold Forge
    try:
        cf_path = os.path.join(app_path, "PrdnImport_CF.xlsx")
        cf = pd.read_excel(cf_path)
        cf = cf[cf['Unnamed: 0'] == "/SPL CF"]
        cf = cf.dropna(subset=['RecordId'])
        cf = cf.set_index('RecordId')
        if cf.shape[1] is not temp.shape[1]:
            print("WARNING! CF Columns numbers do not match")
        temp.update(cf, overwrite=False)
        print("...Cold Forge data transferred")
    except:
        print("...Cold Forge File not found")

    # Stamping
    try:
        stp_path = os.path.join(app_path, "PrdnImport_STP.xlsx")
        stp = pd.read_excel(stp_path)
        stp = stp[stp['Unnamed: 0'] == "/SPL STP"]
        stp = stp.set_index('RecordId')
        if stp.shape[1] is not temp.shape[1]:
            print("WARNING! STP Columns numbers do not match")
        temp.update(stp, overwrite=False)
        print("...Stamping data transferred")
    except:
        print("...Stamping File not found")

    # Machining
    try:
        ms_path = os.path.join(app_path, "PrdnImport_MS.xlsx")
        ms = pd.read_excel("PrdnImport_MS.xlsx")
        ms = ms[ms['Unnamed: 0'] == "/SPL MS"]
        ms = ms.dropna(subset=['RecordId'])
        ms = ms.set_index('RecordId')
        if ms.shape[1] is not temp.shape[1]:
            print("WARNING! MS Columns numbers do not match")
        temp.update(ms, overwrite=False)
        print("...Machining data transferred")
    except:
        print("...Machining File not found")

    # Final prep
    final_drop = ['Unnamed: 0']
    temp = temp.drop(columns=final_drop)
    temp.reset_index(inplace=True) # This allows RecordID column to be swapped
    temp = swap(temp, 'RecordId', 'Unnamed: 1')
    temp = swap(temp, 'RecordId', 'Object')

    # Replace blank with 0 in these columns
    replace_nan = ['QttyRes', 'TotalNc', 'QttyNonNc']
    for col in replace_nan:
        temp[col].fillna(0, inplace=True)

    # Remove zeroes in these columns
    replace_zero = ['QttyPlan', 'Availability', 'Performance', 'OEE']
    for col in replace_zero:
        temp[col].replace(0, np.nan, inplace=True)

    # Generate Report
    time = datetime.datetime.now()
    time_str = time.strftime("%Y%m%d%H%M")
    file = 'PrdnImport_' + time_str + '_SPL_ALL_M.csv'
    temp.to_csv(file, index=False)
    print("Report generated")
    input("Press Enter to exit")
