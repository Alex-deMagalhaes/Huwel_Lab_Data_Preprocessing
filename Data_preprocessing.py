"""Data Preprocessing."""

import numpy as np
import pandas as pd
import os
import openpyxl

"""
A start and end path are defined below, so that they can be changed as
necessary.
start_path: str, the absolute file path of the experimental data files (e.g.
"0620" or "0700").
end_path: str, the absolute file path (including the file name) for the excel
file produced by this script as an output.
"""

# start_path = "C:/Users/Pedro/Desktop/2019_07_11/0620"
start_path = "C:/Users/Pedro/Desktop/2019-08-09_saline_reproducibility/700"
end_path = "C:/Users/Pedro/Desktop/Test_Excel_Sheet.xlsx"

###############################################################################
"""
Define the functions that will be needed to preprocess the data. These
functions extract the data from the .asc files and create pandas dataframes
that will become the individual excel sheets in the excel file.
"""


def asc_read(file_path):
    """
    Helper function: Reads .asc files and converts them into a Pandas dataframe
    with two columns: "Pixel" and "value".

    file_path: path for data files
    returns pandas dataframe
    """
    col_names = ["Pixel", "value"]
    data = pd.read_csv(file_path, sep="\t", names=col_names)
    return data


def add_run_numbers(dataframe):
    """
    Add a new row containing the run numbers (e.g. BG, run 1, run 2, etc.)
    to the top of the input dataframe.

    dataframe: pandas dataframe
    returns pandas dataframe
    """
    # BG is background
    run_numbers = [None, "BG"] + [x for x in range(1, dataframe.shape[1] - 1)]
    dataframe.loc[-1, :] = run_numbers  # adding a row
    dataframe.index = dataframe.index + 1  # shifting index
    dataframe.sort_index(inplace=True)

    return dataframe


def iterate_over_runs(directory):
    """
    Create pandas dataframe containing a pixel column and all the runs as
    columns. Also adds the run number and "BG", for background run, as the
    first row of the dataframe.

    directory: filepath, the folder containing the individual runs
    returns pandas dataframe
    """
    # Initialize dataframe with one column: pixel number
    df = pd.DataFrame({
        "Pixel": [x for x in range(1, 1025)]
    })

    # Add each run as a column
    for filename in os.listdir(directory):
        if filename.endswith(".asc"):
            asc_file = asc_read(directory + filename)
            df[filename] = asc_file.iloc[:, 1]

    # Add a row for run number and "BG"
    df = add_run_numbers(df)

    return df


def subtract_background(dataframe):
    """
    Append new columns to the dataframe, in which the background is subtracted
    from all the runs.

    dataframe: pandas dataframe.
    returns tuple, (pandas dataframe, number of new columns)
    """
    i = 1
    for column in dataframe.iloc[:, 2:]:
        column_name = "Run: {}".format(i)
        dataframe[column_name] = df.iloc[1:, i+1] - df.iloc[1:, 1]
        i += 1

    return (dataframe, i - 1)


def get_averaged_data_column(dataframe, number_of_runs):
    """
    Append column to dataframe, which is the average of all the
    background-subtracted runs.

    dataframe: pandas dataframe.
    number_of_runs: number of background-subtracted runs
    returns pandas dataframe
    """
    dataframe["Mean"] = dataframe.iloc[:, -1 * number_of_runs:].mean(axis=1)

    return dataframe

###############################################################################


"""
With the functions defined above, we create a loop that preprocesses the data
into pandas dataframes and stores all the dataframes into a list. Then we use
pandas.ExcelWriter to convert that list into the final Excel file.

Before creating the loop we need to extract the paths for the sub-folders
within the start path and slightly modify them so that Python can read the
paths correctly.
"""

# From the start path, create a list containing all sub-folders in start path.
# Replace "\\" in file paths with "/" so that Python can read the file paths
# properly.
all_sub_directories = [
    sub_directories.replace("\\", "/")
    for sub_directories, dirs, files in os.walk(start_path)
]

# The first sub-folder in all_sub_directories is just the start path,
# so we don't need it. Let's keep all the other sub-folders.
raw_sub_directory_list = all_sub_directories[1:]

sub_directory_list = []
sub_directory_names = []
for sub_directory in raw_sub_directory_list:
    # It's necessary to add a "/" to the end of each sub-folder file path
    sub_directory_list.append(sub_directory + "/")

    # The names in sub_directory_names become the names for the individual
    # Excel sheets
    sub_directory_names.append(sub_directory.split("/")[-1])

# Create the loop
dataframes = []
for sub_folder in sub_directory_list:
    df = iterate_over_runs(sub_folder)
    subtracted_tuple = subtract_background(df)

    subtracted_df = subtracted_tuple[0]
    number_of_runs = subtracted_tuple[1]

    averaged_df = get_averaged_data_column(subtracted_df, number_of_runs)

    # Reset the index to be the "Pixel" column
    final_df = averaged_df.set_index("Pixel")
    dataframes.append(final_df)

# Create Excel file
writer = pd.ExcelWriter(end_path)

i = 0
for dataframe in dataframes:
    dataframe.to_excel(writer, sheet_name=sub_directory_names[i])
    i += 1

writer.save()
