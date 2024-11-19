import pandas as pd
import os

# Directory where the Excel files are stored
input_directory = 'output'  # Change this to the path where your files are located
output_file = 'complete.xlsx'  # Change this to your desired output file path

# List all Excel files in the directory
excel_files = [f for f in os.listdir(input_directory) if f.endswith('.xlsx')]

# Initialize an empty list to store dataframes
dataframes = []

# Loop through all the Excel files and read them into dataframes
for file in excel_files:
    file_path = os.path.join(input_directory, file)
    
    # Read the Excel file into a pandas dataframe
    df = pd.read_excel(file_path)
    
    # Optionally: You can inspect the first few rows of each file
    # print(df.head())
    
    # Append the dataframe to the list
    dataframes.append(df)

# Concatenate all dataframes into a single dataframe
merged_df = pd.concat(dataframes, ignore_index=True)

# Write the merged dataframe to a new Excel file
merged_df.to_excel(output_file, index=False)

print(f"Successfully merged {len(excel_files)} files into {output_file}")
