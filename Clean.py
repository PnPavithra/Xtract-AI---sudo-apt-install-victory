import pandas as pd
import os
from glob import glob

# Directory containing the .xls files
input_dir = "./Data"  # Change to your actual directory path
output_dir = "CleanedFiles"  # Directory to save cleaned files
os.makedirs(output_dir, exist_ok=True)  # Create output directory if it doesn't exist

# Get all .xls files in the directory
xls_files = glob(os.path.join(input_dir, "*.xls"))

# Function to clean each file
def clean_data(file_path):
    df = pd.read_excel(file_path)

    # Standardize column names
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

    # Drop completely empty rows
    df.dropna(how='all', inplace=True)

    # Fill missing values
    df['title'].fillna("Unknown Title", inplace=True)
    df['author'].fillna("Unknown Author", inplace=True)

    # Drop rows where essential columns are missing
    essential_cols = ['date', 'card_number', 'transaction', 'barcode', 'title']
    df.dropna(subset=essential_cols, inplace=True)

    # Convert date column to datetime
    df['date'] = pd.to_datetime(df['date'], errors='coerce')

    # Drop duplicates
    df.drop_duplicates(inplace=True)

    return df

# Process each .xls file
for file in xls_files:
    cleaned_df = clean_data(file)

    # Save cleaned file
    output_file = os.path.join(output_dir, f"cleaned_{os.path.basename(file).replace('.xls', '.xlsx')}")
    cleaned_df.to_excel(output_file, index=False)
    print(f"Cleaned data saved: {output_file}")

print("All files cleaned and saved.")
