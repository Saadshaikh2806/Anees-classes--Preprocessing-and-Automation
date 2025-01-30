import os
import pandas as pd
from pathlib import Path

def read_excel_safe(file_path):
    """Safely read Excel files with error handling."""
    try:
        return pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return pd.DataFrame()

def categorize_files(folder_path):
    """Categorize files based on their names."""
    categories = {
        '11th': [],
        '12th': [],
        '11th_JEE': [],
        '11th_NEET': [],
        '12th_JEE': [],
        '12th_NEET': []
    }
    
    for file in Path(folder_path).glob('*.xlsx'):
        name = file.name.upper()
        if '11TH' in name and 'JEE' in name:
            categories['11th_JEE'].append(file)
        elif '11TH' in name and 'NEET' in name:
            categories['11th_NEET'].append(file)
        elif '12TH' in name and 'JEE' in name:
            categories['12th_JEE'].append(file)
        elif '12TH' in name and 'NEET' in name:
            categories['12th_NEET'].append(file)
        elif '11TH' in name:
            categories['11th'].append(file)
        elif '12TH' in name:
            categories['12th'].append(file)
    
    return categories

def combine_files(file_list):
    """Combine multiple Excel files into one DataFrame."""
    if not file_list:
        return pd.DataFrame()
    
    dfs = [read_excel_safe(f) for f in file_list]
    combined_df = pd.concat(dfs, ignore_index=True)
    return combined_df.drop_duplicates()

def process_folder(folder_path= "aligner"):
    """Main function to process the folder."""
    try:
        # Get categorized file lists
        categories = categorize_files(folder_path)
        
        # Define output mapping
        outputs = {
            '11th': '11th_combined.xlsx',
            '12th': '12th_combined.xlsx',
            '11th_JEE': '11th_JEE_combined.xlsx',
            '11th_NEET': '11th_NEET_combined.xlsx',
            '12th_JEE': '12th_JEE_combined.xlsx',
            '12th_NEET': '12th_NEET_combined.xlsx'
        }
        
        # Process each category
        for category, files in categories.items():
            if files:
                print(f"Processing {category} files...")
                combined_df = combine_files(files)
                if not combined_df.empty:
                    output_path = os.path.join(folder_path, outputs[category])
                    combined_df.to_excel(output_path, index=False)
                    print(f"Created {outputs[category]}")
                
    except Exception as e:
        print(f"Error processing folder: {e}")

if __name__ == "__main__":
    folder_path = "./aligner"
    process_folder(folder_path)