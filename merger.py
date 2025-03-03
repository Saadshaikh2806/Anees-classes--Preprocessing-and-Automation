import os
import pandas as pd

def clean_file_path(file_path):
    # Remove leading '&' and space if present (PowerShell drag and drop artifact)
    if file_path.startswith('& '):
        file_path = file_path[2:]
    
    # Remove quotes if present
    if (file_path.startswith('"') and file_path.endswith('"')) or \
       (file_path.startswith("'") and file_path.endswith("'")):
        file_path = file_path[1:-1]
    
    return file_path

def get_file_path(prompt):
    while True:
        file_path = input(prompt).strip()
        file_path = clean_file_path(file_path)
        
        if os.path.exists(file_path):
            return file_path
        else:
            print(f"File not found: {file_path}")
            print("Please try again or press Ctrl+C to exit.")

def merge_excel_files(file_paths, output_folder):
    all_dfs = []
    for file_path in file_paths:
        try:
            df = pd.read_excel(file_path)
            all_dfs.append(df)
            print(f"Added: {file_path}")
        except Exception as e:
            print(f"Error reading {file_path}: {str(e)}")
    
    if not all_dfs:
        print("No valid Excel files to merge. Exiting.")
        return
    
    # Merge all DataFrames
    merged_df = pd.concat(all_dfs, ignore_index=True)
    
    # Group by 'Name' and merge other columns
    merged_df = merged_df.groupby('Name', as_index=False).agg(lambda x: ', '.join(map(str, x.dropna().unique())))

    # Ensure output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Save the merged DataFrame to Excel
    output_path = os.path.join(output_folder, "merged_file.xlsx")
    merged_df.to_excel(output_path, index=False)
    print(f"Merged file saved as: {output_path}")

def main():
    print("Welcome to the Excel File Merger!")
    print("Please drag and drop your Excel files when prompted.")
    
    file_paths = []
    file_paths.append(get_file_path("Drag and drop the first Excel file here: "))
    file_paths.append(get_file_path("Drag and drop the second Excel file here: "))
    
    while True:
        another = input("Do you want to add another file? (yes/no): ").lower().strip()
        if another != 'yes':
            break
        file_paths.append(get_file_path("Drag and drop the next Excel file here: "))
    
    output_folder = "Output"
    merge_excel_files(file_paths, output_folder)

if __name__ == "__main__":
    main()


