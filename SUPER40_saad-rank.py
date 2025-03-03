import pandas as pd
import sys
import os
from collections import defaultdict
import traceback

def process_exam_data(file_path):
    # Determine file type and read accordingly
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    elif file_path.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file format. Please use CSV or Excel file.")

    print("Columns in the file:")
    print(df.columns.tolist())

    print("\nFirst few rows of the data:")
    print(df.head())

    # Function to get top 4 scores
    def get_top_4_total(row, columns):
        scores = []
        for col in columns:
            if pd.notna(row[col]):
                # Check if the score is in fractional format
                if '/' in str(row[col]):
                    numerator, _ = map(float, str(row[col]).split('/'))
                    scores.append(numerator)
                else:
                    scores.append(float(row[col]))
        return sum(sorted(scores, reverse=True)[:4])

    # Separate GAT and MATHS columns
    gat_columns = [col for col in df.columns if 'Total Marks' in col and 'GAT' in str(df[col.replace('Total Marks', 'Exam')].iloc[0])]
    maths_columns = [col for col in df.columns if 'Total Marks' in col and 'MATHS' in str(df[col.replace('Total Marks', 'Exam')].iloc[0])]

    print("\nGAT columns:", gat_columns)
    print("MATHS columns:", maths_columns)

    # Process GAT Marks
    df['Total GAT (Top 4)'] = df.apply(lambda row: get_top_4_total(row, gat_columns), axis=1)

    # Process MATHS Marks
    df['Total MATHS (Top 4)'] = df.apply(lambda row: get_top_4_total(row, maths_columns), axis=1)

    # Print sample results for debugging
    print("\nSample results:")
    print(df[['Total GAT (Top 4)', 'Total MATHS (Top 4)']].head())

    # Process Subjective Marks
    subj_columns = [f'Subjective Marks{i}' for i in range(1, 6)]  # Assuming there are 5 subjective exams
    df['Total Subjective (Top 4)'] = df.apply(lambda row: get_top_4_total(row, subj_columns), axis=1)

    # Calculate Grand Total
    df['Grand Total'] = df['Total GAT (Top 4)'] + df['Total MATHS (Top 4)'] + df['Total Subjective (Top 4)']

    # Calculate Average Score
    df['Average Score'] = df['Grand Total'] / 3  # Equal weight for GAT, MATHS, and Subjective

    return df

def select_top_students(df, total_students=40, exclude_class="XI- JEE/NEET"):
    # Exclude the specified class
    df_filtered = df[df['CLASS'] != exclude_class]

    # Group students by class
    class_groups = df_filtered.groupby('CLASS')
    
    # Calculate the number of students to select from each class
    class_sizes = class_groups.size()
    total_class_size = class_sizes.sum()
    students_per_class = (class_sizes / total_class_size * total_students).round().astype(int)
    
    # Adjust if the total is not exactly 40
    while students_per_class.sum() != total_students:
        if students_per_class.sum() < total_students:
            students_per_class[students_per_class.idxmin()] += 1
        else:
            students_per_class[students_per_class.idxmax()] -= 1
    
    # Select top students from each class
    selected_students = []
    for class_name, group in class_groups:
        top_n = students_per_class[class_name]
        top_students = group.nlargest(top_n, 'Average Score')
        selected_students.append(top_students)
    
    # Combine all selected students
    result = pd.concat(selected_students)
    return result.sort_values('Average Score', ascending=False)

def clean_file_path(file_path):
    # Remove leading '&' if present
    file_path = file_path.lstrip('&').strip()
    
    # Remove surrounding quotes if present
    file_path = file_path.strip("'\"")
    
    # Replace escaped spaces if present
    file_path = file_path.replace("\\ ", " ")
    
    return file_path

def get_file_path():
    if len(sys.argv) > 1:
        # If file path is provided as command-line argument
        return clean_file_path(' '.join(sys.argv[1:]))
    else:
        # Ask user to drag and drop the file
        print("Please drag and drop your CSV or Excel file into the console and press Enter:")
        file_path = input().strip()
        return clean_file_path(file_path)

def main():
    file_path = get_file_path()
    
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' does not exist.")
        print("Please make sure you've entered the correct path and the file exists.")
        return

    try:
        df = process_exam_data(file_path)
        top_students = select_top_students(df)
        
        print("Super 40 students (excluding XI- JEE/NEET class):")
        print(top_students[['NAME', 'CLASS', 'Total GAT (Top 4)', 'Total MATHS (Top 4)', 'Total Subjective (Top 4)', 'Grand Total', 'Average Score']])
        
        # Save the processed data to a new file
        output_file = os.path.splitext(file_path)[0] + "_super40_students.xlsx"
        top_students.to_excel(output_file, index=False)
        print(f"\nSuper 40 students data saved to: {output_file}")

        # Print excluded class data separately
        excluded_class_data = df[df['CLASS'] == "XI- JEE/NEET"].sort_values('Average Score', ascending=False)
        print("\nXI- JEE/NEET class data (excluded from Super 40):")
        print(excluded_class_data[['NAME', 'CLASS', 'Total GAT (Top 4)', 'Total MATHS (Top 4)', 'Total Subjective (Top 4)', 'Grand Total', 'Average Score']])

        # Save excluded class data to a separate file
        excluded_output_file = os.path.splitext(file_path)[0] + "_XI_JEE_NEET_students.xlsx"
        excluded_class_data.to_excel(excluded_output_file, index=False)
        print(f"\nXI- JEE/NEET class data saved to: {excluded_output_file}")

    except Exception as e:
        print(f"An error occurred while processing the file: {str(e)}")
        print("Error details:")
        traceback.print_exc()
        print("\nPlease ensure that the file is not open in another program and that you have permission to access it.")
        print("Also, check if the file structure matches the expected format.")

if __name__ == "__main__":
    main()