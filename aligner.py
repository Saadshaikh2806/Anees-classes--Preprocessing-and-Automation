import pandas as pd
import os

def load_data(file_path):
    file_path = clean_file_path(file_path)
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ['.xls', '.xlsx']:
        return pd.read_excel(file_path)
    elif ext == '.csv':
        return pd.read_csv(file_path)
    else:
        raise ValueError("Unsupported file format. Please provide a .csv or .xls/.xlsx file.")

def filter_columns(df):
    # Define allowed columns including up to 100 exams and their total marks
    allowed_columns = {'Roll No', 'Name', 'Student Contact No.', 'Father/Guardian Contact No.', 'Mother/Guardian Contact No.'}
    allowed_columns.update({f'Exam{i}' for i in range(1, 101)})
    allowed_columns.update({f'Total Marks{i}' for i in range(1, 101)})
    
    # Filter the DataFrame to keep only the allowed columns
    filtered_df = df[[col for col in df.columns if col in allowed_columns]]
    return filtered_df

def rename_exam_columns(df):
    if 'Exam' in df.columns and 'Total Marks' in df.columns:
        df = df.rename(columns={'Exam': 'Exam1', 'Total Marks': 'Total Marks1'})
    return df

def gather_contacts(data_folder):
    contacts_df = pd.DataFrame()
    for filename in os.listdir(data_folder):
        file_path = os.path.join(data_folder, filename)
        if os.path.isdir(file_path) or filename.startswith('~') or filename.startswith('.'):
            continue
        try:
            current_df = load_data(file_path)
            current_df = filter_columns(current_df)
            required_columns = ['Name', 'Student Contact No.', 'Father/Guardian Contact No.', 'Mother/Guardian Contact No.']
            if all(column in current_df.columns for column in required_columns):
                contacts_df = pd.concat([contacts_df, current_df[required_columns]], ignore_index=True)
            else:
                print(f"File '{filename}' does not contain the required columns and will be skipped.")
        except Exception as e:
            print(f"Error processing file '{filename}': {e}")
    contacts_df.drop_duplicates(subset=['Name'], keep='last', inplace=True)
    return contacts_df

def merge_contacts(names_file, contacts_df, output_file):
    names_df = load_data(names_file)
    names_df = rename_exam_columns(names_df)
    names_df = filter_columns(names_df)
    
    merged_df = names_df.merge(
        contacts_df[['Name', 'Student Contact No.', 'Father/Guardian Contact No.', 'Mother/Guardian Contact No.']],
        on='Name', how='left'
    )
    
    output_dir = 'Output'
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, output_file)
    
    if output_file.lower().endswith('.xlsx'):
        merged_df.to_excel(output_path, index=False)
    elif output_file.lower().endswith('.xls'):
        merged_df.to_excel(output_path, index=False, engine='xlwt')
    else:
        merged_df.to_csv(output_path, index=False)
    
    print(f"Data has been merged and saved to {output_path}")
    return merged_df

def merge_exam_data(merged_df, exam_file, exam_index):
    exam_df = load_data(exam_file)
    print(f"Exam file columns: {exam_df.columns}")
    required_columns = ['Name', 'Exam', 'Total Marks']
    for column in required_columns:
        if column not in exam_df.columns:
            raise ValueError(f"Column '{column}' not found in the exam file '{exam_file}'.")
    new_columns = {
        'Exam': f'Exam{exam_index}',
        'Total Marks': f'Total Marks{exam_index}'
    }
    exam_df = exam_df.rename(columns=new_columns)
    merged_df = merged_df.merge(exam_df[['Name'] + list(new_columns.values())], on='Name', how='left')
    
    # Ensure proper column ordering
    exam_columns = [col for col in merged_df.columns if col.startswith('Exam') or col.startswith('Total Marks')]
    exam_columns_sorted = sorted(exam_columns, key=lambda x: (int(''.join(filter(str.isdigit, x))), x))
    
    contact_columns = ['Student Contact No.', 'Father/Guardian Contact No.', 'Mother/Guardian Contact No.']
    other_columns = ['Roll No', 'Name'] + [col for col in merged_df.columns if col not in exam_columns + contact_columns + ['Roll No', 'Name']]
    merged_df = merged_df[other_columns + exam_columns_sorted + contact_columns if any(col in merged_df.columns for col in contact_columns) else []]
    
    return merged_df

def clean_file_path(file_path):
    # Remove leading & trailing spaces and special characters
    cleaned_path = file_path.strip().strip('\'"& ')
    return cleaned_path

def main():
    data_folder = 'Data'
    names_file = input("Drag and drop the names file into the terminal or type 'next' to skip: ").strip()
    names_file = clean_file_path(names_file)
    
    if names_file.lower() != 'next':
        output_file = input("Enter the name for the output file (include .csv or .xlsx extension): ").strip()
        output_file = clean_file_path(output_file)
        try:
            contacts_df = gather_contacts(data_folder)
            merged_df = merge_contacts(names_file, contacts_df, output_file)
        except Exception as e:
            print(f"Error: {e}")
            return
    else:
        names_file = input("Drag and drop the main file to be updated into the terminal: ").strip()
        names_file = clean_file_path(names_file)
        names_df = load_data(names_file)
        names_df = rename_exam_columns(names_df)
        names_df = filter_columns(names_df)
        if 'Exam1' not in names_df.columns or 'Total Marks1' not in names_df.columns:
            print("Please add a file with the proper columns 'Exam' or 'Exam1' and 'Total Marks1'.")
            return
        output_file = input("Enter the name for the output file (include .csv or .xlsx extension): ").strip()
        output_file = clean_file_path(output_file)
        merged_df = names_df
        exam_index = len([col for col in merged_df.columns if col.startswith('Exam')]) + 1

        exam_file = input("Drag and drop the exam file into the terminal: ").strip()
        exam_file = clean_file_path(exam_file)
        try:
            merged_df = merge_exam_data(merged_df, exam_file, exam_index)
            if output_file.lower().endswith('.xlsx'):
                merged_df.to_excel(os.path.join('Output', output_file), index=False)
            elif output_file.lower().endswith('.xls'):
                merged_df.to_excel(os.path.join('Output', output_file), index=False, engine='xlwt')
            else:
                merged_df.to_csv(os.path.join('Output', output_file), index=False)
            print(f"Exam data has been merged and saved to {os.path.join('Output', output_file)}")
        except Exception as e:
            print(f"Error processing exam file '{exam_file}': {e}")
            return
    
    exam_index = len([col for col in merged_df.columns if col.startswith('Exam')]) + 1

    while True:
        more_exam_data = input("Do you want to add data for another exam? (yes/no): ").strip().lower()
        if more_exam_data == 'yes':
            exam_file = input("Drag and drop the exam file into the terminal: ").strip()
            exam_file = clean_file_path(exam_file)
            try:
                merged_df = merge_exam_data(merged_df, exam_file, exam_index)
                if output_file.lower().endswith('.xlsx'):
                    merged_df.to_excel(os.path.join('Output', output_file), index=False)
                elif output_file.lower().endswith('.xls'):
                    merged_df.to_excel(os.path.join('Output', output_file), index=False, engine='xlwt')
                else:
                    merged_df.to_csv(os.path.join('Output', output_file), index=False)
                print(f"Exam data has been merged and saved to {os.path.join('Output', output_file)}")
                
                exam_index += 1
            except Exception as e:
                print(f"Error processing exam file '{exam_file}': {e}")
        elif more_exam_data == 'no':
            break
        else:
            print("Invalid input. Please enter 'yes' or 'no'.")

if __name__ == "__main__":
    main()