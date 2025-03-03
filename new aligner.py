import pandas as pd
import os


def load_data(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ['.xls', '.xlsx']:
        try:
            return pd.read_excel(file_path)
        except Exception as e:
            raise ValueError(f"Error loading {file_path}: {e}")
    elif ext == '.csv':
        return pd.read_csv(file_path)
    else:
        raise ValueError(f"Unsupported file format: {file_path}")


def filter_columns(df, is_gat_exam=False):
    if is_gat_exam:
        allowed_columns = {'Name', 'Exam', 'ENGLISH', 'GAT'}
    else:
        allowed_columns = {'Name', 'Exam', 'Total Marks'}
    return df[[col for col in df.columns if col in allowed_columns]]


def gather_contacts(contact_folder):
    contacts_df = pd.DataFrame()
    for filename in os.listdir(contact_folder):
        file_path = os.path.join(contact_folder, filename)
        if os.path.isdir(file_path) or filename.startswith('~') or filename.startswith('.'):
            continue
        try:
            df = load_data(file_path)
            required_columns = ['Name', 'Student Contact No.', 'Father/Guardian Contact No.', 'Mother/Guardian Contact No.']
            if all(column in df.columns for column in required_columns):
                contacts_df = pd.concat([contacts_df, df[required_columns]], ignore_index=True)
            else:
                print(f"File '{filename}' does not contain the required columns and will be skipped.")
        except Exception as e:
            print(f"Error processing file '{filename}': {e}")

    contacts_df.drop_duplicates(subset=['Name'], keep='last', inplace=True)
    return contacts_df


def merge_exam_data(merged_df, exam_file, exam_index, is_gat_exam=False):
    exam_df = load_data(exam_file)
    if 'Name' not in exam_df.columns:
        raise ValueError(f"Missing 'Name' column in {exam_file}")

    # Filter and rename columns for merging
    exam_df = filter_columns(exam_df, is_gat_exam)
    
    if is_gat_exam:
        new_columns = {
            'Exam': f'Exam{exam_index}',
            'ENGLISH': f'ENGLISH{exam_index}',
            'GAT': f'GAT{exam_index}'
        }
    else:
        new_columns = {
            'Exam': f'Exam{exam_index}',
            'Total Marks': f'Total Marks{exam_index}'
        }
    
    exam_df = exam_df.rename(columns=new_columns)

    # Merge data and remove duplicate names
    merged_df = pd.merge(merged_df, exam_df[['Name'] + list(new_columns.values())], on='Name', how='outer')
    merged_df.drop_duplicates(subset=['Name'], keep='last', inplace=True)
    return merged_df


def merge_by_keyword(folder, grade, keywords, output_folder):
    merged_df = pd.DataFrame(columns=['Name'])
    exam_index = 1

    grade_files = [file for file in os.listdir(folder) if grade in file.upper()]
    for keyword in keywords:
        keyword_files = [os.path.join(folder, file) for file in grade_files if keyword in file.upper()]
        for file_path in keyword_files:
            print(f"Processing {file_path}...")
            try:
                # Check if current file is a GAT exam
                is_gat_exam = 'GAT' in keyword.upper()
                merged_df = merge_exam_data(merged_df, file_path, exam_index, is_gat_exam)
                exam_index += 1
            except Exception as e:
                print(f"Error processing file {file_path}: {e}")

    if not merged_df.empty:
        output_file = os.path.join(output_folder, f"{grade}_{'_'.join(keywords).upper()}_merged.xlsx")
        merged_df.to_excel(output_file, index=False)
        print(f"Merged data for {grade} {' and '.join(keywords)} saved to {output_file}")
        return output_file
    else:
        print(f"No data found for {grade} {' and '.join(keywords)} in the folder.")
        return None


def append_all_files(files_to_append, output_folder, contact_folder):
    appended_df = pd.DataFrame()
    for file in files_to_append:
        try:
            df = pd.read_excel(file)
            appended_df = pd.concat([appended_df, df], ignore_index=True)
        except Exception as e:
            print(f"Error appending file {file}: {e}")

    # Merge with contact details
    try:
        contacts_df = gather_contacts(contact_folder)
        appended_df = pd.merge(
            appended_df,
            contacts_df,
            on='Name',
            how='left'
        )
        print("Contacts successfully merged into appended data.")
    except Exception as e:
        print(f"Error merging contacts: {e}")

    # Remove duplicate names from the final merged DataFrame
    appended_df.drop_duplicates(subset=['Name'], keep='last', inplace=True)

    if not appended_df.empty:
        output_file = os.path.join(output_folder, "all_merged_data.xlsx")
        appended_df.to_excel(output_file, index=False)
        print(f"Appended data saved to {output_file}")
    else:
        print("No data to append.")


def auto_merge_evalbee(folder, contact_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)

    all_files = []
    grades = ['11TH', '12TH']
    exam_types = [['GAT', 'MATHS', 'MHTCET'], ['JEE'], ['NEET']]

    for grade in grades:
        for keywords in exam_types:
            merged_file = merge_by_keyword(folder, grade, keywords, output_folder)
            if merged_file:
                all_files.append(merged_file)

    # Append all files if user agrees
    append_choice = input("Do you want to append all merged files into one? (yes/no): ").strip().lower()
    if append_choice == "yes":
        append_all_files(all_files, output_folder, contact_folder)


def main():
    evalbee_folder = 'Evalbee'
    contact_folder = 'Data'
    output_folder = 'Output'

    try:
        auto_merge_evalbee(evalbee_folder, contact_folder, output_folder)
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()