import os
from faker import Faker
import pandas as pd
from docx import Document
from tqdm import tqdm

def setup_directories(base_path="Testing_PII_Data"):
    """
    Sets up base and format-specific directories for file export.
    :param base_path: The base directory for saving the files.
    :return: A tuple containing paths for Excel and Word directories.
    """
    excel_dir = os.path.join(base_path, "Excel Files")
    word_dir = os.path.join(base_path, "Word Files")

    # Create directories if they do not exist
    for directory in [excel_dir, word_dir]:
        if not os.path.exists(directory):
            os.makedirs(directory)

    return excel_dir, word_dir

def generate_data(record_count):
    """
    Generates a list of dictionaries, each representing a customer response.
    :param record_count: Number of records to generate.
    :return: List of dictionaries.
    """
    fake = Faker()
    data = [{
        'Full Name': fake.name(),
        'Address': fake.address().replace("\n", ", "),
        'Phone Number': fake.phone_number(),
        'Email Address': fake.email(),
        'Customer Feedback': fake.sentence()
    } for _ in tqdm(range(record_count), desc="Generating Data")]
    return data

def export_to_excel(data, filename, directory):
    """
    Exports the given data to an Excel file.
    :param data: List of dictionaries to export.
    :param filename: Filename for the Excel file.
    """
    df = pd.DataFrame(data)
    filepath = os.path.join(directory, filename)
    df.to_excel(filepath, index=False)
    print(f"Data exported to Excel file: {filepath}")

def export_to_word(data, filename, directory):
    """
    Exports the given data to a Word document.
    :param data: List of dictionaries to export.
    :param filename: Filename for the Word document.
    """
    doc = Document()
    for item in data:
        doc.add_paragraph(f"Full Name: {item['Full Name']}\n"
                          f"Address: {item['Address']}\n"
                          f"Phone Number: {item['Phone Number']}\n"
                          f"Email Address: {item['Email Address']}\n"
                          f"Customer Feedback: {item['Customer Feedback']}\n"
                          "----------------------------------------")
    filepath = os.path.join(directory, filename)
    doc.save(filepath)
    print(f"Data exported to Word file: {filepath}")

def user_input_options():
    print("Choose input option:")
    print("1. Specify number of records directly")
    print("2. Specify desired file size in MB")
    choice = input("Enter your choice (1 or 2): ").strip()

    if choice not in ['1', '2']:
        print("Invalid choice. Please enter 1 or 2.")
        return None, None  # Return None values to indicate error or cancellation

    file_naming_preference = input("Name the file based on: \n1. Number of Records\n2. File Size\nEnter choice (1 or 2): ").strip()

    try:
        if choice == '1':
            num_records = int(input("Enter the number of records you'd like to generate: "))
            file_label = f"{num_records}_records" if file_naming_preference == '1' else "custom_size"
        elif choice == '2':
            file_size_mb = float(input("Enter your desired file size in MB (e.g., 1.5 for 1.5MB): "))
            average_bytes_per_record = 450  # Adjust based on your data
            num_records = int((file_size_mb * 1024 * 1024) / average_bytes_per_record)
            file_label = f"{file_size_mb}MB" if file_naming_preference == '2' else f"{num_records}_records"
    except ValueError:
        print("Invalid number entered.")
        return None, None

    # Warning for large operations
    if num_records >= (11700):  # Roughly equivalent to 5MB based on the example average
        proceed = input("Warning: This operation might take a long time. Do you want to continue? (yes/no): ").strip().lower()
        if proceed != "yes":
            print("Operation canceled.")
            return None, None

    return num_records, file_label

# Example usage:
if __name__ == "__main__":

    num_records, file_label = user_input_options()
    if num_records is not None and file_label is not None:
        # Set up directories for file export
        excel_dir, word_dir = setup_directories()
        data = generate_data(num_records)
        export_to_excel(data, f'customer_responses_{file_label}_records.xlsx', excel_dir)
        export_to_word(data, f'customer_responses_{file_label}_records.docx', word_dir)
