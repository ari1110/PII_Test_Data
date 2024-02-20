import os
from faker import Faker
import pandas as pd
from docx import Document
from tqdm import tqdm
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def setup_directories(base_path="Testing_PII_Data"):
    """
    Sets up base and format-specific directories for file export.
    :param base_path: The base directory for saving the files.
    :return: A tuple containing paths for Excel and Word directories.
    """
    excel_dir = os.path.join(base_path, "Excel Files")
    word_dir = os.path.join(base_path, "Word Files")
    pdf_dir = os.path.join(base_path, "PDF Files")   # Adding PDF directory
    text_dir = os.path.join(base_path, "Text Files")  # Adding Text directory

    # Create directories if they do not exist
    for directory in [excel_dir, word_dir, pdf_dir, text_dir]:
        if not os.path.exists(directory):
            os.makedirs(directory)

    return excel_dir, word_dir, pdf_dir, text_dir  # Return all four directories

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
    
    filepath = os.path.join(directory, filename)
    counter = 1
    while os.path.exists(filepath):
        base, extension = os.path.splitext(filename)
        filepath = os.path.join(directory, f"{base}({counter}){extension}")
        counter += 1
    df = pd.DataFrame(data)
    df.to_excel(filepath, index=False)
    print(f"Data exported to Excel file: {filepath}")

def export_to_word(data, filename, directory):
    """
    Exports the given data to a Word document.
    :param data: List of dictionaries to export.
    :param filename: Filename for the Word document.
    """
    filepath = os.path.join(directory, filename)
    counter = 1
    while os.path.exists(filepath):
        base, extension = os.path.splitext(filename)
        filepath = os.path.join(directory, f"{base}({counter}){extension}")
        counter += 1
    doc = Document()
    for item in data:
        doc.add_paragraph(f"Full Name: {item['Full Name']}\n"
                          f"Address: {item['Address']}\n"
                          f"Phone Number: {item['Phone Number']}\n"
                          f"Email Address: {item['Email Address']}\n"
                          f"Customer Feedback: {item['Customer Feedback']}\n"
                          "----------------------------------------")
    doc.save(filepath)
    print(f"Data exported to Word file: {filepath}")

def export_to_pdf(data, filename, directory):
    c = canvas.Canvas(os.path.join(directory, filename), pagesize=letter)
    counter = 1
    while os.path.exists(os.path.join(directory, filename)):
        base, extension = os.path.splitext(filename)
        filename = f"{base}({counter}){extension}"
        counter += 1
    height = letter[1] - 30  # Start height for the first line
    for item in data:
        text = (f"Full Name: {item['Full Name']}\n" \
               f"Address: {item['Address']}\n" \
               f"Phone Number: {item['Phone Number']}\n" \
               f"Email Address: {item['Email Address']}\n" \
               f"Customer Feedback: {item['Customer Feedback']}\n" \
               "----------------------------------------")
        c.drawString(30, height, text)
        height -= 15 * 6  # Adjust height for next entry, assuming 6 lines per entry
        if height < 100:  # New page if less than 100 units of height remain
            c.showPage()
            height = letter[1] - 30
    c.save()
    print(f"Data exported to PDF file: {filename}")

def export_to_text(data, filename, directory):
    counter = 1
    while os.path.exists(os.path.join(directory, filename)):
        base, extension = os.path.splitext(filename)
        filename = f"{base}({counter}){extension}"
        counter += 1
    with open(os.path.join(directory, filename), 'w') as f:
        for item in data:
            f.write(f"Full Name: {item['Full Name']}\n"
                    f"Address: {item['Address']}\n"
                    f"Phone Number: {item['Phone Number']}\n"
                    f"Email Address: {item['Email Address']}\n"
                    f"Customer Feedback: {item['Customer Feedback']}\n"
                    "----------------------------------------\n")
    print(f"Data exported to text file: {filename}")

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
            file_size_mb = float(input("Enter your desired file size in MB (e.g., 1.5 for 1.5MB, or 0.5 for 500KB): "))
            average_bytes_per_record = 125  # Adjust based on your data
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

    # Ask for file types to generate
    print("Select the file types to generate:")
    print("0. Select All\n1. Excel\n2. Word\n3. PDF\n4. Text")
    file_types = input("Enter your choices separated by commas, or enter '0' to select all (e.g., 1,3 for Excel and PDF, or 0): ").strip()
    
    if file_types == '0':
        selected_file_types = [1, 2, 3, 4]
    else:
        selected_file_types = [int(choice.strip()) for choice in file_types.split(',') if choice.strip().isdigit() and int(choice.strip()) in range(1, 5)]

    return num_records, file_label, selected_file_types

# Example usage:
if __name__ == "__main__":
    num_records, file_label, selected_file_types = user_input_options()
    if num_records is not None and file_label is not None and selected_file_types:
        excel_dir, word_dir, pdf_dir, text_dir = setup_directories()  # Ensure it matches here
        data = generate_data(num_records)
        file_types_functions = {
            1: (export_to_excel, excel_dir, 'xlsx'),
            2: (export_to_word, word_dir, 'docx'),
            3: (export_to_pdf, pdf_dir, 'pdf'),  # Ensure this matches the updated setup_directories output
            4: (export_to_text, text_dir, 'txt')
        }
        
        for file_type in selected_file_types:
            if file_type in file_types_functions:
                export_func, directory, extension = file_types_functions[file_type]
                filename = f'customer_responses_{file_label}.{extension}'
                export_func(data, filename, directory)
