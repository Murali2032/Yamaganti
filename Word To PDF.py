import comtypes.client
import os


def word_to_pdf(input_file):
    # Define the output PDF path as "INVOICE.pdf" in the same folder
    output_file = os.path.join(os.path.dirname(input_file), "INVOICE.pdf")

    # Initialize the Word application
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False  # Run in the background without opening Word

    try:
        # Open the Word document
        doc = word.Documents.Open(input_file)

        # Save the document as a PDF with the fixed name "INVOICE.pdf"
        doc.SaveAs(output_file, FileFormat=17)  # 17 is the code for PDF format
        print(f"Converted: {input_file} → {output_file}")
    except Exception as e:
        print(f"Error converting {input_file}: {e}")
    finally:
        # Clean up resources
        doc.Close()
        word.Quit()


def convert_all_docs_in_directory(input_dir):
    # Check if the specified directory exists
    if not os.path.exists(input_dir):
        print(f"Directory not found: {input_dir}")
        return

    # Iterate over all files in the directory
    for file in os.listdir(input_dir):
        # Process only Word documents (.doc and .docx)
        if file.lower().endswith(('.doc', '.docx')):
            input_path = os.path.join(input_dir, file)
            word_to_pdf(input_path)


# Specify the input directory
input_directory = r"C:\Users\YamagantiMuraliKrish\Downloads\Invoices"

# Convert all Word files in the directory
convert_all_docs_in_directory(input_directory)
