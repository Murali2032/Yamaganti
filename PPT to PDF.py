import os
import comtypes.client

# Directories
ppt_directory = r"C:\Users\YamagantiMuraliKrish\Music\October 2024"  # Input folder with PPT files
pdf_output_directory = r"C:\Users\YamagantiMuraliKrish\Music\PDF output"  # Output folder for PDF files

# Create the output directory if it doesn't exist
if not os.path.exists(pdf_output_directory):
    os.makedirs(pdf_output_directory)

# Function to convert PPT to Portrait A4 PDF
def convert_ppt_to_portrait_a4_pdf(ppt_file_path):
    try:
        # Generate the output PDF path
        pdf_file_path = os.path.join(pdf_output_directory, os.path.basename(ppt_file_path).replace('.pptx', '.pdf'))

        # Create PowerPoint application object
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1  # Run PowerPoint in the background

        # Open the PowerPoint file
        presentation = powerpoint.Presentations.Open(ppt_file_path)

        # Set slide size to A4 (Portrait: 21 cm x 29.7 cm)
        presentation.PageSetup.SlideOrientation = 1  # 1 = Portrait, 2 = Landscape
        presentation.PageSetup.SlideWidth = 21 * 28.35  # 21 cm in points
        presentation.PageSetup.SlideHeight = 29.7 * 28.35  # 29.7 cm in points

        # Save as PDF
        presentation.SaveAs(pdf_file_path, 32)  # 32 is the format type for PDF
        presentation.Close()
        powerpoint.Quit()

        print(f"Converted {ppt_file_path} to Portrait A4 PDF: {pdf_file_path}")

    except Exception as e:
        print(f"Failed to convert {ppt_file_path}: {e}")

# Loop through all the PPT files in the folder and subfolders
for root, dirs, files in os.walk(ppt_directory):
    for file in files:
        if file.endswith('.pptx'):  # Only process PPTX files
            ppt_file_path = os.path.join(root, file)
            convert_ppt_to_portrait_a4_pdf(ppt_file_path)

print("PPT to Portrait A4 PDF conversion completed for all files.")
