from PyPDF2 import PdfReader, PdfWriter

def merge_pdfs(pdf_list, output_filename):
    """
    Merge multiple PDF files from the given list and save as a new PDF file.

    Parameters:
        pdf_list (list): A list of file paths to the PDF files to be merged.
        output_filename (str): The output file path and name for the merged PDF.

    Example:
        merge_pdfs(["out1.pdf", "out2.pdf"], "merged_output.pdf")
    """
    # Create a PdfWriter object to store the merged PDF content
    output = PdfWriter()
    for pdf_path in pdf_list:
        # Open each PDF file and read all its pages
        with open(pdf_path, "rb") as pdf_file:
            pdf_reader = PdfReader(pdf_file)
            # Iterate through each page of the current PDF and add it to the output
            for page_num in range(len(pdf_reader.pages)):
                output.add_page(pdf_reader.pages[page_num])

    with open(output_filename, "wb") as merged_file:
        output.write(merged_file)
    print(f"PDF merging completed. The merged file has been saved as '{output_filename}'.")

pdf_files = ["259096_updated.pdf", "300600_updated.pdf"]
merge_pdfs(pdf_files, "final_merged_output.pdf")
