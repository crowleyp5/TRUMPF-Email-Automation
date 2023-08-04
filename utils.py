import PyPDF2 as pdf2
import os

def get_pdf_content(attachment):
    # We need to extract text from pdfs quite a bit. This saves a temporary file
    # for the pdf to be read, and then the file is removed
    if attachment.FileName.endswith('.pdf'):
        temp_file_path = save_attachment_to_tempfile(attachment)
        with open(temp_file_path, 'rb') as pdf_file:
            pdf_reader = pdf2.PdfReader(pdf_file)
            content = ''
            for page in pdf_reader.pages:
                content += page.extract_text()
        os.remove(temp_file_path)
    return content

def save_attachment_to_tempfile(attachment):
    # Saving temporary files to read pdfs
    temp_dir = os.path.join(os.path.expanduser('~'), 'temp_attachments')
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    temp_file_path = os.path.join(temp_dir, attachment.FileName)
    attachment.SaveAsFile(temp_file_path)
    return temp_file_path
