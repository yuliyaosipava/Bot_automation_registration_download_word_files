import os
import win32com.client as win32

def convert_rtf_to_docx(input_dir, output_dir):
    word = win32.Dispatch("Word.Application")
    for filename in os.listdir(input_dir):
        if filename.endswith('.rtf'):
            input_file = os.path.join(input_dir, filename)
            output_file = os.path.join(output_dir, filename.replace('.rtf', '.docx'))
            doc = word.Documents.Open(input_file)
            doc.SaveAs(output_file, FileFormat=16)  # FileFormat=16 means docx format
            doc.Close()
    word.Quit()

input_dir = 'c:\\Users\\user\\Documents\\doci\\docs\\'
output_dir = 'c:\\Users\\user\\Documents\\doci\\docs\\'
convert_rtf_to_docx(input_dir, output_dir)

