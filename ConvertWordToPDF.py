from pdfminer.high_level import extract_text
import os

def pdf_to_text(filename):
    text = extract_text(filename)
    return text

# Define the directory where the PDF file is located
directory = "C:\\Users\\YourWindowsUsernameHere\\Desktop\\python\\convert_pfd_to_word"

filename = os.path.join(directory, 'PDF_Files.pdf')  # replace with your PDF file name
text = pdf_to_text(filename)

# Remove extra spaces and empty lines
lines = text.split('\n')
lines = [line.strip() for line in lines if line.strip() != '']
text = '\n'.join(lines)

output_file = os.path.join(directory, 'output.txt')
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(text)
