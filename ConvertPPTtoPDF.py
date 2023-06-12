import os
import glob
from comtypes import client

def ppt_to_pdf(inputFileName, outputFileName, formatType = 32):
    powerpoint = client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType) # formatType = 32 for pdf
    deck.Close()
    powerpoint.Quit()

folderpath = "C:\\Users\\YourWindowsUserNameHere\\Desktop\\python\\convert_ppt_to_pdf"
files = glob.glob(folderpath + "\*.pptx")

for file in files:
    output_file = os.path.splitext(file)[0]
    ppt_to_pdf(file, output_file)
    os.remove(file) # Remove the file after conversion
