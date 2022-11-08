# https://stackoverflow.com/questions/31487478/how-to-convert-a-pptx-to-pdf-using-python
import win32com.client
from tkinter.filedialog import askdirectory
import os
from pathlib import Path


def PPTtoPDF(inputFileName, outputFileName, formatType=32):
    powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
    powerpoint.Visible = 1

    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType)  # formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()


pptxFolder = askdirectory(title='Select Folder To Save Songs In')

for subdir, dirs, files in os.walk(pptxFolder):
    for i in files:
        if i.endswith(".pdf"):
            continue
        else:
            fileName = Path(i).stem
            PPTtoPDF(fileName, fileName)

    for i in files:
        if i.endswith(".pdf"):
            continue
        else:
            os.remove(os.path.join(pptxFolder, i))
