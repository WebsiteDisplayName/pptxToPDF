# https://stackoverflow.com/questions/31487478/how-to-convert-a-pptx-to-pdf-using-python

# https://stackoverflow.com/questions/20891787/an-efficient-way-to-convert-document-to-pdf-format


# https://youtu.be/mlOO9Rp5s8A
# Python 3 Script To Convert Single Or Multiple Powerpoint (.PPTX) Files to PDF Using ppt2pdf Module
# pip install ppt2pdf
# ppt2pdf dir ppt
# for the ppt directory or: ppt2pdf dir .
#  ^ for current working directory
# change drive https://stackoverflow.com/questions/11065421/command-prompt-wont-change-directory-to-another-drive


from tkinter.filedialog import askdirectory
import os
import win32com.client


def PPTtoPDF(inputFileName, outputFileName, formatType=32):
    powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
    powerpoint.Visible = 1

    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType)  # formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()


pptxFolder = askdirectory(title='Select Folder with .pptx')

for subdir, dirs, files in os.walk(pptxFolder):
    for i in files:
        if i.endswith(".pdf"):
            continue
        else:
            # pptxFilePath = os.path.join(pptxFolder, i)
            pptxFilePath = pptxFolder + "/" + i
            pptxFilePath = pptxFilePath.replace("/", "\\")
            # https://stackoverflow.com/questions/55227428/opening-a-powerpoint-presentation-saving-as-pdf-and-closing-the-application-usi
            # must use two backslashes as separation
            print(pptxFilePath)
            PPTtoPDF(pptxFilePath, pptxFilePath)

for subdir, dirs, files in os.walk(pptxFolder):
    for i in files:
        if i.endswith(".pdf"):
            os.rename(os.path.join(pptxFolder, i), os.path.join(
                pptxFolder, i.replace(".pptx", "")))
            continue
        elif i.endswith(".pptx"):
            os.remove(os.path.join(pptxFolder, i))
