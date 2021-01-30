import argparse
import glob
import os
from pathlib import Path
import logging as l

import win32com.client as win32client


def Docx2PDFConvert(input, output):
    wdFormatPDF = 17
    word = win32client.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(str(input))
    doc.SaveAs(str(output), FileFormat=wdFormatPDF)
    word.Quit()


def isSupportedDocFile(filepath):
    supportedDocFormat = ('.doc', '.docx')
    ext = os.path.splitext(str(filepath))[1].lower()
    return False if ext not in supportedDocFormat else True

def resolvePaths(file, outputExt, outputpath=''):
    filename = os.path.splitext(file)[0]
    absPath = Path(file).resolve()
    inputfile = str(absPath)
    if outputpath == '':
        outputfile = str(absPath.cwd()) + "\\" + os.path.splitext(file)[0] + outputExt
    else:
        outputfile = str(os.path.abspath(outputpath)) + "\\" + filename + outputExt
    return (inputfile, outputfile)

def argumentParserSetup():
    parser = argparse.ArgumentParser(description="PDF conversion tool")
    parser.add_argument('--input', action='store', type=str, nargs='+',required=True, help="input file path") #compulsory
    parser.add_argument('--output', action='store', type=str, help="output file path") #optional
    parser.add_argument('-outputname', action="store", type=str, help="set output file name")
    args = parser.parse_args()
    return args
'''
Returns input and output filepaths provided by the user.
incase of output is None an empty string '' will be returned
'''
def getArgumentValues(args):
    input = args.input
    output = args.output if args.output != None else ''
    return input, output

'''
Return the list of filepath. Uses glob() to solve wildcards.
'''

def validFilepath(inputfilenames):
    filepaths=[]
    for inputfile in inputfilenames:
        for file in glob.glob(str(inputfile)):
            filepaths.append(str(file))
    return filepaths

if __name__  ==  "__main__":
    args = argumentParserSetup()
    filepaths = validFilepath(getInputfilepaths(args=args))
    exit(1) if len(filepaths)  <=  0 else True
    val=filter(isSupportedDocFile,filepaths)

    for file in list(val):
        if not (args.output is None):
            inputfile, outputfile = resolvePaths(file,outputpath=args.output, outputExt=".pdf")
        else:
            inputfile, outputfile = resolvePaths(file,outputExt=".pdf")
        Docx2PDFConvert(input=str(inputfile), output=outputfile)
