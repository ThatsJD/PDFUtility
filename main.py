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
        outputfile = str(absPath.cwd()) + "\\" + filename + outputExt
    else:
        outputfile = str(outputpath) + "\\" + filename + outputExt
    return (inputfile, outputfile)

def resolveOutputPath(outputPath):
    absPath = Path(outputPath).resolve()
    return str(absPath) if os.path.isdir(str(absPath)) else ''

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
    input, output = getArgumentValues(args)
    output = resolveOutputPath(output)
    filepaths = validFilepath(input)
    exit(1) if len(filepaths)  <=  0 else True
    files=filter(isSupportedDocFile,filepaths)
    for file in list(files):
        inputfile, outputfile = resolvePaths(file,outputpath=output,outputExt=".pdf")
        Docx2PDFConvert(input=str(inputfile), output=outputfile)
