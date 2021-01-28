import argparse
import glob
import os
from pathlib import Path

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
    ext = os.path.splitext(file)[1].lower()
    return False if ext not in supportedDocFormat else True

def resolvePaths(file, outputExt, outputpath=''):
    filename = os.path.splitext(file)[0]
    absPath = Path(file).resolve()
    inputfile = str(absPath)
    if outputpath == '':
        outputfile = str(absPath.cwd()) + "\\" + os.path.splitext(file)[0] + outputExt
    else:
        outputfile = str(os.path.abspath(outputpath)) + "\\" + filename + outputExt
        print(outputfile, "Check this value")
        print("This path exsists YEAH", os.path.exists(os.path.abspath(outputpath)))
    return (inputfile, outputfile)

def argumentParserSetup():
    parser = argparse.ArgumentParser(description="PDF conversion tool")
    parser.add_argument('--input', action='store', type=str, nargs='+',required=True, help="input file path") #compulsory
    parser.add_argument('--output', action='store', type=str, help="output file path") #optional
    args = parser.parse_args()
    print(args)
    return args

if __name__  ==  "__main__":
    args = argumentParserSetup()
    print("Path exsits", os.path.exists(str(Path(args.output).resolve()))) #important ".\hello\hello" Check for All
    print("Abs path", os.path.abspath(args.output))
    print("is abs path", os.path.isabs(os.path.abspath(args.output)))
    print("path resolve", str(Path(args.output).resolve()))
    print(args.input[0])
    for file in glob.glob(args.input[1]):
        if not isSupportedDocFile(file):
            continue
        if not (args.output is None):
            inputfile, outputfile = resolvePaths(file,outputpath=args.output, outputExt=".pdf")
            print(outputfile, args.output, "outname is provided")
        else:
            inputfile, outputfile = resolvePaths(file,outputExt=".pdf")
            print(outputfile, "output name is not provided")
        print("The final path: ", outputfile)
        Docx2PDFConvert(input=str(inputfile), output=outputfile)
