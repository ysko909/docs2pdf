import sys
import comtypes.client
import glob
import pathlib
import PyPDF2
import time

start = time.time()

wdFormatPDF = 17


def convert(in_file, out_file):
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


def pdf_merger(out_pdf, pdfs):
    merger = PyPDF2.PdfFileMerger()

    for pdf in pdfs:
        print(pdf)
        merger.append(pdf)

    merger.write(out_pdf)
    merger.close()


argvs = sys.argv
arg_count = len(argvs)

if arg_count > 1:
    file_path = argvs[1]
else:
    file_path = 'C:\\app\\work\\docs\\*.docx'

parent_folder = pathlib.Path(file_path).parent
files = glob.glob(file_path)

pdfs = []

for f in files:
    file_p = pathlib.Path(f)
    file_pdf = f.replace(file_p.suffix, '.pdf')
    pdfs.append(file_pdf)
    convert(f, file_pdf)

out_file = str(pathlib.Path(parent_folder).joinpath('out.pdf'))

pdf_merger(out_file, pdfs)

process_time = time.time() - start

print(f'Process time is : {process_time}')
