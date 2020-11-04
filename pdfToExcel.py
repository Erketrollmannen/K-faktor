import tabula
import pandas as pd
import os


# -----[ Global Variables ]-----
pdffiles = []
directory = ("F:\\Scripts\\Kfaktorlogg\\MSA_1")

# Make list of pdf files
for fileName in os.listdir(directory):
    if fileName.endswith(".pdf"):
        pdffiles.append(fileName)

print(pdffiles)
#tabula.convert_into(pdffiles, directory, output_format="tsv", pages='all')
#print(path)

#tabula.convert_into_by_batch(inputDir, output_format="tsv", pages='all')

#df = tabula.read_pdf('MSA_linje 1_1.pdf', pages = 'all')[0]
#df.values = [x.replace(',', '.') for x in df[values]]
#df.to_tsv('test.csv', sep=' ', encoding='utf-8', quotechar='"', decimal=',')
#print(df.values)

