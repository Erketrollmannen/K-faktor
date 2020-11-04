import tabula
import pandas as pd

#tabula.convert_into("MSA_linje 1_1.pdf", "output2.csv", output_format="tsv", pages='all')

df = tabula.read_pdf('MSA_linje 1_1.pdf', pages = 'all')[0]
#df.values = [x.replace(',', '.') for x in df[values]]
#df.to_tsv('test.csv', sep=' ', encoding='utf-8', quotechar='"', decimal=',')
print(df.values)

