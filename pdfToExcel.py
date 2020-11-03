import tabula
#tabula.convert_into("MSA_linje 1_1.pdf", "output2.csv", output_format="tsv", pages='all')

df = tabula.read_pdf('MSA_linje 1_1.pdf', pages = 'all', lattice = True)[0]
#df.head()
df.values = df.values.str.replace('.', ',')
print(df.values)
