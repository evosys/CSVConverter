import tabula

df = tabula.read_pdf("test.pdf", spreadsheet=True)

print(df)