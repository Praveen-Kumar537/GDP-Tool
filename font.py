from pdfreader import PDFDocument
fd = open("sample.pdf", "rb")
doc = PDFDocument(fd)


page = next(doc.pages())
Sorted = sorted(page.Resources.Font.keys())
print(Sorted)

font = page.Resources.Font['F1']
print(font.Subtype, font.BaseFont, font.Encoding)