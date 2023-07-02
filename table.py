from docx import Document

def read_tables_from_docx(file_path):
    doc = Document(file_path)
    tables = doc.tables

    empty_rows = []
    for table in tables:
        for row in table.rows:
            if all(cell.text.strip() == '' for cell in row.cells):
                empty_rows.append(row.cells)

    return empty_rows

# Example usage
docx_file = 'sample.docx'
empty_rows = read_tables_from_docx(docx_file)
for row in empty_rows:
    print([cell.text.strip() for cell in row])
