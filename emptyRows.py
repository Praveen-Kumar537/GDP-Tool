from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

# Load the Word document
doc = Document('sample.docx')

# Iterate over the tables in the document
for table in doc.tables:
    # Process each row in the table
    for row in table.rows:
        # Process each cell in the row
        for cell in row.cells:
            # Check if the cell is empty
            if cell.text.strip() == "":
                # Set the background color of the empty cell to yellow
                shading = parse_xml('<w:shd {} w:fill="FFFF00"/>'.format(nsdecls('w')))
                cell._tc.get_or_add_tcPr().append(shading)

# Save the modified document with highlighted empty cells
doc.save('modified_document.docx')
