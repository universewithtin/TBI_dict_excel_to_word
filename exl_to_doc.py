from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement

# Python-docx does not support font spacing, so here we go, using StackOverflow code
def run_set_spacing(run, value: int):
    """Set the font spacing for `run` to `value` in twips.

    A twip is a "twentieth of an imperial point", so 1/1440 in.
    """

    def get_or_add_spacing(rPr):
        # --- check if `w:spacing` child already exists ---
        spacings = rPr.xpath("./w:spacing")
        # --- return that if so ---
        if spacings:
            return spacings[0]
        # --- otherwise create one ---
        spacing = OxmlElement("w:spacing")
        rPr.insert_element_before(
            spacing,
            *(
                "w:w",
                "w:kern",
                "w:position",
                "w:sz",
                "w:szCs",
                "w:highlight",
                "w:u",
                "w:effect",
                "w:bdr",
                "w:shd",
                "w:fitText",
                "w:vertAlign",
                "w:rtl",
                "w:cs",
                "w:em",
                "w:lang",
                "w:eastAsianLayout",
                "w:specVanish",
                "w:oMath",
            ),
        )
        return spacing

    rPr = run._r.get_or_add_rPr()
    spacing = get_or_add_spacing(rPr)
    spacing.set(qn('w:val'), str(value))


# Load the Excel file. LOCAL VERSION, not TG
workbook = load_workbook('test.xlsx')
sheet = workbook.active

# Create a Word document
doc = Document()
doc.styles['Normal'].font.name = 'Times New Roman'
doc.styles['Normal'].font.size = Pt(14)
    
# Iterate through rows in Excel and add to the Word document
if __name__ == '__main__':
	table = doc.add_table(rows=1, cols=1)
	table.style = 'Table Grid'
	cell = table.cell(0, 0)
	cell.text = str(sheet['B1'].value)
	cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
	doc.add_paragraph()
	
	for row in sheet.iter_rows(min_row=3, values_only=True):
	    if row[0] is None:
	        continue
	    paragraph = doc.add_paragraph()
	    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
	
	    for i, cell_value in enumerate(row):
	        run = paragraph.add_run(str(cell_value))
		        
	        if i == 0:
	            run.font.bold = True
	            run.font.all_caps = True
	            if not str(cell_value).endswith(' '):
	                run.text += ' '  
	        elif i == 1:
	            if not str(cell_value).endswith(' '):
	                run.text += ' '  
	            run_set_spacing(run, 60)
	        elif i == 2:
	            run.italic = True
	            if not str(cell_value).endswith('.'):
	                run.text += '. '  
	        else:
	            if not str(cell_value).endswith(' '):
	                run.text += ' ' 

# LOCAL VERSION, not sending name here yet. 
doc.save('test.docx')