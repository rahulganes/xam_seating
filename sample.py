from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_ORIENT

year = input("Enter the year : ")
rollno = int(input("Enter the Starting Roll no :"))
filename = input("Enter Hall Name : ")

def merge_cell(cell1,cell2):
	cell1.merge(cell2)


doc = Document()
table = doc.add_table(rows = 7, cols = 10)
table.allow_autofit = False
table.style = 'TableGrid'

new_width = Inches(11.69)
new_height = Inches(8.27)

section = doc.sections[-1]
section.orientation = WD_ORIENT.LANDSCAPE
section.page_width = new_width
section.page_height = new_height


hdr_cells = table.rows[0].cells

merge_cell(table.rows[0].cells[0],table.rows[0].cells[1])
merge_cell(table.rows[0].cells[2],table.rows[0].cells[3])
merge_cell(table.rows[0].cells[4],table.rows[0].cells[5])
merge_cell(table.rows[0].cells[6],table.rows[0].cells[7])
merge_cell(table.rows[0].cells[8],table.rows[0].cells[9])

c = 1
for i in range(1,10):
	hdr_cells[i-1].text = 'Col'+ str(int(c))
	c+=0.5

	
for j in range(0,10,2):
	hdr_cells = table.columns[j].cells
	for i in range(1,7):
		hdr_cells[i].text = year+str(rollno).zfill(3)
		rollno+=1

doc.save(filename+'.docx')