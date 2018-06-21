print "Initialized"

from docx import Document
import re

filename = 'InputFiles/GST.docx'

document = Document('InputFiles/GST.docx')

print 'Tables:\t{}'.format(document.tables.__len__())

for table_index in range(document.tables.__len__()):

	print '\nTable:\t{}'.format(table_index)

	document_table = document.tables[table_index]

	print 'Rows:\t{}'.format(document_table.rows.__len__())
	table_rows = document_table.rows.__len__()
	print 'Cols:\t{}\n'.format(document_table.columns.__len__())
	table_cols = document_table.columns.__len__()

	table_text = ""
	for i in range(table_rows):
		for j in range(table_cols):

			try:
				print '({},{}):\t{}'.format(i,j,document_table.cell(i,j).text)
				table_text += document_table.cell(i,j).text
			except IndexError:
				print '\t\tIndex Error at ({},{})'.format(i,j)
				pass


	print '\nTable Text:\n{}'.format(table_text)

	gst_cell = document_table.cell(0,0).text
	gst_index = next(((index, i) for (index,i) in enumerate(gst_cell.split()) if re.search(r'GSTIN', gst_cell.split()[index], flags=re.M|re.S) is not None), (-1,-1))[0]

	if gst_index != -1:
		print '\nGST Index:\t{}'.format(gst_index)
		print '\nGST Number:\t{}'.format(gst_cell.split()[gst_index + 1])
	else: 
		print 'GST Index not found in {}'.format(filename)

# =====================================
# FUNCTIONS
# =====================================