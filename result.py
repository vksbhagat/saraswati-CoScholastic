#!/bin/bash
#
#Encoding Unicode (UTF-8)
#
#Author
#   Vivek Bhagat
#GNU General Public Licence v3.0
#Copyright (c) 2017 Vivek Bhagat
#

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
#from docx.enum.style import WD_STYLE_TYPE
from docx.shared import RGBColor

#Creating an object for Document()
result = Document()

#Adding a picture school letter head
#result.add_picture('monty-truth.png', width=Inches(1.25))

#Modifying the page layout
sections = result.sections[0]
sections.top_margin = Inches(0)
sections.bottom_margin = Inches(0)
#sections.header_distance = Pt(0)
#sections.footer_distance = Pt(0)

#Heading
heading = result.add_paragraph()
heading.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
heading.add_run('Co-Scholastic Grade Certificate Class X 2017').bold = True
#head1=result.add_heading('Co-Scholastic Grade Certificate Class X 2017',3).paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

#Body setting
paragraph_format = result.styles['Normal'].paragraph_format
paragraph_format.space_before = Pt(0)
paragraph_format.line_spacing = Pt(10)
#paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
style = result.styles['Normal']
font = style.font
font.name = 'Calibri'
font.size = Pt(10)
#font.color.rgb = RGBColor(0, 0, 0)

#Body (Paragraph)
para1 = result.add_paragraph('\n')
para1.add_run('Name of Student: \n')
para1.add_run('Admission No: \t\t\t\tSection: \n')
para1.add_run('Roll No: \n')
para1.add_run("Mother's name: \n")
para1.add_run("Father's name: ")

para2a = result.add_paragraph()
para2a.add_run('2(A) Life Skills:')
#Creating table with 4 rows and 4 columns
table2a = result.add_table(rows=4, cols=4)
table2a.style = 'TableGrid'
#table2a.autofit = True
#table2a.allow_autofit = True
#Setting the column 1 width
table2a.columns[1].width = Inches(1.6)
table2a.cell(0,0).text = 'Life Skills'
table2a.cell(0,1).text = 'Descriptive Indicator'
table2a.cell(0,2).text = 'Grade'
table2a.cell(0,3).text = 'Grade Point'
table2a.cell(1,0).text = 'Thinking Skills'
table2a.cell(2,0).text = 'Social Skills'
table2a.cell(3,0).text = 'Emotional Skills'

para2b = result.add_paragraph('\n')
para2b.add_run('2(B)')
#Creating table with 2 rows and 4 columns
table2b = result.add_table(rows=2, cols=4)
table2b.style = 'TableGrid'
#Setting the column 1 width
table2b.columns[1].width = Inches(1.6)
table2b.cell(0,0).text = 'Work Education'
table2b.cell(0,1).text = 'Descriptive Indicator'
table2b.cell(0,2).text = 'Grade'
table2b.cell(0,3).text = 'Grade Point'
table2b.cell(1,0).text = 'Work Education'

para2c = result.add_paragraph('\n')
para2c.add_run('2(C)')
#Creating table with 2 rows and 4 columns
table2c = result.add_table(rows=2, cols=4)
table2c.style = 'TableGrid'
#Setting the column 1 width
table2c.columns[1].width = Inches(1.6)
table2c.cell(0,1).text = 'Descriptive Indicator'
table2c.cell(0,2).text = 'Grade'
table2c.cell(0,3).text = 'Grade Point'
table2c.cell(1,0).text = 'Visual and Performing Arts'

para2d = result.add_paragraph('\n')
para2d.add_run('2(D) Attitudes and Values:')
#Creating table with 5 rows and 4 columns
table2d = result.add_table(rows=5, cols=4)
table2d.style = 'TableGrid'
#Setting the column 1 width
table2d.columns[1].width = Inches(1.6)
table2d.cell(0,0).text = 'Attitude towards'
table2d.cell(0,1).text = 'Descriptive Indicator'
table2d.cell(0,2).text = 'Grade'
table2d.cell(0,3).text = 'Grade Point'
table2d.cell(1,0).text = 'Teachers'
table2d.cell(2,0).text = 'Schoolmates'
table2d.cell(3,0).text = 'School Programmes & Environment'
table2d.cell(4,0).text = 'Value Systems'

para3a = result.add_paragraph('\n')
para3a.add_run('3(A) Co-Curricular Activities:')
#Creating table with 3 rows and 4 columns
table3a = result.add_table(rows=3, cols=4)
table3a.style = 'TableGrid'
#Setting the column 1 width
table3a.columns[1].width = Inches(1.6)
table3a.cell(0,0).text = 'Activity'
table3a.cell(0,1).text = 'Descriptive Indicator'
table3a.cell(0,2).text = 'Grade'
table3a.cell(0,3).text = 'Grade Point'

para3b = result.add_paragraph('\n')
para3b.add_run('3(B)Physical and Health Education:')
#Creating table with 3 rows and 4 columns
table3b = result.add_table(rows=3, cols=4)
table3b.style = 'TableGrid'
#Setting the column 1 width
table3b.columns[1].width = Inches(1.6)
table3b.cell(0,0).text = 'Activity'
table3b.cell(0,1).text = 'Descriptive Indicator'
table3b.cell(0,2).text = 'Grade'
table3b.cell(0,3).text = 'Grade Point'

Para4 = result.add_paragraph('\n\t\t\t\t\t\t\t\t\tTotal = ')

#Adding page break
result.add_page_break()
#Saving Document
result.save('result.docx')
