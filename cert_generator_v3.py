from openpyxl import load_workbook
from docx.shared import Pt
import translator
import shutil
import docx
import glob

#Starter docx document search
starter_doc = glob.glob('./*.docx')[0].strip('.\\')
#Excel sheet setup
wb = load_workbook('grelha.xlsx', data_only = True)
sheet = wb['workshops']
cert_amount_range = (12, 34) #Change the amount here (see excel sheet, equals to the rows interval)

for i in range(*cert_amount_range):
    #individual_data >> certification id, name, grade, evaluation
    individual_data = (sheet.cell(row=i, column=1).value,
                       sheet.cell(row=i, column=2).value,
                       sheet.cell(row=i, column=20).value,
                       sheet.cell(row=i, column=21).value)
    new_filename = str(individual_data[1].strip(' ')) + '.docx'

    #Doc and text settings setup
    doc = docx.Document(starter_doc)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Museo 300'
    font.size = Pt(10)

    for paragraph in doc.paragraphs:
        if '...' in paragraph.text:
            temp = paragraph.text
            paragraph.text = temp.replace('...', individual_data[1])
            
        if '10 (dez valores)' in paragraph.text:
            #Temp vars and variables for the different versions of the grades
            temp = paragraph.text
            grade = str(individual_data[2])
            separated_grade = grade.split('.')
            if len(separated_grade) > 1:
                separated_grade[1] = separated_grade[1][:2]
            translated_grade = [translator.translate(elem) for elem in separated_grade]
            
            if len(translated_grade) > 1:
                formatted_grade = str(grade) + ' (' + translated_grade[0] + ' valores e ' + translated_grade[1] + ' decimais) '
                paragraph.text = temp.replace('10 (dez valores), ', formatted_grade)
            else:
                if translated_grade == ['']:
                    paragraph.text = temp.replace('10 (dez valores), ', 'não classificável ')
                else:
                    formatted_grade = str(grade) + ' (' + translated_grade[0] + ' valores) '
                    paragraph.text = temp.replace('10 (dez valores), ', formatted_grade)

        if 'Excelente' in paragraph.text:
            temp = paragraph.text
            paragraph.text = temp.replace('Excelente', individual_data[3])
                
        if '780' in paragraph.text:
            temp = paragraph.text
            paragraph.text = temp.replace('780', str(individual_data[0]))

    doc.save(individual_data[1] + '.docx')

    print(individual_data, ' -> Certificate Generated')
