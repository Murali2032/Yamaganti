from pathlib import Path
import xlwings as xw
import os

with xw.App() as app:
    app.visible = False
    book = app.books.open('C:\\Users\\Murali\Videos\\Copy of Operations Leader.xlsx')
    sheet = book.sheets[0]
    sheet.page_setup.print_area = '$A$1:$O$140'
    sheet.range("A1")
    current_work_dir = os.getcwd()
    pdf_file_name = "pdf_workbook_printout.pdf"
    pdf_path = Path('C:\\Users\\Murali\\Downloads\\report.pdf')
    
    sheet.to_pdf(path=pdf_path, show=True)
    
    
    


import xlwings as xw

app = xw.App(visible=False)
book = app.books[0]
sheet = book.sheets[0]

sheet.range('A1').value = 73913

book.save('book.xlsx')
app.kill()
