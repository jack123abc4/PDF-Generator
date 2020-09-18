from fpdf import FPDF # PDF handler
import xlrd # Excel handler
import os


pdf_w=210 # A4 width
pdf_h=297 # A4 height

# PDF class
class PDF(FPDF):
    # border
    def lines(self):
        self.rect(5.0, 5.0, 205.9,269.4)
    
    # title function
    def titles(self,s):
      self.set_xy(0.0,0.0)
      self.set_font('Arial', 'B', 16)
      self.set_text_color(0,68,255)
      self.cell(w=210.0, h=40.0, align='C', txt=s, border=0)

    # text function
    def writeLine(self,s):
      self.set_xy(10.0,30.0) 
      self.set_text_color(0,0,0)
      self.set_font('Arial', '', 12)
      self.multi_cell(0,10,s)
      self.lines()

def clear():
  print("\033[H\033[J")

# tracks spreadsheets
sheetsInDirectory = []
termBody = ""
for file in os.listdir():
  if file.endswith(".xlsx"):
    sheetsInDirectory.append(str(os.path.basename(file)))

if len(sheetsInDirectory) == 0:
  print("No spreadsheets found.\nFinished.")
  quit()
elif len(sheetsInDirectory) == 1:
  termBody += "1 spreadsheet found.\n"
else:
  termBody += str(len(sheetsInDirectory)) + " spreadsheets found.\n"

termBody += "\n"
for p in sheetsInDirectory:
  termBody += '- "' + p + '"\n'
termBody += "\nConvert all? (Y/N) "

choice = ""
sheetsToConvert = []
while choice != "y" and choice != "n":
  clear()
  print (termBody,end="")
  choice = input().lower()
termBody += choice + "\n"
if choice == "n":
  for p in sheetsInDirectory:
    choice = ""
    while choice != "y" and choice != "n":
      clear()
      print (termBody,end="")
      choice = input('Convert "' + p + '"? (Y/N) ')
    termBody += 'Convert "' + p + '"? (Y/N) ' + choice + "\n"
    if choice == "y":
      sheetsToConvert.append(p)
else:
  for p in sheetsInDirectory:
    sheetsToConvert.append(p)


for p in sheetsToConvert:
  # imports Excel sheet
  # current format - first two columns are Last Name, First Name
  # first row contains headers and is ignored when parsing
  book = xlrd.open_workbook(p)
  sheet = book.sheet_by_index(0)
  width = sheet.ncols
  height = sheet.nrows
  print("------------------------------------")
  print ('Spreadsheet "' + p + '" opened.')
  print ("Creating " + str(height-1) + " PDF files...\n")

  # creates list of rows from Excel sheet
  clients = []
  for y in range(1,height):
    clients.append(sheet.row_slice(y,0))

  for sl in clients:
    # generates PDF title
    pdfTitle = str(sl[1].value + " " + sl[0].value)
    


    # generates PDF body
    pdfBody = ""
    headerIndex = 0
    for c in sl:
      pdfBody += str(sheet.cell(0,headerIndex).value) + ": "
      headerIndex += 1

      
      if (c.ctype == 3): # date formatting
          year, month, day, hour, minute, second = xlrd.xldate_as_tuple(c.value, book.datemode) 
          pdfBody += str(month) + "/" + str(day) + "/" + str(year) + "\n"
      else:
        pdfBody += str(c.value) + "\n"
      
    # creates PDF
    pdf = PDF(format='Letter')
    pdf.add_page()

    # formatting
    pdf.lines()

    # creates text from Excel data
    pdf.titles(pdfTitle)
    pdf.writeLine(pdfBody)

    
    # outputs PDF to folder
    if not os.path.exists('PDFs'):
      os.makedirs('PDFs')
    if not os.path.exists('PDFs/' + p + ' BATCH'):
      os.makedirs('PDFs/' + p + ' BATCH')
    pdf.output("PDFs/" + p + " BATCH/" + sl[0].value+","+sl[1].value+".pdf",'F')
    print('File exported: "' + sl[0].value+','+sl[1].value+'.pdf"')

print("\nFinished.")
