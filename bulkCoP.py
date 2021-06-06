import io,sys,csv,copy,os
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from os import path
import urllib.request

x = [40, 190, 420]  # Offsets of "Name","Address","Class" columns
y0 = 705            # offset of 1st line
fontName,fontSize = "Helvetica",9.5
yoffs = -15.5       # gap between lines
pageCount = 0
countX, countY = 350, 150
origPDFurl = "https://www.royalmail.com/sites/default/files/Bulk_Certificate_Posting_Standard%20MAR14.pdf"
origPDFfname = "./Bulk_Certificate_Posting_Standard%20MAR14.pdf"
newPageExt = "_BulkCoP.pdf"

def newPage():
    global can,y,count,fontName,fontSize
    can.setFont(fontName,fontSize)
    y = y0
    count = 0

def flushPage():
    global can,count,pageCount, countX, countY
    if count > 0:
        pageCount += 1
        print("Writing page {}".format(pageCount))
        can.drawString(countX, countY, str(count))  # Add the "Number of Items"
        can.showPage()
        newPage()

packet = io.BytesIO()

# create a new PDF with Reportlab
can = canvas.Canvas(packet, pagesize=A4)

try:
    csvFname = sys.argv[1]
except IndexError:
    print("Must provide a CSV filename")
    exit(1)

with open(csvFname, newline='') as csvfile:
    print("Reading from {}".format(csvFname))
    reader = csv.reader(csvfile)
    newPage()
    for row in reader:
        if row[0] == "Name": # skip an actual "name" label
            continue

        for j in range(0, 3):
            can.drawString(x[j], y, row[j])
        y += yoffs

        count += 1
        if(count == 30):
           flushPage()

flushPage() # push any remaining stuff in canvas out
can.save()  # close the canvas

output = PdfFileWriter()
basename,ext = os.path.splitext(csvFname)
outfile = basename + newPageExt

# move to the beginning of the StringIO buffer
packet.seek(0)
new_pdf = PdfFileReader(packet)

pdfpath = path.abspath(path.join(path.dirname(__file__), origPDFfname))

while True:
    try:
        existing_pdf = PdfFileReader(open(pdfpath, "rb"))
        print("Reading template from {}".format(origPDFfname))
        break
    except FileNotFoundError:
        print("Trying to fetch {}".format(origPDFurl))
        urllib.request.urlretrieve(origPDFurl,origPDFfname)
        print("Success. Retrying file read...")
        # try:
        #     urllib.retrieve(origPDFurl,origPDFfname)
        #     continue
        # except:
        #     print("Could not fetch {}".format(origPDFurl))
        #     exit(1)


origpage = existing_pdf.getPage(0)

for i in range(0,pageCount):
    page = copy.copy(origpage) # make a copy of the origpage to avoid trashing it
    page.mergePage(new_pdf.getPage(i))
    output.addPage(page)

print("Writing " + outfile)
outputStream = open(outfile, "wb")
output.write(outputStream)
outputStream.close()


