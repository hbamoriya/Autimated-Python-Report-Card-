from fpdf import FPDF
from main import arr
#print(arr)
for x in arr:
    document = FPDF()
    document.add_page()
    document.set_font('helvetica', size=12)
    document.cell(txt="Round - "+x["Round"])
    document.cell.ln(10)
    document.cell(txt="First Name - "+x["First Name"])
    document.cell(txt="Registration Number - "+x["Registration Number"])
    document.cell(txt="Grade - "+x["Grade"])
    document.cell(txt="Name Of School - "+x["Name Of School"])
    document.cell(txt="Gender - "+x["Gender"])
    document.cell(txt="Date Of Birth - "+x["Date Of Birth"])
    document.cell(txt="City Of Residence - "+x["City Of Residence"])
    #document.cell(txt=x["Last Name"])
    document.output(x["First Name"]+".pdf")

