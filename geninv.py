"""
Options:
- PDF (reportlab, pyPDF, pyPDF2, pyfpdf, pycairo)
- SVG
- template + custom vs complete


Cairo apparently produces ugly text.

Use:
  * https://code.google.com/p/pyfpdf/
  * https://github.com/mstamy2/PyPDF2/


Would make more sense to use a svg template rather than pdf but cairo sucks to setup

todo:
- auto generate invoice date, number and period
- make the various greys constants
- delete tmp file (or use in memory stream, http://www.blog.pythonlibrary.org/2013/07/16/pypdf-how-to-write-a-pdf-to-memory/)
"""

import sys
import fpdf
import string
import os
import datetime, dateutil.parser
from PyPDF2 import PdfFileReader, PdfFileWriter
import sqlite3
import re

db = sqlite3.connect('myinvoices.db')

inpfn = sys.argv[1]
tmpfn = "tmp.pdf"
#pdffn = os.path.splitext(inpfn)[0]+'.pdf'

def parse_line(f):
    l = f.readline()
    return l[l.find(':')+1:].strip()

def expand_client(name,db):
    id, name, addr = [x for x in db.execute('select id,fullname,mailAddress from clients where name=?',(name,))][0]
    return name+', '+addr
    
### load data
data = []
infile = open(inpfn,'r')
client = parse_line(infile) # this should exactly match a client name in myinvoices.db:clients.name
client_and_address = expand_client(client,db)
attn_person = parse_line(infile)
invoice_period = parse_line(infile)
invoice_date = None
client_ref =  parse_line(infile) # purchase order
addGST = True # should GST be calculated and added?
prelim = []
try:
    m = re.match("(.*)--(.*)",invoice_period)
    period_start = dateutil.parser.parse(m.group(1))
    period_stop = dateutil.parser.parse(m.group(2))
    invoice_period = (period_start.strftime('%d/%m/%Y')+' \x96 '+period_stop.strftime('%d/%m/%Y'))#.encode('latin-1','replace')
except Exception as err: 
    #print(err)
    # presumably the period is just a date...
    invoice_date = dateutil.parser.parse(invoice_period)
    period_start = invoice_date
    period_stop = invoice_date
for l in infile:
    if l.strip()=="":
        continue
    if l.strip()[0:5]=="=====":
        break
    try:
        descr, value = map(str.strip,l.split('|'))
    except ValueError:
        # send text directly to output
        prelim.append(l.strip())
        continue
    #print(value)
    gst = 'excl'
    currency='NZD'
    if value[-1]=='i': # this item is gst inclusive
        value = value[:-1]
        gst = 'incl'
    if value[-1]=='x':
        value = value[:-1]
        gst = 'excl'
    if value[:3] in ('AUD', 'USD'):
        currency = value[:3]
        value = value[3:]
        addGST = False
    if value[-1]=='n': # no GST component (e.g. for work done in Aus but in NZD or oncharged internet expenses)
        value = value[:-1]
        currency = 'NZD'
        gst = 'none'
    data.append( (descr, value, gst, currency) )
    #print(data)
print data
def get_next_invoice_number(db):
    return int([x for x in db.execute('select max(number) from invoices')][0][0])+1
if invoice_date==None:
    invoice_date = datetime.date.today()
invoice_number = get_next_invoice_number(db)


### generate items block
pdf = fpdf.FPDF(format='a4',unit='mm')
pdf.set_margins(left=15,top=100)
pdf.add_page()

def line_item(descr,value,widths,height=8,fill=None):
    if fill!=None:
        pdf.set_fill_color(fill)
        fill = 1
    if value is None:
        pdf.cell(widths[0],height,descr,border=0, ln=1, align='L', fill=fill)
        return    
    else:
        pdf.cell(widths[0],height,descr,border=1, ln=0, align='L', fill=fill)
    #pdf.cell(widths[1],height,"$%08,.2f" % value, border=1, ln=1, align='R', fill=0)
    if value>0:
        if currency=='NZD':
            pdf.cell(widths[1],height,"${value:,.2f}".format(value=value), border=1, ln=1, align='R', fill=fill)
        else:
            pdf.cell(widths[1],height,"{currency}${value:,.2f}".format(currency=currency,value=value), border=1, ln=1, align='R', fill=fill)
    else:
        pdf.cell(widths[1],height,"", border=1, ln=1, align='R', fill=fill)    



col_widths = (140,40)
pdf.set_font("Arial", size=12)
pdf.set_draw_color(150)
fill=210

# prelim info
for l in prelim:
    line_item(l, value=None, widths=(180,))

gst_rate = .15 
subtotal = 0
gsttotal = 0
for descr, value, gst, currency in data:
    value = float(value)
    if gst=='none': # no GST to collect for this item
        subtotal += value
    elif gst=='incl':
        subtotal += value/(1 + gst_rate)
        gsttotal += value * gst_rate/(1 + gst_rate)
    elif gst=='excl':
        subtotal += value
        gsttotal += value * gst_rate
    if descr=='GST': # this item is entirely GST
        gsttotal += value
    line_item(descr, value, col_widths, fill=fill)
    if fill==210:
        fill=230
    else:
        fill=210
pdf.set_font("Arial", size=12, style='B')
if addGST and gsttotal>0:
    line_item("Sub Total", subtotal, col_widths, fill=190)
    line_item("GST", gsttotal, col_widths, fill=190)
    total = subtotal + gsttotal
else:
    total = subtotal
line_item("Total", total, col_widths, fill=190)


### insert a record into the invoices table; if it exists already 
def insert_into_invoices(db,client,periodStart,periodStop,invoice_number,amount,currency):
    c = db.cursor()
    clientId = c.execute('select id from clients where name=?',(client,)).fetchone()[0]
    c.execute('insert into invoices(clientId,number,startDate,stopDate,invoiceDate,amount,currency) values(?,?,?,?,?,?,?)',
              (clientId,invoice_number,periodStart.strftime('%Y-%m-%d'),periodStop.strftime('%Y-%m-%d'),invoice_date.strftime('%Y-%m-%d'),amount,currency))
    c.close()
    db.commit()


try:
    insert_into_invoices(db,client,period_start,period_stop,invoice_number,subtotal,currency)
except sqlite3.IntegrityError:
    print("Invoice already exists in db")
    clientId = [x for x in db.execute('select id from clients where name=?',(client,))][0][0]
    #print clientId
    #res = [x for x in db.execute('select number, invoiceDate from invoices where clientId=? and startDate=? and stopDate=? and amount=?', (clientId,period_start.strftime('%Y-%m-%d'),period_stop.strftime('%Y-%m-%d'),subtotal))]
    res = [x for x in db.execute('select number, invoiceDate from invoices where clientId=? and startDate=? and stopDate=?', (clientId,period_start.strftime('%Y-%m-%d'),period_stop.strftime('%Y-%m-%d')))]
    invoice_number, invoice_date = res[0]
    invoice_date = datetime.datetime.strptime(invoice_date,'%Y-%m-%d')
    #print(invoice_date)
    #import pdb; pdb.set_trace()
    #exit()


#### invoice #, client, etc
pdf.set_font("Arial", size=13)
svgwidth = 744.09448819
svgheight = 1052.3622047
pdfwidth = 210
pdfheight = 297
border = 0

# x="57.187263" y="254.76782"
# id="tspan3029">client-and-address
pdf.set_xy(53.0/svgwidth*pdfwidth,240./svgheight*pdfheight)
pdf.cell(120,5,client_and_address,border=border, ln=0, align='L', fill=0)

# x="57.187263" y="274.76782"
# id="tspan3057">period-that-this-invoice-covers
pdf.set_xy(53.0/svgwidth*pdfwidth,260.0/svgheight*pdfheight)
pdf.cell(120,5,invoice_period,border=border, ln=0, align='L', fill=0)

# x="57.187263" y="294.76782"
# id="tspan3031">Attn: attn-person-x
pdf.set_xy(53.0/svgwidth*pdfwidth,280.0/svgheight*pdfheight)
pdf.cell(120,5,'Attn: %s' % attn_person,border=border, ln=0, align='L', fill=0)

# x="684.44379" y="140.99284"
# id="tspan3005">date-of-this-invoice
pdf.set_xy(510.0/svgwidth*pdfwidth,130.0/svgheight*pdfheight)
pdf.cell(50,5,invoice_date.strftime('%d %B %Y'),border=border, ln=0, align='R', fill=0)

# x="684.44379" y="160.99284"
# id="tspan3007">Invoice #number-of-this-invoice
pdf.set_xy(510.0/svgwidth*pdfwidth,150.0/svgheight*pdfheight)
pdf.cell(50,5,'Invoice #%s' % invoice_number,border=border, ln=0, align='R', fill=0)

#### wrap up and emit
pdf.output(tmpfn)
pdf.close()

### merge with template and write
details = PdfFileReader(open(tmpfn,"rb")).getPage(0)
template = PdfFileReader(open("invoice-template.pdf","rb")).getPage(0)
template.mergePage(details)
combined = PdfFileWriter()
combined.addPage(template)

pdffn = "jharrington-%i-%s-%s.pdf" % (invoice_number,client.lower(),period_stop.strftime('%b%y'))

combined.write(open(pdffn,"wb"))
print("%s -> %s" % (inpfn, pdffn))


