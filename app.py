#library to import the excel file
import openpyxl
import shutil

import sendEmail



#libraries to create the pdf file and add text to it
from reportlab.pdfgen import canvas

from reportlab.lib.styles import (ParagraphStyle, getSampleStyleSheet)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
#libraries to merge pdf files

from reportlab.lib.units import inch, cm
import os
from PyPDF2 import PdfFileReader, PdfFileMerger, PdfFileWriter

from reportlab.lib import pdfencrypt

from reportlab.platypus import Table, TableStyle, Paragraph

from reportlab.platypus import SimpleDocTemplate, Spacer, Image
from reportlab.lib.pagesizes import A4, A5, letter

from reportlab.platypus import TableStyle
from reportlab.lib import colors

from flask import Flask, render_template, request, jsonify, send_file
from zipfile import ZipFile
from io import BytesIO
import os
import requests
import logging

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)


def count_rows_to_send(sheet, start_row=3):
    count = 0
    for row_number in range(start_row, sheet.max_row + 1):
        if sheet.cell(row=row_number, column=1).value is None:
            break
        count += 1
    return count

@app.route('/')
def index():
  return render_template("index.html")

@app.route('/healthz')
def healthz():
    return "ok", 200

@app.route('/extract', methods=["POST"])

def create_payslip():

    app.logger.info(
        "POST /extract received: form_keys=%s file_keys=%s content_length=%s",
        list(request.form.keys()),
        list(request.files.keys()),
        request.content_length,
    )

    #this was changed
    #convert the font so it is compatible
    pdfmetrics.registerFont(TTFont('Arial','arial.ttf'))
    file_header = ""
    bottom_label = ""
    email_file_body = ""
    is_encrypted = request.form.get('encrypt')

    option = request.form.get('options')
    smtp_email = request.form.get('smtpEmail', "")
    smtp_password = request.form.get('smtpPassword', "")
    app.logger.info("Selected option=%s encrypt=%s", option, is_encrypted)
    app.logger.info("SMTP login supplied=%s sender_email=%s", bool(smtp_email), smtp_email)
    
    # Handle the uploaded file and form inputs here
    # Example: Save the file, process it, etc.


    r = request.files.get('excelFile')
    if r is None or r.filename == "":
        app.logger.error("No excelFile uploaded. file_keys=%s", list(request.files.keys()))
        return "No Excel file was uploaded.", 400
    app.logger.info("Uploaded excelFile filename=%s content_type=%s", r.filename, r.content_type)
    

    if option == 'custom':
        file_header = request.form['field1']
        bottom_label = request.form['field2']
        email_file_body = request.form['field3']
        # Handle custom fields data
    
    if option == "emailOnly":

        image_file = request.files.get('image')
        img = image_file.filename if image_file and image_file.filename else ""
        sender = request.form.get('field11', "")
        sender_title = request.form.get('field12', "")
        company = request.form.get('field13', "")
        com_address = request.form.get('field14', "")
        ph_number = request.form.get('field15', "")
        com_email = request.form.get('field16', "")
        app.logger.info(
            "emailOnly fields: image_filename=%s sender=%s sender_title=%s company=%s website=%s",
            img,
            sender,
            sender_title,
            company,
            com_email,
        )
    
    if option == 'salarySlip':
        file_header = "Salary slip for the month of"
        email_file_body = "salary slip for the month of"
        bottom_label = "The figures in the salary slip are confidential and not to be disclosed.\
              Signature is not required for this payslip"

    
    app.logger.info("Starting workbook load for option=%s", option)


    

    #import the sheet from the excel file
    try:
        wb = openpyxl.load_workbook(r, data_only=True)
        sheet = wb[wb.sheetnames[0]]
        # Print or log information to inspect the workbook and sheet
        app.logger.info(
            "Workbook loaded: sheets=%s active_sheet=%s dimensions=%s",
            wb.sheetnames,
            sheet.title,
            sheet.dimensions,
        )
        # Add more print statements to inspect the data
    except Exception as e:
        app.logger.exception("Error loading workbook")
        return f"Error loading workbook: {e}", 400

    app.logger.info("Workbook processing started for option=%s", option)
    if option == "emailOnly":
        email_send_total = count_rows_to_send(sheet, start_row=1)
        email_send_index = 0
        for row_number, row in enumerate(sheet, start=1):
            name = row[0]
            email = row[1]
            app.logger.info("emailOnly row=%s name=%s email=%s", row_number, name.value, email.value)
            try:
                email_send_index += 1
                app.logger.info(
                    "About to send emailOnly message row=%s recipient=%s sender=%s",
                    row_number,
                    email.value,
                    smtp_email,
                )
                app.logger.info(
                    "EmailOnly progress %s/%s row=%s smtp_host=%s smtp_port=%s",
                    email_send_index,
                    email_send_total,
                    row_number,
                    sendEmail.SMTP_HOST,
                    sendEmail.SMTP_PORT,
                )
                sendEmail.sendEmailWithImage(
                    str(name.value),
                    img,
                    sender,
                    sender_title,
                    company,
                    com_address,
                    ph_number,
                    com_email,
                    str(email.value),
                    smtp_user=smtp_email,
                    smtp_password=smtp_password,
                )

            except Exception as e:
                app.logger.exception("emailOnly send failed at row=%s", row_number)
                return f"Email send failed at row {row_number}: {e}", 400
        return "done", 200






    #print(sheet.cell(4,1).value)
    #Page information
    page_width = 2156
    page_height = 3050
    spread = 50
    start = 200
    start_2 = 700

    #Payslip variables
    company_name = 'Conacent Consulting'

    logo = Image("logo and address.png")
    logo._restrictSize(6*inch,8*inch)
    logo.hAlign = "LEFT"
    logo.VALIGN = "TOP"
    col_names = []
    for name in sheet[2]:
        if name.value is not None and name.value != "email":
            col_names.append(name.value)
        

    super_columns = []
    for super_column in sheet[1]:
        super_columns.append(super_column.value)

    year = ""

    styles = getSampleStyleSheet()
    i = 3
    emails = []
    pdf_send_total = count_rows_to_send(sheet, start_row=3)
    pdf_send_index = 0
    while (i):
        vals = []
        password = ""
        if sheet.cell(row = i, column = 1).value is None:
            break
        for j in range(1,sheet.max_column+1):
           
            if sheet.cell(row = i, column = j).value is not None:

                

                inp = str(sheet.cell(row = i, column = j).value)
        
                if "00:00:00" in inp:

                    inp = inp.replace("00:00:00","")    
                      
                    y = inp[:inp.find("-")]
                    # inp = inp[::-1]
                    month = inp[inp.find("-") + 1: inp.rfind("-")]
                    date = inp[inp.rfind("-") + 1: ]
                    inp = date + "/" + month + "/" + y
                    inp = inp.replace(" ","")             
                if ("#" in inp):
                    print("this date of joining is weird")
                if "@" in inp:
                    emails.append(inp)
                else :
                    vals.append(inp)
            else:
                vals.append("N/A")
        name =  str(vals[1])+ ' ' + str(vals[0])  + '.pdf' 
        email = emails[i - 3]
        pdf_buffer = BytesIO()
        pdf = SimpleDocTemplate(
                    pdf_buffer,
                    pagesize= A4,
                    )

        if str(vals[2]) != 'x' and str(vals[2]) != 'N/A' and type(vals[0]) == str:
            data = []
            elements = [logo]
            for j in range(len(vals)):
                if super_columns[j] is None:
                    if str(vals[j]) == 'N/A':
                        continue
                    data.append([col_names[j], str(vals[j])])
                elif super_columns[j] == "LEAVE STATEMENT":
                    table = Table(data, rowHeights = len(data) * [13], colWidths=inch*3)
                    style = TableStyle([
                            ('BACKGROUND', (0,0), (3,0), colors.steelblue),
                            ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),

                            ('ALIGN',(0,0),(-1,-1),'CENTER'),

                            ('FONTNAME', (0,0), (-1,0), 'Courier-Bold'),
                            ('FONTSIZE', (0,0), (-1,0), 10),

                            ('BOTTOMPADDING', (0,0), (-1,0), 2),

                            ('BACKGROUND',(0,1),(-1,-1),colors.beige),
                        ])
                    table.hAlign = "LEFT"
                    
                    table.setStyle(style)


                    elements.append(table)
                    data = []
                    data.append([super_columns[j], ""])
                    data.append([col_names[j], str(vals[j])])
                else:
                    if len(data) == 1:
                        continue
                    #This changes main table
                    #13 IS BEST FONT SIZE
                    table = Table(data, rowHeights = len(data) * [13], colWidths=inch*3)
                    if data[0][1] != "":

                        #First part
                        
                        data[0] = [Paragraph(f"<b> {file_header} " + " " + data[0][0].upper() +  "</b>")]
                        table = Table(data)
                        style = TableStyle([
                        ('BACKGROUND', (0,0), (3,0), colors.white),
                        ('TEXTCOLOR',(0,0),(-1,0),colors.black),

                        ('ALIGN',(0,0),(-1,-1),'LEFT'),

                        ('VALIGN',(-200,-200),(-100,-100),'TOP'),

                        ('FONTNAME', (0,0), (-1,0), 'Courier-Bold'),
                        ('FONTSIZE', (0,0), (-1,0), 100),
                        ('TEXTFONT', (0, 1), (-1, 1), 'Times-Bold'),

                        ('BOTTOMPADDING', (0,0), (-1,0), 2),

                        ('BACKGROUND',(0,1),(-1,-1),colors.white),
                        ])
                        table.hAlign = "LEFT"
                    else:         
                        
                        style = TableStyle([
                            ('BACKGROUND', (0,0), (3,0), colors.steelblue),
                            ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),

                            ('ALIGN',(0,0),(-1,-1),'CENTER'),

                            ('FONTNAME', (0,0), (-1,0), 'Courier-Bold'),
                            #('FONTSIZE', (0,0), (-1,0), 12),
                            ('TEXTSIZE', (0,0), (-1,0), 10),
                            #('TOPPADDING', (0,0), (-1,0), 2),
                            ('BOTTOMPADDING', (0,0), (-1,0), 2),

                            ('BACKGROUND',(0,1),(-1,-1),colors.beige),
                        ])
                        table.hAlign = "LEFT"
                    table.setStyle(style)
                    elements.append(table)
                    elements.append(Spacer(1,5))
                    data = []
                    data.append([super_columns[j], ""])
                    elements.append(Spacer(1,10))
                    data.append([col_names[j], vals[j]])
            
            #this is for leave 
            elements.append(Spacer(1,5))
            table = Table(data, rowHeights = len(data) * [13], colWidths=inch*3)
            style = TableStyle([
                                        ('BACKGROUND', (0,0), (3,0), colors.steelblue),
                                        ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),

                                        ('ALIGN',(0,0),(-1,-1),'CENTER'),

                                        ('FONTNAME', (0,0), (-1,0), 'Courier-Bold'),
                                        ('BOTTOMPADDING', (0,0), (-1,0), 2),

                                        ('BACKGROUND',(0,1),(-1,-1),colors.beige),
                                    ])
            table.hAlign = "LEFT"
            table.setStyle(style)
            elements.append(Spacer(1,10))
            elements.append(table)
            elements.append(Spacer(1,100))
            style_new = getSampleStyleSheet()
            yourStyle = ParagraphStyle('yourtitle',
                           fontName="Helvetica",
                           fontSize=8,
                           parent=style_new['Heading2'],
                           alignment=1,
                           spaceAfter=2)
            elements.append(Paragraph(f"<i>{bottom_label}</i>",  yourStyle))
            elements.append(Paragraph("<i> Registered Office:  P-94/95, Bangur Avenue, BL-C, Kolkata - 700055 </i>",  yourStyle))
            pdf.build(elements)
            pdf_buffer.seek(0)

            # create a PdfFileWriter object
            out = PdfFileWriter()

            # Read the ge