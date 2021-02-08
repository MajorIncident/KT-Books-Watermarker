#KT Books Watermarker
#Author: Shane Chagpar
#Inception Date: 2020 12 08
#Formats PDFs with a watermark and emails them to a mail merge list from excel
#Requires Libraties: Pypdf2, fpdp, pypiwin32
#Inputs: list.csv, cards.pdf, cases.pdf, notes.pdf, extra.pdf in same folder
#Outputs: zip files sent to email address in outoook 
#Limitations: Outlook must be open when run

#Libraries for CSV Manipulation
from csv import reader

#Libraries for File Handling
import os.path
import os, errno

#Libraries for Creating a Watermark
import fpdf #pip install fdpf
from datetime import date

#Libraries for Merging PDFs
import PyPDF2 #pip install pypdf2

#Libraies for CreateMail
import win32com.client as win32 #pip install pypiwin32

#ZIP Library
from zipfile import ZipFile

filePathDir = r"C:\Users\SCHAGPAR\OneDrive\Documents\Personal\Filing Cabinet\Coding Projects\KT Books Watermarker"

def read_list(listFileName = ''):
    # open file in read mode
    with open(listFileName, 'r') as read_obj:
        csv_reader = reader(read_obj)
        header = next(csv_reader)
        # Check file as empty
        if header != None:
            
            #Set today's date for the files
            today = date.today()

            # Iterate over each row after the header in the csv
            for row in csv_reader:
                # row variable is a list that represents a row in csv
                print(row)

                #DO THE WORK
                create_watermark('These materials produced for ' + row[0] + ' ' + row[1] + ' on ' + today.strftime("%B %d, %Y") + ' and may not be redistributed')
                merge_pdf(row[0] + row [1], 'notes.pdf')
                merge_pdf(row[0] + row [1], 'cases.pdf')
                merge_pdf(row[0] + row [1], 'cards.pdf')
                merge_pdf(row[0] + row [1], 'extra.pdf')
                zip_pdf(row[0] + row [1])
                create_mail("Hi " + row[0] + " " + row [1] + " attached please find your KT Books Digital Materials. Enjoy your training session!", "KT Digital Books for your upcoming class", row[2], row[0] + row [1] + '.zip', send=False)
                cleanup_folder(row[0] + row [1])

def create_watermark(watermarkText = ''):
    pdf = fpdf.FPDF(format='letter') #pdf format
    pdf.add_page() #create new page
    pdf.set_font("Arial", size=12) # font and textsize
    pdf.cell(200, 10, txt=watermarkText, ln=1, align="L")
    pdf.cell(200, 10, txt='Copyright (C) Kepner-Tregoe All Rights Reserved', ln=2, align="L")
    
    pdf.output("watermark.pdf")

def merge_pdf(participantName, pdf_file = ''):
    watermark = "watermark.pdf"
    merged_file = participantName + pdf_file

    #Open and read original file
    try:
        input_file = open(pdf_file,'rb')
        input_pdf = PyPDF2.PdfFileReader(input_file)

        #Open and read watermark file
        watermark_file = open(watermark,'rb')
        watermark_pdf = PyPDF2.PdfFileReader(watermark_file)

        #Get first page of original file
        pdf_page = input_pdf.getPage(0)

        #Get first page of watermark file
        watermark_page = watermark_pdf.getPage(0)

        #Perform Merge
        pdf_page.mergePage(watermark_page)

        #Save Output in memory
        output = PyPDF2.PdfFileWriter()
        output.addPage(pdf_page)

        #Encrypt the PDF with password on this line - Future Use
        #output.encrypt('KTPassword')

        #Save output from memory to disk
        merged_file = open(merged_file,'wb')
        output.write(merged_file)

        #Cleanup
        merged_file.close()
        watermark_file.close()
        input_file.close()
    except IOError:
        print ("Error: File does not exist")
    return 0

def create_mail(text, subject, recipient, attachmentName, send=True):

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    attachment1 = os.path.join(filePathDir, attachmentName)
    mail.Attachments.Add(Source=attachment1)
    
    if send:
        mail.send()
    else:
        mail.save()

def zip_pdf(participantName = ''):
    # create a ZipFile object
    zipObj = ZipFile(participantName + '.zip', 'w')

    # Add multiple files to the zip
    if os.path.isfile(participantName + 'notes.pdf'):
        zipObj.write(participantName + 'notes.pdf')
    if os.path.isfile(participantName + 'cases.pdf'):
        zipObj.write(participantName + 'cases.pdf')
    if os.path.isfile(participantName + 'cards.pdf'):
        zipObj.write(participantName + 'cards.pdf')
    if os.path.isfile(participantName + 'extra.pdf'):
        zipObj.write(participantName + 'extra.pdf')
    # close the Zip File
    zipObj.close()

def cleanup_folder(participantName = ''):
    cleanup_file(participantName + 'cards.pdf')
    cleanup_file(participantName + 'notes.pdf')
    cleanup_file(participantName + 'cases.pdf')
    cleanup_file(participantName + 'extra.pdf')
    cleanup_file(participantName + '.zip')
    cleanup_file ('watermark.pdf')

def cleanup_file(filename):
    try:
        os.remove(filename)
    except OSError as e: # this would be "except OSError, e:" before Python 2.6
        if e.errno != errno.ENOENT: # errno.ENOENT = no such file or directory
            raise # re-raise exception if a different error occurred

#Begin execution
read_list('list.csv')