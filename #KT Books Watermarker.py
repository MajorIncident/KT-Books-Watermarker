#KT Books Watermarker
#Author: Shane Chagpar
#Inception Date: 2020 12 08
#Formats PDFs with a watermark and emails them to a mail merge list from excel
#Requires Libraries: Pypdf2, fpdf, pypiwin32, pikepdf --> Requirements.txt generated from pipreqs library
#Installation: pip install -r requirements.txt
#Inputs: list.csv, cards.pdf, cases.pdf, notes.pdf, extra.pdf in same folder
#Outputs: zip files sent to email address in outoook 
#Limitations: Outlook must be open when run

#Future: Use setuptools to build package distribtor 

#CUSTOMIZATIONS ----------------------------------------------------------
copyrightLine = "Copyright (C) Kepner-Tregoe All Rights Reserved"
emailSubject = "KT Digital Books for your upcoming class"
emailBody = "Attached please find your KT Books Digital Materials. Enjoy your training session!"
ccEmail = "jneylan@kepner-tregoe.com"
auditEmail = "creppy@kepner-tregoe.com"
auditSubjectEmail = "KT Digital Materials Audit Email"
auditCCEmail = ""
pdfPassword = "KTPassword"

#GLOBAL VARIABLES -------------------------------------------------------------
auditReport = ""

#LIBRARY IMPORTS ---------------------------------------------------------------
#Libraries for CSV Manipulation
from csv import reader

#Libraries for File Handling
import os.path
import os, errno
from pathlib import Path

#Libraries for Creating a Watermark
import fpdf #pip install fdpf
from datetime import date, datetime

#Libraries for Merging PDFs
import PyPDF2 #pip install pypdf2

#for Encryption 
import pikepdf
from pikepdf import Pdf

#Libraies for CreateMail
import win32com.client as win32 #pip install pypiwin32

#ZIP Library
from zipfile import ZipFile

#SUBROUTINES -------------------------------------------------------------
def read_list(listFileName = ''):
    # open file in read mode
    p = Path(__file__).with_name(listFileName)
    with open(p, 'r') as read_obj:
        csv_reader = reader(read_obj)
        header = next(csv_reader)
        # Check file as empty
        if header != None:
            
            #Set today's date for the files
            today = date.today()
            audit_log("Materials Production Begins")

            # Iterate over each row after the header in the csv
            for row in csv_reader:
                # row variable is a list that represents a row in csv
                print("START PROCESS - Found ", row)
            
                #DO THE WORK
                #audit_log("Materials for " + row[0] + ' ' + row[1] + " started")
                print("1/9 Watermarking for", row[0] + ' ' + row[1])
                create_watermark('These materials produced for ' + row[0] + ' ' + row[1] + ' on ' + today.strftime("%B %d, %Y") + ' and may not be redistributed.')
                print("2/9 Starting on Notes for", row[0] + ' ' + row[1])
                merge_pdf(row[0] + row [1], 'notes.pdf') 
                print("3/9 Starting on Cases for", row[0] + ' ' + row[1])
                merge_pdf(row[0] + row [1], 'cases.pdf')
                print("4/9 Starting on Cards for", row[0] + ' ' + row[1])
                merge_pdf(row[0] + row [1], 'cards.pdf')
                print("5/9 Starting on Extras for", row[0] + ' ' + row[1])
                merge_pdf(row[0] + row [1], 'extra.pdf')
                #audit_log("Materials for " + row[0] + ' ' + row[1] + " Watermarked")
                print("6/9 Creating compressed package for", row[0] + ' ' + row[1])
                zip_pdf(row[0] + row [1])
                #audit_log("Materials for " + row[0] + ' ' + row[1] + " Compressed")
                print("7/9 Emailing", row[0] + ' ' + row[1])
                create_mail("Hi " + row[0] + " " + row [1] + ", " + emailBody, emailSubject, row[2], row[0] + row [1] + '.zip', send=False)
                print("8/9 Documenting Audit Log for", row[0] + ' ' + row[1])
                audit_log("Materials for " + row[0] + ' ' + row[1] + " Sent")
                print("9/9 Cleaning up temporary files for", row[0] + ' ' + row[1])
                cleanup_folder(row[0] + row [1])
                #audit_log("Materials for " + row[0] + ' ' + row[1] + " Removed from Local Drive")
                
def create_watermark(watermarkText = ''):
    pdf = fpdf.FPDF(format='letter') #pdf format
    pdf.add_page() #create new page
    pdf.set_font("Arial", 'B', size=8) # font and textsize
    pdf.cell(0, 10, txt=watermarkText + " " + copyrightLine, ln=1, align="C")
    #pdf.cell(0, 10, txt=copyrightLine, ln=2, align="L")
    
    pdf.output("watermark.pdf")

def merge_pdf(participantName, pdf_file = ''):
    watermark = "watermark.pdf"
    merged_file = participantName + pdf_file

    #Open and read original file
    try:
        p = Path(__file__).with_name(pdf_file)
        input_file = open(p,'rb')
        input_pdf = PyPDF2.PdfFileReader(input_file)
        output = PyPDF2.PdfFileWriter()
        
        #Open and read watermark file
        p = Path(__file__).with_name(watermark)
        watermark_file = open(p,'rb')
        watermark_pdf = PyPDF2.PdfFileReader(watermark_file)

        #Watermark each page in the PDF
        for i in range(input_pdf.getNumPages()):
            #Get first page of original file
            pdf_page = input_pdf.getPage(i)

            #Perform Merge with first page of watermark PDF
            pdf_page.mergePage(watermark_pdf.getPage(0))

            #Save Output in memory
            output.addPage(pdf_page)

        #Save output from memory to disk
        p = Path(__file__).with_name(merged_file)
        merged_file = open(p,'wb')
        output.write(merged_file)

        #Cleanup
        merged_file.close()
        
        #Encrypt with PikePDF
        encrypt_PDF(participantName + pdf_file)
        
        #Cleanup
        watermark_file.close()
        input_file.close()

    except IOError:
        print ("Error: File does not exist")
    return 0

def encrypt_PDF(sourceFileName = ''):
    p = Path(__file__).with_name(sourceFileName)
    pdf = pikepdf.Pdf.open(p, allow_overwriting_input=True)
    pdfRestrictions = pikepdf.Permissions(accessibility=False, extract=False, print_lowres=False, print_highres=False, modify_annotation=False, modify_assembly=False, modify_form=False, modify_other=False)
    pdf.save(sourceFileName, encryption=pikepdf.Encryption(user="", owner=pdfPassword, allow=pdfRestrictions))

def create_mail(text, subject, recipient, attachmentName, send=True):

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.CC = ccEmail
    mail.Subject = subject
    mail.HtmlBody = text
    
    filePathDir = os.getcwd()
    attachment1 = os.path.join(filePathDir, attachmentName)
    #attachment1 = Path(__file__).with_name(attachmentName) <--- how everything else is setup but returns windowspath so not used
    
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
        p = Path(__file__).with_name(filename)
        os.remove(filename)
    except OSError as e: # this would be "except OSError, e:" before Python 2.6
        if e.errno != errno.ENOENT: # errno.ENOENT = no such file or directory
            raise # re-raise exception if a different error occurred

def audit_log(auditEvent=""):   
    global auditReport 
    
    auditReport = (auditReport + "\n" + str(datetime.now()) + " - " + auditEvent)

def send_audit_log(send=False):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = auditEmail
    mail.CC = auditCCEmail
    mail.Subject = auditSubjectEmail
    mail.HtmlBody = auditReport
    
    if send:
        mail.send()
    else:
        mail.save()

#MAIN EXECUTION ------------------------------------------------------------
read_list('list.csv')
send_audit_log(False)