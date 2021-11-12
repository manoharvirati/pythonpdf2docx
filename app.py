from flask import Flask
from flask import request,render_template,redirect,url_for,send_file
import os
from docx2pdf import convert
from pdf2docx import parse
from typing import Tuple
import sys
import tkinter import _tkinter

#import win32com.client
#import pythoncom
#pythoncom.CoInitialize()

UPLOADER_FOLDER=''
AllOWED_EXTENSIONS={'docx'}

app = Flask(__name__)
app.config['UPLOADER_FOLDER']=UPLOADER_FOLDER

@app.route('/')
@app.route('/index',methods=['GET','POST'])
def index():
    if request.method == "POST":
        #pythoncom.CoInitialize()
        def convert_pdf2docx(input_file: str, output_file: str, pages: Tuple = None):
                 """Converts pdf to docx"""
                 if pages:
                     pages = [int(i) for i in list(pages) if i.isnumeric()]
                 result = parse(pdf_file=input_file,
                   docx_with_path=output_file, pages=pages)
                 summary = {
                     "File": input_file, "Pages": str(pages), "Output File": output_file
                     }
                 # Printing Summary
                 print("## Summary ########################################################")
                 print("\n".join("{}:{}".format(i, j) for i, j in summary.items()))
                 print("###################################################################")
                 return result

        file=request.files['filename']
        if file.filename !='':
            file.save(os.path.join(app.config['UPLOADER_FOLDER'],file.filename))
        print(file.filename)
        #convert(file.filename)
        #convert(file.filename,"document.pdf")
        #return redirect('/pdf')
        #return send_file(file.filename, as_attachment=True)
        wdFormatPDF = 17
        input_file = file.filename
        output_file = r"hell.docx"
        convert_pdf2docx(input_file, output_file)
        #inputFile = os.path.abspath(r"C:\Users\mvirati\Downloads\fie.docx")
        #outputFile = os.path.abspath(r"document.pdf")
        #word = win32com.client.Dispatch('Word.Application')
        #doc = word.Documents.Open(file.filename)
        #doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
        #doc.Close()
        #word.Quit()
        #return render_template("pdf.html")
        return send_file(output_file, as_attachment=True)

    return render_template("index.html")

@app.route('/pdf',methods=['GET','POST'])
def pdf():
    if request.method =="GET":
        #return send_file(file.filename,as_attachment=True)
       return send_file("document.pdf",as_attachment=True)
    print('wrong')
if __name__ == "__main__":
    app.debug=True
    app.run()

