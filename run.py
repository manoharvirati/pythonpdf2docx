from flask import Flask
from flask import request,render_template,redirect,url_for,send_file
import os
import win32com.client
import pythoncom
pythoncom.CoInitialize()

UPLOADER_FOLDER=''
AllOWED_EXTENSIONS={'docx'}

app = Flask(__name__)
app.config['UPLOADER_FOLDER']=UPLOADER_FOLDER

@app.route('/')
@app.route('/index',methods=['GET','POST'])
def index():
    if request.method == "POST":
        pythoncom.CoInitialize()

        file=request.files['filename']
        if file.filename !='':
            file.save(os.path.join(app.config['UPLOADER_FOLDER'],file.filename))
        print(file.filename)
        #return redirect('/pdf')
        #return send_file(file.filename, as_attachment=True)
        wdFormatPDF = 17

        #inputFile = os.path.abspath(r"C:\Users\mvirati\Downloads\fie.docx")
        outputFile = os.path.abspath(r"document.pdf")
        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Open(file.filename)
        doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        return render_template("pdf.html")
        #return send_file("document.pdf", as_attachment=True)

    return render_template("index.html")

@app.route('/pdf',methods=['GET','POST'])
def pdf():
    if request.method =="GET":
       return send_file("document.pdf",as_attachment=True)
    print('wrong')
if __name__ == "__main__":
    app.debug=True
    app.run()

