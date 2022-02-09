import flask
import xlsxwriter as xlsxwriter
from flask import Flask, render_template, request

app = Flask(__name__)
app.secret_key = "super secret key"
isAttendanceOn = False
sheetName = ''
counter = 0
wb = ''
list_Of_IP = []


# Just interfaces
@app.route('/')
def home():
    return render_template('HomePage.html')

@app.route('/thankYou')
def thankYou():
    return render_template('ThankYouPage.html')

@app.route('/QR')
def QRcode():
    return render_template('QRCode.html')

@app.route('/form')
def attendance():
    if (isAttendanceOn == True):
        return render_template('index.html')
    else:
        return render_template('ThankYouPage.html')

# backend
@app.route('/', methods=['POST'])
def home_post():
    if request.method == 'POST':
        # take the date and course ID to name the excel sheet
        date = request.form['attendance-start']
        classID = request.form['ClassName']
        # should be changed with the cloud location
        location = r'C:\Users\20115\PycharmProjects\flaskProject\ '
        global sheetName
        sheetName = location + classID + '-' + date + '.xlsx'
        # Workbook is created
        global wb
        wb = xlsxwriter.Workbook(sheetName)
        # add_sheet is used to create sheet.
        sheet1 = wb.add_worksheet('Attendance')
        sheet1.write(0, 0, 'IP Address')
        sheet1.write(0, 1, 'Student Name')
        sheet1.write(0, 2, 'Student ID')
        sheet1.write(0, 3, 'Cheating Case')
        global isAttendanceOn
        isAttendanceOn = True
        if request.form['submit_button'] == 'Create Attndance':
            return flask.redirect("/QR")

@app.route('/QR', methods=['POST'])
def QR_post():
    if request.method == 'POST':
        if request.form['submit_button'] == 'Finish Attndance':
            global isAttendanceOn
            global wb
            isAttendanceOn = False
            wb.close()
            return flask.redirect('/thankYou')

@app.route('/form', methods=['POST'])
def my_form_post():
    global list_Of_IP
    global counter
    global sheetName
    global wb
    name = request.form['Name']
    ID = request.form['ID']
    ip_address = flask.request.remote_addr
    counter += 1
    existingWorksheet = wb.get_worksheet_by_name('Attendance')
    existingWorksheet.write(counter, 0, ip_address)
    existingWorksheet.write(counter, 1, name)
    existingWorksheet.write(counter, 2, ID)
    if (ip_address in list_Of_IP):
        list_Of_IP.append(ip_address)
        existingWorksheet.write(counter, 3, 'yes')
    else:
        list_Of_IP.append(ip_address)
        existingWorksheet.write(counter, 3, 'no')
    return render_template('ThankYouPage.html')

if __name__ == '__main__':
    app.run()
