from tkinter import filedialog
import capsolver
import flask
import requests
from bs4 import BeautifulSoup
import csv
import os
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
cors = CORS(app, resources={r"/*": {"origins": "*"}})



currentFolder = os.getcwd()+"/NCS/QRS"

# Capsolver
capsolver.api_key = "CAP-09961469DA29A03381421590A9C5550E"
websitekey = "6Lcreg8UAAAAANPsXxX59qwh76mUDFzohw-FnNqs"
websiteURL = "https://erp.aktu.ac.in/webpages/oneview/oneview.aspx"
type = "ReCaptchaV2TaskProxyLess"


def saveSemResultToCommonCSV(result):
    rows = result.subjectData
    soup = BeautifulSoup(rows, 'html.parser')

    headings = ["Roll No.","Name","DOB", "Total Marks", "SGPA"]
    subheading = ["","","","",""]
    data = [str(result.rollno),str(result.dob),str(result.name), str(result.totalmarks), str(result.sgpa)]
    print("Total marks : ", str(result.totalmarks), "Total SGPA ", result.sgpa)

    subname_pattern = 'subName'
    subjectNames = soup.find_all('span', id=lambda x: x and subname_pattern in x)

    for subName in subjectNames:
        if subName == subjectNames[0]:
            headings.append(str(subName.text))
            subheading.append("Ext")
            subheading.append("Int")
        else:
            headings.append("")
            headings.append(str(subName.text))
            subheading.append("Ext")
            subheading.append("Int")

    marks = soup.select('tr')
    for row in marks:
        cols = row.find_all(['td', 'th'])
        print("Columns",cols)
        if cols[0].name == 'td':
            external_marks = cols[4].text.strip()
            internal_marks = cols[3].text.strip()
            data.append(str(external_marks))
            data.append(str(internal_marks))

    # print(data)
    csv_file_path = f"{currentFolder}//Results/{result.classcode}/Result-{result.classcode}.csv"

    fileexists = checkExists(csv_file_path)
    print("CSV Exists : ",fileexists)

    with open(csv_file_path, 'a', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)

        if not fileexists or not is_header_present(csv_file_path):
            csv_writer.writerow(headings)
            csv_writer.writerow(subheading)

        csv_writer.writerow(data)

def is_header_present(file_path):
    with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile)

        # Read the first row
        first_row = next(csvreader, None)

        # Check if the first row contains headers
        if first_row is not None:
            # Assuming headers are present if the first element of the row is a string
            return isinstance(first_row[0], str)
        else:
            return False


def readResultFromFile(student, keySem):
    classcode = student.classcode
    rollno = student.rollno

    with open(f"{currentFolder}//Results//{classcode}//Individual//{rollno}//{rollno}.html", "r", encoding='utf-8') as file:
        content = file.read()

        htmlResponse = content
        soup = BeautifulSoup(htmlResponse, 'html.parser')

        semester_id_pattern = 'grdViewSubjectMarksheet'
        sgpa_id_pattern = 'lblSGPA'
        total_marks_id_pattern = 'SemesterTotalMarksObtained'
        studentFullname = soup.find('span', {'id': 'lblFullName'}).text

        if studentFullname:
            student.name = studentFullname

        semesters = soup.find_all('table', id=lambda x: x and semester_id_pattern in x)
        resultSGPAs = soup.find_all('span', id=lambda x: x and sgpa_id_pattern in x)
        totalmarks = soup.find_all('span', id=lambda x: x and total_marks_id_pattern in x)
        print("Student name : ", student.name)
        totalMarksIdx = 1
        semIdx= 1
        for semester in semesters:


            rows = semester.find_all('tr')
            saveToStudentCSV(classcode, rollno, semIdx, rows)

            print('Total marks Debug:',totalmarks[totalMarksIdx].text)
            #print('Total marks Debug Struct :', totalmarks )

            if semIdx == keySem:
                totalmarksfinal = totalmarks[totalMarksIdx].text
                result = Result(classcode=classcode, name=student.name,
                                sgpa=resultSGPAs[semIdx - 1].text,
                                rollno=student.rollno,
                                totalmarks=totalmarksfinal,
                                subjectData=str(rows), dob= student.dob)
                print("Preparing to save required sem result in common csv..")
                saveSemResultToCommonCSV(result)

            print(f"Semester {semIdx} Result Done..")
            semIdx = semIdx + 1
            totalMarksIdx = totalMarksIdx + 2


def saveToStudentCSV(classcode, rollno, sem, rows):
    print("Saving Sem ", sem, " Result..")
    with open(f"{currentFolder}//Results//{classcode}//Individual//{rollno}//Result-Sem-{sem}.csv", 'w', newline='',
              encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)

        for row in rows:
            columns = row.find_all(['th', 'td'])
            row_data = [column.text.strip() for column in columns]
            csv_writer.writerow(row_data)
    print("Done..")


def saveHTMLFile(student):
    classcode = student.classcode
    rollno = student.rollno
    result = student.result

    directory = f"{currentFolder}//Results//{classcode}//Individual//{rollno}"
    os.makedirs(directory, exist_ok=True)
    file_path = os.path.join(directory, f'{rollno}.html')

    with open(file_path, "w+",encoding='utf-8') as file:
        file.write(result)


def hasErrors(content):
    check = "एoकेoटीoयूo-एसडीसी द्वारा संचालित"
    return content.__contains__(check)


def getResultofStudent(index, student, sem):
    try:
        print(f"{index}: Running on {student.rollno}...")

        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'Accept-Language': 'en-US,en;q=0.9',
            'Cache-Control': 'max-age=0',
            'Connection': 'keep-alive',
            'Content-Type': 'application/x-www-form-urlencoded',
            'DNT': '1',
            'Origin': 'https://erp.aktu.ac.in',
            'Referer': 'https://erp.aktu.ac.in/webpages/oneview/oneview.aspx',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-User': '?1',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'sec-ch-ua': '"Not_A Brand";v="8", "Chromium";v="120"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': 'macOS',
        }

        print("Solving captcha (can take 3-15s)...")

        solution = capsolver.solve({
            "type": type,
            "websiteKey": websitekey,
            "websiteURL": websiteURL
        })['gRecaptchaResponse']

        print("Got solution... ", solution)

        data = {
            '__EVENTTARGET': '',
            '__EVENTARGUMENT': '',
            '__VIEWSTATE': '/wEPDwULLTExMDg0MzM4NTIPZBYCAgMPZBYEAgMPZBYEAgkPDxYCHgdWaXNpYmxlaGRkAgsPDxYCHwBnZBYCAgEPDxYCHwBnZBYEAgMPDxYCHgdFbmFibGVkZ2RkAgUPFgIfAWdkAgkPZBYCAgEPZBYCZg9kFgICAQ88KwARAgEQFgAWABYADBQrAABkGAEFEmdyZFZpZXdDb25mbGljdGlvbg9nZJh0WBmmwJNZgNhWJGDYiHOZt4n7',
            '__VIEWSTATEGENERATOR': '309C3D50',
            'txtRollNo': student.rollno,
            'txtDOB': student.dob,
            'g-recaptcha-response': solution,
            'btnSearch': 'खोजें',
            'hidForModel': '',
        }

        print("Fetching response from AKTU OneView...")
        response = requests.post('https://erp.aktu.ac.in/webpages/oneview/oneview.aspx', headers=headers, data=data)
        print("Status code : ", response.status_code)

        if response.status_code == 200 or response.status_code == 301 or response.status_code == 302:
            student.result = response.text

            if not hasErrors(student.result):
                saveHTMLFile(student)
                readResultFromFile(student, sem)
                print(f"Result saved successfully for {student.rollno}...\n")
                print("------------------\n")
                return "Done"
            else:
                print(f"Failure : Invalid details for {student.rollno}...")
                return "Failed"


        else:
            print(f"Failure : ${requests.status_codes}")
            return "Failed"

    except Exception as e:
        print(f"Problem : \n{e}")
        return "Failed"


def checkExists(file_path):
    if os.path.exists(file_path):
        return True
    else:
        return False


def getProblems(classname, sem, source):
    problems = []
    if classname == None or classname == "":
        problems.append("> Error : Enter valid class name..")

    if sem > 8 or sem < 1:
        problems.append("> Error: Enter valid semester number..")

    if not checkExists(source):
        problems.append("> Error : Selected source file doesn't exist..")

    return problems


class Student:
    def __init__(self, classcode, rollno, dob, name=None, result=None):
        self.name = name
        self.classcode = classcode
        self.rollno = rollno
        self.dob = dob
        self.result = result


class Result:
    def __init__(self, classcode, name, rollno, sgpa, totalmarks, subjectData, dob):
        self.classcode = classcode
        self.name = name
        self.rollno = rollno
        self.sgpa = sgpa
        self.totalmarks = totalmarks
        self.subjectData = subjectData
        self.dob = dob



def read_csv_and_filter_status(input_file, status_to_filter):
    filtered_list = []

    with open(input_file, 'r', newline='', encoding='utf-8') as csv_file:
        csv_reader = csv.DictReader(csv_file)

        for row in csv_reader:
            if 'Status' in row and row['Status'] == status_to_filter:
                filtered_list.append(row['Roll number'])

    return filtered_list


def delete_file(file_path):
    try:
        os.remove(file_path)

    except OSError as e:
        print("No log file, this is a new session..")


def exec():
    index = 1
    doneRecords = []
    success = 0

    print("\n\n<==== JSS Quick Result Scraper (QRS) by Nibble Computer Society, CSE Dept. ====> \n")

    classcode = input("Class name (Ex: 5CS1) : ")
    directoryPath = f"{currentFolder}//Results//{classcode}"
    logFile = f"{directoryPath}//{classcode}-log.csv"

    if not os.path.exists(directoryPath):
        os.makedirs(directoryPath)

    if (checkExists(logFile)):
        print("Log file found.. will rerun on failed records..")
        doneRecords = read_csv_and_filter_status(logFile, "Done")

    sem = int(input("Semester number (1-8) : "))
    print("Source CSV File : ", end="")

    source = filedialog.askopenfilename(
        title="Select a source csv file",
        initialdir=os.getcwd(),
        filetypes=[("Text files", "*.csv"), ("All files", "*.*")])
    print(source)
    print("\n")

    problems = getProblems(classcode, sem, source)
    if len(problems) == 0:

        delete_file(logFile)

        with open(source, 'r+', encoding='utf-8') as csv_file, open(logFile, 'w', newline='',
                                                                    encoding='utf-8') as logfile:
            print("Getting results for ", classcode)

            # Preparing csv file
            csv_reader = csv.reader(csv_file)
            next(csv_reader)

            # Preparing log file
            header = ['Roll number', 'Status']
            csv_writer = csv.writer(logfile)
            csv_writer.writerow(header)

            for row in csv_reader:
                rollno, dob = row
                if not doneRecords.__contains__(rollno):
                    student = Student(classcode, str(rollno), str(dob))
                    status = getResultofStudent(index, student, sem)
                    csv_writer.writerow([rollno, status])
                    index = index + 1
                    success = success + 1
                else:
                    print(f"{index}: Skipping on {rollno} as already done..\n")
                    csv_writer.writerow([rollno, 'Done'])
                    index = index + 1
            print("\nFetched ", success, " result.")
    else:
        for problem in problems:
            print(problem)
        print("Resolve errors and try again..")
        exit(0)


@app.route('/hello', methods=['GET'])
def helloworld():
    return 'Hello World'


@app.route('/fetch_results', methods=['POST'])
def fetch_results():
    try:

        print("Got a request..", request.files)

        class_name = request.form['class_name']
        print(class_name)

        semester_number = int(request.form['semester_number'])
        print(semester_number)

        sourcefile = request.files['source_csv_file']
        print('Got a source file : ', sourcefile.name)

        doneRecords = []
        directoryPath = f"{currentFolder}//Results//{class_name}"
        logFile = f"{directoryPath}//{class_name}-log.csv"

        if not os.path.exists(directoryPath):
            os.makedirs(directoryPath)
            source_csv_file = f'{currentFolder}//Results//{class_name}//source.csv'
            sourcefile.save(source_csv_file)

        else :
            source_csv_file = f'{currentFolder}//Results//{class_name}//source.csv'
            sourcefile.save(source_csv_file)

        if (checkExists(logFile)):
            print("Log file found.. will rerun on failed records..")
            doneRecords = read_csv_and_filter_status(logFile, "Done")

        problems = getProblems(class_name, semester_number, source_csv_file)

        if len(problems) == 0:
            global index
            global success
            index = 1
            success = 0

            delete_file(logFile)

            with open(source_csv_file, 'r+', encoding='utf-8') as csv_file, open(logFile, 'w', newline='',
                                                                                 encoding='utf-8') as logfile:
                csv_reader = csv.reader(csv_file)
                next(csv_reader)

                header = ['Roll number', 'Status']
                csv_writer = csv.writer(logfile)
                csv_writer.writerow(header)

                for row in csv_reader:
                    rollno, dob = row
                    if not doneRecords.__contains__(rollno):
                        student = Student(class_name, str(rollno), str(dob))
                        status = getResultofStudent(index, student, semester_number)
                        csv_writer.writerow([rollno, status])
                        index = index + 1
                        success = success + 1
                    else:
                        print(f"{index}: Skipping on {rollno} as already done..\n")
                        csv_writer.writerow([rollno, 'Done'])
                        index = index + 1

            responsefilepath= f"{currentFolder}//Results//{class_name}//Result-{class_name}.csv"
            return flask.send_file(responsefilepath, as_attachment=True)

        else:
            return jsonify({"success": 0, "message": "Resolve errors and try again.", "errors": problems})
    except Exception as e:
        return jsonify({"success": 0, "message": f"Error: {str(e)}"})


print("\n\nRunning QRS Server v2.0 No Cors ...")
print("CWD : ",currentFolder)
if not os.path.exists(currentFolder):
    os.makedirs(currentFolder)

if __name__ == '__main__':
    app.run(debug=True)
