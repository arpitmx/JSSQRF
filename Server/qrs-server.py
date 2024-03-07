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

currentFolder = os.getcwd() + "/NCS/QRS"

# Capsolver
capsolver.api_key = "CAP-09961469DA29A03381421590A9C5550E"
websitekey = "6Lcreg8UAAAAANPsXxX59qwh76mUDFzohw-FnNqs"
websiteURL = "https://erp.aktu.ac.in/webpages/oneview/oneview.aspx"
type = "ReCaptchaV2TaskProxyLess"

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


# --------------------------------------- Helper Functions ---------------------------------------------------------

def checkExists(file_path):
    if os.path.exists(file_path):
        return True
    else:
        return False


def hasErrors(content):
    check = "एoकेoटीoयूo-एसडीसी द्वारा संचालित"
    return content.__contains__(check)


def getProblems(classname, sem, source):
    problems = []
    if classname == None or classname == "":
        problems.append("> Error : Enter valid class name..")

    if sem > 8 or sem < 1:
        problems.append("> Error: Enter valid semester number..")

    if not checkExists(source):
        problems.append("> Error : Selected source file doesn't exist..")

    return problems


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


class Student:
    def __init__(self, classcode, rollno, dob, name=None, result=None):
        self.name = name
        self.classcode = classcode
        self.rollno = rollno
        self.dob = dob
        self.result = result
        self.hasBack = False


class Result:
    def __init__(self, semester, sgpa, totalmarks, subjectData, intMarks, extMarks, backPaper, grades, subCodes,
                 subjectNames, status, dod):
        self.sgpa = sgpa
        self.totalmarks = totalmarks
        self.subjectData = subjectData
        self.intMarks = intMarks
        self.extMarks = extMarks
        self.subjectNames = subjectNames
        self.status = status
        self.dod = dod
        self.semester = semester
        self.subCode = subCodes
        self.backPaper = backPaper
        self.grades = grades


# ------------------------------------------------------------------------------------------------

pending_back_students_results_data = []


def appendPendingBackSutdentsToCommonCSV(classname):
    csv_file_path = f"{currentFolder}//Results/{classname}/Result-{classname}.csv"

    isCommonCSVExists = checkExists(csv_file_path)
    print("CSV Exists : ", isCommonCSVExists)

    with open(csv_file_path, 'a', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)

        for results in pending_back_students_results_data:

            # Set heading, subheading if not exists in common csv
            if not isCommonCSVExists or not is_header_present(csv_file_path):
                csv_writer.writerow(["Roll No.", "Name", "DOB"])
                csv_writer.writerow(results)

            else:
                csv_writer.writerow(results)


semTemp = -1

def modify_semTemp(change):
    global semTemp
    semTemp = change

def getSemTemp():
    return semTemp

def saveResultToMasterMarksheetCSV(student, marksheet):
    # Setting heading and values for Roll no., Name, DOB, Total Marks, SGPA
    margin = ["   "]
    headings = ["Semester","Roll No.", "Name", "DOB", "Declaration", "Total Marks", "SGPA", "Status"]
    subheading = ["", "", "", "", "", "", "", ""]
    marks_value = [marksheet.semester,str(student.rollno), str(student.name), str(student.dob), marksheet.dod,
                   str(marksheet.totalmarks),
                   str(marksheet.sgpa), marksheet.status]

    # Setting subject marks in heading and adding external and internal in subheading list
    for subject in range(0, len(marksheet.subjectNames)):
        marks_value.append(str(marksheet.intMarks[subject]))
        marks_value.append(str(marksheet.extMarks[subject]))
        marks_value.append(marksheet.backPaper[subject])
        marks_value.append(marksheet.grades[subject])

    csv_file_path = f"{currentFolder}//Results/{student.classcode}/Master-Marksheet-{student.classcode}.csv"

    isCommonCSVExists = checkExists(csv_file_path)
    print("CSV Exists : ", isCommonCSVExists)

    with open(csv_file_path, 'a', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)

        # Set heading, subheading if not exists in common csm
        if getSemTemp() != int(marksheet.semester):

            for sub in range(0, len(marksheet.subjectNames)):
                headings.append(f"{marksheet.subCode[sub]} ({marksheet.subjectNames[sub]})")
                subheading.append("Int")
                subheading.append("Ext")
                subheading.append("Back")
                subheading.append("Grade")
                headings.append("")
                headings.append("")
                headings.append("")

            csv_writer.writerow(margin)
            csv_writer.writerow(headings)
            csv_writer.writerow(subheading)

            modify_semTemp(int(marksheet.semester))

        csv_writer.writerow(marks_value)


# if not isCommonCSVExists or not is_header_present(csv_file_path):
#     # Setting subject name in heading and adding external and internal in subheading list
#     for sub in range (0, len(marksheet.subjectNames)):
#
#             headings.append(f"{marksheet.subCode[sub]} ({marksheet.subjectNames[sub]})")
#             subheading.append("Int")
#             subheading.append("Ext")
#             subheading.append("Back")
#             subheading.append("Grade")
#             headings.append("")
#             headings.append("")
#             headings.append("")
#
#     csv_writer.writerow(headings)
#     csv_writer.writerow(subheading)


def saveSemResultToCommonCSV(student, marksheet):
    # Setting heading and values for Roll no., Name, DOB, Total Marks, SGPA
    headings = ["Roll No.", "Name", "DOB", "Semester", "Declaration", "Total Marks", "SGPA", "Status"]
    subheading = ["", "", "", "", "", "", "", ""]
    marks_value = [str(student.rollno), str(student.name), str(student.dob), marksheet.semester, marksheet.dod,
                   str(marksheet.totalmarks),
                   str(marksheet.sgpa), marksheet.status]

    print("Total marks : ", str(marksheet.totalmarks), "Total SGPA ", marksheet.sgpa)

    # Setting subject marks in heading and adding external and internal in subheading list
    for subject in range(0, len(marksheet.subjectNames)):
        marks_value.append(str(marksheet.intMarks[subject]))
        marks_value.append(str(marksheet.extMarks[subject]))
        marks_value.append(marksheet.backPaper[subject])
        marks_value.append(marksheet.grades[subject])

    csv_file_path = f"{currentFolder}//Results/{student.classcode}/Result-{student.classcode}.csv"

    isCommonCSVExists = checkExists(csv_file_path)
    print("CSV Exists : ", isCommonCSVExists)

    with open(csv_file_path, 'a', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)

        # Set heading, subheading if not exists in common csv
        if not isCommonCSVExists or not is_header_present(csv_file_path):
            # Setting subject name in heading and adding external and internal in subheading list
            for sub in range(0, len(marksheet.subjectNames)):
                headings.append(f"{marksheet.subCode[sub]} ({marksheet.subjectNames[sub]})")
                subheading.append("Int")
                subheading.append("Ext")
                subheading.append("Back")
                subheading.append("Grade")
                headings.append("")
                headings.append("")
                headings.append("")

            csv_writer.writerow(headings)
            csv_writer.writerow(subheading)

        csv_writer.writerow(marks_value)


SEM_TOTAL_MARKS = 'totalMarks'
SEM_DOD = 'DOD'
SEM_RESULT_STATUS = 'resultStatus'
SEM_SGPA = 'sgpa'
SEM_SUBJECTWISE_MARKS = 'result'
SEM_RAW = 'raw'
SEMESTER_IND = 'semester'

YEAR_SESSION = 'yearSession'
YEAR_SEMESTERS = 'yearSemesters'
YEAR_RESULT_STATUS = 'yearStatus'
YEAR_TOTAL_MARKS = 'yearTotalMarks'
YEAR_COP = 'yearCOP'
YEAR_AUDIT1 = 'yearAudit1'
YEAR_AUDIT2 = 'yearAudit2'
YEAR_DATA = 'yearData'


def getSemResultsFromHTML(student):
    classcode = student.classcode
    rollno = student.rollno

    semResults = {}

    with (open(f"{currentFolder}//Results//{classcode}//Individual//{rollno}//{rollno}.html", "r",
               encoding='utf-8') as file):

        print("Getting semester results...")

        content = file.read()
        htmlResponse = content
        soup = BeautifulSoup(htmlResponse, 'html.parser')

        semester_id_pattern = 'grdViewSubjectMarksheet'
        sgpa_id_pattern = 'lblSGPA'
        semester_total_marks_id_pattern = 'SemesterTotalMarksObtained'
        semester_result_status_pattern = 'lblResultStatus'
        semester_index_pattern = "lblSemesterId"
        semester_dod_patter = 'lblDateOfDeclaration'

        # Year patterns
        year_status_pattern = 'lblResult'
        year_total_marks_pattern = 'lblMarks'
        year_cop_pattern = 'lblCOP'
        year_audit1 = 'lblIsAUCPassed'
        year_audit2 = 'lblIsCybPassed'
        year_sem_pattern = "lblSem"
        year_session_pattern = "lblSession"

        # Saving semester data, semester index, date of declaration, SGPA, totalmarks, full name, result status
        semesters = soup.find_all('table', id=lambda x: x and semester_id_pattern in x)
        semesterIndexes = [int(span.text) for span in soup.find_all('span') if
                           span.get('id') and span.get('id').endswith('_lblSemesterId')]
        dods = [span.text for span in soup.find_all('span') if
                span.get('id') and span.get('id').endswith(semester_dod_patter)]
        resultSGPAs = [span.text for span in soup.find_all('span') if
                       span.get('id') and span.get('id').endswith(sgpa_id_pattern)]
        totalmarks = soup.find_all('span', id=lambda x: x and semester_total_marks_id_pattern in x)
        studentFullname = soup.find('span', {'id': 'lblFullName'}).text
        result_status = [span.text for span in soup.find_all('span') if
                         span.get('id') and span.get('id').endswith(semester_result_status_pattern)]

        # Year details : Status, Total Marks, COP
        year_result_status = [span.text for span in soup.find_all('span') if
                              span.get('id') and span.get('id').endswith(year_status_pattern)]
        year_total_marks = [span.text for span in soup.find_all('span') if
                            span.get('id') and span.get('id').endswith(year_total_marks_pattern)]
        year_cop = [span.text for span in soup.find_all('span') if
                    span.get('id') and span.get('id').endswith(year_cop_pattern)]
        year_audit1 = [span.text for span in soup.find_all('span') if
                       span.get('id') and span.get('id').endswith(year_audit1)]
        year_audit2 = [span.text for span in soup.find_all('span') if
                       span.get('id') and span.get('id').endswith(year_audit2)]
        year_sem = [span.text for span in soup.find_all('span') if
                    span.get('id') and span.get('id').endswith(year_sem_pattern)]
        year_session = [span.text for span in soup.find_all('span') if
                        span.get('id') and span.get('id').endswith(year_session_pattern)]

        year_data = {}

        print("Year result : ", year_result_status)
        print("Year total marks: ", year_total_marks)
        print("Year COP: ", year_cop)
        print("Year Audit 1: ", year_audit1)
        print("Year Audit 2: ", year_audit2)
        print("Year Semester: ", year_sem)
        print("Year Session: ", year_session)

        yearidx = 0
        for sem in year_sem:
            year_data[sem] = {
                YEAR_SESSION: year_session[yearidx],
                YEAR_SEMESTERS: sem,
                YEAR_RESULT_STATUS: year_result_status[yearidx],
                YEAR_TOTAL_MARKS: year_total_marks[yearidx],
                YEAR_COP: year_cop[yearidx],
                YEAR_AUDIT1: year_audit1[yearidx],
                YEAR_AUDIT2: year_audit2[yearidx]

            }
            yearidx = yearidx + 1

        print("Year data : ", year_data)

        # Saving student name
        if studentFullname:
            student.name = studentFullname
            print("Student name : ", student.name)

        spanIndex = 1
        semIdx = 0

        for semester in semesters:
            print("Index : ", semIdx)
            rows = semester.find_all('tr')
            row_data = []

            for row in rows:
                columns = row.find_all(['th', 'td'])
                row_data.append([column.text.strip() for column in columns])

            semResults[semesterIndexes[semIdx]] = {
                SEM_TOTAL_MARKS: totalmarks[spanIndex].text,
                SEM_DOD: dods[spanIndex],
                SEM_RESULT_STATUS: result_status[spanIndex],
                SEM_SGPA: resultSGPAs[semIdx],
                SEM_SUBJECTWISE_MARKS: row_data,
                SEM_RAW: semester,
                SEMESTER_IND: semesterIndexes[semIdx]
            }

            for key, value in year_data.items():
                if key.__contains__(str(semesterIndexes[semIdx])):
                    print("Session:", key, " contains :", semesterIndexes[semIdx])
                    print("Year data : ", value)
                    semResults[semesterIndexes[semIdx]][YEAR_DATA] = value

            semIdx = semIdx + 1
            spanIndex = spanIndex + 2

    return semResults


# --------------------------------------- TESTING ---------------------------------------------------------

def getResult(semResult):
    semester = semResult.get(SEMESTER_IND)
    status = semResult.get(SEM_RESULT_STATUS)
    dod = semResult.get(SEM_DOD)
    subData = semResult.get(SEM_RAW)
    totalMarks = semResult.get(SEM_TOTAL_MARKS)
    markList = semResult.get(SEM_SUBJECTWISE_MARKS)
    sgpa = semResult.get(SEM_SGPA)

    subjectNames = []
    subCodes = []
    intMarks = []
    extMarks = []
    backPaper = []
    grades = []

    # Subject details
    for row in range(1, len(markList)):
        subCodes.append(markList[row][0])
        subjectNames.append(markList[row][1])
        intMarks.append(markList[row][3])
        extMarks.append(markList[row][4])
        backPaper.append(markList[row][5])
        grades.append(markList[row][6])

    # print(subCodes)
    # print(subjectNames)
    # print(intMarks)
    # print(extMarks)
    # print(backPaper)
    # print(grades)

    return Result(semester=semester, sgpa=sgpa, totalmarks=totalMarks, subjectData=subData, dod=dod, status=status,
                  subCodes=subCodes, subjectNames=subjectNames, intMarks=intMarks, extMarks=extMarks, grades=grades,
                  backPaper=backPaper)


def createMasterMarksheetCSV(student_list):

    # studentx = Student(classcode="Demo", rollno="1900910100001", dob="25/03/2001")
    # student2 = Student(classcode="Demo", rollno="1900910100002", dob="07/01/2001")
    # student3 = Student(classcode="Demo", rollno="1900910100003", dob="26/03/2000")
    # student4 = Student(classcode="Demo", rollno="1900910100004", dob="26/03/2000")
    # student5 = Student(classcode="Demo", rollno="1900910100005", dob="26/03/2000")
    # student6 = Student(classcode="Demo", rollno="1900910100006", dob="26/03/2000")
    # student7 = Student(classcode="Demo", rollno="1900910100007", dob="26/03/2000")
    # student8 = Student(classcode="Demo", rollno="1900910100008", dob="26/03/2000")
    # student9 = Student(classcode="Demo", rollno="1900910100009", dob="26/03/2000")
    # student10 = Student(classcode="Demo",rollno="1900910100010", dob="26/03/2000")
    #
    # li = [studentx, student2, student3, student4, student5, student6, student7,student8,student9,student10]

    modify_semTemp(-1)

    for sem in range(1, 9):
        for student in student_list:
            marksheet = getSemResultsFromHTML(student)
            semester = marksheet.get(sem)
            semResult = getResult(semester)
            saveResultToMasterMarksheetCSV(student, semResult)


def testing():
    studentx = Student(classcode="8thSemTest", rollno="1900910100001", dob="25/03/2001")
    student2 = Student(classcode="8thSemTest", rollno="1900910100002", dob="07/01/2001")
    student3 = Student(classcode="8thSemTest", rollno="1900910100003", dob="26/03/2000")
    student4 = Student(classcode="8thSemTest", rollno="1900910100004", dob="26/03/2000")
    student5 = Student(classcode="8thSemTest", rollno="1900910100005", dob="26/03/2000")
    li = [studentx, student2, student3, student4, student5]

    for student in li:
        keySem = 3
        marksheet = getSemResultsFromHTML(student)

        for sem in range(1, 9):
            semester = marksheet.get(sem)
            semesterRaw = semester.get(SEM_RAW)
            semResult = getResult(semester)

            saveToStudentIndividualCSV(classcode=student.classcode, rows=semesterRaw.find_all('tr'), sem=sem,
                                       rollno=student.rollno)
            saveToStudentIndividualCommulativeCSV(classcode=student.classcode, semData=semester,
                                                  rows=semesterRaw.find_all('tr'), sem=sem, rollno=student.rollno)

            if sem == keySem:
                saveSemResultToCommonCSV(student=student, marksheet=semResult)

    createMasterMarksheetCSV([])


# ------------------------------------------------------------------------------------------------


def readAndSaveResultFromIndividualSavedHTMLFile(student, keySem):

    marksheet = getSemResultsFromHTML(student)


    for sem in range(1, 9):
        semester = marksheet.get(sem)
        semesterRaw = semester.get(SEM_RAW)
        semResult = getResult(semester)

        saveToStudentIndividualCSV(classcode=student.classcode, rows=semesterRaw.find_all('tr'), sem=sem,
                                   rollno=student.rollno)
        saveToStudentIndividualCommulativeCSV(classcode=student.classcode, semData=semester,
                                              rows=semesterRaw.find_all('tr'), sem=sem, rollno=student.rollno)

        if sem == keySem:
            saveSemResultToCommonCSV(student=student, marksheet=semResult)



def saveToStudentIndividualCSV(classcode, rollno, sem, rows):
    print("Saving Sem ", sem, " Result..")
    with open(f"{currentFolder}//Results//{classcode}//Individual//{rollno}//Result-Sem-{sem}.csv", 'w', newline='',
              encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)

        for row in rows:
            columns = row.find_all(['th', 'td'])
            row_data = [column.text.strip() for column in columns]
            csv_writer.writerow(row_data)
    print("Done..")


def saveToStudentIndividualCommulativeCSV(classcode, semData, rollno, sem, rows):
    yearData = semData.get(YEAR_DATA)
    print("Saving Sem ", sem, " Result..")
    with open(f"{currentFolder}//Results//{classcode}//Individual//{rollno}//Common-Result.csv", 'a', newline='',
              encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)
        row_data = []
        for row in rows:
            columns = row.find_all(['th', 'td'])
            row_data.append([column.text.strip() for column in columns])

        row_data.append([f'{yearData.get(YEAR_SESSION)}',
                         f'TOTAL MARKS (SEM): {semData.get(SEM_TOTAL_MARKS)}', f'SGPA : {semData.get(SEM_SGPA)}',
                         f'STATUS : {semData.get(SEM_RESULT_STATUS)}',
                         f'DOD : {semData.get(SEM_DOD)}',
                         f'TOTAL MARKS (YEAR) : {yearData.get(YEAR_TOTAL_MARKS)}',
                         f'{yearData.get(YEAR_COP)}',
                         yearData.get(YEAR_AUDIT1),
                         yearData.get(YEAR_AUDIT2)
                         ]

                        )

        row_data.append(["   "])

        for i in range(0, len(row_data)):
            if i == 0:
                row_data[i].insert(0, 'Semester')
            elif i == len(row_data) - 2:
                row_data[i].insert(0, 'RESULT -> ')
            elif i < len(row_data) - 2:
                row_data[i].insert(0, semData.get(SEMESTER_IND))

            csv_writer.writerow(row_data[i])

        print(row_data)


def saveResultHTMLFile(student):
    classcode = student.classcode
    rollno = student.rollno
    result = student.result

    directory = f"{currentFolder}//Results//{classcode}//Individual//{rollno}"
    os.makedirs(directory, exist_ok=True)
    file_path = os.path.join(directory, f'{rollno}.html')

    with open(file_path, "w+", encoding='utf-8') as file:
        file.write(result)


def getResultofStudent(index, student, sem):
    try:

        print(f"{index}: Running on {student.rollno}...")

        solution = capsolver.solve({
            "type": type,
            "websiteKey": websitekey,
            "websiteURL": websiteURL
        })['gRecaptchaResponse']

        print("Got solution... ")

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

                saveResultHTMLFile(student)
                readAndSaveResultFromIndividualSavedHTMLFile(student, sem)


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



# --------------------------------- FLASK ENDPOINTS ---------------------------------------------------------------

student_list=[]

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

        else:
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

                        student_list.append(student)

                        csv_writer.writerow([rollno, status])
                        index = index + 1
                        success = success + 1
                    else:
                        print(f"{index}: Skipping on {rollno} as already done..\n")
                        csv_writer.writerow([rollno, 'Done'])
                        index = index + 1


                createMasterMarksheetCSV(student_list)

            responsefilepath = f"{currentFolder}//Results//{class_name}//Result-{class_name}.csv"
            return flask.send_file(responsefilepath, as_attachment=True)

        else:
            return jsonify({"success": 0, "message": "Resolve errors and try again.", "errors": problems})
    except Exception as e:
        return jsonify({"success": 0, "message": f"Error: {str(e)}"})


print("\n\nRunning QRS Server v3.0 No Cors ...")
print("CWD : ", currentFolder)

if not os.path.exists(currentFolder):
    os.makedirs(currentFolder)

#testing()

if __name__ == '__main__':
    app.run(debug=True, port=5000)
