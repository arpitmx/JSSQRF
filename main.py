from tkinter import filedialog

import capsolver
import requests
from bs4 import BeautifulSoup
import csv
import os

# Capsolver
capsolver.api_key = "CAP-09961469DA29A03381421590A9C5550E"
websitekey = "6Lcreg8UAAAAANPsXxX59qwh76mUDFzohw-FnNqs"
websiteURL = "https://erp.aktu.ac.in/webpages/oneview/oneview.aspx"
type = "ReCaptchaV2TaskProxyLess"


def saveSemResultToCommonCSV(result):
    rows = result.subjectData
    soup = BeautifulSoup(rows, 'html.parser')

    headings = ["Roll No.", "Name", "Total Marks", "SGPA"]
    data = [str(result.rollno), str(result.name), str(result.totalmarks), str(result.sgpa)]
    # print("Total marks : ", result.totalmarks, "Total SGPA ", result.sgpa)

    subname_pattern = 'subName'
    subjectNames = soup.find_all('span', id=lambda x: x and subname_pattern in x)

    for subName in subjectNames:
        headings.append(str(subName.text))

    marks = soup.select('tr')
    for row in marks:
        cols = row.find_all(['td', 'th'])
        if cols[0].name == 'td':
            external_marks = cols[4].text.strip()
            data.append(str(external_marks))

    # print(data)
    csv_file_path = f"Results//{result.classcode}//Result-{result.classcode}.csv"

    # Write data to the CSV file
    with open(csv_file_path, 'a', newline='') as csv_file:
        csv_writer = csv.writer(csv_file)

        if (not is_header_present(csv_file_path)):
            csv_writer.writerow(headings)

        csv_writer.writerow(data)


def is_header_present(file_path):
    with open(file_path, 'r', newline='') as csvfile:
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

    with open(f"Results//{classcode}//Individual//{rollno}//{rollno}.html", "r") as file:
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
        semIdx = 1
        for semester in semesters:
            rows = semester.find_all('tr')

            saveToStudentCSV(classcode, rollno, semIdx, rows)

            if semIdx == keySem:

                result = Result(classcode=classcode, name=student.name,
                                sgpa=resultSGPAs[semIdx - 1].text,
                                rollno=student.rollno,
                                totalmarks=totalmarks[semIdx + 1].text,
                                subjectData=str(rows))
                print("Preparing to save required sem result in common csv..")
                saveSemResultToCommonCSV(result)


            print(f"Semester {semIdx} Result Done..")
            semIdx = semIdx + 1


def saveToStudentCSV(classcode, rollno, sem, rows):
    print("Saving Sem ", sem, " Result..")
    with open(f"Results//{classcode}//Individual//{rollno}//Result-Sem-{sem}.csv", 'w', newline='',
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

    directory = f"Results//{classcode}//Individual//{rollno}"
    os.makedirs(directory, exist_ok=True)
    file_path = os.path.join(directory, f'{rollno}.html')

    with open(file_path, "w+") as file:
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
            else:
                print(f"Failure : Invalid details for {student.rollno}...")

        else:
            print(f"Failure : ${requests.status_codes}")

    except Exception as e:
            print(f"Problem : \n{e}")


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
    def __init__(self, classcode, name, rollno, sgpa, totalmarks, subjectData):
        self.classcode = classcode
        self.name = name
        self.rollno = rollno
        self.sgpa = sgpa
        self.totalmarks = totalmarks
        self.subjectData = subjectData


def exec():
    index = 1
    print("<==== JSS Result Scraper by Nibble Computer Society, CSE Dept. ====> \n")
    classcode = input("Class name (Ex: 5CS1) : ")
    sem = int(input("Semester number (1-8) : "))
    print("Source CSV File : ", end="")
    source = filedialog.askopenfilename(title="Select a source csv file", initialdir=os.getcwd(),
                                        filetypes=[("Text files", "*.csv"), ("All files", "*.*")])
    print(source)
    print("\n")

    problems = getProblems(classcode, sem, source)
    if len(problems) == 0:

        with open(source, 'r+') as csv_file:
            print("Getting results for ", classcode)
            csv_reader = csv.reader(csv_file)
            next(csv_reader)
            for row in csv_reader:
                rollno, dob = row
                student = Student(classcode, str(rollno), str(dob))
                getResultofStudent(index, student, sem)
                index = index + 1

            print("\nSuccessfully got ",index, " results.")
    else:
        for problem in problems:
            print(problem)
        print("Resolve errors and try again..")
        exit(0)


exec()

