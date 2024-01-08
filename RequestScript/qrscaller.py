import requests
# Flask app URL
flask_app_url = 'http://127.0.0.1:5000/fetch_results'

data = {
    'class_name': '5CSE1',
    'semester_number': 1,
}

with open('students.csv', 'rb') as file:
    sourcefile = {'source_csv_file': file}
    response = requests.post(flask_app_url, params=data, files=sourcefile)

    if response.status_code == 200:
        downloaded_csv_path = 'response.csv'
        with open(downloaded_csv_path, 'wb') as downloaded_file:
            downloaded_file.write(response.content)
        print(f"Output file saved: {downloaded_csv_path}")
    else:
        print(f"Error: {response.status_code} - {response.text}")
