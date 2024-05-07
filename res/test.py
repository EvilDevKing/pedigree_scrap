import requests

body_data = {
    "CustomerEmailAddress": "pascalmartin973@gmail.com",
    "CustomerID": 1,
    "HorseName": "Brown",
    "RecordOutputTypeCode": "P",
    "RegistrationNumber": "0419888",
    "ReportCode": 21,
    "ReportId": 10008
}
r1 = requests.post("https://aqhaservices2.aqha.com/api/FreeRecords/SaveFreeRecord", json=body_data)
if r1.status_code == 200:
    if r1.json():
        print("Sent!")
    else:
        print("Failed!")