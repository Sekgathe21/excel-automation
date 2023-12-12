from openpyxl import Workbook

# Writing some data into new Excel spreadsheet
infor = {
   "Dlamini": {
      "Initials": "KT",
      "math": 67,
      "science": 70,
      "english": 92,
      "art": 85,
      "bursary(R)": 8900
   },
   "Nkosi": {
         "Initials": "A",
         "math": 56,
         "science": 77,
         "english": 80,
         "art": 93,
         "bursary(R)": 7400
    },
   "Ndlovu": {
         "Initials": "TM",
         "math": 21,
         "science": 45,
         "english": 30,
         "art": 17,
         "bursary(R)": 5700
   },
   "Matlala": {
         "Initials": "S",
         "math": 87,
         "science": 78,
         "english": 89,
         "art": 90,
         "bursary(R)": 8900
   },
   "Maleka": {
         "Initials": "M",
         "math": 88,
         "science": 90,
         "english": 87,
         "art": 80,
         "bursary(R)": 6400
   },
   "Khumalo": {
         "Initials": "HJ",
         "math": 15,
         "science": 37,
         "english": 12,
         "art": 0,
         "bursary(R)": 5200
   },
   "Mahlangu": {
         "Initials": "N",
         "math": 26,
         "science": 30,
         "english": 30,
         "art": 15,
         "bursary(R)": 7400
   },
   "Mokoena": {
         "Initials": "P",
         "math": 100,
         "science": 95,
         "english": 68,
         "art": 68,
         "bursary(R)": 5400
   },
   "Khoza": {
         "Initials": "R",
         "math": 27,
         "science": 30,
         "english": 48,
         "art": 49,
         "bursary(R)": 7050
   },
   "Mofokeng": {
         "Initials": "TP",
         "math": 72,
         "science": 81,
         "english": 53,
         "art": 70,
         "bursary(R)": 7700
   },
   "Smith": {
         "Initials": "KP",
         "math": 44,
         "science": 82,
         "english": 51,
         "art": 92,
         "bursary(R)": 5200
   },
   "Baloyi": {
         "Initials": "TJ",
         "math": 72,
         "science": 45,
         "english": 62,
         "art": 61,
         "bursary(R)": 6400
   },
   "Pillay": {
         "Initials": "KC",
         "math": 55,
         "science": 83,
         "english": 65,
         "art": 76,
         "bursary(R)": 7050
   },
   "Mathebula": {
         "Initials": "SP",
         "math": 36,
         "science": 99,
         "english": 91,
         "art": 40,
         "bursary(R)": 6500
   },
   "Zwane": {
         "Initials": "ME",
         "math": 92,
         "science": 79,
         "english": 43,
         "art": 53,
         "bursary(R)": 7200
   },
   "Tshabalala": {
         "Initials": "Z",
         "math": 66,
         "science": 88,
         "english": 89,
         "art": 72,
         "bursary(R)": 6500
   },
   "van Wyk": {
         "Initials": "WF",
         "math": 83,
         "science": 99,
         "english": 64,
         "art": 100,
         "bursary(R)": 8200
   },
   "Williams": {
         "Initials": "TT",
         "math": 90,
         "science": 38,
         "english": 73,
         "art": 70,
         "bursary(R)": 8000
   },
   "Chauke": {
         "Initials": "DI",
         "math": 78,
         "science": 45,
         "english": 38,
         "art": 68,
         "bursary(R)": 7200
   },
   "Cele": {
         "Initials": "B",
         "math": 59,
         "science": 35,
         "english": 71,
         "art": 69,
         "bursary(R)": 5300
   },
   "van der Merwe": {
         "Initials": "ZV",
         "math": 14,
         "science": 18,
         "english": 16,
         "art": 15,
         "bursary(R)": 7400
   },
   "Maluleke": {
         "Initials": "Y",
         "math": 53,
         "science": 45,
         "english": 88,
         "art": 65,
         "bursary(R)": 10000
   },
   "Molefe": {
         "Initials": "HJK",
         "math": 42,
         "science": 98,
         "english": 53,
         "art": 90,
         "bursary(R)": 7800
   },
   "Mnisi": {
         "Initials": "SD",
         "math": 95,
         "science": 56,
         "english": 58,
         "art": 79,
         "bursary(R)": 7700
   },
   "Motaung": {
         "Initials": "QW",
         "math": 87,
         "science": 92,
         "english": 36,
         "art": 99,
         "bursary(R)": 5200
   },
   "Moodley": {
         "Initials": "OP",
         "math": 13,
         "science": 30,
         "english": 28,
         "art": 19,
         "bursary(R)": 6600
   },
   "Mohlala": {
         "Initials": "BNM",
         "math": 76,
         "science": 81,
         "english": 70,
         "art": 63,
         "bursary(R)": 6600
   },
   "Mothiba": {
         "Initials": "TR",
         "math": 40,
         "science": 64,
         "english": 50,
         "art": 98,
         "bursary(R)": 10000
   },
   "Mabasa": {
         "Initials": "EK",
         "math": 53,
         "science": 43,
         "english": 80,
         "art": 84,
         "bursary(R)": 5200
   }
}

# Creating new workbook with a Student worksheet
wb = Workbook()
ws = wb.active
ws.title = "Students"

# Writing in the data
headings = ['Surname'] + list(infor['Matlala'].keys())
ws.append(headings)

for students in infor:
    data = list(infor[students].values())
    ws.append([students] + data)

wb.save("dataResultsMain.xlsx")
