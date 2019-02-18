# Dumps all the data to the file
import json
import requests
import xlwt
import xlrd
from xlutils.copy import copy

response = requests.get("...")                     #Google API key
todos = json.loads(response.text)                  #converts json to python dictionary
book1 = xlwt.Workbook("data.xls")
sheet2 = book1.add_sheet("Sheet 1")
#print(todos.keys())
for result in todos.keys():
        if result == 'results':
                #print(todos[i])
                data = {k: v for dct in todos[result] for k, v in dct.items()}                   #convert list to dictionary
                i = 0
                for heading in data.keys():
                        sheet2.write(0, i, heading)                                 #write headings
                        i = i+1
book1.save("data.xls")
book = xlrd.open_workbook("data.xls")
sheet1 = book.sheet_by_name("Sheet 1")                       #read mode
book1 = copy(book)
sheet2 = book1.get_sheet("Sheet 1")                          #write mode
res = todos["results"]
length = len(res)                                           #number of details = row

for l in range(length):                                               #row loop
    #data1 = {k: v for dct in res for k, v in dct.items()}
    #print(data1)
    for col in range(len(res[l].keys())):                             #col loop
        try:
            fetch = sheet1.cell_value(0, col)
            lol = str(res[l][fetch])  # convert dictionary to string
            print("Key = ", fetch, "Value= ", lol)
            sheet2.write(l + 1, col, lol)
        except KeyError:
            lol = "skipped"
            print("Key= ", fetch,"row no.= ", l)
            sheet2.write(l + 1, col, lol)
    print(l, col)

book1.save("data.xls")

