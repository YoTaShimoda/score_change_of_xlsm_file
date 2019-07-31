from pymongo import MongoClient # エクセルファイルに書き込み用
import openpyxl
import score
client = MongoClient('localhost', 27017)
db = client.score_database
collection = db.test_collection
file_name = '標準得点換算表.xlsx'
sheet_name = '2018年OSIデータ'
file = openpyxl.load_workbook(file_name)
sheet = file[sheet_name]

colm = 3
for i in range(0,14):
    number = score.number[i]
    for a in range(2,57):
        value = sheet.cell(row=a, column=colm).value
        score_value = collection.find_one({'$and': [{'number':number},{'origin':value}]})
        print(score_value['score'])
        sheet.cell(row=a, column = colm).value = score_value['score']
    colm += 1

file.save(file_name)


