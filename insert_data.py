from pymongo import MongoClient #データ投入用
import score

client = MongoClient('localhost', 27017)
db = client.score_database
collection = db.test_collection


number_count = 1
for a in range(0,14):
    e = 9
    for i in range(0,41):
        value = score.score[a][i]
        e += 1
        insert_data = {'number':score.number[a],'score':value,'origin':e}
        result = collection.insert_one(insert_data)