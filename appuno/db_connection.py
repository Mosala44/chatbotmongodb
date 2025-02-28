import pymongo
URI = 'mongodb://localhost:27017'
client= pymongo.MongoClient(URI)
db = client['analisisdb']