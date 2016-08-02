#! -*-coding:utf-8-*-
import pymongo
import xlrd
import threading
import time
import datetime
import re
import sys

reload(sys)
sys.setdefaultencoding('utf8')

class countTimeThread(threading.Thread):

    def __init__(self):
        super(countTimeThread, self).__init__()
        self.currentTime = time.time()
        self.runFlag = True

    def run(self):
        while self.runFlag:
            actualTime = time.time()
            print 'Loading for: %f s' % (actualTime-self.currentTime)
            time.sleep(1)

    def stop(self):
        self.runFlag = False

thread = countTimeThread()
thread.start()

workbook = xlrd.open_workbook('../Price_Beijing.xls')

thread.stop()
thread.join()

client = pymongo.MongoClient('localhost', 27017)

db = client.test_database
db.posts.remove({})
# db.add_user('priceBeijing', '123')

tagName = []
for booksheet in workbook.sheets():
    row = 0
    for col in xrange(booksheet.ncols):
        cell = booksheet.cell(row, col)
        if col == 9:
            tagName.append("元/平米")
        else:
            tagName.append(cell.value)

for booksheet in workbook.sheets():
    for row in xrange(1, booksheet.nrows):
        currentCol = []
        for col in xrange(booksheet.ncols):
            cell = booksheet.cell(row, col)
            if col == 9:
                txt = re.sub(r'\n', '', cell.value)
                unitPrice = re.findall(r'[1-9][0-9]{0,}', txt)[1]
                currentCol.append(unitPrice)
            else:
                currentCol.append(cell.value)

        # Insert to mongodb
        post_dict = {}
        for i in range(len(tagName)):
            key = tagName[i]
            value = currentCol[i]
            post_dict[key] = value

        posts = db.posts
        post_id = posts.insert_one(post_dict).inserted_id

for post in db.posts.find():
    print post