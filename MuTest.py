import pymysql
import json
import requests
import time
import pymongo
import pandas as pd
import pymongo
from flask import Flask, render_template, request, send_file
from flask_cors import CORS
from openpyxl import load_workbook

myclient = pymongo.MongoClient("mongodb://guvidbmongo:guviGEEK9Mongo7@localhost")
mydb = myclient["userData"]
mycol = mydb["userObject"]

app = Flask(__name__)
CORS(app)


def Course_lister(email):
    mydb = pymysql.connect(
        host="localhost",
        user="guvi",
        passwd="guviDEV#007",
        db="guvi")
    mycursor = mydb.cursor()
    query = "SELECT `course` FROM `coupon_user` WHERE `userid` = '" + email + "'"
    mycursor.execute(query)
    a = []
    for x in mycursor:
        a.append(x[0])
    print(a)


def hash_fetch(email):
    mydb = pymysql.connect(
        host="localhost",
        user="guvi",
        passwd="guviDEV#007",
        db="guvi")
    mycursor = mydb.cursor()
    query = "SELECT `HASH` FROM `userdetails` WHERE `userid` = '" + email + "'"
    mycursor.execute(query)
    a = []
    for x in mycursor:
        a.append(x[0])
    print(a)
    return a


def progress(hash):
    mydb = pymysql.connect(
        host="localhost",
        user="guvi",
        passwd="guviDEV#007",
        db="guvi")
    mycursor = mydb.cursor()
    query = "SELECT `course` FROM `course_point` WHERE `hash` = '" + hash[0] + "'"
    mycursor.execute(query)
    a = []
    for x in mycursor:
        a.append(x[0])
    query = "SELECT `score` FROM `course_point` WHERE `hash` = '" + hash[0] + "'"
    mycursor.execute(query)
    b = []
    for x in mycursor:
        b.append(x[0])
    result = []
    ins = 0
    for i in a:
        result.append({i: b[ins]})
        ins = ins + 1
    return a, json.dumps(result)

    # query = "SELECT 'course', 'score' FROM 'course_point' WHERE `hash` = %s"
    # mycursor = mydb.cursor()
    # mycursor.execute(query, (hash[0],))
    # result = mycursor.fetchall()
    # print(result)


def CouponLister(mail, coupon):
    mydb = pymysql.connect(
        host="localhost",
        user="guvi",
        passwd="guviDEV#007",
        db="guvi")
    mycursor = mydb.cursor()
    query = "SELECT `course` FROM `coupon_user` WHERE `coupon_code` = '" + coupon + "' AND `userid` = '" + mail + "'"
    mycursor.execute(query)
    a = []
    for x in mycursor:
        a.append(x[0])
    return a


def CouponEmail(coupon):
    mydb = pymysql.connect(
        host="localhost",
        user="guvi",
        passwd="guviDEV#007",
        db="guvi")
    mycursor = mydb.cursor()
    query = "SELECT `userid` FROM `coupon_user` WHERE `coupon_code` = '" + coupon + "'"
    mycursor.execute(query)
    a = []
    for x in mycursor:
        a.append(x[0])
    return a


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        user = request.args.get('coupon')
        coupons = user
        limit = request.args.get('limit')
        es = CouponEmail(coupons)
        tempF = []
        temp_col = []
        indexesF = []
        covid_ListF = []
        emails = es

        for email in emails:
            if len(covid_ListF) > 1000:
                     break
            print(email)
            covidList = CouponLister(email, coupons)
            myquery = {"email": email}
            x = mycol.find_one(myquery)
            try:
                college_name = x["collegeName"]
            except:
                college_name = "Not Available"

            try:
                progress_list, json_data = progress(hash_fetch(email))
            except:
                continue

            for i in covidList:
                try:
                    j = c.index(a[0])
                    # indexes.append(progress_list[j])
                    indexesF.append(progress_list[j])
                except:
                    indexesF.append("0")

            for i in covidList:
                covid_ListF.append(i)
            for i in range(len(covidList)):
                # temp.append(email)
                tempF.append(email)
                temp_col.append(college_name)

        print(len(tempF), len(indexesF), len(covid_ListF))
        data = {'email': tempF, 'courses': covid_ListF, 'progress': indexesF, "college_name": temp_col}
        df = pd.DataFrame(data)
        df.to_excel("output.xlsx")
        file_name = 'output.xlsx'

        return send_file(file_name, attachment_filename='output.xls', as_attachment=True)



if __name__ == "__main__":
        app.run(port=5559, host="0.0.0.0", debug=True)
