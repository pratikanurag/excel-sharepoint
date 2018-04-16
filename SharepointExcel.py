import sharepy as sharepy
import json
import requests
from requests.auth import HTTPBasicAuth
import os
import pandas as pd
from os import listdir
from os.path import isfile, join
import argparse
import urllib.parse
import datetime
import math  

# (1) Authenticate
# #The below code Authenticate the user by passing the user email id and password
#replace example.sharepoint.com with your sharepoint website followed by user name and password

s = sharepy.connect("https://example.sharepoint.com",\
username='username@domain.com', password='password')

def create_dir(dirPath):
    if not os.path.exists(dirPath):
        os.makedirs(dirPath)

d=datetime.date.today()
#d=d+timedelta(days=23)
c_year=str(d.year)
#print('Year :'+c_year)
Path='D:\new\docs\\'+c_year #the documents will be downloaded to the path described here, for instance D:\new\docs
create_dir(Path)
Q=math.ceil(d.month/3.)
c_qtr='Q '+str(Q)
Path=Path+'\\'+c_qtr
create_dir(Path)
#print('Quarter : Quarter'+c_qtr)
c_month=str(d.strftime("%B"))
#print('Month :'+c_month)
Path=Path+'\\'+c_month
create_dir(Path)
def current_week(date):
    weeks=["","Week1","Week2","Week3","Week4","Week5"]
    dayNumber=math.ceil(d.day/7.)
    return weeks[dayNumber]
c_week=current_week(d)
#print('Week : '+c_week)
Path=Path+'\\'+c_week
create_dir(Path)

def download_file(fname):
    fname1=fname[fname.find('ts_'):len(fname)]
    #fname=fname.strip('/')
    #print('fname1 : '+fname1)
    #print('fname : '+fname)
    link="https://example.sharepoint.com/sites/folder/_api/web/GetFileByServerRelativeUrl('"+fname+"')/$value" #modify the site name and folder directory structure in appropriate way
    #print('link : '+link)
    r = s.getfile(link, filename = Path+'\\' + fname1)
    #print(fname1[fname1.find('ts_'):len(fname1)])
    #print(r.status_code)
      

r = s.get("https://example.sharepoint.com/sites/folder/_api/web\
/GetFolderByServerRelativeUrl('path to the excel files')/Files")#replace path to the excelfiles with the actual site path
data = r.json()
file = open("downloadFolder_Json.json", "w")
file.write(json.dumps(data, indent=4))
print("json file has been generated")
#print(data)
for item in data['d']['results']:
    print(item['ServerRelativeUrl'])
    download_file(item['ServerRelativeUrl'])

print("Files Downloaded to "+Path)
os.remove(Path +"\\x")  