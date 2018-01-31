#!/usr/bin/env python
import json
from urllib.request import urlopen
import requests
import xlrd
import math
from apscheduler.schedulers.blocking import BlockingScheduler
from xlutils.copy import copy    
from xlrd import open_workbook
import pandas as pd
import openpyxl
import datetime
import os
import pathlib

def prog():
	today = datetime.date.today()
	# print("jfvn");
	workbook1 = xlrd.open_workbook('Stoppage_425CLDown.xlsx', on_demand=True)
	sheet1 = workbook1.sheet_by_index(0)
	arrayofnodelat = sheet1.col_values(3)
	arrayofnodelong = sheet1.col_values(4)
	# nearlat = sheet1.col_values(2)
	# nearlong = sheet1.col_values(3)
	speedval1=[]
	Node = []
	Time =[]
	for i in range(1,42):
		Node.append(i)
	url = 'https://maps.googleapis.com/maps/api/distancematrix/json?origins='+str(arrayofnodelat[1])+','+str(arrayofnodelong[1])+'&destinations='+str(arrayofnodelat[1]+0.0011)+','+str(arrayofnodelong[1]+0.0011)+'&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyD6dV47_SQcIPLLkuDiTW1WRI3LjIgqHXE'
	res=requests.get(url)
	result1=res.json()
	# print(result1)
	a = result1['rows'][0]['elements'][0]['duration_in_traffic']['value']
	time=float(a)
	a = result1['rows'][0]['elements'][0]['distance']['value']
	dist = float(a)
	speed = dist/(time + 1)
	speedval1.append(speed)
	
	time2 = str(datetime.datetime.now().time())
	# speedval3.append(speed)
	Time.append(time2)
	for i in range(2,len(arrayofnodelong)):
		url = 'https://maps.googleapis.com/maps/api/distancematrix/json?origins='+str(arrayofnodelat[i])+','+str(arrayofnodelong[i])+'&destinations='+str(arrayofnodelat[i-1])+','+str(arrayofnodelong[i-1])+'&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyD6dV47_SQcIPLLkuDiTW1WRI3LjIgqHXE'
		res=requests.get(url)
		result1=res.json()
		#print(result1)
		a = result1['rows'][0]['elements'][0]['duration_in_traffic']['value']
		time=float(a)
		a = result1['rows'][0]['elements'][0]['distance']['value']
		dist = float(a)
		speed = dist/(time + 1)
		#time2 = str(datetime.datetime.now().time())
		speedval1.append(speed)
		
		time2 = str(datetime.datetime.now().time())
		# speedval3.append(speed)
		Time.append(time2)
	path = pathlib.Path('425_googleapi_down'+str(today)+'.xlsx')
	if(path.is_file()==False):
		#fo.close()
		#print("hello2")
		df = pd.DataFrame({'Node':Node,'Best Guess Speed':speedval1, 'Time':Time})
		book_ro = pd.ExcelWriter('425_googleapi_down'+str(today)+'.xlsx', engine='xlsxwriter')
		df.to_excel(book_ro, sheet_name='Sheet1')
		book_ro.save()

	else:
		#fo.close()
		#print("hello")
		book_ro = openpyxl.load_workbook('425_googleapi_down'+str(today)+'.xlsx')
		#book = copy(book_ro)  # creates a writeable copy
		sheet1 = book_ro.active # get a first sheet
		#max1 = sheet1.max_row
		for i in range(0,41):
			sheet1.append({3:Node[i],2:speedval1[i],4:Time[i]})

		book_ro.save('425_googleapi_down'+str(today)+'.xlsx')
		
scheduler = BlockingScheduler()
scheduler.add_job(prog, 'cron', hour='*', minute='0')
scheduler.start()
