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

def prog():
	today = datetime.date.today()
	workbook1 = xlrd.open_workbook('nodes.xlsx', on_demand=True)
	sheet1 = workbook1.sheet_by_index(0)
	arrayofnodelat = sheet1.col_values(0)
	arrayofnodelong = sheet1.col_values(1)
	nearlat = sheet1.col_values(2)
	nearlong = sheet1.col_values(3)
	speedval1=[]
	speedval2 = []
	speedval3 = []
	Node = []
	Time =[]
	for i in range(1,15):
		Node.append(i)
	url = 'https://maps.googleapis.com/maps/api/distancematrix/json?origins='+str(arrayofnodelat[i])+','+str(arrayofnodelong[i])+'&destinations='+str(nearlat[i])+','+str(nearlong[i])+'&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyBrSfoJB6DHg798S-ZT7iecYMQo_7IVrcY'
	res=requests.get(url)
	result1=res.json()
	print(result1)
	a = result1['rows'][0]['elements'][0]['duration_in_traffic']['value']
	time=float(a)
	a = result1['rows'][0]['elements'][0]['distance']['value']
	dist = float(a)
	speed = dist/(time + 1)
	speedval1.append(speed)
	url = 'https://maps.googleapis.com/maps/api/distancematrix/json?origins='+str(arrayofnodelat[i])+','+str(arrayofnodelong[i])+'&destinations='+str(nearlat[i])+','+str(nearlong[i])+'&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyBrSfoJB6DHg798S-ZT7iecYMQo_7IVrcY'
	res=requests.get(url)
	result1=res.json()
	#print(result1)
	a = result1['rows'][0]['elements'][0]['duration_in_traffic']['value']
	time=float(a)
	a = result1['rows'][0]['elements'][0]['distance']['value']
	dist = float(a)
	speed = dist/(time + 1)
	speedval2.append(speed)
	url = 'https://maps.googleapis.com/maps/api/distancematrix/json?origins='+str(arrayofnodelat[i])+','+str(arrayofnodelong[i])+'&destinations='+str(nearlat[i])+','+str(nearlong[i])+'&departure_time=now&traffic_model=optimistic&mode=driving&key=AIzaSyBrSfoJB6DHg798S-ZT7iecYMQo_7IVrcY'
	res=requests.get(url)
	result1=res.json()
	#print(result1)
	a = result1['rows'][0]['elements'][0]['duration_in_traffic']['value']
	time=float(a)
	a = result1['rows'][0]['elements'][0]['distance']['value']
	dist = float(a)
	speed = dist/(time + 1)
	time2 = str(datetime.datetime.now().time())
	speedval3.append(speed)
	Time.append(time2)
	for i in range(2,len(arrayofnodelong)):
		url = 'https://maps.googleapis.com/maps/api/distancematrix/json?origins='+str(arrayofnodelat[i])+','+str(arrayofnodelong[i])+'&destinations='+str(arrayofnodelat[i-1])+','+str(arrayofnodelong[i-1])+'&departure_time=now&traffic_model=pessimistic&mode=driving&key=AIzaSyBrSfoJB6DHg798S-ZT7iecYMQo_7IVrcY'
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
		url = 'https://maps.googleapis.com/maps/api/distancematrix/json?origins='+str(arrayofnodelat[i])+','+str(arrayofnodelong[i])+'&destinations='+str(arrayofnodelat[i-1])+','+str(arrayofnodelong[i-1])+'&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyBrSfoJB6DHg798S-ZT7iecYMQo_7IVrcY'
		res=requests.get(url)
		result1=res.json()
		#print(result1)
		a = result1['rows'][0]['elements'][0]['duration_in_traffic']['value']
		time=float(a)
		a = result1['rows'][0]['elements'][0]['distance']['value']
		dist = float(a)
		speed = dist/(time + 1)
		speedval2.append(speed)
		url = 'https://maps.googleapis.com/maps/api/distancematrix/json?origins='+str(arrayofnodelat[i])+','+str(arrayofnodelong[i])+'&destinations='+str(arrayofnodelat[i-1])+','+str(arrayofnodelong[i-1])+'&departure_time=now&traffic_model=optimistic&mode=driving&key=AIzaSyBrSfoJB6DHg798S-ZT7iecYMQo_7IVrcY'
		res=requests.get(url)
		result1=res.json()
		#print(result1)
		a = result1['rows'][0]['elements'][0]['duration_in_traffic']['value']
		time=float(a)
		a = result1['rows'][0]['elements'][0]['distance']['value']
		dist = float(a)
		speed = dist/(time + 1)
		time2 = str(datetime.datetime.now().time())
		speedval3.append(speed)
		Time.append(time2)
		print("current time: ", time2)
		#print("time: ", time)
		#print("distance " ,dist)
	fo = open('value2.txt','r')
	#print(speedval)
	#if(int(fo.read())==0):
	if(~os.path.isfile('419_googleapi'+str(today)+'.xlsx')):
		fo.close()
		df = pd.DataFrame({'Node':Node,'Best Guess Speed':speedval1,'Optimistic Speed':speedval2,'Pessimistic Speed':speedval3, 'Time':Time})
		book_ro = pd.ExcelWriter('419_googleapi'+str(today)+'.xlsx', engine='xlsxwriter')
		df.to_excel(book_ro, sheet_name='Sheet1')
		book_ro.save()
#for row, entry in enumerate(data1,start=1):     
			#book_ro = open_workbook('position_eve'+str(i)+'.xlsx', on_demand=True)
			#book = copy(book_ro)  # creates a writeable copy
			
		fo = open('value2.txt','w')
		fo.write('1')
		fo.close()
	else:
		fo.close()
		book_ro = openpyxl.load_workbook('419_googleapi'+str(today)+'.xlsx')
		#book = copy(book_ro)  # creates a writeable copy
		sheet1 = book_ro.active # get a first sheet
		#max1 = sheet1.max_row
		for i in range(0,14):
			sheet1.append({3:Node[i],2:speedval1[i],4:speedval2[i],5:speedval3[i],6:Time[i]})
		#sheet1.write(row=max1+1, column=0, value=result['response'][i]['latitude'])
		#sheet1.write(row=max1+1, column=1, value=result['response'][i]['longitude'])
		#sheet1.write(row=max1+1, column=2, value=result['response'][i]['speed'])
		book_ro.save('419_googleapi'+str(today)+'.xlsx')

		
	#fo = open('bus1.txt', 'a')
	#fo.writelines(result['response'][11]['speed'])
	#fo.writelines('\n')
	#fo.close()

	#fo = open('bus2.txt', 'a')
	#fo.writelines(result['response'][12]['speed'])
	#fo.writelines('\n')
	#fo.close()

	#fo = open('bus3.txt', 'a')
	#fo.writelines(result['response'][13]['speed'])
	#fo.writelines('\n')
	#fo.close()

	#list1=[]
	#urls = {'a': 'https://maps.googleapis.com/maps/api/distancematrix/json?origins=28.598415, 77.063717&destinations=28.598683, 77.063955&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyDIvYBsEta5VXQXUrNfoJ2h5h_dD0cJDNk', 
	#		'b': 'https://maps.googleapis.com/maps/api/distancematrix/json?origi2ns=28.598683, 77.063955&destinations=28.598936, 77.064261&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyDIvYBsEta5VXQXUrNfoJ2h5h_dD0cJDNk',
	#		'c': 'https://maps.googleapis.com/maps/api/distancematrix/json?origins=28.598936, 77.064261&destinations=28.599932, 77.064952&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyDIvYBsEta5VXQXUrNfoJ2h5h_dD0cJDNk',
	#		'd': 'https://maps.googleapis.com/maps/api/distancematrix/json?origins=28.599932, 77.064952&destinations=28.598936, 77.064261&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyDIvYBsEta5VXQXUrNfoJ2h5h_dD0cJDNk',
	#		'e': 'https://maps.googleapis.com/maps/api/distancematrix/json?origins=28.598936, 77.064261&destinations=28.599990, 77.065165&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyDIvYBsEta5VXQXUrNfoJ2h5h_dD0cJDNk',
	#		'f': 'https://maps.googleapis.com/maps/api/distancematrix/json?origins=28.599990, 77.065165&destinations=28.600338, 77.065884&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyDIvYBsEta5VXQXUrNfoJ2h5h_dD0cJDNk',
	#		'g': 'https://maps.googleapis.com/maps/api/distancematrix/json?origins=28.600338, 77.065884&destinations=28.601007, 77.066999&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyDIvYBsEta5VXQXUrNfoJ2h5h_dD0cJDNk',
	#		'h': 'https://maps.googleapis.com/maps/api/distancematrix/json?origins=28.601007, 77.066999&destinations=28.601450, 77.067772&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyDIvYBsEta5VXQXUrNfoJ2h5h_dD0cJDNk',
	#		'i': 'https://maps.googleapis.com/maps/api/distancematrix/json?origins=28.601450, 77.067772&destinations=28.602288, 77.068577&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyDIvYBsEta5VXQXUrNfoJ2h5h_dD0cJDNk',
	#		'j': 'https://maps.googleapis.com/maps/api/distancematrix/json?origins=28.602288, 77.068577&destinations=28.6030218, 77.0700495&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyDIvYBsEta5VXQXUrNfoJ2h5h_dD0cJDNk'
	#		 }
	#for who in urls.keys():
	   #url = urlopen(urls[who])
	#   url = requests.get(urls[who])
	#   result = url.json()  # result is now a dict
	#   #print(result)
	#   list1.append(result['rows'][0]['elements'][0]['duration_in_traffic']['text'])
	#list2=[]
	#for i in list1:
	#	a,b=i.split(mnbg)
	#	a=int(a)
	#	list2.append(a)
	#avg=sum(list2)/10.0
	#r_earth=6378
	#pi=3.14
	#l = [11,12,13]
	#k=1
	


	# Convert the dataframe to an XlsxWriter Excel object.
		#df.to_excel(writer, sheet_name='Sheet1')
	
	# Close the Pandas Excel writer and output the Excel file.
		#writer.save()
		#fo = open('busmorn'+str(k)+'.txt', 'a')
		#fo.writelines(result['response'][i]['speed'])
		#fo.writelines('\n')
		#fo.close()
		#fo = open('latmorn'+str(k)+'.txt', 'a')
		#fo.writelines(result['response'][i]['latitude'])
		#fo.writelines('\n')
		#fo.close()
		#fo = open('longmorn'+str(k)+'.txt', 'a')
		#fo.writelines(result['response'][i]['longitude'])
		#fo.writelines('\n')
		#fo.close()
		#k=k+1
		#new_latitude2 = float(result['response'][i]['latitude'])  + (0.25 / r_earth) * (180 / pi);
		#new_longitude2 = float(result['response'][i]['longitude']) + (0.25 / r_earth) * (180 / pi) / math.cos(float(result['response'][i]['latitude']) * pi/180);
		#new_latitude1 = float(result['response'][i]['latitude'])  - (0.25 / r_earth) * (180 / pi);
		#new_longitude1 = float(result['response'][i]['longitude']) - (0.25 / r_earth) * (180 / pi) / math.cos(float(result['response'][i]['latitude']) * pi/180);
		#url=requests.get('https://roads.googleapis.com/v1/nearestRoads?points='+str(new_latitude1)+','+str(new_longitude1)+'|'+str(new_latitude2)+','+str(new_longitude2)+'&key=AIzaSyDrkMiU2F4LmAIuT-eX4epqbzhl7NMv_4U')
		#result1 = url.json()
		#new_latitude1 = result1['snappedPoints'][0]['location']['latitude']
		#new_longitude1 = result1['snappedPoints'][0]['location']['longitude']
		#new_latitude2 = result1['snappedPoints'][1]['location']['latitude']
		#new_longitude2 = result1['snappedPoints'][1]['location']['longitude']
		#url = requests.get('https://maps.googleapis.com/maps/api/distancematrix/json?origins='+str(new_latitude1)+','+str(new_longitude1)+'&destinations='+str(new_latitude2)+','+str(new_longitude2)+'&departure_time=now&traffic_model=best_guess&mode=driving&key=AIzaSyDrkMiU2F4LmAIuT-eX4epqbzhl7NMv_4U')
		#result1 = url.json()
		#fo = open('timevalues'+str(k)+'.txt', 'a')
		#avg=str(avg)
		#a,b=result1['rows'][0]['elements'][0]['duration_in_traffic']['text'].split()
		#time=float(a)
		#a,b=result1['rows'][0]['elements'][0]['distance']['text'].split()
		#dist=float(a)
		#speed = str(dist/time)
		#fo.writelines(speed)
		#fo.writelines('\n')
		#fo.close()
		#fo = open('lat'+str(k)+'.txt', 'a')
		#fo.writelines(result['response'][i]['latitude'])
		#fo.writelines('\n')
		#fo.close()
		#fo = open('long'+str(k)+'.txt', 'a')
		#fo.writelines(result['response'][i]['longitude'])
		#fo.writelines('\n')
		#fo.close()
		#k=k+1

#prog()
scheduler = BlockingScheduler()
scheduler.add_job(prog, 'cron', hour='*', minute='15', second='*')
scheduler.start()
