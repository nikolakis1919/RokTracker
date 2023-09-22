from ppadb.client import Client
import cv2
import sys
import os
from PIL import Image
import pytesseract
import numpy as np
import time
from matplotlib import pyplot as plt
import xlwt 
from xlwt import Workbook
from datetime import date
import tkinter as tk
from tkinter import messagebox
import keyboard
from neural_network import read_ocr
import requests
import webbrowser
import traceback

version = "RokTracker-v8.2"
def tointcheck(element):
	try:
		return int(element)
	except ValueError:
		return element
		
def tointprint(element):
	try:
		return str(f'{int(element):,}')
	except ValueError:
		return str(element)

#Initiliaze paths and variables
today = date.today()

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' #Change to your installation path folder.

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)
os.system("")

##### Update Checker #####
response = requests.get("https://api.github.com/repos/nikolakis1919/RokTracker/releases/latest")
if (response.json()["name"]) != version:
	bo = tk.Tk()
	bo.withdraw()
	messagebox.showinfo("Tool is outdated", "New version is available on github repository. It is highly recommended to update the tool!")
	bo.destroy()


####### Tkinter Section ##########
#Create input gui
root=tk.Tk()

#Tkinter title
root.title('RokTracker')

#Tkinter window size
root.geometry("400x250")

#Tkinter function
def link():
	webbrowser.open_new(r"https://www.paypal.com/donate/?hosted_button_id=55G95MNYPVX72")

#Initialize Options for dropdown box
OPTIONS = []
for i in range(38):
	OPTIONS.append(50+i*25)
	
#Variables
variable = tk.StringVar(root)
variable.set('')
variable2 = tk.IntVar(root)
variable2.set(OPTIONS[0]) # default value
var1 = tk.IntVar()

#Labels
kingdom_label = tk.Label(root, text = 'Kingdom', font=('calibre',10, 'bold'))  
search_top_label = tk.Label(root, text = 'Search Amount', font=('calibre',10, 'bold'))
#Copyrights
copyright=u"\u00A9"
l1=tk.Label(root,text=copyright + ' nikolakis1919', font = ('calibre',10,'bold')) 
l2=tk.Button(root, foreground='Green', text='Donate', command=link, font = ('calibre',10,'bold'))
l3=tk.Label(root,text='Find me on discord: nikos#4469', font = ('calibre',10,'bold'))

#Input Fields
kingdom_entry = tk.Entry(root,textvariable = variable, font=('calibre',10,'normal'))
w = tk.OptionMenu(root, variable2, *OPTIONS)
resume_scan =tk.Checkbutton(root, text="Resume Scan", variable=var1, font=('calibre',10,'bold'))

def search():
	if variable.get():
		global kingdom
		kingdom = variable.get()
		global search_range
		search_range = variable2.get()
		root.destroy()
		global resume_scanning
		resume_scanning = var1.get()
		print("Scanning Started...")
	else:
		print("You need to fill Kingdom number!")
		kingdom_entry.focus_set()
		
button = tk.Button(root, text="Search", command=search)

#Positions in tkinter Grid
kingdom_label.grid(row=0,column=0)
kingdom_entry.grid(row=0,column=1)
search_top_label.grid(row=1,column=0)
w.grid(row=1,column=1)
resume_scan.grid(row=2,column=1,pady=4)
button.grid(row=3,column=1,pady=5)
l1.grid(row=4,column=1,pady=10)
l2.grid(row=8,column=1,pady=10)
l3.grid(row=5,column=1,pady=10)
root.mainloop()

#######RokTracker
#Initialize the connection to adb
adb = Client(host='localhost', port=5037)
devices = adb.devices()

if len(devices) == 0:
    print('no device attached')
    quit()

#Prolly a good idea to have only 1 device while running this
device = devices[0]

######Excel Formatting
wb = Workbook()
sheet1 = wb.add_sheet(str(today))

#Make Head Bold
style = xlwt.XFStyle()
font = xlwt.Font()
font.bold = True
style.font = font

#Initialize Excel Sheet Header
sheet1.write(0, 0, 'Governor Name', style)
sheet1.write(0, 1, 'Governor ID', style)
sheet1.write(0, 2, 'Power', style)
sheet1.write(0, 3, 'Kill Points', style)
sheet1.write(0, 4, 'Deads', style)
sheet1.write(0, 5, 'Tier 1 Kills', style)
sheet1.write(0, 6, 'Tier 2 Kills', style)
sheet1.write(0, 7, 'Tier 3 Kills', style)
sheet1.write(0, 8, 'Tier 4 Kills', style)
sheet1.write(0, 9, 'Tier 5 Kills', style)
sheet1.write(0, 10,'Rss Assistance', style)
sheet1.write(0, 11,'Alliance', style)

#Position for next governor to check
Y =[285, 390, 490, 590, 605]

#Resume Scan options. Refine the loop
j = 0
if resume_scanning:
	j = 4
	search_range = search_range + j
#The loop in TOP XXX Governors in kingdom - It works both for power and killpoints Rankings
#MUST have the tab opened to the 1st governor(Power or Killpoints)


##### Save button listener#####
stop = False
def onkeypress(event):
	global stop
	if event.name == '\\':
		print("Your scan will be terminated when current governor scan is over!")
		stop = True

keyboard.on_press(onkeypress)




try:
	for i in range(j,search_range):
		if stop:
			print("Scan Terminated! Saving the current progress...")
			break
		if i>4:
			k = 4
		else:
			k = i
			
		gov_dead = 0
		gov_kills_tier1 = 0
		gov_kills_tier2 = 0
		gov_kills_tier3 = 0
		gov_kills_tier4 = 0
		gov_kills_tier5 = 0
		gov_rss_assistance = 0
		#Open governor
		device.shell(f'input tap 690 ' + str(Y[k]))
		time.sleep(1.8)
		
		##### Ensure that governor tab is open #####
		gov_info = False
		count = 0
		while not (gov_info):
			image_check = device.screencap()
			with open(('check_more_info.png'), 'wb') as f:
						f.write(image_check)
			image_check = cv2.imread('check_more_info.png',cv2.IMREAD_GRAYSCALE)
			roi = (294, 786, 116, 29)
			im_check_more_info = image_check[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]
			check_more_info = pytesseract.image_to_string(im_check_more_info,config="-c tessedit_char_whitelist=MoreInfo")
			if 'MoreInfo' not in check_more_info :
				device.shell(f'input swipe 690 605 690 540')
				device.shell(f'input tap 690 ' + str(Y[k]))
				count += 1
				time.sleep(2)
				if count == 5:
					break
			else:
				gov_info = True
				break
		
		#nickname copy
		device.shell(f'input tap 654 245')
		time.sleep(1.3)
		
		##### Governor main page capture #####
		image = device.screencap()
		with open(('gov_info.png'), 'wb') as f:
					f.write(image)
		image = cv2.imread('gov_info.png')
		#Power and Killpoints
		roi = (733, 192, 200, 35)
		im_gov_id = image[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]
		image = cv2.imread('gov_info.png')
		roi = (874, 327, 224, 40)
		im_gov_power = image[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]
		roi = (1106, 327, 224, 40)
		im_gov_killpoints = image[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]
		gov_name = tk.Tk().clipboard_get()


		#New image for alliance tag
		image = cv2.imread('gov_info.png')
		kernel = np.ones((2, 2), np.uint8)
	 
		image = cv2.erode(image, kernel) 
		roi = (598, 331, 250, 40) #alliance tag
		im_alliance_tag = image[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]
		
		#kills tier
		device.shell(f'input tap 1118 314')
        
		#1st image OCR
		gov_id = read_ocr(im_gov_id)
		gov_killpoints2 = pytesseract.image_to_string(im_gov_killpoints,config="-c tessedit_char_whitelist=0123456789")
		gov_power2 = pytesseract.image_to_string(im_gov_power,config="-c tessedit_char_whitelist=0123456789")
		gov_power = read_ocr(im_gov_power)
		gov_killpoints = read_ocr(im_gov_killpoints)
		gov_killpoints = gov_killpoints2 if (len(''.join(str(gov_killpoints).split()))-1 > len(str(gov_killpoints))) else gov_killpoints
		gov_power = gov_power2 if (len(''.join(str(gov_power2).split()))-1 > len(str(gov_power))) else gov_power

		time.sleep(1)
		##### Kill tier Capture #####
		image = device.screencap()
		with open(('kills_tier.png'), 'wb') as f:
					f.write(image)
		image2 = cv2.imread('kills_tier.png')
		image2 = cv2.fastNlMeansDenoisingColored(image2,None,20,20,7,3)
		_,image2 = cv2.threshold(image2,180,255,cv2.THRESH_BINARY)
		roi = (862, 430, 215, 26) #tier 1
		im_kills_tier1 = image2[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]

		roi = (862, 475, 215, 26) #tier 2
		im_kills_tier2 = image2[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]

		roi = (862, 516, 215, 26) #tier 3
		im_kills_tier3 = image2[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]

		roi = (862, 561, 215, 26) #tier 4
		im_kills_tier4 = image2[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]

		roi = (862, 606, 215, 26) #tier 5
		im_kills_tier5 = image2[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]

		#More info tab
		device.shell(f'input tap 350 740') 
		
		##### Kill tier OCR #####
		gov_kills_tier1 = pytesseract.image_to_string(im_kills_tier1,config="-c tessedit_char_whitelist=0123456789")
		gov_kills_tier2 = pytesseract.image_to_string(im_kills_tier2,config="-c tessedit_char_whitelist=0123456789")
		gov_kills_tier3 = pytesseract.image_to_string(im_kills_tier3,config="-c tessedit_char_whitelist=0123456789")
		gov_kills_tier4 = pytesseract.image_to_string(im_kills_tier4,config="-c tessedit_char_whitelist=0123456789")
		gov_kills_tier5 = pytesseract.image_to_string(im_kills_tier5,config="-c tessedit_char_whitelist=0123456789")
		time.sleep(1)
		
		
		##### More Info Page Capture #####
		image = device.screencap()
		with open(('more_info.png'), 'wb') as f:
					f.write(image)
		image3 = cv2.imread('more_info.png')
		kernel = np.ones((2, 2), np.uint8)
		image3 = cv2.dilate(image3, kernel)
		roi = (1130, 443, 183, 40) #dead
		im_dead = image3[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]
		roi = (1130, 668, 183, 40) #rss assistance
		im_rss_assistance = image3[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]
		
		#2nd check for deads with more filters to avoid some errors
		roi = (1130, 443, 183, 40) #dead
		thresh = 127
		thresh_image = cv2.threshold(image3, thresh, 255, cv2.THRESH_BINARY)[1]
		im_dead2 = thresh_image[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]
		roi = (1130, 668, 183, 40) #rss assistance
		im_rss_assistance2 = thresh_image[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]
		
		#3rd check for deads with more filters to avoid some errors
		roi = (1130, 443, 183, 40) #dead
		blur_img = cv2.GaussianBlur(image3, (3, 3), 0)
		im_dead3 = blur_img[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]
		roi = (1130, 668, 183, 40) #rss assistance
		im_rss_assistance3 = blur_img[int(roi[1]):int(roi[1]+roi[3]), int(roi[0]):int(roi[0]+roi[2])]
		
		##### More info page OCR #####
		gov_dead = read_ocr(im_dead)
		gov_dead2 = read_ocr(im_dead2)
		gov_dead3 = read_ocr(im_dead3)
		gov_rss_assistance = read_ocr(im_rss_assistance)
		gov_rss_assistance2 = read_ocr(im_rss_assistance2)
		gov_rss_assistance3 = read_ocr(im_rss_assistance3)
		
		
		##### Alliance tag #####
		gray = cv2.cvtColor(im_alliance_tag,cv2.COLOR_BGR2GRAY)
		threshold_img = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
		alliance_tag = pytesseract.image_to_string(threshold_img)
		
		
		#Just to check the progress, printing in cmd the result for each governor
		if gov_power == '':
			gov_power = 'Unknown'
		if gov_killpoints =='':
			gov_killpoints = 'Unknown'
		if gov_dead == '' :
			if gov_dead2 == '':
				if gov_dead3 =='':
					gov_dead = 'Unknown'
				else:			
					gov_dead = gov_dead3
			else:
				gov_dead = gov_dead2
		if gov_kills_tier1 == '' :
			gov_kills_tier1 = '0'
		if gov_kills_tier2 == '' :
			gov_kills_tier2 = '0'
		if gov_kills_tier3 == '' :
			gov_kills_tier3 = '0'
		if gov_kills_tier4 == '' :
			gov_kills_tier4 = '0'
		if gov_kills_tier5 == '' :
			gov_kills_tier5 = '0'
		if gov_rss_assistance == '' :
			if gov_rss_assistance2 =='':
				if gov_rss_assistance3 =='':
					gov_rss_assistance = 'Unknown'
				else: 
					gov_rss_assistance = gov_rss_assistance3
			else:
				gov_rss_assistance= gov_rss_assistance2
		print('Governor ID: ' + str(gov_id) + '\nGovernor Name: ' + gov_name + '\nGovernor Power: ' + tointprint(gov_power) + '\nGovernor Killpoints: ' + tointprint(gov_killpoints) + '\nTier 1 kills: ' + tointprint(gov_kills_tier1) + '\nTier 2 kills: ' + tointprint(gov_kills_tier2) + '\nTier 3 kills: ' + tointprint(gov_kills_tier3) + '\nTier 4 kills: ' + tointprint(gov_kills_tier4) + '\nTier 5 kills: ' + tointprint(gov_kills_tier5) + '\nGovernor Dead Troops: ' + tointprint(gov_dead) + '\nGovernor RSS Assistance: ' + tointprint(gov_rss_assistance) +'\nAlliance: ' + str(alliance_tag) + '\n')
		device.shell(f'input tap 1396 58') #close more info
		time.sleep(0.5)
		device.shell(f'input tap 1365 104') #close governor info
		time.sleep(1)

		#Write results in excel file
		sheet1.write(i+1-j, 0, gov_name)
		sheet1.write(i+1-j, 1, tointcheck(gov_id))
		sheet1.write(i+1-j, 2, tointcheck(gov_power))
		sheet1.write(i+1-j, 3, tointcheck(gov_killpoints))
		sheet1.write(i+1-j, 4, tointcheck(gov_dead))
		sheet1.write(i+1-j, 5, tointcheck(gov_kills_tier1))
		sheet1.write(i+1-j, 6, tointcheck(gov_kills_tier2))
		sheet1.write(i+1-j, 7, tointcheck(gov_kills_tier3))
		sheet1.write(i+1-j, 8, tointcheck(gov_kills_tier4))
		sheet1.write(i+1-j, 9, tointcheck(gov_kills_tier5))
		sheet1.write(i+1-j, 10, tointcheck(gov_rss_assistance))
		sheet1.write(i+1-j, 11, alliance_tag)
except:
	print('An issue has occured. Please rerun the tool and use "resume scan option" from where tool stopped. If issue seems to remain, please contact me on discord!')
	#Save the excel file in the following format e.g. TOP300-2021-12-25-1253.xls or NEXT300-2021-12-25-1253.xls
	traceback.print_exc()
	pass


if resume_scanning :
	file_name_prefix = 'NEXT'
else:
	file_name_prefix = 'TOP'
wb.save(file_name_prefix + str(search_range-j) + '-' +str(today)+ '-' + kingdom +'.xls')
