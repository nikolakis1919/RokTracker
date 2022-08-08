import cv2
from glob import glob
import numpy as np
import pickle
from sklearn.preprocessing import normalize
from sklearn.neural_network import MLPClassifier
import sys
from collections import Counter
import os
import csv
from difflib import SequenceMatcher
from time import gmtime, strftime

def read_ocr(image):

	show_img = False
	#list for wrongly classified 
	img_list = []

	#takes an image and returns array with all digits pixels in it
	def digits_read(im, show_img=False, check=False):
		#copy becuase will be edited
		img = im.copy()

		#make an output, convert to grayscale and apply thresh-hold
		out = np.zeros(im.shape,np.uint8)
		gray = cv2.cvtColor(im,cv2.COLOR_BGR2GRAY)
		thresh = cv2.adaptiveThreshold(gray,255,1,1,11,2)

		#find contours
		contours,hierarchy = cv2.findContours(thresh,cv2.RETR_LIST,cv2.CHAIN_APPROX_SIMPLE)
		#create empty return list
		samples =  np.empty((0,100))

		#for every contour if area large enoug to be digit add the box to list
		li = []
		for cnt in contours:
			if cv2.contourArea(cnt)>20:
				[x,y,w,h] = cv2.boundingRect(cnt)
				li.append([x,y,w,h])
		#sort list so it read from right to left
		li = sorted(li,key=lambda x: x[0], reverse=True)
		#loop over all digits
		for i in li:
			#unpack data
			x,y,w,h = i[0], i[1], i[2], i[3]

			#check if large enough to be digit but small enough to ignore rest
			if  h>15 and h<30 and w<40 and w>7:

				#draw rectangle with thresh-hold and shape correct form
				cv2.rectangle(im,(x,y),(x+w,y+h),(0,255,0),2)
				roi = thresh[y:y+h,x:x+w]
				roismall = cv2.resize(roi,(10,10))
				sample = roismall.reshape((1,100))
				samples = np.append(samples,sample,0)
		#return all digits found
		return samples

	def similar(a, b):
		return SequenceMatcher(None, a, b).ratio()

	#get list of found digits and runs it through NN
	def classify(data):
		clas = []
		#run every f ound digit through NN
		if isinstance(data, bool):
			return 0
		if isinstance(data, int):
			return data
		for i in data:
			a = int(digits_model.predict([i])[0])
			if a == 11:
				a = 44
			clas.append(a)

		#reverse list and add all together in 1 integer to find final power
		clas.reverse()
		clas = map(str, clas)
		clas = ''.join(clas)
		return clas

	try:
		digits_model = pickle.load(open('digits_model.sav', 'rb'))
	except:
		print("No models found")
		sys.exit()
	return int(classify(digits_read(image)))
