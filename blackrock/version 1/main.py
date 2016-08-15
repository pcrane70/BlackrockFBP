#!/usr/bin/python
#coding: utf-8
import os, sys

# Blackrock Facebook Profiler
# by Muiz Yusuff
# Version 1.0

from req import *
from PIL import Image
from unidecode import unidecode
import datetime

class FacebookProfiler:

	def __init__(self, excel_directory):
		# Create blackrock facebook
		# for new hires.

		# excel directory is where the excel file is located and 
		# picture_directpry is where the pictures are located

		self.excel_dir = excel_directory
		self.picture_dir = excel_directory + '/Photos'
		self.excel_dict = getExcelNames(self.excel_dir)

		
	# Step 1: Get excel list
	def get_from_excel(self):
		#excel_file = self.excel_dir + '/' + 

		#getExcelNames(self.excel_dir)

		for name in self.excel_dict:
			if name['type'] != 'question':
				excel_file = self.excel_dir + '/' + name['name']
			else:
				question_file = self.excel_dir + '/' + name['name']

		# List of Dictionaries with first names and last names
		excel_names = getExcelContents(excel_file, question_file)
		image_names = getImageContents(self.picture_dir, excel_names)

		# Return getImageExcel
		self.complete_list = combineImageExcel(image_names, excel_names)

		# Encode all elements in list to utf-8


		return self.complete_list

	def list_to_pptx(self):
		
		# Call prev. function
		self.get_from_excel()

		self.updated_list = imgToThumb(self.picture_dir, self.complete_list)

		print('Making presentaiton ... ')

		self.organization = makePres(self.updated_list, self.excel_dir)

		print(self.organization)
		return 'Process Complete \n Your presentations can be located in '



excel = 'C:/users/muizyusuff/desktop'
testFb = FacebookProfiler(excel)

# exec
print(testFb.list_to_pptx())

# Version 1.0
# If the correct excel file has been identified as containing questions, the program should run smoothly.