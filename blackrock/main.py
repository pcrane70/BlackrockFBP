
# Blackrock Facebook Profiler
# by Muiz Yusuff
# Version 1.0

from req import *
import os
import sys
from PIL import Image


class FacebookProfiler:

	def __init__(self, excel_directory, picture_directory):
		# Create blackrock facebook
		# for new hires.

		# excel directory is where the excel file is located and 
		# picture_directpry is where the pictures are located

		self.excel_dir = excel_directory
		self.picture_dir = picture_directory

		
	# Step 1: Get excel list
	def get_from_excel(self):
		#excel_file = self.excel_dir + '/' + 

		#getExcelNames(self.excel_dir)
		for name in getExcelNames(self.excel_dir):
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
		# Create powerpoint
		pass



excel = 'C:/users/basir/desktop'
photos = 'C:/Users/basir/Desktop/python/blackrock/photos'
testFb = FacebookProfiler(excel, photos)

# print(testFb.get_from_excel())

with open('final-object.txt', 'w') as f:
	f.write(str(testFb.get_from_excel()))

print(testFb.get_from_excel())

# Version 1.0
# If the correct excel file has been identified as containing questions, the program should run smoothly.