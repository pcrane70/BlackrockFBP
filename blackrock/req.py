# Preliminary Functions

import xlrd, sys, os, re
from pptx import Presentation

# Gets name of excel file from directory
def getExcelNames(directory):
	files = os.listdir(directory)
	pattern = r'(xls|xlsx|xlw){1}'

	file_name = []

	for file in files:
		if re.search(pattern, file):
			add_file = {
				'name': str(file),
			}
			file_name.append(add_file)

	for file in file_name:
		# Prompt the files that have ben selected
		print('\n \n \n ')
		print(str(file_name[0]['name']) + ' and ' + str(file_name[1]['name']) + ' have been selected.')
		print('Please identify which excel sheet contains the questions by typing 1, for the first \n file, and 2 for the second file \n')
		question_index = str(input('Answer: '))

		while True:
			if question_index == 1 or question_index == '1':
				file_name[0]['type'] = 'question'
				file_name[1]['type'] = 'non-question'
				break

			elif question_index == 2 or question_index == '2':
				file_name[0]['type'] = 'non-question'
				file_name[1]['type'] = 'question'
				break

			else:
				print('\n \n \n')
				print('Invalid answer. Please enter 1 for ' + str(file_name[0]['name']) + ' or 2 for ' + str(file_name[1]['name']) + '. \n')
				question_index = input('Answer: ')

	return(file_name)

# Get contents from excel file
def getExcelContents(excel_file_name, question_file_name):
	
	# Step 1:  Get data from excel file name

	workbook = xlrd.open_workbook(excel_file_name, on_demand=True, encoding_override="cp1252")
	worksheet = workbook.sheet_by_index(0)

	workbook2 = xlrd.open_workbook(question_file_name, on_demand=True, encoding_override="cp1252")
	worksheet2 = workbook2.sheet_by_index(0)

	# Empty List
	info = []

	for i in range(1, 1500):
		if worksheet.cell(i, 0).value == '' or worksheet.cell(i, 0).value == None:
			break

		else:
			data = {
				'first_name': worksheet.cell_value(i, 0),
				'last_name': worksheet.cell_value(i, 1),
				'organization': worksheet.cell_value(i, 3),
				'location': worksheet.cell_value(i, 29),
				'university': worksheet.cell_value(i, 19),
			}

			# Check if the name is equal to the name on the question
			for x in range(1, 376):
				if data['first_name'] == worksheet2.cell(x, 2).value and data['last_name'] == worksheet2.cell(x, 1).value:
					

					interests = worksheet2.cell_value(x, 22).encode("utf-8")

					dinner = worksheet2.cell_value(x, 23).encode("utf-8")
					
					film = worksheet2.cell_value(x, 24).encode("utf-8")

					# Add answers to question
					data['questions'] = {
						'interest': str(interests),
						'dinner': str(dinner),
						'film': str(film),
					}

			# Add the data to info list
			info.append(data)

	# Step 2: Get data from question file name

	return info

# Get image contents
def getImageContents(image_directory, excel_cont):

	# Excel contents should return a list of dictionaries
	for person in excel_cont:
		# Each dictionary

		name = person['last_name'] + ', ' + person['last_name']
		
		photo_dir = os.listdir(image_directory)
		profiles = []

		pattern = r'(\.){1}(jpeg|jpg|gif|jfif|tiff|png){1}$'

		# Check if it is an image
		for file in photo_dir:
			if re.search(pattern, file):

				# Add to profiles list
				profiles.append(file)

	return profiles


def combineImageExcel(image_list, excel_list):

	# Make pattern to get name
	lastpatt = r'(?P<last>[a-zA-Z]+)(\,){1}(\s)?(?P<first>[a-zA-Z]+)(\s)?(\.){1}(?P<ext>[a-zA-Z]+)'
	firstpatt = r'(?P<first>[a-zA-Z]+)(\s)+(?P<last>[a-zA-Z]+)(\.){1}(?P<ext>[a-zA-Z]+)'

	for img in image_list:

		if re.match(lastpatt, img):
			last = re.match(lastpatt, img).group('last')
			first = re.match(lastpatt, img).group('first')
			ext = re.match(lastpatt, img).group('ext')
		
		elif re.match(firstpatt, img):
			last = re.match(firstpatt, img).group('last')
			first = re.match(firstpatt, img).group('first')
			ext = re.match(firstpatt, img).group('ext')

		# Check if image name matches that in excel_list
		for person in excel_list:

			first_name = person['first_name']
			last_name = person['last_name']

			if first_name == first and last_name == last:

				person['image'] = str(img)

	return excel_list
