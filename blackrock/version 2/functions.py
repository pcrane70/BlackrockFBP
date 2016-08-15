#!/usr/bin/env python
# -*-coding=utf-8 -*-

# Name: BlackRock Facebook Profiler v2.0
# Author: Muiz Yusuff (https://www.linkedin.com/in/muiz-yusuff-148572b8)

import sys, os, xlrd, re
import pandas as pd

def get_excel_files( directory ):
	# name = name of csv file

	err = ['ERROR']

	# directory = input(' [ > ] Enter the name of the directory your excel files are located in. Referense instructions.txt for more info \n > ')

	all_files = os.listdir(directory)
	excel_files = []

	xlrd_patt = r'^.+(\.){1}(xlsx|xlsm|xlsb|xltm|xls){1}$'

	# Chec kto make sure it is an excel file.
	for file in all_files:
		if re.match(xlrd_patt, file):
			file = {
				'name':	file,
			}
			excel_files.append(file)

	# Make sure there are 2 excel files
	if len(excel_files) != 2:
		if len(excel_files) < 2:
			msg = ' [ ! ] ERR: There are less than 2 excel files in this directory. There needs to be exactly 2 excel files \n One excel file with the questions and another without the questions.'
			err.append(msg)
		elif len(excel_files) > 2:
			msg = ' [ ! ] ERR: There are more than 2 excel files in this directory. There needs to be exactly 2 excel files \n One excel file with the questions and another without the questions.'
			err.append(msg)

	# Check which excel file has the questions
	file_identified = False
	while not file_identified:

		# Ask which file contains the questions
		question = int(input(' \n [ ? ] Which file contains the employee\'s responses to the questions? \n Type 1 for [ {0} ] and 2 for [ {1} ]. \n [ > ] '.format(excel_files[0]['name'], excel_files[1]['name'])))

		# Check if questions is a valid response:
		if question == 1 or question == 2:
			# Valid response given
			# Add label to dictionary
			if question == 1:
				excel_files[0]['type'] = 'questions'
				excel_files[1]['type'] = 'global'
			else:
				excel_files[0]['type'] = 'global'
				excel_files[1]['type'] = 'questions'

			file_identified = True
			break

		else:
			# Invalid response given
			print(' [ ! ] Invalid response given. Enter 1 for [ {0} ] and 2 for [ {1} ] \n [ > ] '.format(excel_files[0]['name'], excel_files[1]['name']))

	# Check if any errors
	if len(err) > 1:
		return err
	else:
		return excel_files

def xl_to_df(directory, file_dict):

	# Get excel file

	file_path = ''

	for file in file_dict:
		if not file['type'] == 'questions':
			file_path = str(directory) + '\\' + file['name']

		else:
			file_path2 = str(directory) + '\\' + file['name']

	main1 = pd.read_excel(file_path, pd.ExcelFile(file_path).sheet_names[0], encoding='utf-8')
	main2 = pd.read_excel(file_path2, pd.ExcelFile(file_path2).sheet_names[0], encoding='utf-8')

	# xls_file = pd.ExcelFile(file_path)
	# main1 = xls_file.parse( xls_file.sheet_names[0] )

	# xls_file2 = pd.ExcelFile(file_path2)
	# main2 = xls_file2.parse( xls_file2.sheet_names[0] )

	full_df = main1.append(main2)

	return full_df