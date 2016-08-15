#!/usr/bin/env python
# -*-coding=utf-8 -*-

# Name: BlackRock Facebook Profiler v2.0
# Author: Muiz Yusuff (https://www.linkedin.com/in/muiz-yusuff-148572b8)

from functions import *
import os, sys, pandas as pd

class BlackrockProfiler:

	def __init__(self):

		# Get directory
		self.directory = input(' [ > ] Enter the name of the directory your excel files are located in. Referense instructions.txt for more info \n > ')

	def create_df(self):

		# Get excel filenames

		get_excel_results = get_excel_files( self.directory )
		if not 'ERROR' in get_excel_results:

			print(' - Creating dataframe file... \n') # Status Update
			# Create CSV

			self.excel_files = get_excel_results
			# Create csv file
			self.final_df = xl_to_df(self.directory, self.excel_files)
			print(' Dataframe file created. ')

		else: 
			for err in get_excel_files(self.directory):
				print(err + '\n')

			print(' Please retry. ')

	def create_product(self):
		# Call create_df()
		self.create_df()

		# Take dataframe and manipulate
		print(self.final_df)


newProfile = BlackrockProfiler()
newProfile.create_product()