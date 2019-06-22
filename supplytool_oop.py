__author__ = 'Mark Mon Monteros'

import os
import pandas as pd
from xlrd import XLRDError
from tkinter import *
from tkinter import messagebox

class SupplyTool():

	def __init__(self):
		self.path = os.path.dirname(os.path.realpath(__file__))
		self.xlsLs = []
		self.csvLs = []
		self.supplyCSV = None
		self.hcCSV = None
		self.forCSV = None

		self.convert()
		self.del_col()
		self.filter_rows()
		self.vlookup()
		self.save_output()

	def convert(self):
		#Renaming file extension case
		for f in os.listdir(self.path):
			if f.endswith('.XLS') or f.endswith('.XLSX'):
				os.rename(f, f.replace(f[f.r('.'):], f[f.rindex('.'):].lower()))

		print('Converting excel files to CSV UTF-8 format...')
		for f in os.listdir(self.path):
			if f.endswith('.xls') or f.endswith('.xlsx'):
				self.xlsLs.append(f)
		for sheet in self.xlsLs:
			if 'Deleted' in sheet:
				print('  [*]', sheet)
				supplyF = pd.read_excel(sheet, 'Sheet1', index_col=None)
				supplyF.to_csv(sheet.replace(sheet[sheet.rindex('.'):], '') + '.csv', sep=',', encoding='UTF-8')
			elif 'HC' in sheet:
				print('  [*]', sheet)
				hcF = pd.read_excel(sheet, 'Sheet1', index_col=None)
				hcF.to_csv(sheet.replace(sheet[sheet.rindex('.'):], '') + '.csv', sep=',', encoding='UTF-8')
			elif 'reporting' in sheet:
				print('  [*]', sheet, end='  ----> ')
				try:
					forF = pd.read_excel(sheet, 'Conso', index_col=None)
					forF.to_csv(sheet.replace(sheet[sheet.rindex('.'):], '') + '.csv', sep=',', encoding='UTF-8')
				except XLRDError:
					print('sheet contains password, ignore this error if you already converted manually.')
		print('  ', len(self.xlsLs), 'Excel files found...')

		print('\nCSV UTF-8 Conversion Completed...')
		for f in os.listdir(self.path):
			if f.endswith('.csv'):
				self.csvLs.append(f)
		#Assign CSV sheets variable
		for sheet in self.csvLs:
			print('  [*]', sheet)
			if 'Deleted' in sheet:
				self.supplyCSV = sheet
			elif 'HC' in sheet:
				self.hcCSV = sheet
			elif 'reporting' in sheet:
				self.forCSV = sheet
		print('  ', len(self.csvLs), 'CSV files found...')

	def del_col(self):
		print('\nDeleting columns from', self.supplyCSV, '...\n')
		with open(self.forCSV, 'r', encoding='UTF-8') as for_csv:
			df = pd.read_csv(for_csv, low_memory=False)
			with open(self.supplyCSV, 'r', encoding='UTF-8') as supply_csv:
				df2 = pd.read_csv(supply_csv)
				for i in df2.columns:
					try:
						del df[i]
					except KeyError:
						pass
				df.to_csv('draft.csv', index=False)

	def filter_rows(self):
		#latest sheet(#FILTER 1,597 records)
		print('\nFiltering Rows...')
		with open('draft.csv', 'r', encoding='UTF-8') as draft_csv:
			df = pd.read_csv(draft_csv, low_memory=False)
			df = df.loc[df['IG'].isin(['SFDC IPS', 'Oracle IPS', 'Workday IPS']) | df['Resources Reqd From'].isin(['Salesforce IPS', 'Oracle IPS', 'Workday IPS'])]
			output = df.drop('Technology', axis=1)
			output.to_csv('draft.csv', index=False)

	def vlookup(self):
		print('\nVLOOKUP', self.hcCSV, '&\n', ' ', self.forCSV, '\n')
		with open('draft.csv', 'r', encoding='UTF-8') as draft_csv:
			df = pd.read_csv(draft_csv)
			with open(self.hcCSV, 'r', encoding='UTF-8') as hc_csv:
				df2 = pd.read_csv(hc_csv)
				df2 = df2[['Name', 'Technology']].drop_duplicates()
				output = df.merge(df2, on=['Name'], how='left')
				output.to_csv('output.csv', index=False)
		#take not null values or drop null values and 
		#replace all NULL to Non-Cloud in Technology Column
		with open('output.csv', 'r', encoding='UTF-8') as output_csv:
			df = pd.read_csv(output_csv)
			df['Technology'] = df['Technology'].fillna('Non-Cloud')
			output = df[pd.notnull(df['Personnel No'])]
			output.to_csv('output.csv', index=False)



	def save_output(self):
		os.remove('draft.csv')
		os.remove(self.supplyCSV)
		os.remove(self.hcCSV)
		#os.remove(self.forCSV)

		print('Enter output filename :')
		fn = input('	')
		print('Saving output file as : \n', ' [*]', fn + str('.csv'))
		try:
			os.rename('output.csv', fn + '.csv')
		except FileExistsError:
			os.remove(fn + '.csv')
			os.rename('output.csv', fn + '.csv')

		root = Tk()
		root.withdraw()
		messagebox.showinfo(title='NOTE: ', message='\nPlease find and replace all characters "Ã±" to "ñ" manually...\nClick OK to open the output file.')

		print('Opening output file : \n', ' [*]', self.path + str('\\') + fn + str('.csv\n'))
		os.startfile(fn + '.csv')

if __name__ == '__main__':

	print('\n+-----------+-----------+')
	print(' Supply Automation Tool')
	print('   by Mark Mon Monteros')
	print('+-----------+-----------+')
	print(' Coded in Python ver 3.7*\n')
	print('NOTE:', 'Please convert manually password-protected files\n        to CSV-UTF8 before executing this program.\n')
	
	SupplyTool()

	print('\nD O N E !!!')
