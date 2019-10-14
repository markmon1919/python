__author__ = 'Mark Mon Monteros'

#from PyQt5 import QtCore, QtGui, QtWidget
import os, shutil, time, sys
import pandas as pd
import numpy as np

class SalesforceMapCreator():
	clock_start = time.time() #Time before the operations start

	def __init__(self):
		self.path = os.path.dirname(os.path.realpath(__file__))
		self.maps_path = 'C:\\Users\\mark.mon.e.monteros\\OneDrive - Accenture\\Documents\\Projects\\Marsh\\Dovetail\\Exports\\maps\\'
		self.object = None

		self.set_object()

	def set_object(self):
		checker = False

		print(str('\n[1] -- ACCOUNT'))
		print(str('[2] -- CONTACT'))

		try:
			input_obj = int(input('\nSelect object to create field mapping: '))
		except (ValueError, TypeError):
			checker = False
			input_obj = int(0)

		if input_obj == 1:
			checker = True
			self.object = self.maps_path + 'Account.txt'
		elif input_obj == 2:
			checker = True
			self.object = self.maps_path + 'Contact.txt'
		else:
			checker = False

		if (checker):
			print(str('Object: '), os.path.basename(self.object).replace(os.path.basename(self.object)[os.path.basename(self.object).rindex('.'):], ''))
			self.create_map()
		else:
			print('\nNumber not in choice.')
			self.set_object()

	def create_map(self):
		print('\nCreating Map File...')
		with open(self.object, 'r') as fields_txt:
			txt_f = pd.read_csv(fields_txt, low_memory=False, names=['FIELDS'])

		field_name = list()
		to_upper_txt = list()
		for i in list(txt_f['FIELDS']):
			field_name.append(i)
			to_upper_txt.append(i.upper())

		output = open(self.object.replace(self.object[self.object.rindex('.'):], ' (mapping)') + '.sdl', 'w')
		for i in range(len(field_name)):
			output.writelines(to_upper_txt[i] + '=' + field_name[i] + '\n') 
		output.close()

		print(str('Saving...'), os.path.basename(self.object.replace(self.object[self.object.rindex('.'):], ' (mapping)') + '.sdl'))

if __name__ == '__main__':
	print('\nSALESFORCE MAP CREATOR')
	print('\nCreated by: Mark Mon Monteros')
	SalesforceMapCreator()
	print('\nDONE!!!')
