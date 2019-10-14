__author__ = 'Mark Mon Monteros'

#from PyQt5 import QtCore, QtGui, QtWidget
import os, shutil, time, sys
import pandas as pd
import numpy as np

class SalesforceDataloadGenerator():

	clock_start = time.time() #Time before the operations start

	def __init__(self):
		self.path = os.path.dirname(os.path.realpath(__file__))
		self.dataload_path = 'C:\\Users\\mark.mon.e.monteros\\OneDrive - Accenture\\Documents\\Projects\\Marsh\\Dovetail\\Exports\\'
		
		self.dataload_file = self.dataload_path + 'Update Blank Contact Addresses 10092019.csv'
		self.exported_dataload1 = self.dataload_path + 'Account (PROD) - Oct 14 2019.csv'
		self.input_fields1 = ('Id', 'Name', 'ShippingStreet', 'ShippingCity', 'ShippingState',  'ShippingPostalCode', 'ShippingStateCode', 'ShippingCountry', 'ShippingCountryCode', 'BillingStreet', 'BillingCity', 'BillingState', 'BillingPostalCode', 'BillingStateCode', 'BillingCountry', 'BillingCountryCode')
		self.exported_dataload2 = self.dataload_path + 'Contact (PROD) - Oct 14 2019.csv'
		self.input_fields2 = ('Id', 'AccountId', 'Name', 'OtherStreet', 'OtherCity', 'OtherState', 'OtherPostalCode', 'OtherStateCode', 'OtherCountry', 'OtherCountryCode', 'MailingStreet', 'MailingCity', 'MailingState', 'MailingPostalCode', 'MailingStateCode', 'MailingCountry', 'MailingCountryCode')

		self.set_fields1 = list()
		self.set_fields2 = list()
		self.get_acct_ids = list()
		self.get_contact_ids = list()

		self.get_headers()
		self.match_ids()
		self.generate()
		self.mapping()

	def get_headers(self):
		for fields in self.input_fields1:
			self.set_fields1.append(fields.upper())

		for fields in self.input_fields2:
			self.set_fields2.append(fields.upper())

	def match_ids(self):
		with open(self.dataload_file, 'r') as dataload_file:
			df = pd.read_csv(dataload_file, low_memory=False)

		with open(self.exported_dataload1, 'r') as export_file1:
			ef1 = pd.read_csv(export_file1, low_memory=False)
		
		with open(self.exported_dataload2, 'r') as export_file2:
			ef2 = pd.read_csv(export_file2, low_memory=False)		

		for ids in list(df['ACCOUNTID'].drop_duplicates()):
			if ids in list(ef1['ID']):
				self.get_acct_ids.append(ids)

		for ids in list(df['ID']):
			if ids in list(ef2['ID']):
				self.get_contact_ids.append(ids)

		account_extract = ef1.loc[ef1['ID'].isin(self.get_acct_ids)]
		account_extract = account_extract[self.set_fields1]

		account_extract.to_csv(os.path.join(self.dataload_path, r'account_extract.csv'), sep=',', index=False, encoding='UTF-8')

	def generate(self):
		with open(self.dataload_file, 'r') as dataload_file:
			df = pd.read_csv(dataload_file, low_memory=False)
			df = df[['ID', 'ACCOUNTID', 'NAME']]

		with open(os.path.join(self.dataload_path, r'account_extract.csv'), 'r') as account_extract:
			ae = pd.read_csv(account_extract, low_memory=False)
			ae.rename(columns={'ID':'ACCOUNTID'}, inplace=True)

		output = df.merge(ae, on=['ACCOUNTID'], how='left')
		output.to_csv(os.path.join(self.dataload_path, r'dataload_draft.csv'), sep=',', index=False, encoding='UTF-8')

	def mapping(self):	
		#MAPPING here
		'''
		Contact Field -> Account Field
		OTHERSTREET -> SHIPPINGSTREET
		OTHERCITY -> SHIPPINGCITY
		OTHERSTATE -> SHIPPINGSTATE
		OTHERPOSTALCODE -> SHIPPINGPOSTALCODE
		MAILINGSTREET -> BILLINGSTREET
		MAILINGCITY -> BILLINGCITY
		MAILINGSTATE -> BILLINGSTATE
		MAILINGPOSTALCODE -> BILLINGPOSTALCODE
		MAILINGSTATECODE -> BILLINGSTATECODE
		OTHERSTATECODE -> SHIPPINGSTATECODE
		#MAPPING here (Account to Contact fields)
		mapping = {
			'SHIPPINGSTREET': 'OTHERSTREET',
			'SHIPPINGCITY': 'OTHERCITY', 
			'SHIPPINGSTATE': 'OTHERSTATE',
			'SHIPPINGPOSTALCODE': 'OTHERPOSTALCODE',
			'BILLINGSTREET': 'MAILINGSTREET',
			'BILLINGCITY': 'MAILINGCITY', 
			'BILLINGSTATE': 'MAILINGSTATE',
			'BILLINGPOSTALCODE': 'MAILINGPOSTALCODE',
			'BILLINGSTATECODE': 'MAILINGSTATECODE', 
			'SHIPPINGSTATECODE': 'OTHERSTATECODE',
			'NAME_x':'NAME',
			'NAME_y':'ACCOUNTNAME'
		}
		'''
		with open(os.path.join(self.dataload_path, r'dataload_draft.csv'), 'r') as dataload_draft:
			dd = pd.read_csv(dataload_draft, low_memory=False)
		
		final_fields = list(dd.columns)

		final_fields.remove('ID')
		final_fields.remove('ACCOUNTID')
		final_fields.remove('NAME_x')
		final_fields.remove('NAME_y')

		self.set_fields2.remove('ID')
		self.set_fields2.remove('ACCOUNTID')
		self.set_fields2.remove('NAME')

		#print(str('final_fields --> '), final_fields)
		#print(str('set_fields2 --> '), self.set_fields2)

		for i in range(len(final_fields)):
			dd.rename(columns=lambda x: x.replace(final_fields[i], self.set_fields2[i]), inplace=True)
		dd.rename(columns=lambda x: x.replace('NAME_x', 'NAME'), inplace=True)
		dd.rename(columns=lambda x: x.replace('NAME_y', 'ACCOUNTNAME'), inplace=True)

		dd.to_csv(os.path.join(self.dataload_path, r'dataload_output.csv'), sep=',', index=False, encoding='UTF-8')

		self.housekeeping()

	def housekeeping(self):
		#os.remove(os.path.join(self.dataload_path, r'account_extract.csv'))
		os.remove(os.path.join(self.dataload_path, r'dataload_draft.csv'))

if __name__ == '__main__':
	print('\nSALESFORCE DATALOAD GENERATOR')
	print('\nCreated by: Mark Mon Monteros')
	SalesforceDataloadGenerator()
	