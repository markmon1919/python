__author__ = 'Mark Mon Monteros'

import os, shutil, sys
import pandas as pd
import numpy as np

class DemandTool():

	def __init__(self):
		self.path = os.path.dirname(os.path.realpath(__file__))
		self.dump = self.path + '\\dump'
		self.dcsopath = self.path + '\\dcso'
		self.xls = []
		self.dcso = []
		self.dcso_error = []
		self.dcso_error_substr = []
		self.rhlyesdrop = []
		self.rhlnodrop = []
		self.du = []
		self.ps = []
		self.fe = []
		self.rhlrrd = []
		self.rrdskills = {}
		#self.fn = None

		self.convert()
		self.get_dcso_csv()
		self.cut_dcso_rows()
		self.cut_dcso_columns()
		self.combine_rhl()
		self.filter_rows()
		self.get_rhl_rrd()
		self.get_dcso_rrd_skills()
		self.add_skills_column()
		self.append_skills()
		self.case_time_format()
		self.end()

	def convert(self):
		try:
			os.remove(os.path.join(self.path, r'final.csv'))
		except FileNotFoundError:
			pass
		try:
			os.mkdir(self.dump)
			os.mkdir(self.dcsopath)
		except FileExistsError:
			shutil.rmtree(self.dump)
			shutil.rmtree(self.dcsopath)
			os.mkdir(self.dump)
			os.mkdir(self.dcsopath)
		#Renaming file extension case
		for f in os.listdir(self.path):
			if f.endswith('.XLS') or f.endswith('.XLSX') or f.endswith('.XLSM'):
				os.rename(f, f.replace(f[f.rindex('.'):], f[f.rindex('.'):].lower()))
		print('\nConverting excel files to CSV UTF-8 format...')
		for f in os.listdir(self.path):
			if f.endswith('.xls') or f.endswith('.xlsx') or f.endswith('.xlsm'):
				self.xls.append(f)
		for sheet in self.xls:
			if 'ATCP Open and Overdue Demand' in sheet:
				print('  [*]', sheet)
				rhl_yes = pd.read_excel(sheet, 'RHL=Yes', index_col=None)
				rhl_no = pd.read_excel(sheet, 'RHL=No', index_col=None)
				rhl_yes.to_csv(os.path.join(self.dump, r'RHL Yes.csv'), sep=',', index=False, encoding='UTF-8')
				rhl_no.to_csv(os.path.join(self.dump, r'RHL No.csv'), sep=',', index=False, encoding='UTF-8')
			elif 'Demand Columns To Be Included' in sheet:
				print('  [*]', sheet)
				demand_cols = pd.read_excel(sheet, 'Sheet1', index_col=None)
				demand_cols.to_csv(os.path.join(self.dump, (sheet.replace(sheet[sheet.rindex('.'):], '') + '.csv')), sep=',', index=False, encoding='UTF-8')
			elif 'Data loader' in sheet:
				print('  [*]', sheet)
				data_loader = pd.read_excel(sheet, 'Data loader - Salesforce', index_col=None)
				data_loader.to_csv(os.path.join(self.dump, (sheet.replace(sheet[sheet.rindex('.'):], '') + '.csv')), sep=',', index=False, encoding='UTF-8')
			else:
				print('  [*]', sheet)
				dcso_csv = pd.read_excel(sheet, 'Resource Requirements Detail', index_col=None)
				dcso_csv.to_csv(os.path.join(self.dcsopath, (sheet.replace(sheet[sheet.rindex('.'):], '') + '.csv')), sep=',', index=False, encoding='UTF-8')
		print('  ', len(self.xls), 'Excel files found...')
		print('\nCSV UTF-8 Conversion Completed...')

	def get_dcso_csv(self):
		for f in os.listdir(self.dcsopath):
			if f.endswith('.csv'):
				self.dcso.append(f)
		print('\nTotal DCSO files: ' + str(len(self.dcso)))

	def cut_dcso_rows(self):
		for dcso in self.dcso:
			with open(os.path.join(self.dcsopath, dcso), 'r', encoding='UTF-8') as dcso_csv:
				df = pd.read_csv(dcso_csv, low_memory=False)
				df.drop(df.index[:28], inplace=True)
				df.to_csv(os.path.join(self.dcsopath, r'#edited - ' + dcso), header=None, index=False)
		del self.dcso[:]
		for f in os.listdir(self.dcsopath):
			if f.startswith('#edited'):
				self.dcso.append(f)
			else:
				os.remove(os.path.join(self.dcsopath, f))

	def cut_dcso_columns(self):
		print('\nDeleting Columns...')
		with open(os.path.join(self.dump, r'Demand Columns To Be Included.csv'), 'r', encoding='UTF-8') as del_colums:
			df = pd.read_csv(del_colums, low_memory=False)
			df.columns = df.columns.str.upper() 
		with open(os.path.join(self.dump, r'RHL Yes.csv'), 'r', encoding='UTF-8') as rhl_yes:
			df2 = pd.read_csv(rhl_yes, low_memory=False)
			df2.columns = df2.columns.str.upper()
		with open(os.path.join(self.dump, r'RHL No.csv'), 'r', encoding='UTF-8') as rhl_no:
			df3 = pd.read_csv(rhl_no, low_memory=False)
			df3.columns = df3.columns.str.upper()
		#drop RHL YES columns
		for col in list(df2.columns):
			if col not in list(df.columns) or not list(df3.columns):
				self.rhlyesdrop.append(col)
		#drop RHL NO columns
		for col in list(df3.columns):
			if col not in list(df.columns) or not list(df2.columns):
				self.rhlnodrop.append(col)

		df2.drop(self.rhlyesdrop, axis=1, inplace=True)
		df2.to_csv(os.path.join(self.dump, r'RHL-Y.csv'), index=False, encoding='UTF-8')

		df3.drop(self.rhlnodrop, axis=1, inplace=True)
		df3.to_csv(os.path.join(self.dump, r'RHL-N.csv'), index=False, encoding='UTF-8')

	def combine_rhl(self):
		with open(os.path.join(self.dump, r'RHL-Y.csv'), 'r', encoding='UTF-8') as rhl_yes:
			df = pd.read_csv(rhl_yes, low_memory=False)
		with open(os.path.join(self.dump, r'RHL-N.csv'), 'r', encoding='UTF-8') as rhl_no:
			df2 = pd.read_csv(rhl_no, low_memory=False)
			output = pd.concat([df, pd.DataFrame(df2)], join='outer')
			output.to_csv(os.path.join(self.dump, r'RHL combined.csv'), index=False, encoding='UTF-8')

		for f in os.listdir(self.dump):
			if not f.endswith('combined.csv') and not f.startswith('Data loader'):
				os.remove(os.path.join(self.dump, f))

	def filter_rows(self):
		print('\nFiltering Rows...')
		with open(os.path.join(self.dump, r'Data loader - Salesforce.csv'), 'r', encoding='UTF-8') as data_loader:
			df = pd.read_csv(data_loader, low_memory=False)
			du = df['Delivery Unit/DU'].dropna()
			ps = df['Primary Skill'].dropna()
			fe = df['Fulfillment Entity'].dropna()

		for i in du:
			self.du.append(i)
		for i in ps:
			self.ps.append(i)
		for i in fe:
			self.fe.append(i)

		with open(os.path.join(self.dump, r'RHL combined.csv'), 'r', encoding='UTF-8') as rhl:
			df = pd.read_csv(rhl, low_memory=False)
			output = df.loc[df['DU'].isin(self.du) | df['PRIMARY SKILL'].isin(self.ps) | df['FULFILLMENT ENTITY'].isin(self.fe)]
			output.to_csv(os.path.join(self.dump, r'#edited - RHL combined.csv'), index=False, encoding='UTF-8')

		for f in os.listdir(self.dump):
			if not f.startswith('#edited'):
				os.remove(os.path.join(self.dump, f))

	def get_rhl_rrd(self):
		with open(os.path.join(self.dump, '#edited - RHL combined.csv'), 'r', encoding='UTF-8') as rhl:
			df = pd.read_csv(rhl, low_memory=False)
			df = df['RRD NUMBER']

			for i in df:
				self.rhlrrd.append(i)

	def get_dcso_rrd_skills(self):
		for dcso in self.dcso:
			try:
				with open(os.path.join(self.dcsopath, dcso), 'r', encoding='UTF-8') as dcso_csv:
					df = pd.read_csv(dcso_csv, low_memory=False, index_col='RRD\nNumber')
			except ValueError:
				self.dcso_error.append(dcso.split('- ', 1)[1])

		for i in self.dcso_error:
			self.dcso_error_substr.append(i.rsplit('.', 1)[0])

		if not self.dcso_error_substr:
			pass
		else:
			self.error()
					
		for i in self.rhlrrd:
			try:
				self.rrdskills[i] = df.loc[i, 'Skill']
			except KeyError:
				pass

		with open(os.path.join(self.dump, r'rrd_skills.csv'), 'w', encoding='UTF-8') as rrdskills_csv:
			for i in self.rrdskills.keys():
				rrdskills_csv.write("%s,%s\n"%(i,self.rrdskills[i]))

		with open(os.path.join(self.dump, r'rrd_skills.csv'), 'r', encoding='UTF-8') as rrdskills_csv:
			df = pd.read_csv(rrdskills_csv, low_memory=False, names=['RRD NUMBER', 'ADDITIONAL SKILLS (NICE TO HAVE THIS)'])
			df.to_csv(os.path.join(self.dump, r'#rrd_skills.csv'), index=False, encoding='UTF-8')

	def add_skills_column(self):
		with open(os.path.join(self.dump, r'#edited - RHL combined.csv'), 'r', encoding='UTF-8') as rhl:
			df = pd.read_csv(rhl, low_memory=False)
			#FILL NULL VALUES TO COLUMN
			df['ADDITIONAL SKILLS (NICE TO HAVE THIS)'] = [np.nan for _ in range(len(df))]
			df.to_csv(os.path.join(self.dump, r'#final.csv'), index=False, encoding='UTF-8')

	def append_skills(self):
		with open(os.path.join(self.dump, r'#final.csv'), 'r', encoding='UTF-8') as rhl:
			df = pd.read_csv(rhl, low_memory=False)
			df = df[['RRD NUMBER', 'ADDITIONAL SKILLS (NICE TO HAVE THIS)']]
		with open(os.path.join(self.dump, r'#rrd_skills.csv'), 'r', encoding='UTF-8') as rrdskills_csv:
			df2 = pd.read_csv(rrdskills_csv, low_memory=False)

		output = df.merge(df2, on=['RRD NUMBER'], how='left')
		output = output.drop('ADDITIONAL SKILLS (NICE TO HAVE THIS)_x', axis=1)
		output.rename(columns={'RRD NUMBER':'RRD NUMBER', 'ADDITIONAL SKILLS (NICE TO HAVE THIS)_y':'ADDITIONAL SKILLS (NICE TO HAVE THIS)'}, inplace=True)
		output.to_csv(os.path.join(self.dump, r'#map.csv'), index=False, encoding='UTF-8')

		with open(os.path.join(self.dump, r'#final.csv'), 'r', encoding='UTF-8') as rhl:
			df3 = pd.read_csv(rhl, low_memory=False)
		with open(os.path.join(self.dump, r'#map.csv'), 'r', encoding='UTF-8') as map_csv:
			df4 = pd.read_csv(map_csv, low_memory=False)
			df4 = df4['ADDITIONAL SKILLS (NICE TO HAVE THIS)']

		df3 = df3.drop('ADDITIONAL SKILLS (NICE TO HAVE THIS)', axis=1)
		output = pd.concat([df3, df4], axis=1)
		output.to_csv(os.path.join(self.path, r'output.csv'), index=False, encoding='UTF-8')

	def case_time_format(self):
		with open(os.path.join(self.path, r'output.csv'), 'r', encoding='UTF-8') as output_csv:
			df = pd.read_csv(output_csv, low_memory=False)
			df.columns = df.columns.str.title()

			changeUpper = ['Id', 'Sg', 'Du', 'Og', 'Dcso', 'Rrd', 'Dg', 'Sid', 'Sr', 
			'Ig', 'Drd', 'Rsd', 'Gu', 'Dcn', 'Gcp', 'Ou', 'Csg', 'It', 'Rhly']
			
			#ID, #SG, #DU, #OG, #DCSO, #RRD, #WBSe, #DG, #SID, #SR, #IG,
			#DRD, #RSD, #GU, #DCN, #GCP, #To, #Of, #OU, #CSG, #IT, #RHLY

			for i in changeUpper:
				df.rename(columns=lambda x: x.replace(i, i.upper()), inplace=True)
			df.rename(columns=lambda x: x.replace('Wbse', 'WBSe'), inplace=True)
			df.rename(columns=lambda x: x.replace('To', 'to'), inplace=True)
			df.rename(columns=lambda x: x.replace('Of', 'of'), inplace=True)
	
			for i in df.columns:
				if not i.startswith('No of Times') and not i.startswith('Abacus') and 'Date' in i or 'Created' in i or i.endswith('Changed On'):
					df[i] = pd.to_datetime(df[i])
					df[i] = df[i].dt.strftime('%m/%d/%Y')

			df = df.replace('NaT', np.nan, regex=True)
			df.to_csv(os.path.join(self.path, r'final.csv'), index=False, encoding='UTF-8')

	def end(self):
		#print('\nSaving file as: ', self.fn)
		print('\nDONE!\n')
		shutil.rmtree(self.dump)
		shutil.rmtree(self.dcsopath)
		os.remove(os.path.join(self.path, r'output.csv'))
		os.startfile(os.path.join(self.path, r'final.csv'))

	def error(self):
		print('\nERROR: Please check/review "Resource Requirement Detail" column for these DCSO files..')
		for i in self.dcso_error_substr:
			print('   [*]', i)
		print('     -- Column header should start at row #30 or DCSO sheet template is mismatch  --\n')
		shutil.rmtree(self.dump)
		#shutil.rmtree(self.dcsopath)
		sys.exit()

if __name__ == '__main__':
	print('\nDEMAND TOOL AUTOMATION')
	print('\nCreated by: Mark Mon Monteros')
	DemandTool()