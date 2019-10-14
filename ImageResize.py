
__author__ = 'Mark Mon Monteros'

import os, shutil
from psd_tools import PSDImage
from PIL import Image


class ImageResize():

	def __init__(self):
	    self.path = os.path.dirname(os.path.realpath(__file__))
	    self.source = self.path + '\\psd_pics\\'
	    self.converted = self.path + '\\converted\\'
	    self.output = self.path + '\\output\\'
	    self.psdFiles = list()
	    self.dpi = [75, 300]

	    self.x_small_w = float()
	    self.x_small_h = float()
	    self.small_w = float()
	    self.small_h = float()
	    self.medium_w = float()
	    self.medium_h = float()
	    self.large_w = float()
	    self.large_h = float()
	    self.x_large_w = float()
	    self.x_large_h = float()

	    self.start()
	    self.convert_psd()
		
	def start(self):
	    try:
	    	os.mkdir(self.output)
	    	os.mkdir(self.converted)
	    except FileExistsError:
	    	shutil.rmtree(self.output)
	    	shutil.rmtree(self.converted)
	    	os.mkdir(self.output)
	    	os.mkdir(self.converted)

	def convert_psd(self):
		print('\nConverting PSD to JPEG...')

		for f in os.listdir(self.source):
			if f.endswith('.psd'):
				self.psdFiles.append(f)

		while (True):
			try:
				dpi_input = int(input("Choose DPI [75] or [300]: "))
			except ValueError:
				print('Should be a number..')
				continue
			if dpi_input in self.dpi:
				break
			else:
				print('DPI should be 75 or 300.')
				continue

		while (True):
			try:
				quality_input = int(input("Select QUALITY from range [1-100]: "))
			except ValueError:
				print('Should be a number..')
				continue
			if quality_input < 1:
				continue
			elif quality_input > 100:
				continue
			else:
				break

		for f in self.psdFiles:
			psd = PSDImage.open(self.source + f)
			basename = os.path.basename(f)

			print(str('[*] -- ' + f))
			print(psd.size)

			psd.compose().save(os.path.join(self.converted, basename.replace(basename[basename.rindex('.'):], '') + '.jpg'), dpi=(dpi_input, dpi_input), quality=quality_input)

		if dpi_input == 75:
			self._75here(dpi_input, quality_input)
		if dpi_input == 300:
			self._300here(dpi_input, quality_input)

	def _75here(self, dpi_input, quality_input):
		print('\nSetting image size...')
		print(str('\nDPI set to: '), dpi_input)
		print(str('QUALITY set to: '), quality_input)

		self.set_x_small()
		self.set_small()
		self.set_medium()
		self.set_large()
		self.set_x_large()
		
		self.switch(dpi_input, quality_input)
		'''
		x_small = img.resize((300, 375), Image.ANTIALIAS)
		small = img.resize((375, 468), Image.ANTIALIAS)
		medium = img.resize((600, 750), Image.ANTIALIAS)
		large = img.resize((825, 1031), Image.ANTIALIAS)
		x_large = img.resize((1350, 1687), Image.ANTIALIAS)
		'''
	def _300here(self, dpi_input, quality_input):
		print('\nSetting image size...')
		print(str('\nDPI set to: '), dpi_input)
		print(str('QUALITY set to \n'), quality_input)

		self.set_x_small()
		self.set_small()
		self.set_medium()
		self.set_large()
		self.set_x_large()
		
		self.switch(dpi_input, quality_input)
		'''
		x_small = img.resize((1200, 1500), Image.ANTIALIAS)
		small = img.resize((1500, 1875), Image.ANTIALIAS)
		medium = img.resize((2400, 3000), Image.ANTIALIAS)
		large = img.resize((3300, 4125), Image.ANTIALIAS)
		x_large = img.resize((5400, 6750), Image.ANTIALIAS)
		'''
	def switch(self, dpi_input, quality_input):
		checker = False

		print('\nCheck the sizes below: ')
		print(str('\n[1]-- X-SMALL: '), '({}"x {}")'.format(self.x_small_w, self.x_small_h))
		print(str('[2] -- SMALL: '), '({}"x {}")'.format(self.small_w, self.small_h))
		print(str('[3] -- MEDIUM: '), '({}"x {}")'.format(self.medium_w, self.medium_h))
		print(str('[4] -- LARGE: '), '({}"x {}")'.format(self.large_w, self.large_h))
		print(str('[5] -- X-LARGE: '), '({}"x {}")'.format(self.x_large_w, self.x_large_h))

		try:
			num = float(input('\nSelect a number from [1-5] if you want to edit. \nOr press any key when your good! '))
		except ValueError:
			checker = False
			num = float(0)

		if num == 1:
			checker = True
			self.set_x_small()
		elif num == 2:
			checker = True
			self.set_small()
		elif num == 3:
			checker = True
			self.set_medium()
		elif num == 4:
			checker = True
			self.set_large()
		elif num == 5:
			checker = True
			self.set_x_large()
		else:
			checker = False

		if (checker):
			self.switch(dpi_input, quality_input)
		else:
			self.to_pixels(dpi_input, quality_input)

	def set_x_small(self):
		#X-SMALL - width
		while (True):
			print('\n+-+- Enter number in inches for X-SMALL size (width, height) -+-+')
			try:
				self.x_small_w = float(input('width: '))
			except ValueError:
				print('Should be a number..')
				continue
			if self.x_small_w <=0 :
				print('Width should be greater than 0..')
			else:
				break
		#X-SMALL - height
		while (True):
			print('\n+-+- Enter number in inches for X-SMALL size (width, height) -+-+')
			print(str('width: '), self.x_small_w)
			try:
				self.x_small_h = float(input('height: '))
			except ValueError:
				print('Should be a number..')
				continue
			if self.x_small_h <= 0:
				print('Height should be greater than 0..')
				continue
			else:
				break	

	def set_small(self):
		#SMALL - width
		while (True):
			print('\n+-+- Enter number in inches for SMALL size (width, height) -+-+')
			try:
				self.small_w = float(input('width: '))
			except ValueError:
				print('Should be a number..')
				continue
			if self.small_w <=0 :
				print('Width should be greater than 0..')
			else:
				break
		#SMALL - height
		while (True):
			print('\n+-+- Enter number in inches for SMALL size (width, height) -+-+')
			print(str('width: '), self.small_w)
			try:
				self.small_h = float(input('height: '))
			except ValueError:
				print('Should be a number..')
				continue
			if self.small_h <=0 :
				print('Height should be greater than 0..')
			else:
				break

	def set_medium(self):
		#MEDIUM - width
		while (True):
			print('\n+-+- Enter number in inches for MEDIUM size (width, height) -+-+')
			try:
				self.medium_w = float(input('width: '))
			except ValueError:
				print('Should be a number..')
				continue
			if self.medium_w <= 0:
				print('Width should be greater than 0..')
			else:
				break
		#MEDIUM - height
		while (True):
			print('\n+-+- Enter number in inches for MEDIUM size (width, height) -+-+')
			print(str('width: '), self.medium_w)
			try:
				self.medium_h = float(input('height: '))
			except ValueError:
				print('Should be a number..')
				continue
			if self.medium_h <= 0:
				print('Height should be greater than 0..')
			else:
				break

	def set_large(self):
		#LARGE - width
		while (True):
			print('\n+-+- Enter number in inches for LARGE size (width, height) -+-+')
			try:
				self.large_w = float(input('width: '))
			except ValueError:
				print('Should be a number..')
				continue
			if self.large_w <= 0:
				print('Width should be greater than 0..')
			else:
				break
		#LARGE - height
		while (True):
			print('\n+-+- Enter number in inches for LARGE size (width, height) -+-+')
			print(str('width: '), self.large_w)
			try:
				self.large_h = float(input('height: '))
			except ValueError:
				print('Should be a number..')
				continue
			if self.large_h <= 0:
				print('Height should be greater than 0..')
			else:
				break	

	def set_x_large(self):
		#X-LARGE - width
		while (True):
			print('\n-+- Enter number in inches for X-LARGE size (width, height) -+-+')
			try:
				self.x_large_w = float(input('width: '))
			except ValueError:
				print('Should be a number..')
				continue
			if self.x_large_w <- 0:
				print('Width should be greater than 0..')
			else:
				break
		#X-LARGE - height
		while (True):
			print('\n-+- Enter number in inches for X-LARGE size (width, height) -+-+')
			print(str('width: '), self.x_large_w)
			try:
				self.x_large_h = float(input('height: '))
			except ValueError:
				print('Should be a number..')
				continue
			if self.x_large_h <= 0:
				print('Height should be greater than 0..')
			else:
				break	
		
	def to_pixels(self, dpi_input, quality_input):
		print('\nResizing image...')
		#Formula: Inches X DPI = Pixels

		self.x_small_w = int(self.x_small_w * dpi_input)
		self.x_small_h = int(self.x_small_h * dpi_input)
		self.small_w = int(self.small_w * dpi_input)
		self.small_h = int(self.small_h * dpi_input)
		self.medium_w = int(self.medium_w * dpi_input)
		self.medium_h = int(self.medium_h * dpi_input)
		self.large_w = int(self.large_w * dpi_input)
		self.large_h = int(self.large_h * dpi_input)
		self.x_large_w = int(self.x_large_w * dpi_input)
		self.x_large_h = int(self.x_large_h * dpi_input)

		for f in os.listdir(self.converted):
			if f.endswith('.jpg'):
				img = Image.open(self.converted + f)
				basename = os.path.basename(f)		

				x_small = img.resize((self.x_small_w, self.x_small_h), Image.ANTIALIAS)
				small = img.resize((self.small_w, self.small_h), Image.ANTIALIAS)
				medium = img.resize((self.medium_w, self.medium_h), Image.ANTIALIAS)
				large = img.resize((self.large_w, self.large_h), Image.ANTIALIAS)
				x_large = img.resize((self.x_large_w, self.x_large_h), Image.ANTIALIAS)

				print(str('[*] -- ' + f), str('@'), '({} x {})'.format(x_small.width, x_small.height), str('--'), dpi_input, str('DPI'), '<{} quality>'.format(quality_input))
				print(str('[*] -- ' + f), str('@'), '({} x {})'.format(small.width, small.height), str('--'), dpi_input, str('DPI'), '<{} quality>'.format(quality_input))
				print(str('[*] -- ' + f), str('@'), '({} x {})'.format(medium.width, medium.height), str('--'), dpi_input, str('DPI'), '<{} quality>'.format(quality_input))
				print(str('[*] -- ' + f), str('@'), '({} x {})'.format(large.width, large.height), str('--'), dpi_input, str('DPI'), '<{} quality>'.format(quality_input))
				print(str('[*] -- ' + f), str('@'), '({} x {})'.format(x_large.width, x_large.height), str('--'), dpi_input, str('DPI'), '<{} quality>'.format(quality_input))
				
		if dpi_input == 75:
			x_small.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], ' (75dpi)_x-small') + '.jpg'), dpi=(dpi_input, dpi_input), quality=quality_input)
			small.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], ' (75dpi)_small') + '.jpg'), dpi=(dpi_input, dpi_input), quality=quality_input)
			medium.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], ' (75dpi)_medium') + '.jpg'), dpi=(dpi_input, dpi_input), quality=quality_input)
			large.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], ' (75dpi)_large') + '.jpg'), dpi=(dpi_input, dpi_input), quality=quality_input)
			x_large.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], ' (75dpi)_x-large') + '.jpg'), dpi=(dpi_input, dpi_input), quality=quality_input)
		if dpi_input == 300:
			x_small.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], ' (300dpi)_x-small') + '.jpg'), dpi=(dpi_input, dpi_input), quality=quality_input)
			small.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], ' (300dpi)_small') + '.jpg'), dpi=(dpi_input, dpi_input), quality=quality_input)
			medium.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], ' (300dpi)_medium') + '.jpg'), dpi=(dpi_input, dpi_input), quality=quality_input)
			large.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], ' (300dpi)_large') + '.jpg'), dpi=(dpi_input, dpi_input), quality=quality_input)
			x_large.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], ' (300dpi)_x-large') + '.jpg'), dpi=(dpi_input, dpi_input), quality=quality_input)



if __name__ == '__main__':
	print('\nIMAGE RESIZE AUTOMATION')
	print('\nCreated by: ' + __author__)

	print('\nINSTRUCTIONS:')
	print('[*] -- Create a folder named "psd_pics", make sure this folder and the application are both in the same directory.')
	print('[*] -- Place .PSD image/s inside psd_pics folder.')
	print('[*] -- Execute the application.')
	print('[*] -- A folder named "converted" is created where the converted image/s from PSD to JPEG are stored.')
	print('[*] -- A folder named "output" is created where the resized image/s are stored.')

	ImageResize()

	print('\nDONE!!!\n')