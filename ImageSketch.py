
__author__ = 'Mark Mon Monteros'

import os, shutil
from psd_tools import PSDImage
from PIL import Image


class ImageSketch():

	def __init__(self):
	    self.path = os.path.dirname(os.path.realpath(__file__))
	    self.source = self.path + '\\300dpi\\'
	    self.converted = self.path + '\\converted\\'
	    self.output = self.path + '\\output\\'
	    self.psdFiles = list()
	    self.jpgFiles = list()

	    self.start()
	    self.convert_psd()
	    self.sketch()
		
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

		for f in self.psdFiles:
			psd = PSDImage.open(self.source + f)
			basename = os.path.basename(f)

			print(str('[*] -- ' + f))
			#print(img.size)
			#print(psd.width)
			#print(psd.height)

			psd.compose().save(os.path.join(self.converted, basename.replace(basename[basename.rindex('.'):], '') + '.jpg'), dpi=(300, 300), quality=100)

	def sketch(self):
		print('\nSketching images...')

		for f in os.listdir(self.converted):
			if f.endswith('.jpg'):
				self.jpgFiles.append(f)

		for f in self.jpgFiles:
			img = Image.open(self.converted + f)
			basename = os.path.basename(f)

			print(str('[*] -- ' + f))
			#print(img.size)
			#print(img.width)
			#print(img.height)

			x_small = img.resize((1200, 1800), Image.ANTIALIAS)
			small = img.resize((1500, 2100), Image.ANTIALIAS)
			medium = img.resize((2400, 3000), Image.ANTIALIAS)
			large = img.resize((3300, 4200), Image.ANTIALIAS)
			x_large = img.resize((5400, 7200), Image.ANTIALIAS)

			x_small.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], '-(x_small)') + '.jpg'), dpi=(300, 300), quality=100)
			small.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], '-(small)') + '.jpg'), dpi=(300, 300), quality=100)
			medium.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], '-(medium)') + '.jpg'), dpi=(300, 300), quality=100)
			large.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], '-(large)') + '.jpg'), dpi=(300, 300), quality=100)
			x_large.save(os.path.join(self.output, basename.replace(basename[basename.rindex('.'):], '-(x_large)') + '.jpg'), dpi=(300, 300), quality=100)

			#print(x_small.size)
			#print(small.size)
			#print(medium.size)
			#print(large.size)
			#print(x_large.size)


if __name__ == '__main__':
	print('\nIMAGE SKETCH AUTOMATION')
	print('\nCreated by: Mark Mon Monteros\n')
	ImageSketch()
	print('\nDONE!!!\n')

