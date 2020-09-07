'''
This class is used to represent a MS Office document, based
on its file extension and where it is located. It determines the
location of each component type based on the document type and 
returns the components in their respective lists.
'''

from os import walk
from os.path import basename, normpath, isdir, isfile, join
from xml.dom import minidom


class Document:

	def __init__(self, doc_type, doc_location):
		'''Initialize document type, location, and metadata paths'''
		self.doc_type = doc_type
		self.doc_location = doc_location
		self.path_info = self.handle_doc_type()


	def handle_doc_type(self):
		'''Based on the extension, create metadata path information'''
		if self.doc_type == '.docx':
			return self.process_word_doc(self.doc_location)
		elif self.doc_type == '.xlsx':
			return self.process_excel(self.doc_location)
		elif self.doc_type == '.pptx':
			return self.process_powerpoint(self.doc_location)
		else:
			return {}


	def process_word_doc(self, base):
		'''Defines locations for Word document metadata'''
		word_paths = {
			'XML' : [base + '/customXML/', base + '/docProps/', base + '/word/'],
			'CSS' : [base + '/word/styles.xml'],
			'Images' : [base + '/word/media/'],
			'Content' : [base + '/word/document.xml'],
		}
		return word_paths


	def process_excel(self, base):
		'''Defines locations for Excel document metadata'''
		excel_paths = {
			'XML' : [base + '/docProps/', base + '/xl/'],
			'CSS' : [base + '/xl/styles.xml'],
			'Images' : [base + '/xl/media/'],
			'Content' : [base + '/xl/workbook.xml', base + '/xl/worksheets/'],
		}
		return excel_paths


	def process_powerpoint(self, base):
		'''Defines locations for PowerPoint document metadata'''

		ppt_paths = {
			'XML' : [base + '/docProps/', base + '/ppt/'],
			'CSS' : [base + '/ppt/tableStyles.xml'],
			'Images' : [base + '/ppt/media/'],
			'Content' : [base + '/ppt/presentation.xml', base + '/ppt/slides/', base + '/ppt/slideMasters/'],
		}
		return ppt_paths


	def get_xml(self, component_type):
		'''Gets xml metadata files. Content, CSS, and XML content types are stored as XML'''
		search_locations = self.path_info[component_type] # Get search locations for this document
		retrieved_xml = []
		# Loop through search locations
		for location in search_locations:
			# If location points to a file no need to crawl the directory
			if isfile(location):
				fname = basename(normpath(location))
				retrieved_xml.append(location)
			# If location is directory, loop through it and grab relevant files
			elif isdir(location):
				for (dirpath, dirnames, filenames) in walk(location):
					for file in filenames:
						if file.endswith('.xml'):
							self.format_xml(join(dirpath, file))
							retrieved_xml.append(join(dirpath, file))
		return retrieved_xml


	def get_images(self):
		'''Gets document image metadata'''
		# Get search location(s) for image files
		search_locations = self.path_info['Images']
		retrieved_images = []
		# Loop through all images and return them as a list
		for location in search_locations:
			if isdir(location):
				for (dirpath, dirnames, filenames) in walk(location):
					for file in filenames:
						retrieved_images.append(join(dirpath, file))
		return retrieved_images


	def format_xml(self, input_file):
		'''Formats XML files to make them more readable'''
		xml = minidom.parse(input_file)
		formatted_xml = xml.toprettyxml(indent='	', newl='\r')

		with open(input_file, 'w') as formatted_file:
			formatted_file.write(formatted_xml)

