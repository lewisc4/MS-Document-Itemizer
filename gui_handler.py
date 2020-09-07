'''
This class is used to handle the user input and handle the
selections accordingly. This is where the search, file, and
component options are defined. The itemizer class is instantiated
here to handle computation.
'''

from os import getcwd
from os.path import split
from itertools import zip_longest
from itemizer import Itemizer


class GuiHandler:

	def __init__(self):
		'''All default values for GUI are defined here'''
		# Search options to display to user
		self.search_options = [
			"Search current directory (default)",
			"Search for specific file",
			"Search specific directory (non-recursive)",
			"Search specific directory (recursive)",
		]

		# File type options to display to user
		self.file_options = [
			("Word Documents", ".docx"),
			("Excel Spreadsheets", ".xlsx"),
			("Powerpoint Presentations", ".pptx"),
		]

		# Component type options to display to user
		self.component_options = [
			"XML",
			"CSS",
			"Images",
			"Content",
		]

		# Set default values for the handler
		self.search_type = self.search_options[0]
		self.recursive_search = False

		self.file_types = []
		self.component_types = []

		self.search_directory = getcwd()
		self.save_directory = self.search_directory

		self.summary = None


	def set_search_type(self, search_sel):
		'''Set handler search type based on user input'''
		self.search_type = search_sel

		if self.search_type == self.search_options[0]:
			self.search_directory = getcwd()

		elif self.search_type == self.search_options[3]:
			self.recursive_search = True


	def set_selections(self, ftype_sels, component_sels):
		'''Set handler file and component types based on user input'''
		for f, c in zip_longest(ftype_sels, component_sels):
			if f and ftype_sels[f].get(): 
				self.file_types.append(f)

			if c and component_sels[c].get(): 
				self.component_types.append(c)
			

	def set_directories(self, search_inp, save_inp):
		'''Set handler directories based on user input'''
		if search_inp and self.search_type != self.search_options[0]:
			self.search_directory = search_inp

		if save_inp:
			self.save_directory = save_inp


	def itemize_documents(self, keep_excess_data):
		'''This is where the inputs are evaluated and assigned'''
		file_extensions = []
		# Get file extensions of the MS office file types selected 
		file_extensions = []
		for ftype in self.file_types:
			file_extensions.append(ftype[1])

		# Itemizer class instantiated to handle file searching
		itemizer = Itemizer()
		# Gets the documents based on search type and document type
		documents = itemizer.retrieve_ms_docs(self.search_directory, self.recursive_search, file_extensions)
		# Zips then unzips documents to generate metadata and stores its location(s)
		metadata_locations = itemizer.zip_documents(documents, self.save_directory)
		# Parses the metadata and moves desired component types to save location
		itemizer.extract_metadata(metadata_locations, self.component_types, keep_excess_data)

		self.create_summary(documents, metadata_locations)


	def create_summary(self, documents, metadata_locations):
		'''Creates a simple summary to display to the user'''
		dir_summary = 'Directory searched:\n' + self.search_directory + '\n'
		dir_summary += 'Information saved to:\n' + self.save_directory + '\n'

		doc_summary = 'Summary of documents: \n\n'
		for document, metadata in zip(documents, metadata_locations.keys()):
			path, name = split(document)
			
			doc_summary += 'Document Name: ' + name + '\n'
			doc_summary += 'Document Location: ' + path + '\n'
			doc_summary += 'Components Saved to: ' + metadata[:-4] + '\n\n'

		self.summary = dir_summary + doc_summary

