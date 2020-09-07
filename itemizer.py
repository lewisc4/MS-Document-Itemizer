'''
This class is used to retrieve the MS Office documents based on
the search directory and document type. Once the documents are found,
they are moved to the save directory and zipped in order to extract
their metadata. All instances of this class are called within the GUIHandler
class. This class does not actually parse the metadata. That is handeled by
the Document class, which is instantiated in this class
'''

from os import walk, rename, mkdir, chdir
from os.path import splitext, split, basename, normpath, isdir, isfile, join
from shutil import copy, rmtree
from zipfile import ZipFile
from ms_documents import Document


class Itemizer:

	def __init__(self):
		'''Respective arrays/dicts for metadata are initialized here'''
		self.saved_data = {} # Dict for storing document name and its metadata location
		# Arrays to store files for each component type
		self.xml_files = []
		self.image_files = []
		self.content_files = []
		self.css_files = []
		# Boolean that determines if temp folder with extra metadata is deleted
		self.remove_excess = False


	def retrieve_ms_docs(self, search_directory, recursive_search, doc_ext):
		'''Retrieves documents based on the search directory/type and document type'''
		documents = []
		# If the path points to a specific file no looping/extra searching required
		if isfile(search_directory):
			requested_file = basename(normpath(search_directory)) # Get name/extension of file
			# If the extension matches the document type store it in document list
			if requested_file.endswith(tuple(doc_ext)):
				documents.append(requested_file)
		# Potentially many files in search directory
		else:
			# Loop through all files in the search directory, break if search is not recursive
			for (dirpath, dirnames, filenames) in walk(search_directory):
				for file in filenames:
					if file.endswith(tuple(doc_ext)):
						documents.append(join(dirpath, file))
				if not recursive_search: break
				
		return documents


	def zip_documents(self, documents, save_directory):
		'''Zips the documents to the save directory'''
		doc_metadata = {} # Stores metadata directory and the document type
		# Loop through all retrieved documents
		for doc in documents:
			# Get name, extension, and path for document
			doc_path, doc_full_name = split(doc)
			doc_name, doc_ext = splitext(doc_full_name)

			# Create names for zip file, temp directory, and permanent directory
			zipped_doc = doc_name + '.zip'
			new_directory = save_directory + '/' + doc_name + ' Information'
			temp_directory = new_directory + '/' + 'temp'

			# If directory already exists, no need to make it again
			if isdir(new_directory):
				chdir(temp_directory)
			else:
				chdir(save_directory) # Navigate to save directory
				mkdir(new_directory) # Make directory to store document info
				print(doc_full_name)
				copy(doc, new_directory) # Move the document to its directory
				chdir(new_directory) # Move to new directory
				rename(doc_full_name, zipped_doc) # Rename document to create zip file

				# Unzip document to generate metadata
				with ZipFile(zipped_doc, 'r') as zip_obj:
					zip_obj.extractall(temp_directory)

			# Store metadata location and file type
			doc_metadata[temp_directory] = doc_ext
			self.saved_data[doc_name] = new_directory

		return doc_metadata


	def extract_metadata(self, metadata_locations, components, keep_excess_data):
		'''Gathers document metadata using the Document class'''
		# Loop through each document
		for location, doc_type in metadata_locations.items():
			# Create document based on its type and location
			current_doc = Document(doc_type, location)
			# Save metadata for each component type
			for component_type in components:
				if component_type == 'Images':
					self.__save_metadata(current_doc.get_images(), component_type, location)
				else:
					self.__save_metadata(current_doc.get_xml(component_type), component_type, location)
			if not keep_excess_data:
				rmtree(location)


	def __save_metadata(self, file_list, component_type, location):
		'''Saves metadata to permanent location and removes temp folder that has all metadata'''
		if location.endswith('/temp'):
			# Make permanent directory based on temp directory
			permanent_location = location[:-5] + '/' + component_type
		# Check if permanent directory exists
		if not isdir(permanent_location):
			mkdir(permanent_location) # Make permanent save directory
			# Save all component files to respective folders
			for file in file_list:
				copy(file, permanent_location)

