'''
This class is used to set the user interface for the itermizer
script. It defines all of the frames, buttons, and text
inputs/outputs. It also defines the functions to call when
the user selects buttons.
'''

from os import getcwd
from tkinter import *
from gui_handler import GuiHandler


class ItemizerGUI:

	def __init__(self, root):
		'''This is where all components are initialized'''
		# Instance of GUI handler to deal with all user inputs accordingly
		self.handler = GuiHandler()

		# Defines the root view and all of its attributes
		self.root = root
		self.root.title('Microsoft Office Document Itemizer')
		self.root.geometry("810x800")
		self.root.grid_columnconfigure(0, weight=1)
		self.root.grid_rowconfigure(0, weight=1)

		# Here is where all frames are initialized to hold buttons/inputs/outputs
		self.main_frame = self.initialize_frame(self.root, 0, 0)
		self.dropdown_frame = self.initialize_frame(self.main_frame, 0, 0)
		self.directory_frame = self.initialize_frame(self.main_frame, 1, 0)
		self.button_frame = self.initialize_frame(self.main_frame, 2, 0)
		self.summary_frame = self.initialize_frame(self.main_frame, 3, 0)

		# Dictonaries to store file/component type choices
		self.ftype_choices = {}
		self.component_choices = {}	

		# Creates all radiobuttons to place on frames
		self.displ_search_options()
		self.displ_radio_btns(self.handler.file_options, self.ftype_choices, "File Type Options", 1, self.select_all_ftypes)
		self.displ_radio_btns(self.handler.component_options, self.component_choices, "Component Type Options", 2, self.select_all_components)

		# Create directory inputs and their default values
		def_value = getcwd()
		self.search_input = self.create_dir_input("Search Directory", [0,0], [0,1], def_value)
		self.save_input = self.create_dir_input("Save Directory", [1,1], [0,1], def_value)

		# Create compute button and summary output
		self.create_comp_btn()
		self.create_summary_outp()


	def initialize_frame(self, parent_frame, r, c):
		'''Creates a frame based on parent frame and the rows/columns for layout'''
		initial_frame = Frame(parent_frame, padx=8, pady=8)
		initial_frame.grid(row=r, column=c, sticky=(NSEW))

		return initial_frame


	def displ_search_options(self):
		'''Displays the available search types'''
		search_options = self.handler.search_options # Search types are defined in GUI handler class
		self.search_frame = LabelFrame(self.dropdown_frame, text="Searching Options", padx=10, pady=10)
		self.search_frame.grid(row=0, column=0, sticky=NSEW, padx=10)

		# Loop through all search types to display them with associated radiobutton
		self.selected_search_opt = StringVar()
		for i, option in enumerate(search_options):
			search_btn = Radiobutton(self.search_frame, text=option, value=option, var=self.selected_search_opt)
			search_btn.grid(row=i, column=0, sticky=NW)
			if i == 0: self.selected_search_opt.set(option)


	def displ_radio_btns(self, value_list, choice_list, label, c, all_func):
		'''Given a list of values, creates section of checkbuttons for user selection'''
		option_list = value_list # List of options to display
		radio_frame = LabelFrame(self.dropdown_frame, text=label, padx=10, pady=10)
		radio_frame.grid(row=0, column=c, sticky=NSEW, padx=10)

		# Checkbutton for "Select All" functionality
		all_values = IntVar(value=0)
		select_all_btn = Checkbutton(radio_frame, text="Select All", variable=all_values, onvalue=1, offvalue=0,
										command=all_func)
		select_all_btn.grid(row=0, column=0, sticky=NW)

		# Create checkbutton for each option
		for i, option in enumerate(option_list):
			choice_list[option] = IntVar(value=0)
			btn = Checkbutton(radio_frame, text=option, variable=choice_list[option], onvalue=1, offvalue=0)
			btn.grid(row=i+1, column=0, sticky=NW)


	def select_all_ftypes(self):
		'''Selects all file types when select all button is clicked'''
		for ftype in self.ftype_choices:
			if self.ftype_choices[ftype].get():
				self.ftype_choices[ftype].set(0)
			else:
				self.ftype_choices[ftype].set(1)


	def select_all_components(self):
		'''Selects all component types when select all button is clicked'''
		for component in self.component_choices:
			if self.component_choices[component].get():
				self.component_choices[component].set(0)
			else:
				self.component_choices[component].set(1)


	def create_dir_input(self, label, r, c, def_value):
		'''Creates an input text box given a default value, label, and positions'''
		dir_label = Label(self.directory_frame, text=label, pady=10)
		dir_label.grid(row=r[0], column=c[0])

		dir_input = Entry(self.directory_frame, width=81)
		dir_input.grid(row=r[1], column=c[1], sticky=NSEW, pady=10)
		dir_input.insert(0, def_value)

		return dir_input


	def create_comp_btn(self):
		'''Creates the compute button and checkbutton to include/exclude additional data'''
		self.compute_btn = Button(self.button_frame, text="Compute", command=self.get_selections)
		self.compute_btn.grid(row=0, column=0, padx=(280, 10))

		self.delete_val = IntVar(value=1)
		self.delete_btn = Checkbutton(self.button_frame, text="Keep excess metadata", variable=self.delete_val,
										onvalue=1, offvalue=0)
		self.delete_btn.grid(row=0, column=1)


	def create_summary_outp(self):
		'''Creates text box for the summary output'''
		summary_label = LabelFrame(self.summary_frame, text="Summary of Files", padx=10, pady=10)
		summary_label.grid(row=0, column=0, sticky=NSEW)

		self.summary_txt = Text(summary_label, width=105)
		self.summary_txt.grid(row=0, column=0, sticky=NSEW)


	def get_selections(self):
		'''Validates all user inputs and gets selected options when compute button is clicked'''
		# Validate/set search type selected
		self.handler.set_search_type(self.selected_search_opt.get())
		# Validate/set file and component types selected
		self.handler.set_selections(self.ftype_choices, self.component_choices)
		# Validate/set directories entered
		self.handler.set_directories(self.search_input.get(), self.save_input.get())

		# Automatically display current directory for default search type
		if self.handler.search_type == self.handler.search_options[0]:
			self.search_input.delete(0, END)
			self.search_input.insert(0, self.handler.search_directory)

		# If no file types were selected, select all of them
		if not self.handler.file_types:
			self.handler.file_types = self.handler.file_options
			self.select_all_ftypes()
		# If no component types were selected, select all of them
		if not self.handler.component_types:
			self.handler.component_types = self.handler.component_options
			self.select_all_components()
		# Default if search directory is empty
		if self.search_input.get() == '':
			self.search_input.insert(0, getcwd())
		# Default is save directory is empty
		if self.save_input.get() == '':
			self.save_input.insert(0, getcwd())

		# Call handler to itemize based on selections and display a summary
		self.handler.itemize_documents(self.delete_val.get())
		self.summary_txt.insert('1.0', self.handler.summary)

