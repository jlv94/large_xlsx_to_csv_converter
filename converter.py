import os
import pandas as pd
import shutil
import colorama
from colorama import Fore, Style, Back

colorama.init(autoreset=True)


def definePath():
	"""This function sets and defines paths of file to convert."""
	
	global tmp_folder, done_folder, csv_folder

	os.makedirs('tmp', exist_ok=True)
	os.makedirs('done', exist_ok=True)
	os.makedirs('csv', exist_ok=True)

	tmp_folder = r'tmp'
	done_folder = r'done'
	csv_folder = r'csv'


def colorPrintText():
	"""This function initializes colors for better guidance in the Command Prompt."""
	
	global color_error, color_input, color_continuing, color_reset
	
	colorama.init(autoreset = True)
	color_error = Fore.WHITE + Back.RED
	color_input = Fore.WHITE + Back.BLUE
	color_continuing = Fore.GREEN
	color_reset = Style.RESET_ALL 


def getFiles():
	"""This function reads the files located in tmp folder and asks user which one to convert."""
	
	print(f'\nList of Excel files to convert in TMP folder:')
	dir_files = [file for file in os.listdir(tmp_folder) if file.endswith('.xlsx')]

	#Exit program if tmp folder is empty of xlsx files
	if len(dir_files) == 0:
		print(f'{color_error}No Excel files to convert found in tmp folder.{color_reset}\nTerminating {'.' *5}')
		exit()

	else:
		dict_index_file = {}

		for index, file in enumerate(dir_files):
			dict_index_file[index] = file
			print(f'{index} - {file}')

		while True:
			index_file = input(f'\n{color_input}Enter the index of the file to convert:{color_reset} ')

			try:
				index_file = int(index_file)

				if index_file in dict_index_file.keys():
					print(f'The index selected is {index_file} - {dict_index_file.get(index_file)}')
					break

				else:
					print(f'{color_error}Invalid input. Please enter a valid index integer within range.{color_reset}')

			except ValueError:
				print(f'{color_error}Invalid input. Please enter a valid index integer within range.{color_reset}')

		file_path = os.path.join(tmp_folder, dict_index_file.get(index_file))
		file_ext = dict_index_file.get(index_file)
		file_name = dict_index_file.get(index_file).replace('.xlsx', '')
		print(f'{color_continuing}Continuing {'.' *5} {color_reset}Selected file: {file_name}')

		return file_path, file_ext, file_name


def readExcelFile(file_path, file_name, start=None, n=0):
	"""This function initializes a list of chunks then reads n rows of Excel files and appends them together."""
	
	#Confirming parameters
	if start is not None and n is not None:
		print(f'\nParameters to read the file are currently setup to:\nStarting line {'.' *5} {start}\nChunk size {'.' *5} {n}')

		while True:
			parameters_confirmation = input(f'\n{color_input}Continue with these parameters [Y] or set new parameters [N]?:{color_reset} ').upper()

			if parameters_confirmation in ['Y', 'N']:

				if parameters_confirmation == 'Y':
					print(f'{color_continuing}Continuing {'.' *5} {color_reset}Script uses current parameters.')
					break

				else:
					start = None
					n = None
					break

			else:
				print(f'{color_error}Invalid input. Please enter [Y] to continue with these parameters or [N] to set new parameters.{color_reset}')

	#Setup start row
	if start is None:
		start_row = input(f'{color_input}Set a starting row:{color_reset} ')
		start_row = int(start_row)

	else:
		start_row = start

	print(f'{color_continuing}Continuing {'.' *5} {color_reset}the file will start being read at line: {start_row}')

	#Setup chunk size
	if n is None:
		n = input(f'{color_input}Set a chunk size:{color_reset} ')
		n = int(n)

	else:
		n = n

	print(f'{color_continuing}Continuing {'.' *5} {color_reset}The file will be appended by chunk size of: {n}')

	#Initialize list of chunks 
	chunks = []

	# Read the first chunk to get the column names
	print(f'\n##### Script starts reading and appending the Excel File: {file_name}')
	
	df = pd.read_excel(file_path, engine='openpyxl', skiprows=start_row, nrows=n)
	chunks.append(df)

	print(f'\t{color_continuing}Continuing {'.' *5} {color_reset}Starting at line {start_row}. First chunk of {n} rows read and appended')

	# Update the starting row for the next chunk
	start_row += n

	# Loop to read the remaining chunks
	while True:
		try:
			chunk = pd.read_excel(file_path, engine='openpyxl', skiprows=start_row, nrows=n, header=None)

			if chunk.empty:
				break

			chunk.columns = df.columns  # Set the column names
			chunks.append(chunk)
			start_row += n
			print(f'\t{color_continuing}Continuing {'.' *5} {color_reset}Next chunk added - {start_row}')

		except Exception as e:
			print(f'\t{color_error}An error occurred: {e}{color_reset}')
			break

	df = pd.concat(chunks, ignore_index=True)
	print(f'\t{color_continuing}Continuing {'.' *5} {color_reset}All chunks have been appended.')
	
	return df

def moveFileTmpToDone(file_path, file_ext):
	"""This function moves the Excel file that has been converted into CSV from the tmp folder into the done folder."""
	
	shutil.move(file_path, os.path.join(done_folder, file_ext))

def main():
	"""This is the main function."""

	definePath()
	colorPrintText()
	
	print('\n########## CONVERTER: LARGE EXCEL FILE TO CSV ##########')
	print('\n\n########## START OF SCRIPT ##########')

	file_path, file_ext, file_name = getFiles()
	df = readExcelFile(file_path, file_name, start=0, n=50000)

	print(f'\nExporting CSV file {'.' *5}')
	df.to_csv(os.path.join(csv_folder, f'{file_name}.csv'), index=False)
	
	print(f'\t{color_continuing}Continuing {'.' *5} {color_reset}The CSV file has been exported and is to be found in folder ./csv/{file_name}.csv')

	print(f'\nMoving Excel file {'.' *5}')
	moveFileTmpToDone(file_path, file_ext)

	print(f'\t{color_continuing}Continuing {'.' *5} {color_reset}The Excel file has been moved from ./tmp to ./done.')
	print('\n\n########## END OF SCRIPT ##########')

	exit()

if __name__ == '__main__':
	main()