import glob, os, re, sys, datetime, argparse
import pandas as pd
from nltk import word_tokenize
from collections import defaultdict

def masterchecklist_to_dictionary(uniqueID_column, uniqueDescription_column):
	"""
	TODO: DOCSTRING
	"""
	uniqueID2description = defaultdict(str)	

	for key, value in zip(uniqueID_column, uniqueDescription_column):
		key, value = str(key), str(value)
		value = re.sub("[-,.]", "", value)

		if key in uniqueID2description: # Need to add whitespace before concatting 
			value = " " + value

		uniqueID2description[key] += value

	return uniqueID2description

def original_checklist_to_dictionary(originalID_column, originalDescription_column):
	"""
	TODO: DOCSTRING
	"""
	originalID2descriptionDict = defaultdict(str)
	
	_previous_row_item = None
	for key, value in zip(originalID_column, originalDescription_column):
		key, value = str(key), str(value)
		if (key == "nan") and (value.isupper() == False):
			key = _previous_row_item
			_previous_row_item = _previous_row_item
		else:
			_previous_row_item = key

		value = re.sub("[-,.]", "", value) 
		if key in originalID2descriptionDict: # Need to add whitespace before concatenation 		
			value = " " + value

		originalID2descriptionDict[key] += value

	return originalID2descriptionDict

def load_masterchecklist(path_to_masterchecklist):
	"""
	Loads the HGB masterchecklist from given path
	Must be excel file (usually with extions xlsm)
	returns: dictionary mapping with (key = "Description", value = "Unique ID") or None if error
	"""

	try:
		df = pd.read_excel(path_to_masterchecklist, header=None, skiprows=1)
		uniqueID_column = df.iloc[:,0]
		uniqueDescription_column = df.iloc[:,2]
		uniqueID2description = masterchecklist_to_dictionary(uniqueID_column, uniqueDescription_column)

		return uniqueID2description

	except Exception as e:
		print("Error loading masterchecklist from path: {} -->: {}}".format(path_to_masterchecklist, e))
		return None

def load_original_checklist(path_to_checklist, sheetname = "Ergebnis"):
	"""
	Loads the original xlsm checklist from given path
	File is read by pandas
	$path_to_checklist holds path to the checklist
	$sheetname is by default the tab with the name "Ergebnis"
	returns: dictionary mapping with (key = "original ID", value = "Description") or None if error
	"""

	if path_to_checklist.split(".")[-1:][0] != "xlsm":
		print("Wrong file extension. File extension is {} but should be xlsm".format(path_to_checklist.split(".")[-1:][0]))
		return None

	try:
		df = pd.read_excel(path_to_checklist, header=None, sheet_name=sheetname, skiprows=1)
		originalID_column = df.iloc[:,0]
		originalDescription_column = df.iloc[:,1]
		originalID2descriptionDict = original_checklist_to_dictionary(originalID_column, originalDescription_column)

		return originalID2descriptionDict

	except Exception as e:
		print("Error loading original checklist from path: {} -->: {}".format(path_to_checklist, e))
		return None
		
def load_original_annotation(path_to_annotation):
	"""
	Usually one sheet in excel file but for safety we use the first sheet
	"""

	try: 
		return pd.read_excel(path_to_annotation, header = None, sheet_name=0, skiprows=1)
	except Exception as e:
		print("Error loading annotation file: {}".format(e))
		return None

def map_originalID_to_uniqueID(uniqueID2description, originalID2description):
	"""
	Maps the original checklist ID to the unique ID of the masterchecklist
	The function compares the description of the two checklists and returns the respetive IDs

	$uniqueID2description is a dictionary with (key = description, value = unique ID) of the masterchecklist
	$originalID2description is a dictionary with (key = originalID, value = description) of the original checklist

	returns: dictionary with mapping (key = "originalID", value = "uniqueID")
	"""

	originalID2uniqueID = defaultdict(str)
	for uniqueID, description_uniq in uniqueID2description.items():
		for originalID, description_orig in originalID2description.items():
			if description_uniq == description_orig:
				originalID2uniqueID[originalID] = uniqueID

	return originalID2uniqueID

def map_original_annotation_to_uniqueID(originalID2uniqueID, path_to_original_annotation):
	"""
	Maps the original annotation to unique ID to have a common class system
	
	$originalID2uniqueID: dictionary mapping of originalID (key) to uniqueID (value)
	$original_annotation_dataframe: excel (preferred) or csv file of original annotation
	
	returns: a dataframe with blob, annotation_original and annotation_unique
	"""

	df = load_original_annotation(path_to_original_annotation)
	blob_column = df.iloc[:,0]
	annotation_column = df.iloc[:,1]

	uniqueID_column = []
	uniqueID_extra_column = []
	for blob, annotation in zip(blob_column, annotation_column):
		# Split on comma because a cell can contain multiple requirement classes
		blog, annotation = str(blob), str(annotation) # Convert to string because types are not always similar in excel file
		list_of_annotations = annotation.split(",") # If cell contains multiple annotations split them (separated by comma)
		
		list_of_unique_annotations = []
		list_of_unique_annotations_extra = [] # unique ID with the additional information 
		for annotation in list_of_annotations:
			annotation_clean = annotation.split(".")[0].replace(" ", "")  # Strip away the additional information. Otherwise we cannot match with masterchecklist
			additional_information = re.findall("\.\d{1,3}", annotation) # However, store this information because it will be appended to unique ID later
			uniqueID = originalID2uniqueID.get(annotation_clean)

			if uniqueID != None and additional_information != []:
				unique_id_extra = uniqueID + str(additional_information[0]) # Concat the additional information to unique ID Extra
				list_of_unique_annotations_extra.append(unique_id_extra)
			else:
				list_of_unique_annotations_extra.append(uniqueID)

			list_of_unique_annotations.append(uniqueID)


		uniqueID_original_naming_convention = ', '.join(str(item) for item in list_of_unique_annotations) # Translates to same naming convention as original annotation
		uniqueID_column.append(uniqueID_original_naming_convention) 

		uniqueID_extra_original_naming_convention = ', '.join(str(item) for item in list_of_unique_annotations_extra) # Translates to same naming convention as original annotation
		uniqueID_extra_column.append(uniqueID_extra_original_naming_convention)

	return pd.DataFrame({
		'blob': blob_column,
		'annotation_original': annotation_column,
		'annotation_unique': uniqueID_column,
		'annotation_unique_extra': uniqueID_extra_column
		})

def get_day_month_year():
	"""
	"""

	now = datetime.datetime.now()

	day = str(now.day)
	month = str(now.month)
	year = str(now.year)

	return "_".join((day, month, year))

def save_result_to_new_excel_file(blob_annot_orig_annot_unique_dataframe, PATH_ORIGINAL_ANNOTATION):
	"""
	TODO: DOCSTRING
	"""
	try:
		cur_time = get_day_month_year()
		uniqueID_basename = "_".join(("uniqueID", cur_time, os.path.basename(PATH_ORIGINAL_ANNOTATION)))
		uniqueID_basename_no_whitespace = uniqueID_basename.replace(" ", "_")
		dirname = os.path.dirname(PATH_ORIGINAL_ANNOTATION)
		uniqueID_path = "/".join((dirname, uniqueID_basename_no_whitespace))
		writer = pd.ExcelWriter(uniqueID_path)
		blob_annot_orig_annot_unique_dataframe.to_excel(writer, 'output')
		writer.save()
		print("Saved unique IDs to: {}\n".format(uniqueID_path))
	
	except Exception as e:
		print("Error writing results to excel for company {} with error \n{}".format(PATH_ORIGINAL_ANNOTATION, e))

def main(PATH_MASTER_CHECKLIST, PATH_ORIGINAL_CHECKLIST, PATH_ANNOTATION):
	"""
	Combines all functions into main saves result as excel into respective folder
	"""
	uniqueID2description = load_masterchecklist(PATH_MASTER_CHECKLIST)
	originalID2descriptionDict = load_original_checklist(PATH_ORIGINAL_CHECKLIST)
	originalID2uniqueID = map_originalID_to_uniqueID(uniqueID2description, originalID2descriptionDict)
	blob_annot_orig_annot_unique_dataframe = map_original_annotation_to_uniqueID(originalID2uniqueID, PATH_ANNOTATION)

	# Save to new excel file
	save_result_to_new_excel_file(blob_annot_orig_annot_unique_dataframe, PATH_ANNOTATION)


if __name__ == "__main__":


	PATH_MASTER_CHECKLIST = "./Masterliste.xlsx"
	root_path = "./data/"

	#PATH_ORIGINAL_CHECKLIST = "./data/Takata GmbH/Anhangcheckliste HGB_BilRuG_Takata GmbH.xlsm"
	#PATH_ORIGINAL_ANNOTATION = "./data/Takata GmbH/Takata GmbH_Konzernabschluss 2016_2017.xlsx"

	for i, company_folder in enumerate(os.listdir(root_path)):
		print("="*80)
		print([i], company_folder)

		PATH_ORIGINAL_CHECKLIST = None
		PATH_ANNOTATION_LIST = []
		for file in os.listdir(root_path + company_folder):
			basename = os.path.basename(file)

			# Deselect all which are not xlsx (e.g. the csv and xlsm file) and where the uniqueID pattern is not part of the filename
			valid_annotation_file_extension = ".xlsx"
			pattern_not_allowed = "uniqueID"
			if (basename.endswith(valid_annotation_file_extension)) and (pattern_not_allowed not in basename):
				PATH_ANNOTATION_LIST.append(file)

			valid_checklist_file_extension = ".xlsm"
			if basename.endswith(valid_checklist_file_extension):
				PATH_ORIGINAL_CHECKLIST = file

		number_of_files = len(PATH_ANNOTATION_LIST)
		if number_of_files == 1 and PATH_ORIGINAL_CHECKLIST != None:
			PATH_ANNOTATION = "".join((root_path + company_folder + "/" + PATH_ANNOTATION_LIST[0]))
			PATH_ORIGINAL_CHECKLIST = "".join((root_path + company_folder + "/" + PATH_ORIGINAL_CHECKLIST))

			print("Annotation: {}\nChecklist: {}\nMasterCheck: {}".format(
							PATH_ANNOTATION, 
							PATH_ORIGINAL_CHECKLIST, 
							PATH_MASTER_CHECKLIST))
			main(PATH_MASTER_CHECKLIST, PATH_ORIGINAL_CHECKLIST, PATH_ANNOTATION)

		else:
			print("{} files with xlsx format found - cannot proceed with {}".format(number_of_files, company_folder))
