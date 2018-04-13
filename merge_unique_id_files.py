import os, argparse
import pandas as pd

ROOT_PATH = "./data/"
UNIQUE_ID_DATE = "5_4_2018"

def retrieve_uniqueID_files(company_folder_path, date):
	PATTERN_MUST_MATCH = "uniqueID_" + date

	for file in os.listdir(company_folder_path):
		if PATTERN_MUST_MATCH in file:
			return os.path.join(company_folder_path, file)

def loop_through_data_root_folder(ROOT_PATH, UNIQUE_ID_DATE):
	paths_to_uniqueID_list = []
	for company_folder in os.listdir(ROOT_PATH):
		company_folder_path = os.path.join(ROOT_PATH, company_folder)

		path_to_uniqueID = retrieve_uniqueID_files(company_folder_path, UNIQUE_ID_DATE)
		if path_to_uniqueID != None:
			paths_to_uniqueID_list.append(path_to_uniqueID)

	return paths_to_uniqueID_list


paths_to_uniqueID_list = loop_through_data_root_folder(ROOT_PATH, UNIQUE_ID_DATE)
master_dataframe = pd.concat([pd.read_excel(path) for path in paths_to_uniqueID_list], ignore_index=True)
master_dataframe.to_csv("annotated_data_uniqueID_all_"+UNIQUE_ID_DATE+".csv")