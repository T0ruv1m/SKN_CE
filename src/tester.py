from config_tools import DIR, EmptyFileMaker

path_to = DIR()
create_ExtractedData = EmptyFileMaker(path_to.extracted_data).empty_file()
create_Timestamp = EmptyFileMaker(path_to.timestamp).empty_file()
create_json = EmptyFileMaker(path_to.json_object).empty_file()


