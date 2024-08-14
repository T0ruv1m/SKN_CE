from config_tools import DIR, EmptyFileMaker

path_to = DIR()
create_ExtractedData = EmptyFileMaker(path_to.xl_file)
create_Timestamp = EmptyFileMaker(path_to.timestamp)
create_json = EmptyFileMaker(path_to.json_file)

create_ExtractedData.empty_file()
create_Timestamp.empty_file()
create_json.empty_file()

