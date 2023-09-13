import os

from dotenv import dotenv_values

config = dotenv_values(".env")
print(config["INPUT_FILE_PATH"])
print(config["FILE_NAME"])

dir_path = str(config["INPUT_FILE_PATH"])
file_name = str(config["FILE_NAME"])
print(config)
print("dir_path:", dir_path)
print("file_name:", file_name)


file_path = os.path.join(dir_path, file_name)
print("file_path:", file_path)
