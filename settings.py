import os

from dotenv import dotenv_values

# Get the path to he directory this file is in
# BASEDIR = os.path.abspath(os.path.dirname(__file__))
# CONFIG_FILE = ".env"
# Load environment variables
# load_dotenv(os.path.join(BASEDIR, CONFIG_FILE))
config = dotenv_values(".env")


def get_input_file_path():
    dir_path = config["INPUT_FILE_PATH"]
    file_name = config["FILE_NAME"]
    return os.path.join(dir_path, file_name)
