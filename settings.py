
import os

from dotenv import load_dotenv

# Get the path to he directory this file is in
BASEDIR = os.path.abspath(os.path.dirname(__file__))
CONFIG_FILE = 'config.env'
# Load environment variables
load_dotenv(os.path.join(BASEDIR, CONFIG_FILE))

def get_input_file_path():
    dir_path = os.getenv('INPUT_FILE_PATH')
    file_name = os.getenv('FILE_NAME')
    return os.path.join(dir_path, file_name)
