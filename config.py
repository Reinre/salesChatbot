
import os
from dotenv import load_dotenv

basedir = os.path.abspath(os.path.dirname(__file__))
load_dotenv(os.path.join(basedir, '.env'))

API_KEY_FILENAME = os.environ.get('API_KEY_FILENAME')
SHEET_NAME = os.environ.get('SHEET_NAME')
CHANNEL_ACCESS_TOKEN = os.environ.get('CHANNEL_ACCESS_TOKEN')
CHANNEL_SECRET = os.environ.get('CHANNEL_SECRET')
PORT = os.environ.get('PORT', 5000)