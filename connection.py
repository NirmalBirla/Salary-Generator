import os
from slack_sdk import WebClient
from dotenv import load_dotenv
import ssl

# Load variables from .env
load_dotenv("creds.env")

# Get token from environment variable
token = os.getenv("SLACK_BOT_TOKEN")
channel_id = "C020YQEFC64"

# Initialize Slack client
client = WebClient(token=token, ssl=ssl._create_unverified_context())



# from slack_sdk import WebClient
# import ssl
# # bde_application user OAuth token
# token = "" # Put the slack token in creds.env file.
# channel_id = "C020YQEFC64"

# client = WebClient(token=token,ssl=ssl._create_unverified_context())

# Format for  the token inside the creds.env file.
# SLACK_BOT_TOKEN=wegufidnjskfbweirtgjtfglhabesbnlbfgdvsf
