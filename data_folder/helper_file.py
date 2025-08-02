import json
import os
import logging
import re

base_path = os.path.dirname(os.path.abspath(__file__))
email_details_filepath = os.path.join(base_path, 'email_details.json')

def get_emailaddress():
    try:
        with open(email_details_filepath, 'r') as f:
            data = json.load(f)
        return data['EMAIL_ADDRESS']
    except FileNotFoundError:
        logging.error("Error: 'email_details.json' not found.")
    except json.JSONDecodeError:
        logging.error("Error: Could not decode JSON from 'email_details.json'.")

def get_password():
    try:
        with open(email_details_filepath, 'r') as f:
            data = json.load(f)
        return data['APP_PASSWORD']
    except FileNotFoundError:
        logging.error("Error: 'email_details.json' not found.")
    except json.JSONDecodeError:
        logging.error("Error: Could not decode JSON from 'email_details.json'.")

def get_body():
    body_file_path = os.path.join(base_path, 'email_body.txt')
    return body_file_path


def is_valid_email_syntax(email):
    pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    return re.match(pattern, email) is not None

# if __name__ == "__main__":
#     print(get_emailaddress())
#     print(get_password())
