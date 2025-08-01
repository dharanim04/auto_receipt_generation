import json
import os

base_path = os.path.dirname(os.path.abspath(__file__))
email_details_filepath = os.path.join(base_path, 'email_details.json')

def get_emailaddress():
    try:
        print('filepath of emaildetails:', email_details_filepath)
        with open(email_details_filepath, 'r') as f:
            data = json.load(f)
        return data['EMAIL_ADDRESS']
    except FileNotFoundError:
        print("Error: 'email_details.json' not found.")
    except json.JSONDecodeError:
        print("Error: Could not decode JSON from 'email_details.json'.")
def get_password():
    try:
        with open(email_details_filepath, 'r') as f:
            data = json.load(f)
        return data['APP_PASSWORD']
    except FileNotFoundError:
        print("Error: 'email_details.json' not found.")
    except json.JSONDecodeError:
        print("Error: Could not decode JSON from 'email_details.json'.")

def get_body():
    body_file_path = os.path.join(base_path, 'email_body.txt')
    print('filepath', body_file_path)
    return body_file_path


# if __name__ == "__main__":
#     print(get_emailaddress())
#     print(get_password())
