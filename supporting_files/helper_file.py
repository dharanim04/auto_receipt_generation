import json


filename ='supporting_files/email_details.json'

def get_emailaddress():
    try:
        with open(filename, 'r') as f:
            data = json.load(f)
        return data['EMAIL_ADDRESS']
    except FileNotFoundError:
        print("Error: 'email_details.json' not found.")
    except json.JSONDecodeError:
        print("Error: Could not decode JSON from 'email_details.json'.")
def get_password():
    try:
        with open(filename, 'r') as f:
            data = json.load(f)
        return data['APP_PASSWORD']
    except FileNotFoundError:
        print("Error: 'email_details.json' not found.")
    except json.JSONDecodeError:
        print("Error: Could not decode JSON from 'email_details.json'.")

# if __name__ == "__main__":
#     print(get_emailaddress())
#     print(get_password())
