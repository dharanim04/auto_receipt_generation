import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from supporting_files.helper_file import get_emailaddress, get_password

def send_email_with_attachment(receiver_email,subject, body, attachment_path):
    """
    Sends an email with attachments using an SMTP server.
    """
    print("--- SMTP Email Sender with Attachments ---")
    # Get email details from the user
    sender_email = get_emailaddress()
    sender_password = get_password()
 
    # Get SMTP server details
    smtp_server ='smtp.gmail.com' #input("Enter SMTP server address (e.g., smtp.gmail.com): ")
    smtp_port = 587 #int(input("Enter SMTP server port (e.g., 587 for TLS, 465 for SSL): "))
    use_tls = 'yes' # input("Use TLS encryption? (yes/no): ").lower() == 'yes'

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    # Add attachment if provided
    if attachment_path and os.path.exists(attachment_path):
        try:
            with open(attachment_path, "rb") as attachment:
                # Add file as application/octet-stream
                # Email client can usually download this automatically as attachment
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

            # Encode file in base64
            encoders.encode_base64(part)

            # Add header as key/value pair to attachment part
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {os.path.basename(attachment_path)}",
            )

            # Add attachment to message
            message.attach(part)
            print(f"Attachment '{os.path.basename(attachment_path)}' added.")
        except Exception as e:
            print(f"Error adding attachment: {e}")
            print("Proceeding without attachment.")
    elif attachment_path:
        print(f"Warning: Attachment file not found at '{attachment_path}'. Proceeding without attachment.")

    server = None # Initialize server to None
    try:
        # Establish a secure connection with the SMTP server
        if use_tls:
            print(f"Attempting to connect to {smtp_server}:{smtp_port} with TLS...")
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()  # Upgrade the connection to a secure encrypted SSL/TLS connection
        else:
            print(f"Attempting to connect to {smtp_server}:{smtp_port} with SSL...")
            server = smtplib.SMTP_SSL(smtp_server, smtp_port)

        print("Connection established. Attempting to log in...")
        # Login to the SMTP server
        server.login(sender_email, sender_password)
        print("Logged in to SMTP server.")

        # Send the email
        text = message.as_string()
        server.sendmail(sender_email, receiver_email, text)
        print("Email sent successfully!")

    except smtplib.SMTPAuthenticationError:
        print("\nAuthentication Error: Please check your email and password.")
        print("For Gmail, you might need to use an 'App Password' instead of your regular password.")
    except smtplib.SMTPConnectError as e:
        print(f"\nConnection Error: Could not connect to the SMTP server. Details: {e}")
        print("Possible causes: Incorrect server address or port, firewall blocking, or no internet connection.")
    except smtplib.SMTPServerDisconnected as e:
        print(f"\nServer Disconnected Error: The SMTP server unexpectedly closed the connection. Details: {e}")
        print("This often happens due to incorrect server/port/TLS settings, or a firewall.")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
    finally:
        if server:
            server.quit()
            print("Disconnected from SMTP server.")

