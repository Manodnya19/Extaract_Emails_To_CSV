import imaplib
import email
import csv
import os

# Gmail credentials
user = "abc@gmail.com"
password = "abc"

# URL for IMAP connection
imap_url = 'imap.gmail.com'

# Connection with Gmail using SSL
my_mail = imaplib.IMAP4_SSL(imap_url)

# Log in using your credentials
my_mail.login(user, password)

# Select the Inbox to fetch messages
my_mail.select('Inbox')

# Define Key and Value for email search
key = 'SUBJECT'
value = 'Application'  # Search for emails with specific subject value
_, data = my_mail.search(None, key, value)

mail_id_list = data[0].split()  # IDs of all emails that we want to fetch

msgs = []  # empty list to capture all messages

# Iterate through messages and extract data into the msgs list
for num in mail_id_list:
    typ, data = my_mail.fetch(num, '(RFC822)')  # RFC822 returns the whole message
    msgs.append(data)

# Directory to save attachments
attachment_dir = './attachments/'

# Create directory if it doesn't exist
if not os.path.exists(attachment_dir):
    os.makedirs(attachment_dir)

# Open a CSV file to write email data (subject, from, body, attachment filenames)
with open('emails_with_attachments.csv', mode='w', newline='') as file:
    writer = csv.writer(file)
    
    # Write headers to the CSV file
    writer.writerow(['Subject', 'From', 'Body', 'Attachments'])

    # Iterate through fetched messages
    for msg in msgs[::-1]:
        for response_part in msg:
            if isinstance(response_part, tuple):
                my_msg = email.message_from_bytes(response_part[1])
                
                # Extract subject, from, and body text
                subject = my_msg['subject']
                from_ = my_msg['from']
                body = ""
                attachments = []

                # Walk through the message parts to extract body and attachments
                for part in my_msg.walk():
                    # If part is text/plain, it contains the body
                    if part.get_content_type() == 'text/plain':
                        body = part.get_payload(decode=True).decode('utf-8')

                    # If part is an attachment
                    if part.get_content_maintype() == 'multipart':
                        continue

                    if part.get('Content-Disposition') is not None:
                        # Extract the filename
                        filename = part.get_filename()
                        if filename:
                            # Save the attachment to the specified directory
                            filepath = os.path.join(attachment_dir, filename)
                            with open(filepath, 'wb') as f:
                                f.write(part.get_payload(decode=True))
                            attachments.append(filename)

                # Write the email data and attachment names to the CSV
                writer.writerow([subject, from_, body, ', '.join(attachments)])

# Print success message
print(f"Emails and attachments have been saved. Check 'emails_with_attachments.csv' and the 'attachments' directory.")
