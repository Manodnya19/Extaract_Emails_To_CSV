import imaplib
import email
import csv

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
key = 'STBJECT'
value = 'Unfortunately'
_, data = my_mail.search(None, key, value)  # Search for emails with specific key and value

mail_id_list = data[0].split()  # IDs of all emails that we want to fetch 

msgs = []  # empty list to capture all messages

# Iterate through messages and extract data into the msgs list
for num in mail_id_list:
    typ, data = my_mail.fetch(num, '(RFC822)')  # RFC822 returns whole message (BODY fetches just body)
    msgs.append(data)

# Open a CSV file to write data
with open('emails.csv', mode='w', newline='') as file:
    writer = csv.writer(file)
    
    # Write headers to the CSV file
    writer.writerow(['Subject', 'From', 'Body'])

    # Iterate through fetched messages and write them to CSV
    for msg in msgs[::-1]:
        for response_part in msg:
            if isinstance(response_part, tuple):
                my_msg = email.message_from_bytes(response_part[1])
                
                # Extract subject, from, and body text
                subject = my_msg['subject']
                from_ = my_msg['from']
                
                # Extract the body of the email
                body = ""
                for part in my_msg.walk():
                    if part.get_content_type() == 'text/plain':
                        body = part.get_payload()
                
                # Write email data to the CSV file
                writer.writerow([subject, from_, body])

# You can now open 'emails.csv' in Numbers app on your Mac
print("Emails have been saved to emails.csv. You can open this file in the Numbers app.")
