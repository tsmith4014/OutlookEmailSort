# pst is the format we want to work with because it contains attachments.


# this works to print the oldest email in the .pst file
import os
import pypff
import re
import datetime

def parse_email_headers(headers):
    from_pattern = r"From: (.+)"
    to_pattern = r"To: (.+)"
    
    from_address = re.search(from_pattern, headers)
    to_address = re.search(to_pattern, headers)
    
    return from_address.group(1) if from_address else None, to_address.group(1) if to_address else None

def print_email_details(email):
    subject = email.get_subject()
    body = email.get_plain_text_body()
    headers = email.get_transport_headers()

    from_address, to_address = parse_email_headers(headers)
    
    decoded_body = body.decode('utf-8') if body else ''
    first_20_words = " ".join(decoded_body.split()[:20])

    sent_date = email.get_delivery_time()

    print("From:", from_address if from_address else 'N/A')
    print("To:", to_address if to_address else 'N/A')
    print("Subject:", subject if subject else 'N/A')
    print("Sent Date:", sent_date if sent_date else 'N/A')
    print("Body (first 20 words):", first_20_words if first_20_words else 'N/A')
    print("\n")

def get_all_messages(folder, folder_path=''):
    messages = []
    current_folder_path = folder_path + '/' + folder.name
    
    for message in folder.sub_messages:
        messages.append(message)

    for subfolder in folder.sub_folders:
        messages.extend(get_all_messages(subfolder, current_folder_path))

    return messages

file_path = os.path.expanduser("/Users/asmith/Desktop/new_folder/backup.pst")
# file_path = os.path.expanduser("/Users/chadthompsonsmith/Desktop/Projects/amy_projects/OutlookEmailSort/Data_Sort/Inbox/backup.pst")
file_path = os.path.expanduser("/Users/chadthompsonsmith/Desktop/Projects/amy_projects/OutlookEmailSort/Data_Sort/Inbox/oldbackup.pst")

# Open the PST file
pst_file = pypff.file()
pst_file.open(file_path)

# Get the root folder and the inbox folder
root_folder = pst_file.get_root_folder()
inbox_folder = None

for folder in root_folder.sub_folders:
    if folder.name == "Top of Outlook data file":
        inbox_folder = folder
        break

if not inbox_folder:
    print("Inbox folder not found")
    exit()

# Get all the messages in the inbox folder and its subfolders
messages = get_all_messages(inbox_folder)

# Find the oldest email by sorting messages by sent date
oldest_email = sorted(messages, key=lambda x: x.get_delivery_time())[0]

# Print the oldest email
print("Oldest email:")
print_email_details(oldest_email)

# Close the PST file
pst_file.close()




#this works to read the first 5 and last 5 emails in the .pst file
# import os
# import pypff
# import re
# import datetime

# def parse_email_headers(headers):
#     from_pattern = r"From: (.+)"
#     to_pattern = r"To: (.+)"
    
#     from_address = re.search(from_pattern, headers)
#     to_address = re.search(to_pattern, headers)
    
#     return from_address.group(1) if from_address else None, to_address.group(1) if to_address else None

# def print_email_details(email):
#     subject = email.get_subject()
#     body = email.get_plain_text_body()
#     headers = email.get_transport_headers()

#     from_address, to_address = parse_email_headers(headers)
    
#     decoded_body = body.decode('utf-8') if body else ''
#     first_20_words = " ".join(decoded_body.split()[:20])

#     sent_date = email.get_delivery_time()

#     print("From:", from_address if from_address else 'N/A')
#     print("To:", to_address if to_address else 'N/A')
#     print("Subject:", subject if subject else 'N/A')
#     print("Sent Date:", sent_date if sent_date else 'N/A')
#     print("Body (first 20 words):", first_20_words if first_20_words else 'N/A')
#     print("\n")

# def get_all_messages(folder, folder_path=''):
#     messages = []
#     current_folder_path = folder_path + '/' + folder.name
#     print("Current folder:", current_folder_path)
    
#     for message in folder.sub_messages:
#         messages.append(message)

#     for subfolder in folder.sub_folders:
#         messages.extend(get_all_messages(subfolder, current_folder_path))

#     return messages

# file_path = os.path.expanduser("~/Library/Mobile Documents/com~apple~CloudDocs/Desktop/Data_Sort/Inbox/backup.pst")

# # Open the PST file
# pst_file = pypff.file()
# pst_file.open(file_path)

# # Get the root folder and the inbox folder
# root_folder = pst_file.get_root_folder()
# inbox_folder = None

# for folder in root_folder.sub_folders:
#     if folder.name == "Top of Outlook data file":
#         inbox_folder = folder
#         break

# if not inbox_folder:
#     print("Inbox folder not found")
#     exit()

# # Get all the messages in the inbox folder and its subfolders
# messages = get_all_messages(inbox_folder)

# # Print the first 25 emails
# for msg in messages[:25]:
#     print_email_details(msg)

# # Print the last 25 emails
# for msg in messages[-25:]:
#     print_email_details(msg)

# # Close the PST file
# pst_file.close()




# # working code much like above also prints head of the .pst file
# import os
# import pypff
# import re

# def parse_email_headers(headers):
#     from_pattern = r"From: (.+)"
#     to_pattern = r"To: (.+)"
    
#     from_address = re.search(from_pattern, headers)
#     to_address = re.search(to_pattern, headers)
    
#     return from_address.group(1) if from_address else None, to_address.group(1) if to_address else None

# def print_email_details(email):
#     subject = email.get_subject()
#     body = email.get_plain_text_body()
#     headers = email.get_transport_headers()
    
#     from_address, to_address = parse_email_headers(headers)
    
#     decoded_body = body.decode('utf-8')
#     first_20_words = " ".join(decoded_body.split()[:20])

#     print("From:", from_address)
#     print("To:", to_address)
#     print("Subject:", subject)
#     print("Body (first 20 words):", first_20_words)
#     print("\n")


# def get_all_messages(folder, folder_path=''):
#     messages = []
#     current_folder_path = folder_path + '/' + folder.name
#     print("Current folder:", current_folder_path)
    
#     for message in folder.sub_messages:
#         messages.append(message)

#     for subfolder in folder.sub_folders:
#         messages.extend(get_all_messages(subfolder, current_folder_path))

#     return messages

# file_path = os.path.expanduser("~/Library/Mobile Documents/com~apple~CloudDocs/Desktop/Data_Sort/Inbox/backup.pst")

# # Open the PST file
# pst_file = pypff.file()
# pst_file.open(file_path)

# # Get the root folder and the inbox folder
# root_folder = pst_file.get_root_folder()
# inbox_folder = None

# for folder in root_folder.sub_folders:
#     if folder.name == "Top of Outlook data file":
#         inbox_folder = folder
#         break

# if not inbox_folder:
#     print("Inbox folder not found")
#     exit()

# # Get all the messages in the inbox folder and its subfolders
# messages = get_all_messages(inbox_folder)

# # Print the first 25 emails
# for msg in messages[:25]:
#     print_email_details(msg)

# # Print the last 25 emails
# for msg in messages[-25:]:
#     print_email_details(msg)

# # Close the PST file
# pst_file.close()


