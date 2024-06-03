# Request API permissions > Application permissions
# Grant admin consent to API/Permissions
import msal
import requests
import base64
from datetime import datetime

def request_headers (tenant_id, client_id, client_secret):
    authority_url = f'https://login.microsoftonline.com/{tenant_id}'
    app = msal.ConfidentialClientApplication(
        authority = authority_url,
        client_id = client_id,
        client_credential = client_secret
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    access_token = token["access_token"]
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    return headers

def fetch_messages(user_id, today_iso, key_word, headers):
    url = f'https://graph.microsoft.com/v1.0/users/{user_id}/mailFolders/inbox/messages?$filter=receivedDateTime ge {today_iso}T00:00:00Z and contains(subject, \'{key_word}\') and isRead eq false'
    response = requests.get(url, headers=headers)
    if response.status_code!= 200:
        print(f"Failed to fetch messages: {response.text}")
        exit(1)
    messages = response.json().get('value', [])
    return messages

def mark_message_as_read(message_id, user_id, headers):
    url = f'https://graph.microsoft.com/v1.0/users/{user_id}/messages/{message_id}'
    payload = {
        "isRead": True
    }
    response = requests.patch(url, json=payload, headers=headers)
    if response.status_code!= 200:
        print(f"Failed to mark message as read: {response.text}")
        return False
    else:
        print("Message marked as read successfully.")
        return True

def reply_email(user_id, message_id, attachment, comment):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{message_id}/replyALL"
    payload = {
        "message": {
        "attachments": [attachment]
        },
    "comment": f"{comment}"
    }
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code == 202:
        print("Mail sent successfully.")
    else:
        print(f"Failed to send mail. Status code: {response.status_code}, Message: {response.text}")
    

today_iso = datetime.now().date().isoformat()
key_word = ""

# Replace 'user_id' with the actual user ID or userPrincipalName
user_id = ''

# Find items in Microsoft Entra admin center > Appregistrations
client_id = ""
client_secret = ""
tenant_id = ""

fetch_attachments = f'https://graph.microsoft.com/v1.0/users/{user_id}/messages/'

headers = request_headers(tenant_id, client_id, client_secret)
messages = fetch_messages(user_id, today_iso, key_word, headers)

# Process the messages
for message in messages:
    message_id = message['id']
    has_attachments = message['hasAttachments']
    mark_message_as_read(message_id, user_id, headers)
 
    # Check if the message has attachments
    if has_attachments:
        # Fetch the attachments
        attachments_endpoint = f'{fetch_attachments}/{message_id}/attachments'
        attachments_response = requests.get(attachments_endpoint, headers=headers)
        
        if attachments_response.status_code!= 200:
            print(f"Failed to fetch attachments for message {message_id}: {attachments_response.text}")
            continue
        
        attachments = attachments_response.json().get('value', [])
        for attachment in attachments:
            attachment_id = attachment['id']
            attachment_name = attachment['name']
            # print(f"Found attachment with ID: {attachment_id} and name: {attachment_name}")
            if ".xlsx" in attachment_name:
            # Download the attachment
                resource = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{message_id}/attachments/{attachment_id}/$value"
                results = requests.get(resource, headers=headers, stream=True)
                results.content
                # print(details)
                # Save the attachment to a file
                with open(f"./data/{attachment_name}", "wb") as f:
                    f.write(results.content)


                #RPA process

                attachment_path = f'./data/{attachment_name}'

                with open(attachment_path, 'rb') as file:
                    file_content = file.read()
                file_content_base64 =  base64.b64encode(file_content).decode('utf-8')
                
                attachment_to_reply = {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": 'reply ' + attachment_path.split('/')[-1],
                    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "contentBytes": file_content_base64
                }

                comment = ""

                reply_email(user_id, message_id, attachment_to_reply, comment)


