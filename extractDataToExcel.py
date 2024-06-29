import win32com.client
from datetime import datetime, timedelta, date
import re
import win32ui
from openpyxl import load_workbook
import configparser
import tkinter as tk
from tkinter import messagebox
import logging
import codecs
import os

def outlook_is_running():
    try:
        win32ui.FindWindow(None, "Microsoft Outlook")
        return True
    except win32ui.error:
        return False

def connect_to_outlook():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    return outlook

def search_emails_in_folder(folder, content, start_date):
    items = folder.Items
    # filter_criteria = f"@SQL=(urn:schemas:httpmail:datereceived >= '{start_date}') AND (urn:schemas:httpmail:textdescription like '%{content}%')"
    filter_criteria = f"@SQL=(urn:schemas:httpmail:datereceived >= '{start_date}') AND (urn:schemas:httpmail:subject like '%{content}%')"
    filtered_emails = items.Restrict(filter_criteria)
    return filtered_emails

def extract_data(text, project_pattern, amount_pattern, payee_pattern):
    projects = re.findall(project_pattern, text)
    amounts = re.findall(amount_pattern, text)
    payees = re.findall(payee_pattern, text)
    result = []
    for i in range(len(projects)):
        result.append({
            "projectName": projects[i].strip(),
            "amount": amounts[i].strip(),
            "payee": payees[i].strip()
        })
    return result

def read_config():
    config_path = 'config.ini'
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Configuration file {config_path} not found. Please make sure it is in the same directory as the program.")
        
    
    config = configparser.ConfigParser()
    with codecs.open(config_path, 'r', encoding='utf-8') as f:
        config.read_file(f)
    return (
        config['DEFAULT']['project'],
        config['DEFAULT']['amount'],
        config['DEFAULT']['payee']
    )

def main(content_to_search, days):
    logging.basicConfig(filename='email_processing.log', level=logging.INFO, format='%(asctime)s - %(message)s', encoding='utf-8')
    logging.info("Starting email processing")

    try:
        if not outlook_is_running():
            raise RuntimeError("Please launch Outlook!")

        project_pattern, amount_pattern, payee_pattern = read_config()
        
        outlook = connect_to_outlook()
        inbox = outlook.GetDefaultFolder(6)  # 6 corresponds to inbox
        start_date = (datetime.now() - timedelta(days=days)).strftime('%m/%d/%Y %H:%M %p')
        emails = search_emails_in_folder(inbox, content_to_search, start_date)
        today = date.today()
        all_data = []
        if len(emails) == 0:
            logging.info("Can not find the email!")
        for email in emails:
            logging.info(f"Email Subject: {email.Subject}")
            email_body = email.Body.encode('utf-8', errors='ignore').decode('utf-8')
            data = extract_data(email_body, project_pattern, amount_pattern, payee_pattern)
            all_data.extend(data)

        template_path = 'template.xlsx'
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file {template_path} not found. Please make sure it is in the same directory as the program.")
            

        for item in all_data:
            workbook = load_workbook(template_path)
            worksheet = workbook.active
            worksheet['F6'] = f"{today.year}年{today.month}月{today.day}日"
            worksheet['A9'] = f"{today.year}年{today.month}月{today.day}日"
            worksheet['B9'] = item['projectName']
            worksheet['F9'] = item['amount']
            worksheet['B13'] = item['payee']
            file_name = f"{item['payee']}.xlsx"
            workbook.save(file_name)

        logging.info(f"Processed data: {all_data}")
        messagebox.showinfo("Success", "Email processing completed successfully!")
    except FileNotFoundError as e:
        error_message = str(e)
        logging.error(error_message)
        messagebox.showerror("File Not Found", error_message)
    except Exception as e:
        error_message = f"An error occurred: {str(e)}"
        logging.error(error_message)
        messagebox.showerror("Error", error_message)
    finally:
        logging.info("Email processing finished")

def create_gui():
    def on_submit():
        subject = subject_entry.get()
        days = int(days_entry.get())
        root.destroy()
        main(subject, days)

    root = tk.Tk()
    root.title("Extract Data From Email")
    
    window_width = 400
    window_height = 150
    
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    x = (screen_width/2) - (window_width/2)
    y = (screen_height/2) - (window_height/2)
    
    root.geometry('%dx%d+%d+%d' % (window_width, window_height, x, y))
    
    root.resizable(False, False)

    tk.Label(root, text="Content to search:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
    subject_entry = tk.Entry(root, width=40)
    subject_entry.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(root, text="Days to look back:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
    days_entry = tk.Entry(root, width=10)
    days_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    submit_button = tk.Button(root, text="Submit", command=on_submit)
    submit_button.grid(row=2, column=0, columnspan=2, pady=10)

    developer_label = tk.Label(root, text="Developed by Raven", font=("Arial", 8))
    developer_label.place(relx=1.0, rely=1.0, x=-5, y=-5, anchor="se")

    root.mainloop()

if __name__ == "__main__":
    create_gui()
