import win32com.client
import datetime
import numpy as np
import re

def connect_outlook():
    # Create an Outlook application object
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    return outlook_app

def find_folders(inbox, parent_folder_name, sub_parent_folder_name, subfolder_name):
    # We will be moving emails to a few different folders, but always at the subfolder level
    parent_folder = None
    sub_parent_folder = None
    subfolder = None
    for folder in inbox.Folders:
        if folder.Name == parent_folder_name:
            parent_folder = folder
            break

    if parent_folder:
        for folder in parent_folder.Folders:
            if folder.Name == sub_parent_folder_name:
                sub_parent_folder = folder
                break

    if sub_parent_folder:
        for folder in sub_parent_folder.Folders:
            if folder.Name == subfolder_name:
                subfolder = folder
                break
    
    return subfolder

def to_holding_box(outlook_app, email, carrier):
    # We specify the names of the folder we want to navigate through
    # and use the find_folders function to send the email to the holding box.
    # The holding box is for emails pertaining to lasers not ready to have a 
    # tracking email sent. This could be due to incoterms, project status, or
    # that we have not yet received a tracking number from the carrier
    namespace = outlook_app.GetNamespace("MAPI")
    email_account = "PO_submittal@us.TRUMPF.com"
    inbox = namespace.GetSharedDefaultFolder(namespace.CreateRecipient(email_account), 6)
    if carrier == []:
        return
    elif carrier == "UPS":
        parent_folder_name = "Freight Tracking"
        sub_parent_folder_name = "UPS-Air"
        subfolder_name = "Holding box"
    elif carrier == "FedEx":
        parent_folder_name = "Freight Tracking"
        sub_parent_folder_name = "Fedex - Air"
        subfolder_name = "Holding box"
    subfolder = find_folders(inbox, parent_folder_name, sub_parent_folder_name, subfolder_name)
    if subfolder:
        email.Move(subfolder)
    
    return

def to_sharepoint(outlook_app, email, carrier):
    # We do the same thing here, but with the to_sharepoint folder. We move emails
    # here after a tracking email has been sent for the laser to which the email
    # pertains.
    namespace = outlook_app.GetNamespace("MAPI")
    email_account = "PO_submittal@us.TRUMPF.com"
    inbox = namespace.GetSharedDefaultFolder(namespace.CreateRecipient(email_account), 6)
    if carrier == []:
        return
    elif carrier == "UPS":
        parent_folder_name = "Freight Tracking"
        sub_parent_folder_name = "UPS-Air"
        subfolder_name = "In Sharepoint"
    elif carrier == "FedEx":
        parent_folder_name = "Freight Tracking"
        sub_parent_folder_name = "Fedex - Air"
        subfolder_name = "In Sharepoint"
    subfolder = find_folders(inbox, parent_folder_name, sub_parent_folder_name, subfolder_name)
    if subfolder:
        email.Move(subfolder)
    
    return

def get_emails(outlook_app):
    namespace = outlook_app.GetNamespace("MAPI")

    # Get Inbox folder for purchase order submittals
    email_account = "PO_submittal@us.TRUMPF.com"
    inbox = namespace.GetSharedDefaultFolder(namespace.CreateRecipient(email_account), 6)

    # Create time restriction
    current_date = datetime.datetime.now()
    end_time = datetime.datetime(current_date.year, current_date.month, current_date.day, 23, 59, 59)
    start_time = end_time - datetime.timedelta(days=1)
    time_restriction = "[ReceivedTime] >= '" + start_time.strftime('%m/%d/%Y %H:%M %p') + "' AND [ReceivedTime] <= '" + end_time.strftime('%m/%d/%Y %H:%M %p') + "'"
    
    # Get emails from main inbox within time restriction
    inbox_emails = inbox.Items.Restrict(time_restriction)

    # Get emails from holding boxes without time restriction
    subfolder = find_folders(inbox, "Freight Tracking", "UPS-Air", "Holding Box")
    if subfolder:
        ups_holding = subfolder.Items
    subfolder = find_folders(inbox, "Frieght Tracking", "Fedex - Air", "Holding Box")
    if subfolder:
        fedex_holding = subfolder.Items
    
    # Concatonate emails into one list
    emails = list(inbox_emails) + list(ups_holding) + list(fedex_holding)

    return emails

def draft_email(outlook_app, po, so, equipment_number, recipients, cc_recipients, customer_name, link, carrier):
    # We set the path to our external signature, so it can be added at the end of the email body
    signature_path = "C:\Users\crowleypa\AppData\Roaming\Microsoft\Signatures\extern (english).htm"
    # We create a descriptive subject line
    subject = "Shipment notice for PO # " + po + "; Machine S/N # " + equipment_number + "; SO # " + so
    # Set display text for the hyperlink
    display_text = ("Tracking | " 
                f"{carrier}"
                " - United States"
    )
    # Generate email body
    body = (
        f"Dear {customer_name},<br><br>"
        "Please be advised of the shipment of your machine order. "
        "Please see the link to tracking below. As a reminder, the ETA is subject to change due to customs and other shipping events that may occur. "
        "TRUMPF strongly recommends checking back for updates to the anticipated delivery estimate regularly.<br><br>"
        f"<a href='{link}'>{display_text}</a><br><br>"
        "If you have any further questions, please feel free to reach out at any time."
    )

    # Now we create the email and incorporate the details defined above
    draft = outlook_app.CreateItem(0)
    draft.Subject = subject
    font_style = '<span style="font-family: Arial; font-size: 10pt;">{}</span>'.format(body)
    draft.HTMLBody = font_style
    draft.To = ";".join(recipients)
    draft.CC = ";".join(cc_recipients)
    with open(signature_path, "r", encoding="utf-16") as file:
            html_content = file.read()
    draft.HTMLBody += "\n\n" + html_content
    draft.Save()

    return

def find_machine_tools(outlook_app, emails):
    # The inbox can be full of shipments pertaining to machine tools.
    # This is not our responsibility. We move those to 
    # the machine tools folder
    namespace = outlook_app.GetNamespace("MAPI")
    email_account = "PO_submittal@us.TRUMPF.com"
    inbox = namespace.GetSharedDefaultFolder(namespace.CreateRecipient(email_account), 6)

    # Find machine tools
    machine_tools = np.array([])
    not_machine_tools = np.array([])
    for email in emails:
        subject = email.subject
        serial_numbers = re.findall(r"[ABCN]\d{4}[A-Z]\d{4}", subject)
        if serial_numbers:
            machine_tools = np.append(machine_tools, email)
        else:
            not_machine_tools = np.append(not_machine_tools, email)
    
    subfolder = find_folders(inbox, "Freight Tracking", "DB Schenker - Sea", "Machine Tools")
    if subfolder:
        for email in machine_tools:
            email.Move(subfolder)
    
    return
