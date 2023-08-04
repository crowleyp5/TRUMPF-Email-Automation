from shareplum import Site, Office365

def connect_sharepoint():
    sharepoint_url = "*******************************"
    username = "**************"
    password = "*****"
    authcookie = Office365(sharepoint_url, username=username, password=password).GetCookies()
    site = Site(sharepoint_url, authcookie=authcookie)
    return site


def update_and_get_sp_item(site, target_equipment_num, carrier, tracking_num, link):
    # Retrieve the SharePoint list and all list items
    list_title = "Project Synchro"
    sp_list = site.List(list_title)
    items = sp_list.GetListItems()

    # Find line item corresponding to the serial number
    found_item = None
    for item in items:
        equipment_number = item["Equipment #"]
        if equipment_number == target_equipment_num:
            found_item = item
            break
    
    description = f"{carrier}_{tracking_num}"
    hyperlink = f"{link}, {description}"
    data = {"Housebill Number": hyperlink}

    # Update housebill and ship date fields
    sp_list.UpdateListItems(data, ID=found_item["ID"])

    return item

def read_sharepoint(item):
    # Store data
    po = item["Customer PO # (CA)"]
    so = item["SAP Order # (CA)"]
    customer_name = item["Customer Name"]
    incoterms = item["Incoterms (CA)"]
    project_status = item["Project Status PPM and CA"]
    project_manager = item["Project/Product Manager"]
    rsm = item["RSM"]
    commercial_contact = item["E-mail Commercial Contact"]
    project_contact = item["E-Mail Project Contact"]
    distributor = item["Distributor"]

    # Create list for recipients
    if distributor == "No":
        recipients = [f"{commercial_contact}", f"{project_contact}"]
    else:
        recipients = [f"{commercial_contact}"]
    cc_recipients = [f"{project_manager}", f"{rsm}"]

    return(po, so, customer_name, incoterms, project_status, recipients, cc_recipients)
