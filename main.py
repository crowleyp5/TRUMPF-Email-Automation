import re
import outlook
import utils
import tracking
import sharepoint

def main():
    outlook = outlook.connect_outlook()
    emails = outlook.get_emails(outlook)
    lasers = tracking.find_lasers(emails)
    for email in lasers:
        equipment_num = tracking.get_sn(email)
        carrier = tracking.find_carrier(email)
        if carrier == "UPS":
            tracking_num = tracking.find_ups_tracking(email)
            link = tracking.get_ups_link(tracking_num)
        elif carrier == "FedEx":
            tracking_num = tracking.find_fedex_tracking(email)
            link = tracking.get_fedex_link(tracking_num)
        else:
            tracking_num = []
        if tracking_num and link:
            site = sharepoint.connect_sharepoint()
            item = sharepoint.update_and_get_sp_item(site, equipment_num, carrier, tracking_num, link)
            data = sharepoint.read_sharepoint(item)
            # If the conditions are met, we move all emails pertaining to this laser to
            # the in sharepoint folder.
            if re.match(r"^[DC]", data[3]) and data[4] == "5. Approved":
                outlook.draft_email(outlook, data[0], data[1], equipment_num, data[5], data[6], data[2], link, carrier)
                for email2 in lasers:
                    if equipment_num in email2.subject:
                        outlook.to_sharepoint(outlook, email2, carrier)
            # We do not draft an email for F incoterms, and if the project status is still at 4,
            # that means we are waiting on documents from the customer. We move the email to the
            # holding box for now.
            else:
                outlook.to_holding_box(outlook, email, carrier)
        else:
            outlook.to_holding_box(email, carrier)
    outlook.find_machine_tools()

    return

if __name__ == "__main__":
    main()
