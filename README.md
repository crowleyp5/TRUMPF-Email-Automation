# TRUMPF-Email-Automation
As a Business Management Intern at TRUMPF North America, I implemented this automated solution for sending emails to customers so they could track shipments. This process involved the following:
1. Sorting through emails in the inbox to find ones pertaining to laser machine orders.
2. Of those emails, finding those that contain tracking information.
3. Determining the carrier and acquiring tracking links.
4. Updating a SharePoint list with the tracking information.
5. Pulling information from SharePoint relevant to the purchase order.
6. Checking if the incoterms and project status indicate that an email is ready to be sent.
7. Drafting an email with order and tracking information.
8. Adding the appropriate recipients, which varies by laser type and whether we are selling directly to an end user or to a distributor.
9. Organizing the inbox.

The code is split up into five modules. The outlook module is for connecting to outlook, getting email objects, moving emails around folders, and drafting the email. The sharepoint module is for connecting to sharepoint, updating the sharepoint item with tracking information, and reading data from that item. The tracking module is for getting serial numbers, tracking numbers, and tracking links. The utils module is for saving temporary files and reading pdfs. The process flow is in the main module.
