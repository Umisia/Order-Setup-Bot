import datetime
import jinja2
import config
import requests, json
from logger import Logger
import re

log = Logger(__name__).logger

def time_delta(date_from, date_to):
    date_from = datetime.datetime.strptime(date_from, " %d-%m-%Y")
    date_to = datetime.datetime.strptime(date_to, " %d-%m-%Y")
    counter = 0
    for _ in range((date_to - date_from).days):
        weekday = date_to.weekday()
        if weekday < 5:
            counter += 1
        date_to -= datetime.timedelta(days=1)
    return counter

def get_custom_fields_vals(data):
    contact_name = setup_email = impo_info = None
    for cf in data["custom_fields"]:
        if cf['customfield_id'] == '507906000000886011':
            contact_name = cf["value"]
        if cf['customfield_id'] == '507906000016330717':
            setup_email = cf["value"]
        if cf["customfield_id"] == "507906000000886001":
            impo_info = cf["value"]
    return contact_name, setup_email, impo_info

def get_line_items(data):
    all_items = []
    newline = "\n"
    for item in data:
       all_items.append(
            f'{item["name"]} ({item["sku"]}):{item["description"].replace(newline, ",")} ordered:{item["quantity"]}')
    return "|".join(all_items)

def get_items_for_subject(items):
    items_splitted = items.split("|")
    item_count = {8: 0, 4: 0, 30: 0, 1: 0}
    for item in items_splitted:
        ordered = int(item[item.index("ordered") + len("ordered:"):])
        if "set of 8" in item.lower():
            item_count[8] += ordered
        if "set of 4" in item.lower():
            item_count[4] += ordered
        if "set of 30" in item.lower():
            item_count[30] += ordered
        if "standalone headset" in item.lower() or config.standalone_headset.lower() in item.lower():
            item_count[1] += ordered
    text = "+".join("{}x{}".format(v, k) for (k, v) in item_count.items() if v != 0)
    number = eval(text.replace("x", "*"))
    model = config.model_255 if "255" in items else config.model_155
    return text, number, model

def format_address(data):
    nl = "\n"
    address = ', '.join(f'{k}: {v.replace(nl, " ")}' for k, v in data.items() if v != "")
    address = address.replace(',,', ',')
    return address

def get_link_to_ticket(ticketid):
    return config.ticket_link_part + str(ticketid)

def count_duedate(today):
    date_due = today + datetime.timedelta(days=14)
    due_date_to_set = date_due.strftime("%Y-%m-%d") + "T09:00:00.000Z"
    return due_date_to_set

def request_head(token):
    request_head_dict = {
        'orgId': config.org_id,
        'Authorization': 'Zoho-oauthtoken ' + token}
    return request_head_dict

def get_inventory_orders(token):
    req_head = request_head(token)
    response = requests.get("https://inventory.zoho.com/api/v1/salesorders?page=1", headers=req_head)
    json_response = json.loads(response.text)
    log.debug(json_response)
    return json_response["salesorders"]

def get_order_details(token, order_id):
    req_head = request_head(token)
    response = requests.get(f"https://inventory.zoho.com/api/v1/salesorders/{order_id}?", headers=req_head)
    json_response = json.loads(response.text)["salesorder"]
    log.debug(json_response)
    return json_response

def send_click_message(token, text):
    req_head = request_head(token)
    req_head['Content-Type'] = 'application/json'
    my_json = json.dumps({'text': text})
    resp = requests.post("https://cliq.zoho.com/api/v2/channelsbyname/ordernotification/message", headers=req_head,
                  data=my_json)
    log.info(f"sent click message, status: {resp.status_code}")
    return resp.status_code

def create_ticket(token, subject, device_quantity, model):
    req_head = request_head(token)
    data = {
        'subCategory': 'Setup',
        'subject': subject,
        'customFields': {
            'Warranty': '1st Year',
            'Device Quantity': device_quantity,
            'Device Type': model },
        'dueDate': count_duedate(datetime.datetime.now()),
        'departmentId': config.department_id,
        'channel': 'Email',
        'priority': 'Low',
        'classification': 'Setup - Installation',
        'category': 'User Issue',
        'contactId': config.default_contact_id
    }
    my_json = json.dumps(data)
    resp = requests.post("https://desk.zoho.com/api/v1/tickets", headers=req_head, data=my_json)
    json_ticket = json.loads(resp.text)
    log.info(f"create_ticket status: {str(resp.status_code)}")

    return json_ticket["ticketNumber"], json_ticket["id"]

def add_comment(token, ticket_id, message):
    req_head = request_head(token)
    data = {
        "isPublic": "false",
        "contentType": "html",
        "content": message}
    my_json = json.dumps(data)
    resp = requests.post("https://desk.zoho.com/api/v1/tickets/" + str(ticket_id) + "/comments", headers=req_head, data=my_json)
    log.info(f"added comment to {ticket_id}, status: {resp.status_code}")
    return resp.status_code

def delete_ticket(token, ticket_id):
    req_head = request_head(token)
    data = { 'ticketIds': [ticket_id]}
    my_json = json.dumps(data)
    resp = requests.post(f"https://desk.zoho.com/api/v1/tickets/moveToTrash", headers=req_head, data=my_json)
    log.info(f"delete ticket: {ticket_id}, status: {str(resp.status_code)}")
    return resp.status_code

class NewOrder:
    def __init__(self, so_id, token):
        self.so_id = so_id
        self.token = token
        self.request_head = request_head(self.token)
        self.details = get_order_details(self.token, self.so_id)
        self.order_details = self.get_details_dict()
        self.update_order_details()

    @property
    def so_no(self):
        return self.details["salesorder_number"]

    @property
    def status(self):
        return self.details["order_status"].capitalize()

    @property
    def order_data(self):
        return datetime.datetime.strptime(self.details["created_date"], "%Y-%m-%d").strftime(" %d-%m-%Y")

    @property
    def reference(self):
        return self.details["reference_number"]

    @property
    def bill_address(self):
        return format_address(self.details["billing_address"])

    @property
    def ship_address(self):
        return format_address(self.details["shipping_address"])

    @property
    def items_and_desc(self):
        return get_line_items(self.details["line_items"])

    @property
    def company_name(self):
        return self.details["customer_name"]

    @staticmethod
    def search(matchers, values):
        for matcher in matchers:
            if matcher in values:
                return True
        return False

    def get_type(self):
        values = ';'.join(str(x).lower() for x in self.order_details.values())
        matchers = {
            "Setup": ["setup", "set up"],
            "nosetup": ["no setup", "no set up"],
            "Repeat": ["repeat"],
            "License": ["license", "subscription", "portal"],
            "Enroll": ["enroll", "enrol", "enrolled"],
            "Lgfl": ["lgfl"],
            "Inclusive": ["inclusive"],
            "Org": ["workbook", "wb", "vr", "headset"],
            "Repair": ["repair", "ticket"]
        }
        # priority of matching is important!
        if self.search(matchers["Setup"], values) and not self.search(matchers['nosetup'], values):
            return "Setup"
        elif self.search(matchers["Repeat"], values):
            return "Repeat Order"
        elif self.search(matchers["Repair"], values):
            return "Repair"
        elif self.search(matchers["Lgfl"], values):
            return "Setup"
        elif self.search(matchers["Inclusive"], values):
            return "Inclusive"
        elif self.search(matchers["Enroll"], values):
            return "Enroll"
        elif self.search(matchers["License"], values):
            return "License"
        elif self.search(matchers["Org"], values):
            return "Org"
        else:
            self.order_details['Prepared'] = "done"
            return "Other"

    def get_details_dict(self):
        custom_fields = get_custom_fields_vals(self.details)
        details_dict = {
            "Sales Order": self.so_no,
            "Status": self.status,
            "Order Date": self.order_data,
            "Reference#": self.reference,
            "Billing Address": self.bill_address,
            "Shipping Address": self.ship_address,
            "Company Name": self.company_name,
            "Contact Name": custom_fields[0],
            "Setup Contact Email": custom_fields[1],
            "Important Information": custom_fields[2],
            "Items and Description": self.items_and_desc,
            "Confirmed on": "-",
            "Dispatch deadline": "-",
            "Type": "-",
            "CPD": "-",
            "Opening email": "-",
            "Ticket#": "-",
            "Last follow up": "-",
            "Details received": "-",
            "Prepared": "-",
            "Notes": "-",
            "TicketID": "-",
            "SOID": self.so_id,
            "Order link": f"https://inventory.zoho.com/app#/salesorders/{self.so_id}?filter_by=Status.All&per_page=200&sort_column=last_modified_time&sort_order=D"
        }
        return details_dict

    def update_order_details(self):
        self.order_details["Type"] = self.get_type()
        # catch training with 'setup and training' item but no devices on order
        if self.order_details["Type"] == "Setup":
            try:
                get_items_for_subject(self.order_details["Items and Description"])
            except SyntaxError:  # no items on the order
                log.info("Order type error. Changed to Org.")
                self.order_details["Type"] = "Org"
        if "CVR-CPD-1" in self.order_details["Items and Description"] or self.order_details["Type"] == "Setup":
            self.order_details["CPD"] = "todo"
        else:
            self.order_details["CPD"] = "-"

    def find_contact_id(self):
        response_email = requests.get(
            "https://desk.zoho.com/api/v1/contacts/search?limit=1&email=" + self.order_details["Setup Contact Email"],
            headers=self.request_head).text
        name = self.order_details["Contact Name"].split()
        response_name = requests.get(
            f"https://desk.zoho.com/api/v1/contacts/search?limit=1&firstName={name[0]}&lastName={name[1]}",
            headers=self.request_head)
        if response_name.status_code == 422:
            response_name = False
        else:
            response_name = response_name.text
        response = response_email or response_name
        if response:
            json_ticket = json.loads(response)
            contact_id = json_ticket['data'][0]['id']
        else:
            contact_id = config.default_contact_id
        return contact_id

    def assign_ticket_add_contact(self, ticket_id):
        contact_id = self.find_contact_id()
        #setting correct contact_id while creating tickets generates email notification, updating it here to avid that
        data = {
            'assigneeId': config.assignee_id,
            'contactId': contact_id}
        my_json = json.dumps(data)
        req = requests.patch("https://desk.zoho.com/api/v1/tickets/" + ticket_id, headers=self.request_head, data=my_json)
        log.info(f"ticket assigned status: {str(req.status_code)}")
        return req.status_code

    def create_setup_ticket(self):
        items, device_q, model = get_items_for_subject(self.order_details["Items and Description"])
        postcode = self.details["shipping_address"]["zip"] or self.details["billing_address"]["zip"] or None
        subject = f'{config.subj}({items}): {self.company_name}({postcode}), {self.order_details["Sales Order"]}' or None
        log.info(f"subject: {subject}")

        ticket_no, ticket_id = create_ticket(self.token, subject, device_q, model)

        self.order_details["Ticket#"] = ticket_no
        self.order_details["TicketID"] = ticket_id

        assignment_status = self.assign_ticket_add_contact(self.order_details["TicketID"])
        return assignment_status

class SpreadsheetOrder:
    def __init__(self, data: dict, token):
        self.data = data
        self.token = token
        self.request_head = request_head(self.token)

    @property
    def postcode(self):
        try:
            return self.data["Shipping Address"][self.data["Shipping Address"].index("zip:") + 5:self.data["Shipping Address"].index("country") - 2]
        except ValueError:
            return None

    def get_ticket_data(self):
        response = requests.get(
            f"https://desk.zoho.com/api/v1/tickets/search?limit=1&ticketNumber={self.data['Ticket#']}",
            headers=self.request_head)
        json_ticket = json.loads(response.text)
        ticket_status = json_ticket["data"][0]["status"]
        ticket_id = json_ticket["data"][0]["id"]
        log.info(f'get_ticket_data ({response.status_code}): {ticket_status}, {ticket_id}')
        log.info(f'get_ticket_data ({response.status_code}): {ticket_id}')

        return ticket_id, ticket_status

    def get_ticket_resolution(self):
        ticket_id, _ = self.get_ticket_data()
        response = requests.get("https://desk.zoho.com/api/v1/tickets/" + ticket_id + "/resolution",
                                headers=self.request_head)
        json_ticket = json.loads(response.text)
        resolution = json_ticket["content"]
        log.debug(f'getTicketData resolution ({response.status_code}): {resolution}')
        return resolution

    def check_res_for_details(self):
        resolution = self.get_ticket_resolution()
        keywords = ["details", "received", "printed"]
        if resolution:
            for key in keywords:
                if key in resolution and "no" not in resolution:
                    return True
        return False

    def get_repairs_ticket_id(self):
        ticket_no = None
        info = self.data["Important Information"]
        if "#" in info:
            ticket_no = info[info.index("#") + 1:]
        else:
            ticket_no = re.search(r"[0-9]{4,7}", info).group()
        log.info(ticket_no)
        ticket_response = requests.get(
            "https://desk.zoho.com/api/v1/tickets/search?limit=1&ticketNumber=$" + ticket_no, headers=self.request_head)
        json_ticket = json.loads(ticket_response.text)
        log.info(f"get ticket id status: {ticket_response.status_code}")
        ticket_id = json_ticket["data"][0]["id"]
        return ticket_id

    def is_international(self):
        uk_match = ["united kingdom", "wales", "uk"]
        for match in uk_match:
            if match in self.data["Billing Address"].lower() or match in self.data["Shipping Address"].lower():
                return False
        return True

    def do_repair_order(self):
        try:
            ticket_id = self.get_repairs_ticket_id()            
            ticket_link = get_link_to_ticket(ticket_id)
            send_click_message(self.token, f"{self.data['Sales Order']} related to ticket is now Confirmed. \n {ticket_link}")
            # add_comment(f"zsu[@user:{config.repair_confirmed_comment_id}]zsu Order Confirmed.")
            add_comment(self.token, ticket_id, "Order Confirmed.")
            return True
        except AttributeError:
            return False
        
    def do_org_order(self, db):
        if self.postcode:
            if db.find_org(self.postcode, self.data["Company Name"]): #org exists
                return True
            else:
                send_click_message(self.token, f"{self.data['Sales Order']}: Make sure org was created. Update spreadsheet when done.\n {self.data['Order link']}")
                return False

    def do_license_order(self):
        #check if license data added to sales order (items description)
        details = get_order_details(self.token, self.data["SOID"]) #to refresh orders info
        items = get_line_items(details["line_items"])
        if "Customer Organisation Name:" in items:
            return True
        else:
            send_click_message(self.token, f"{self.data['Sales Order']}: Make sure org details were added to the order. Update spreadsheet when done.\n {self.data['Order link']}")
            return False

    def do_cpd(self):
        send_click_message(self.token, f"CPD invite for {self.data['Sales Order']} {self.data['Order link']}\n update spreadsheet {config.cpd_spreadsheet}")
        return 'sent'

    def leave_draft_email(self, email_content):
        data = {
            "channel": "EMAIL",
            "to": self.data["Setup Contact Email"],
            "fromEmailAddress": config.support_email,
            "contentType": "html",
            "content": email_content,
            "isForward": "true"}
        my_json = json.dumps(data)
        req = requests.post(f"https://desk.zoho.com/api/v1/tickets/{self.data['TicketID']}/draftReply", headers=self.request_head, data=my_json)
        log.info(f"{self.data['Sales Order']} draft email left")
        return req.status_code

    def choose_email_template(self, template_type):
        template_dir = "../email_templates"
        jinja_env = jinja2.Environment(loader=jinja2.FileSystemLoader(template_dir))

        template = jinja_env.get_template(template_type + ".html")
        if template_type == "opening" or template_type == "dispatch":
            return template.render()
        elif template_type == "approved":
            return template.render(sales_order=self.data["Sales Order"])
        elif template_type == "confirmed3" or template_type == "confirmed6":
            return template.render(sales_order=self.data["Sales Order"], dispatch_date=self.data["Dispatch deadline"])

    def do_setup_order(self, db):
        # step 1. no opening email sent (order just added to wb)
        if self.data["Opening email"] == "-":
            msg = f"<a href='{self.data['Order link']}'>Sales Order</a><br>Contact email: {self.data['Setup Contact Email']}"
            add_comment(self.token, self.data['TicketID'], msg)
            org_from_db = db.find_org(self.postcode, self.data["Company Name"])
            if org_from_db: #org exists
                add_comment(self.token, self.data['TicketID'], f"Org found: {org_from_db[0]}")
            log.info(f"OPENING EMAIL: {self.data['Sales Order']}")
            self.leave_draft_email(self.choose_email_template("opening"))
            send_click_message(self.token, f"{self.data['Sales Order']}: #{int(self.data['Ticket#'])}: Opening draft left.\n{get_link_to_ticket(self.data['TicketID'])}")
            return "opening"

        # step 2. approved, no follow up or skipped approved and went straight to confirm
        elif (self.data["Status"] == "Approved" and self.data["Last follow up"] == "-") or (self.data["Status"] == "Confirmed" and self.data["Last follow up"] == "-"):
            log.info(f'FOLLOW UP: APPROVED {self.data["Sales Order"]}')
            self.leave_draft_email(self.choose_email_template("approved"))
            send_click_message(self.token, f"{self.data['Sales Order']}: #{int(self.data['Ticket#'])}: Approved draft left.\n{get_link_to_ticket(self.data['TicketID'])}")
            return "approved"

        # step 3. confirmed
        elif self.data["Status"] == "Confirmed":
            today_date = datetime.datetime.now().strftime(" %d-%m-%Y")
            #dispatch deadline passed
            if self.data["Dispatch deadline"] <= today_date:
                log.info(f"FOLLOW UP: DISPATCHING {self.data['Sales Order']}")
                self.leave_draft_email(self.choose_email_template("dispatch"))
                send_click_message(self.token, f"{self.data['Sales Order']}: #{int(self.data['Ticket#'])}: Dispatching draft left.\n{get_link_to_ticket(self.data['TicketID'])}")
                return "dispatch"
            # confirmed+3days or confirmed+6 days
            elif time_delta(self.data["Confirmed on"], today_date) >= 3:
                #confirmed+3
                if time_delta(self.data["Confirmed on"], today_date) < 6 and time_delta(self.data["Last follow up"], today_date) > 2:
                    log.info(f"FOLLOW UP: CONFIRMED +3 DAYS {self.data['Sales Order']}")
                    self.leave_draft_email(self.choose_email_template("confirmed3"))
                    send_click_message(self.token, f"{self.data['Sales Order']}: #{int(self.data['Ticket#'])}: Confirmed+3 draft left.\n{get_link_to_ticket(self.data['TicketID'])}")
                # confirmed+6
                elif time_delta(self.data["Confirmed on"], today_date) >= 6 and time_delta(self.data["Last follow up"], today_date) > 2:
                    log.info(f"FOLLOW UP: CONFIRMED +6 DAYS {self.data['Sales Order']}")
                    self.leave_draft_email(self.choose_email_template("confirmed6"))
                    send_click_message(self.token, f"{self.data['Sales Order']}: #{int(self.data['Ticket#'])}: Confirmed+6 draft left.\n{get_link_to_ticket(self.data['TicketID'])}")
                return "confirmed"

