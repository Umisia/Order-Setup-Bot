import zoho_oauth as zoho_oauth
import config
from spreadsheet_util import Workbook
from logger import Logger
import orders
import send_emails
import time
import datetime
from db_util import DB


def send_logs():
    log.critical("Unexpected exception", exc_info=True)
    send_emails.send_internal_email("Exception occured", config.send_logs_email, "send logs")

def is_work_time(dt_now):
    # dt_now = datetime.datetime.now()
    start_time = dt_now.replace(hour=8, minute=0, second=0, microsecond=0)
    end_time = dt_now.replace(hour=17, minute=0, second=0, microsecond=0)
    return start_time <= dt_now <= end_time

def get_nos_statuses_dicts(orders_dict):
    data = {}
    for one_order in orders_dict:
        data[one_order["salesorder_number"]] = one_order['order_status'].capitalize()
    return data

log = Logger(__name__).logger

if __name__ == "__main__":
    while True:  # 1 hour loop        
        try:
            datetime_now = datetime.datetime.now()
            today_date = datetime_now.strftime(" %d-%m-%Y")
            two_weeks_date = (datetime_now + datetime.timedelta(days=14)).strftime(" %d-%m-%Y")
            today_day = datetime.date.today().weekday()

            # working day and working time
            while today_day < 5 and is_work_time(datetime_now):  # 0.5 hour loop
                token = zoho_oauth.refresh_token(config.inv_token)

                db = DB()
                wb = Workbook(config.spreadhsheet_path)
                open_orders = wb.get_all_spreadsheet_orders()  # list of so numbers
                log.debug(open_orders)
                closed_orders = wb.get_all_spreadsheet_orders("closed")  # list of so numbers
                log.debug(closed_orders)
                website_orders = orders.get_inventory_orders(token)
                log.debug(website_orders)
                web_orders = get_nos_statuses_dicts(website_orders)

                for order in website_orders:  #search for new orders
                    log.info(order["salesorder_number"])
                    if order["order_status"] == "draft" or order["order_status"] == "closed":
                        pass
                    elif order['salesorder_number'] in closed_orders:
                        pass
                    elif order['salesorder_number'] not in open_orders:
                        log.info(f"{order['salesorder_number']}: new order")
                        new_order = orders.NewOrder(order["salesorder_id"], token)
                        log.debug(new_order.order_details)
                        if new_order.order_details["Type"] == "Setup":
                            new_order.create_setup_ticket()
                        wb.write_order(new_order.order_details)

                        open_orders = wb.get_all_spreadsheet_orders()  # update variable
              
                print("-----------------------Spreadhseet------------------------")
                # loop from the bottom to move closed orders to 2nd wb tab
                for row in reversed(range(2, wb.get_last_row() + 1)):
                    wb_status = wb.get_value(row, 'Status')
                    wb_so_no = wb.get_value(row, 'Sales Order')
                    log.info(f"row: {row}, so: {wb_so_no}, status: {wb_status}")

                    if wb_status == "Ignore":
                        log.debug(f"{wb_so_no}: Ignored")
                        continue
                    if wb_so_no not in web_orders.keys():
                        wb.write_data(row, "Status", "DELETED")
                        wb.move_to_closed(row)
                    else:
                        #update order status
                        w_stat = web_orders[wb_so_no]
                        if wb_status != w_stat:
                            wb.write_data(row, "Status", w_stat)
                            log.info(f"{wb_so_no}: updated row:{row} with status: {w_stat}")
                        if w_stat == "Confirmed" and wb.get_value(row, "Confirmed on") == "-":  # catch confirmed date
                            wb.write_data(row, "Confirmed on", today_date)
                            wb.write_data(row, "Dispatch deadline", two_weeks_date)
                    #move closed and completed to 2nd tab
                    if wb_status == "Closed":
                        log.debug(f"{wb_so_no}: Closed")
                        if wb.get_value(row, "Prepared") == "done":
                            if wb.get_value(row, "CPD") == "todo":
                                log.info(f"{wb_so_no}: CPD order")
                                spr_order = orders.SpreadsheetOrder(wb.get_order_data(row), token)  # dict
                                spr_order.do_cpd()
                            else:
                                wb.move_to_closed(row)
                        #mark done if ticket is closed
                        elif wb.get_value(row, "Ticket#") != "-":
                            spr_order = orders.SpreadsheetOrder(wb.get_order_data(row), token)  # dict
                            _, ticket_status = spr_order.get_ticket_data()
                            if ticket_status == "Closed":
                                wb.write_data(row, "Prepared", "done")
                                if spr_order.check_res_for_details():  # True when details received
                                    wb.write_data(row, "Details received", "Yes")
                                else:
                                    wb.write_data(row, "Details received", "No")

                    #not closed and not completed orders
                    elif wb.get_value(row, "Prepared") != "done":
                        spr_order = orders.SpreadsheetOrder(wb.get_order_data(row), token)  # dict

                        if spr_order.data["Type"] == "Setup":
                            log.info(f"{spr_order.data['Sales Order']}: Setup order")
                            step = spr_order.do_setup_order(db)
                            log.info(f"step: {step}")

                            if step == "opening":
                                wb.write_data(row, "Opening email", today_date)
                                related = wb.find_related_orders(row)
                                if related:
                                    orders.add_comment(token, spr_order.data['TicketID'], f"possible related orders: {related}")

                            elif step == "approved" and spr_order.data["Opening email"] != today_date: #not sent today (no spamming)
                                wb.write_data(row, "Last follow up", today_date)

                            elif spr_order.data["Last follow up"] != today_date: #no spamming for further steps
                                if step == "confirmed":
                                    wb.write_data(row, "Last follow up", today_date)

                        elif spr_order.is_international():
                            log.info(f"{spr_order.data['Sales Order']}: International order")
                            wb.write_data(row, "Prepared", "done")

                        elif spr_order.data["Type"] == "Repeat Order":
                            log.info(f"{spr_order.data['Sales Order']}: Repeat Order")
                            wb.write_data(row, "Notes", wb.find_related_orders(row))

                        elif spr_order.data["Type"] == "Repair" and spr_order.data["Status"] == "Confirmed":
                            log.info(f"{spr_order.data['Sales Order']}: Repair Confirmed")                            
                            spr_order.do_repair_order()
                            wb.write_data(row, "Prepared", "done")

                        elif spr_order.data["Type"] == "Org":
                            log.info(f"{spr_order.data['Sales Order']}: Org order")
                            if spr_order.do_org_order(db):
                                wb.write_data(row, "Prepared", "done")

                        elif spr_order.data["Type"] == "License":
                            log.info(f"{spr_order.data['Sales Order']}: License order")
                            if spr_order.do_license_order():
                                wb.write_data(row, "Prepared", "done")
              
                db.close_db()

                log.info("=======Loop completed=======")
                log.info("30 minutes wait\n")
                time.sleep(1800)  # wait 30minutes
                
            if today_day != datetime.date.today().weekday():  # date change
                break

            date = today_date
            time.sleep(3600)  # wait 60minutes
            log.info("60 minutes wait")
           
        except:
            send_logs()
            db.close_db()
            raise

