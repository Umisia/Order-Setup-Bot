import xlwings as xw
from logger import Logger
import string

log = Logger(__name__).logger

#copy and delete ws
# sheet = wb.sheets['Open']
# sheet.api.Copy(Before=sheet.api)
# wb.sheets[0].name = 'testing'
# test_sheet = wb.sheets['testing']
# test_sheet.delete()

class Workbook:
    def __init__(self, wb_path):
        self.wb = xw.Book(wb_path)
        self.ws_open = self.wb.sheets["Open"]
        self.ws_closed = self.wb.sheets["Closed"]
        log.info(f"connected to {wb_path}")

    # last column filled in
    def get_last_column(self, ws='open'):
        if ws == "open":
            return self.ws_open.range(1, 1).end('right').get_address(0, 0)[0]
        elif ws == "closed":
            return self.ws_closed.range(1, 1).end('right').get_address(0, 0)[0]

    # last row filled in
    def get_last_row(self, ws= 'open'):
        if ws == 'open':
            return self.ws_open.range('A' + str(self.ws_open.cells.last_cell.row)).end('up').row
        elif ws == 'closed':
            return self.ws_closed.range('A' + str(self.ws_closed.cells.last_cell.row)).end('up').row

    def get_cell_value(self, row, col):
        return self.ws_open.cells(row, col).value

    def get_column_headers(self):
        return self.ws_open.range(self.ws_open.cells(1, 1), self.ws_open.cells(1, self.get_last_column())).value

    # find column letter by column header
    def get_column_letter(self, label):
        col_index = self.get_column_headers().index(label)
        column_letter = string.ascii_uppercase[col_index]
        return column_letter

    def write_data(self, row, column_name, value):
        self.ws_open.cells(row, self.get_column_letter(column_name)).value = value

    def get_all_spreadsheet_orders(self, ws="open"):
        if ws == "open":
            return self.ws_open.range(self.ws_open.cells(2, 1), self.ws_open.cells(self.get_last_row(), 1)).value
        elif ws == "closed":
            return self.ws_closed.range(self.ws_closed.cells(2, 1),
                                        self.ws_closed.cells(self.get_last_row("closed"), 1)).value

    def find_cell_address(self, values):
        columns = self.get_column_headers()
        finds = []
        for row in range(2, self.get_last_row() + 1):
            for col in range(1, (len(columns) + 1)):
                if self.ws_open.range((row, col)).value == values:
                    finds.append([row, col])
        return finds

    def write_order(self, order_details: dict):
        column_headers = self.get_column_headers()
        last_row_open = self.get_last_row() + 1
        # writing this way allows to add/delete columns from spreadsheet to get more/less data, it is also resilient
        # to order structure changes on Inventory
        for column in column_headers:
            if column not in order_details.keys():
                self.write_data(last_row_open, column, "-")
            for key in order_details.keys():
                if key == column:
                    self.write_data(last_row_open, key, order_details[key])
        log.info(f"written {order_details['Sales Order']}")

    def get_order_data(self, row):      
        columns = self.get_column_headers()
        order = {}

        for col in range(0, (len(columns))):
            order[columns[col]] = self.get_cell_value(row, col + 1)
        log.debug(order)
        return order

    def get_value(self, row, column_name):
        return self.ws_open.cells(row, self.get_column_letter(column_name)).value

    def move_to_closed(self, row):
        log.info(f'moved row {row} ({self.get_value(row, "Sales Order")}: {self.get_value(row, "Status")}) to row: {self.get_last_row("closed")+1}')
        range_copy = self.ws_open.range(self.ws_open.cells(row, self.get_column_letter("Sales Order")),
                                        self.ws_open.cells(row, self.get_last_column()))
        range_paste = self.ws_closed.range(
            self.ws_closed.cells(self.get_last_row('closed') + 1, 1),
            self.ws_closed.cells(self.get_last_row('closed') + 1, self.get_last_column('closed')))
        range_paste.value = range_copy.value
        range_copy.api.Delete()
        range_paste.api.WrapText = False

    def find_related_orders(self, row_passed):
        try:
            postcode = self.get_postcode(row_passed)
        except ValueError:
            return
        name = self.get_value(row_passed, "Company Name")
        this_order = self.get_value(row_passed, "Sales Order")
        related = []
        both_ws = [self.ws_open, self.ws_closed]
        ranges = [self.get_last_row(), self.get_last_row(ws='closed')]
        for ws, rang in zip(both_ws, ranges):
            for row in range(2, rang + 1):
                if postcode in ws.cells(row, self.get_column_letter("Shipping Address")).value or \
                        postcode in ws.cells(row, self.get_column_letter("Billing Address")).value or \
                        name in ws.cells(row, self.get_column_letter("Company Name")).value:

                    if this_order != self.get_value(row, "Sales Order"): #dont want to append the order being checked
                        related.append(ws.cells(row, self.get_column_letter("Sales Order")).value)
                        related.append(ws.cells(row, self.get_column_letter("Type")).value)

        filtered = [i or 'None' for i in related]
        return ", ".join(filtered)

    def get_postcode(self, row):
        shipp_address = self.get_value(row, "Shipping Address")
        return shipp_address[shipp_address.index("zip:")+5:shipp_address.index("country")-2]

    def save(self):
        self.wb.save()

    def close(self):
        self.wb.close()
