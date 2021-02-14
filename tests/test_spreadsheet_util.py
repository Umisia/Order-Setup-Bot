from spreadsheet_util import Workbook
import pytest
from xlwings.constants import DeleteShiftDirection

@pytest.fixture(scope='module')
def wb():
    workbook = Workbook('./testing.xlsm')
    yield workbook
    workbook.save()
    workbook.close()

ws_data = [('open', 'X', 6),
           ('closed', 'X', 4),
           ('incorrect', None, None)
           ]
ws_data_ids = [f'{d[0]}' for d in ws_data]

@pytest.mark.parametrize('ws, expected_col, expected_row', ws_data, ids=ws_data_ids)
def test_get_last_column(ws, expected_col, expected_row, wb):
    """given worksheet name, returns letter of the last filled in column"""
    assert wb.get_last_column(ws) == expected_col

@pytest.mark.parametrize('ws, expected_col, expected_row', ws_data, ids=ws_data_ids)
def test_get_last_row(ws, expected_col, expected_row, wb):
    """given worksheet name, returns number of the last filled in row"""
    assert wb.get_last_row(ws) == expected_row

row_col_data = [(5, 1, 'SO-10001'), (5, 2, 'Confirmed'), (5, 3, ' 17-12-2020'), (5, 10, None), (5, 12, '-')]

@pytest.mark.parametrize('row, col, expected', row_col_data)
def test_get_cell_value(row, col, expected, wb):
    """given row and column gets value of the cell """
    assert wb.get_cell_value(row, col) == expected

def test_get_column_headers(wb):
    """returns a list of column names"""
    headers = ['Sales Order', 'Status', 'Order Date', 'Reference#', 'Billing Address', 'Shipping Address', 'Company Name', 'Contact Name', 'Setup Contact Email', 'Important Information', 'Items and Description', 'Confirmed on', 'Dispatch deadline', 'Type', 'CPD', 'Opening email', 'Ticket#', 'Last follow up', 'Details received', 'Prepared', 'Notes', 'TicketID', 'SOID', 'Order link']
    assert wb.get_column_headers() == headers

labels_data = [
    ('Sales Order', 'A'),
    ('Confirmed on', 'L'),
    ('TicketID', 'V'),
    ('Order link', 'X')
]
@pytest.mark.parametrize('label, expected', labels_data)
def test_get_column_letter(label, expected, wb):
    """given column name returns column letter"""
    assert wb.get_column_letter(label) == expected

def test_write_and_read_order(wb):
    """writed dict to spreadsheet (keys = column names), reads it back to compare data and then deletes from the spreadsheet"""
    spreadsheet_order = {
        'Sales Order': 'SO-11111', 'Status': 'Pending_approval', 'Order Date': ' 15-12-2020', 'Reference#': '1 x CVR4A',
        'Billing Address': 'address: some address', 'Shipping Address': 'some address as well', 'Company Name': 'company name',
        'Contact Name': 'Contact Name', 'Setup Contact Email': 'contact@setup.uk',  'Important Information': 'Repeat order - To be added to current system',
        'Items and Description': 'items', 'Confirmed on': '-', 'Dispatch deadline': '-', 'Type': 'Repeat Order', 'CPD': '-',
        'Opening email': '-', 'Ticket#': '-', 'Last follow up': '-', 'Details received': '-', 'Prepared': '-', 'Notes': None,
        'TicketID': '-', 'SOID': '111111111111111111111111','Order link': 'link'}
    wb.write_order(spreadsheet_order)
    assert wb.get_order_data(wb.get_last_row()) == spreadsheet_order
    range_copy = wb.ws_open.range(wb.ws_open.cells(wb.get_last_row(), wb.get_column_letter("Sales Order")), wb.ws_open.cells(wb.get_last_row(), wb.get_last_column()))
    range_copy.api.Delete(DeleteShiftDirection.xlShiftUp)

get_value_data = [
    (4, 'Order Date', ' 16-12-2020'),
    (5, 'Sales Order', 'SO-10001'),
    (6, 'Status', 'Pending_approval')
]
@pytest.mark.parametrize('row, col_name, expected', get_value_data)
def test_get_value(row, col_name, expected, wb):
    """given row and column name returns cells value"""
    assert wb.get_value(row, col_name) == expected

def test_move_to_closed(wb):
    """given row number copies the row to second worksheet and deletes it from the first one"""
    row = wb.get_last_row()
    #make copy of the row
    range_copy = wb.ws_open.range(wb.ws_open.cells(row, wb.get_column_letter("Sales Order")), wb.ws_open.cells(row, wb.get_last_column()))
    range_paste = wb.ws_open.range(wb.ws_open.cells(row+1, 1), wb.ws_open.cells(row+1, wb.get_last_column()))
    range_paste.value = range_copy.value
    wb.move_to_closed(row+1)
    assert wb.get_value(row+1, 'Sales Order') is None
    range_del = wb.ws_closed.range(wb.ws_closed.cells(wb.get_last_row(), wb.get_column_letter("Sales Order")),
                                  wb.ws_closed.cells(wb.get_last_row(ws='closed'), wb.get_last_column(ws='closed')))
    range_del.api.Delete(DeleteShiftDirection.xlShiftUp)

related_orders_data = [
    (5, ''),
    (6, 'SO-10000, Repeat Order, SO-11222, Repeat Order'),
    (3, None)
]
@pytest.mark.parametrize('row, expected', related_orders_data)
def test_find_related_orders(row, expected, wb):
    """checks both worksheets for other orders with the same customer name or postcode. """
    assert wb.find_related_orders(row) == expected

postcode_rows_data = [
    (4, 'KA22 8EE'),
    (2, '10485'),
    (3, None)
]
@pytest.mark.parametrize('row, expected', postcode_rows_data)
def test_get_postcode(row, expected, wb):
    """extracts postcode from address cell, if not found raises expception"""
    if expected is None:
        with pytest.raises(ValueError):
            wb.get_postcode(row)
    else:
        assert wb.get_postcode(row) == expected

def test_get_all_spreadsheet_orders(wb):
    assert wb.get_all_spreadsheet_orders() == ['SO-10027', 'SO-10042', 'SO-11233', 'SO-10001', 'SO-10000']
    assert wb.get_all_spreadsheet_orders(ws='closed') == ['SO-10027', 'SO-10000', 'SO-11222']

def test_find_cell_address(wb):
    assert wb.find_cell_address('SO-10042') == [[3, 1]]
