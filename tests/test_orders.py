import orders
import pytest
import config

date_data = [(' 30-12-2020', ' 05-01-2021', 4), (' 15-02-2020', ' 20-03-2020', 25), (' 04-05-2020', ' 02-05-2020', 0)]

@pytest.mark.parametrize('data', date_data)
def test_time_delta(data):
    """counts working days between two dates: from, to"""
    assert orders.time_delta(data[0], data[1]) == data[2]

def test_get_custom_fields_vals(order_details_data):
    """returns contact_name, setup_email and important_info fields values"""
    assert orders.get_custom_fields_vals(order_details_data['data']) == tuple(order_details_data['expected']['custom-fields-vals'])

def test_get_line_items(order_details_data):
    """given sequence of dicts extracts items, formats them and returns string"""
    assert orders.get_line_items(order_details_data['data']['line_items']) == order_details_data['expected']['line-items']

def test_get_items_for_subject(order_details_data):
    """given string of ordered items, returns items count for the ticket subject"""
    if order_details_data['expected']['subject'] == "exception":
        with pytest.raises(SyntaxError):
            orders.get_items_for_subject(order_details_data['expected']['line-items'])
    else:
        assert orders.get_items_for_subject(order_details_data['expected']['line-items']) == tuple(order_details_data['expected']['subject'])

def test_format_address(order_details_data):
    """format address received from requests response"""
    assert orders.format_address(order_details_data['data']['billing_address']) == order_details_data['expected']['billing-address']
    assert orders.format_address(order_details_data['data']['shipping_address']) == order_details_data['expected']['shipping-address']

def test_get_link_to_ticket():
    """given ticket id generates correct link to zoho desk ticket"""
    assert orders.get_link_to_ticket("218739000047103001") == config.ticket_link_part + "218739000047103001"

def test_count_duedate():
    """adds 14 days to a date. formats it to zoho desk ticket format"""
    import datetime
    date = datetime.datetime.strptime("2021-02-06", '%Y-%m-%d')
    assert orders.count_duedate(date) == "2021-02-20T09:00:00.000Z"

def test_get_inventory_orders(inv_token):
    assert orders.get_inventory_orders(inv_token)

def test_send_click_message(inv_token):
    """using API request, send a Cliq message"""
    assert orders.send_click_message(inv_token, "message test") == 204

def test_create_ticket_add_comment(inv_token):
    """creates ticket with given subject, device quantity and model.
    adds comment to the ticket.
    deletes the ticket"""
    import config
    ti_no, ti_id = orders.create_ticket(inv_token, 'This is a test ticket', 1, config.model_155)
    assert ti_id, ti_no
    assert orders.add_comment(inv_token, ti_id, "test comment") == 200
    #clean up
    assert orders.delete_ticket(inv_token, ti_id) == 204

class TestNewOrder:
    def test_oder_details(self, inv_token, new_order_data):
        """gathers all instance properties into dic"""
        new_order = orders.NewOrder(new_order_data['so-id'], inv_token)
        assert new_order.order_details == new_order_data['order-data']

    def test_get_contact_id(self, inv_token, new_order_data):
        """returns contact person id for new zoho desk ticket assigment"""
        new_order = orders.NewOrder(new_order_data['so-id'], inv_token)
        assert new_order.find_contact_id() == new_order_data['contact-id']

    def test_create_setup_ticket(self, inv_token):
        """create zoho desk ticket for setup type orders only. assigns it to correct person"""
        new_order = orders.NewOrder(507906000032257844, inv_token)
        assert new_order.create_setup_ticket() == 200
        #clean up after
        assert orders.delete_ticket(inv_token, new_order.order_details['TicketID']) == 204

class TestSpreadsheetOrder:
    @pytest.mark.tegoteraz
    def test_get_postcode(self, inv_token, spr_order_data):
        """extracts postcode"""
        spr_order = orders.SpreadsheetOrder(spr_order_data['data'], inv_token)
        assert spr_order.postcode == spr_order_data['expected']['postcode']

    def test_get_ticket_data(self, inv_token, spr_order_data):
        """gets ticket id and its status"""
        spr_order = orders.SpreadsheetOrder(spr_order_data['data'], inv_token)
        assert spr_order.get_ticket_data() == tuple(spr_order_data['expected']['ticket-data'])

    def test_get_ticket_resolution(self, inv_token, spr_order_data):
        """gets ticket resolution string"""
        spr_order = orders.SpreadsheetOrder(spr_order_data['data'], inv_token)
        assert spr_order.get_ticket_resolution() == spr_order_data['expected']['resolution']

    def test_check_res_for_resolution(self, inv_token, spr_order_data):
        """check if details has been received - looks for keywords in resolution string"""
        spr_order = orders.SpreadsheetOrder(spr_order_data['data'], inv_token)
        assert spr_order.check_res_for_details() == spr_order_data['expected']['details-received']

    def test_get_repairs_ticket_id(self, inv_token, spr_order_data):
        """checks Important Info column for orders of type Repair and returns tickets id"""     
        spr_order = orders.SpreadsheetOrder(spr_order_data['data'], inv_token)
        if spr_order_data['expected']['repairs-ticket-id'] == 'exception':
            with pytest.raises(AttributeError):
                spr_order.get_repairs_ticket_id()
        else:
            assert spr_order.get_repairs_ticket_id() == spr_order_data['expected']['repairs-ticket-id']

    def test_is_international(self, inv_token, spr_order_data):
        """checks if there are keywords in the addresses to indicate international/uk order"""
        spr_order = orders.SpreadsheetOrder(spr_order_data['data'], inv_token)
        assert spr_order.is_international() == spr_order_data['expected']['is-international']

@pytest.fixture(scope='module')
def wb():
    from spreadsheet_util import Workbook
    workbook = Workbook('./test.xlsm')
    yield workbook
    workbook.save()
    workbook.close()

class TestOrdersTypes:
    # test_ticket_no = '#22711'
    # test_ticket_id = 218739000050192407
    def test_do_repair_order(self, inv_token, wb):
        spr_order_fail = orders.SpreadsheetOrder(wb.get_order_data(2), inv_token)
        assert spr_order_fail.do_repair_order() is False
        spr_order_good = orders.SpreadsheetOrder(wb.get_order_data(3), inv_token)
        assert spr_order_good.do_repair_order() is True

    def test_do_org_order(self, inv_token, wb):
        from db_util import DB
        db = DB()
        spr_order_fail = orders.SpreadsheetOrder(wb.get_order_data(4), inv_token)
        assert spr_order_fail.do_org_order(db) is False
        spr_order_good = orders.SpreadsheetOrder(wb.get_order_data(5), inv_token)
        assert spr_order_good.do_org_order(db) is True

    def test_license_order(self, inv_token, wb):
        spr_order_fail = orders.SpreadsheetOrder(wb.get_order_data(6), inv_token)
        assert spr_order_fail.do_license_order() is False
        spr_order_good = orders.SpreadsheetOrder(wb.get_order_data(7), inv_token)
        assert spr_order_good.do_license_order() is True

    def test_do_cpd(self, inv_token, wb):
        spr_order = orders.SpreadsheetOrder(wb.get_order_data(7), inv_token)
        assert spr_order.do_cpd() == 'sent'

    def test_leave_draft_email(self, inv_token, wb):
        spr_order = orders.SpreadsheetOrder(wb.get_order_data(8), inv_token)
        assert spr_order.leave_draft_email("This is a test email content") == 200

    def test_do_setup_order(self, inv_token, wb):
        """
        if there is no opening date in the spreadsheet then send opening email,
        if status is approved and was not follow up, or confirmed and not follow up yet, send approved email,
        if status is confirmed
        if dispatch deadline has passed todays date, send dispatch email,
        if confirmed day was 3 or more days ago, send confirmed email
        if confirmed day was 6 or more days ago, send confirmed email
        """
        from db_util import DB
        from datetime import datetime, timedelta

        db = DB()
        opening_order = orders.SpreadsheetOrder(wb.get_order_data(9), inv_token)
        assert opening_order.do_setup_order(db) == 'opening'
        approved_order = orders.SpreadsheetOrder(wb.get_order_data(10), inv_token)
        confirmed_skipped_order = orders.SpreadsheetOrder(wb.get_order_data(11), inv_token)
        assert approved_order.do_setup_order(db) == 'approved'
        assert confirmed_skipped_order.do_setup_order(db) == 'approved'
        dispatching_order = orders.SpreadsheetOrder(wb.get_order_data(12), inv_token)
        assert dispatching_order.do_setup_order(db) == 'dispatch'

        confirmed3_date = datetime.today() - timedelta(days=5)
        dispatch_date = datetime.today() + timedelta(days=14)
        wb.write_data(13, 'Dispatch deadline', dispatch_date.strftime(" %d-%m-%Y"))
        wb.write_data(13, 'Confirmed on', confirmed3_date.strftime(" %d-%m-%Y"))
        confirmed3_order = orders.SpreadsheetOrder(wb.get_order_data(13), inv_token)
        assert confirmed3_order.do_setup_order(db) == 'confirmed'

        confirmed6_date = datetime.today() - timedelta(days=8)
        wb.write_data(14, 'Dispatch deadline', dispatch_date.strftime(" %d-%m-%Y"))
        wb.write_data(13, 'Confirmed on', confirmed6_date.strftime(" %d-%m-%Y"))
        confirmed6_order = orders.SpreadsheetOrder(wb.get_order_data(14), inv_token)
        assert confirmed6_order.do_setup_order(db) == 'confirmed'

        ntd_date = datetime.today() - timedelta(days=1)
        wb.write_data(15, 'Confirmed on', ntd_date.strftime(" %d-%m-%Y"))
        wb.write_data(15, 'Last follow up', ntd_date.strftime(" %d-%m-%Y"))
        wb.write_data(15, 'Opening email', ntd_date.strftime(" %d-%m-%Y"))
        wb.write_data(14, 'Dispatch deadline', dispatch_date.strftime(" %d-%m-%Y"))
        nothing_to_do_order = orders.SpreadsheetOrder(wb.get_order_data(15), inv_token)
        assert nothing_to_do_order.do_setup_order(db) == None
