import main
import datetime
import pytest

now = datetime.datetime.now()
date_data = [(now.replace(hour=8, minute=0, second=0, microsecond=0), True),
             (now.replace(hour=15, minute=30, second=30, microsecond=30), True),
             (now.replace(hour=17, minute=0, second=0, microsecond=0), True),
             (now.replace(hour=17, minute=0, second=0, microsecond=1), False),
             (now.replace(hour=0, minute=0, second=0, microsecond=0), False)
             ]
date_data_ids = [f'{d[0]}' for d in date_data]

@pytest.mark.parametrize('time, expected', date_data, ids=date_data_ids)
def test_is_work_time(time, expected):
    """work time is between 8am and 5pm"""
    assert main.is_work_time(time) == expected

list_of_orders_dicts_3 = [
    {'salesorder_id': '507906000032673081', 'customer_name': 'test customer 1', 'customer_id': '507906000032265009', 'email': 'test_order1@test.com',  'company_name': 'test company 1', 'salesorder_number': 'SO-10042', 'reference_number': '3 x CVR8A + 1 x CVR2-8 + 1 x CVR4A + 1 x CVR2-4', 'date': '2020-12-17',  'shipment_date': '2021-01-15', 'created_time': '2020-12-17T11:50:29+0000', 'last_modified_time': '2020-12-17T11:50:34+0000', 'order_status': 'closed',   'cf_end_user_email': 'enduser1@test.com'},
    {'salesorder_id': '507906000032608009', 'customer_name': 'test customer 2', 'customer_id': '507906000032265009', 'email': 'test_order2@test.com', 'company_name': 'test company 2', 'salesorder_number': 'SO-10041', 'reference_number': '3 x CVR8A + 1 x CVR2-8 + 1 x CVR4A + 1 x CVR2-4', 'date': '2020-12-17', 'shipment_date': '2021-01-15', 'created_time': '2020-12-17T11:50:29+0000', 'last_modified_time': '2020-12-17T11:50:34+0000', 'order_status': 'closed', 'cf_end_user_email': 'enduser2@test.com'},
    {'salesorder_id': '507906000032586047', 'customer_name': 'test customer 3', 'customer_id': '507906000032265009', 'email': 'test_order3@test.com', 'company_name': 'test company 3', 'salesorder_number': 'SO-10040', 'reference_number': '3 x CVR8A + 1 x CVR2-8 + 1 x CVR4A + 1 x CVR2-4', 'date': '2020-12-17', 'shipment_date': '2021-01-15', 'created_time': '2020-12-17T11:50:29+0000','last_modified_time': '2020-12-17T11:50:34+0000', 'order_status': 'closed', 'cf_end_user_email': 'enduser3@test.com'}
]

def test_get_nos_list():
    """gets list of so_numbers from list of dicts"""
    expected = {'SO-10042': 'Closed',
                'SO-10041': 'Closed',
                'SO-10040': 'Closed'}
    assert main.get_nos_statuses_dicts(list_of_orders_dicts_3) == expected
