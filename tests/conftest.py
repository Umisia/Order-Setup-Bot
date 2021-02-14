import pytest
import json
import zoho_oauth
import config

@pytest.fixture(scope="module")
def inv_token():
    token = config.inv_token
    return zoho_oauth.refresh_token(token)

order_details_data_path = r"./data/order_details.json"
spreadsheet_order_data_path = r"./data/spreadsheet_order_data.json"
new_order_data_path = r"./data/new_order_data.json"

def load_json_data(path):
    with open(path, encoding='UTF-8') as my_data:
        data = json.load(my_data)
        return data

@pytest.fixture(params=load_json_data(order_details_data_path))
def order_details_data(request):
    return request.param

@pytest.fixture(params=load_json_data(spreadsheet_order_data_path))
def spr_order_data(request):
    return request.param

@pytest.fixture(params=load_json_data(new_order_data_path))
def new_order_data(request):
    return request.param



