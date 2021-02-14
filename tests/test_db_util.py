import db_util
import pytest

@pytest.fixture(scope='module')
def mydb():
    """Connect to db before testing, disconnect after"""
    myDB = db_util.DB()
    yield myDB
    myDB.close_db()


test_data = [('testpostcode', 'UlaTest'),  # exact match
             ('test postcode', 'UlaTest'),  # space in postcode
             ('', 'UlaTest'),  # no postcode
             ('testpostcode', ''),  # no name
             ('test postcode', ''),  # space in postcode and no name
             ('', '')  # no values
             ]
data_id = [f'postcode: {data[0]}, name: {data[1]}' for data in test_data]

@pytest.fixture(params=test_data, ids=data_id)
def a_test_data(request):
    return request.param

def test_find_org(mydb, a_test_data):
    """given postcode and/or name it returns org name and org id from database. If there is a space in the postcode it checks both space and spaceless version"""
    postcode = a_test_data[0]
    name = a_test_data[1]

    if postcode == '' and name == '':
        assert mydb.find_org(postcode, name) is False
    else:
        assert mydb.find_org(postcode, name) == [('UlaTest', 22061)]
