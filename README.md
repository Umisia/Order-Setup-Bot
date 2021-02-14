#Order-Setup-Bot
Bot written to automate tedious work.
Every half an hour it checks for new orders and orders' updates on Zoho Inventory.
It arranges orders in categories and takes action when needed.
* Zoho Inventory
  * get list of orders
    
* Zoho Desk tickets
    * create a ticket
    * assign to a person
    * leave comments
    * prepare draft emails
    * check ticket for status and resolution
    
* Zoho Cliq
    * send Cliq messages to notify of actions taken/needed
    
* spreadsheet
    * scrap orders data to a spreadsheet
    * update orders every half an hour
    * move completed orders to second tab
    
* logs
    * gather logs in a file
    * in case of crash send logs via email
    
* pytest coverage<br>
------Name---------------Stmts---Miss---Cover <br>
src\config.py---------------29------0------100%<br>
src\db_util.py--------------27------0------100%<br>
src\logger.py---------------13------0------100%<br>
src\main.py----------------134----113----16%<br>
src\orders.py--------------335-----1------99%<br>
src\send_emails.py-------26-----9------27%<br>
src\spreadsheet_util.py--94-----1------99%<br>
src\zoho_oauth.py--------13-----0------100%<br>
conftest.py-------------------24-----0------100%<br>
test_db_util.py--------------18-----0------100%<br>
test_main.py-----------------13-----0------100%<br>
test_orders.py--------------130----0------100%<br>
test_spreadsheet_util.py-63-----0------100%<br>


