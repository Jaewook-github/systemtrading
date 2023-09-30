from api.Kiwoom import *
import sys

app = QApplication(sys.argv)
kiwoom = Kiwoom()
#kiwoom.get_account_number()

"""kospi_code_list = kiwoom.get_code_list_by_market('0')
print(kospi_code_list)
for code in kospi_code_list:
    code_name = kiwoom.get_master_code_name(code)
    print(code, code_name)

kosdaq_code_list = kiwoom.get_code_list_by_market('10')
print(kosdaq_code_list)
for code in kosdaq_code_list:
    code_name = kiwoom.get_master_code_name(code)
    print(code, code_name)"""

"""df = kiwoom.get_price_data('005930')
print(df)"""
deposit = kiwoom.get_deposit()

position = kiwoom.get_balance()
print(position)


app.exec_()
